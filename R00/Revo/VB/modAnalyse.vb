Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Public Class RevoAnalyse

    Public Sub CreerSurface()

        CreateRegionsForAll() 'CreateRegionsForAll2(GetCurrentLayer())
        ConvertRegionsToPolylines()

    End Sub


    Public Sub AnylseMaxMinZ()

        Dim Ass As New Revo.RevoInfo
        Dim Connect As New Revo.connect
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument




        'Mise à jour des Xdata des bâtiments 
        Dim CollPol As New Collection
        CollPol = SelectPolyline(False, False, "Sélectionner les poylignes des bâtiments")
       
        Dim acCurDb As Database = acDoc.Database
        Dim CollBlock As New Collection

        Try
            Dim NbreOK As Double = 0
            For Each objPolyline As Polyline In CollPol

                Dim strNo As String = 0
                Dim strDes As String = ""
                Dim Verts As Point3dCollection = GetPolyVertices(objPolyline.ObjectId)
                Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "POINT")} 'New TypedValue(DxfCode.LayerName, "MO_CS_BAT_NUM")
                Dim IDs() As ObjectId = SelectWindowPoly(Values, Verts)

                If IDs IsNot Nothing Then

                    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                    Dim db As Database = HostApplicationServices.WorkingDatabase

                    Using tr As Transaction = db.TransactionManager.StartTransaction()


                        'Altitude Max & Min
                        Dim AltMax As Double = -1000000000
                        Dim AltMin As Double = 1000000000


                        ' Traitement des points

                        Dim ci As Double = 0
                        For Each ID As ObjectId In IDs


                            Dim PTent As DBPoint = tr.GetObject(ID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                            ci += 1
                            Connect.Message(Ass.xTitle, "Analyse en cours ...", False, ci, CollBlock.Count, "info")

                            If PTent.Position.Z > AltMax Then AltMax = PTent.Position.Z
                            If PTent.Position.Z < AltMin Then AltMin = PTent.Position.Z

                            NbreOK += 1
                        Next

                        'Effacer les Xdata
                        'ClearAllXData(objPolyline.ObjectId)

                        'Ecriture des Xdata
                        Dim obj As DBObject = tr.GetObject(objPolyline.ObjectId, OpenMode.ForWrite)
                        AddRegAppTableRecord("REVO")
                        Dim rb As New ResultBuffer(New TypedValue(1001, "REVO"), New TypedValue(1000, "ID" & strNo), New TypedValue(1000, AltMin), New TypedValue(1000, AltMax))
                        obj.XData = rb
                        rb.Dispose()
                        tr.Commit()


                        '  MsgBox("Min : " & AltMin & " / Max :" & AltMax)

                    End Using


                End If

            Next

           
        Catch ex As System.Exception
            MsgBox("Erreur nuage de points ou Xdata update : " & ex.Message)
        End Try


        Connect.Message(Ass.xTitle, "Analyse teminée avec succès!", False, 100, 100, "info")
        Connect.Message(Ass.xTitle, "", True, 100, 100)



    End Sub

    Public Sub SuppPointDoubleZ()

        Dim Ass As New Revo.RevoInfo
        Dim Connect As New Revo.connect
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

        'Sélection des objets

        '' Get the current document and database
        Dim acCurDb As Database = acDoc.Database
        Dim CollBlock As New Collection
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Request for objects to be selected in the drawing area
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
            '' If the prompt status is OK, objects were selected
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                '' Step through the objects in the selection set
                For Each acSSObj As SelectedObject In acSSet
                    '' Check to make sure a valid SelectedObject object was returned
                    If Not IsDBNull(acSSObj) Then
                        '' Open the selected object for write
                        Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            If TypeName(acEnt) = "DBPoint" Then
                                CollBlock.Add(acEnt)
                            End If
                        End If
                    End If
                Next
            End If


            ' Traitement des points
            Dim Ptid As New List(Of ObjectId)
            Dim ci As Double = 0
            For Each PTent As DBPoint In CollBlock 'Traitement des objets
                ci += 1
                Connect.Message(Ass.xTitle, "Analyse en cours ...", False, ci, CollBlock.Count, "info")
                For Each PTrech As DBPoint In CollBlock 'Traitement des objets: BlockReference
                    If PTrech.ObjectId <> PTent.ObjectId And PTrech.Position.X = PTent.Position.X And PTrech.Position.Y = PTent.Position.Y Then
                        If PTrech.Position.Z <= PTent.Position.Z Then 'Si recherché est plus bas -> delete
                            Ptid.Add(PTrech.ObjectId)
                        Else ' Si origine est plus bas -> delete
                            Ptid.Add(PTent.ObjectId)
                        End If
                    End If
                Next
            Next

            For i = 0 To Ptid.Count - 1
                Try
                    Dim Ptsupp As DBPoint = DirectCast(acTrans.GetObject(Ptid(i), OpenMode.ForWrite), DBPoint)
                    Ptsupp.Erase()
                Catch
                End Try
            Next

            '' Save the new object to the database
            acTrans.Commit()
            '' Dispose of the transaction
        End Using
        ' End selection of objects

        Connect.Message(Ass.xTitle, "Analyse teminée avec succès!", False, 100, 100, "info")
        Connect.Message(Ass.xTitle, "", True, 100, 100)


    End Sub
    Public Sub IntersPoly()

        Dim Ass As New Revo.RevoInfo
        Dim Connect As New Revo.connect

        Dim PolysAxe As New Collection
        PolysAxe = (SelectPolyline(False, True, "Sélectionner la polyligne définisant le trajectoire"))

        If PolysAxe.Count <> 0 Then
            Dim PolyAxe As Polyline = PolysAxe.Item(1)
            Dim DistPoly As Double = 0 'Longeur de la polyligne
            Dim Polys As New Collection
            Polys = SelectPolyline(False, False, "Sélectionner les polyligne définisant les surfaces")

            Dim CountInters As Double = 0 'Conteur d'intersection à la poyligne
            'Dim AllIntersection As New Collection
            Dim AllIntersection As New SortedList

            Dim TestIntersExiste As Point3dCollection = New Point3dCollection

            'Boucle dans les segments de la polyligne d'axe --------------------------------
            For u As Integer = 0 To PolyAxe.NumberOfVertices - 2 'Fin au dernier segments
                Connect.Message(Ass.xTitle, "Analyse des segments en cours ...", False, u, PolyAxe.NumberOfVertices, "info")

                Dim AxeStart As Double = u
                Dim AxeEnd As Double = u + 1
                Dim LineA As Line = New Line(New Point3d(PolyAxe.GetPoint2dAt(AxeStart).X, PolyAxe.GetPoint2dAt(AxeStart).Y, 0), _
                                             New Point3d(PolyAxe.GetPoint2dAt(AxeEnd).X, PolyAxe.GetPoint2dAt(AxeEnd).Y, 0))


                For Each Poly As Polyline In Polys 'Boucle dans les polylignes ------------------------------
                    For i As Integer = 0 To Poly.NumberOfVertices - 1 'Boucle dans les sommes de la poyligne -----------------------
                        Dim lStart As Double = i
                        Dim lEnd As Double = i + 1
                        If lStart = Poly.NumberOfVertices - 1 Then lEnd = 0
                        Dim LineB As Line = New Line(New Point3d(Poly.GetPoint2dAt(lStart).X, Poly.GetPoint2dAt(lStart).Y, 0), _
                                                     New Point3d(Poly.GetPoint2dAt(lEnd).X, Poly.GetPoint2dAt(lEnd).Y, 0))
                        Dim IntersectPts As Point3dCollection = New Point3dCollection
                        LineA.IntersectWith(LineB, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)
                        CountInters += IntersectPts.Count

                        For Each Pt In IntersectPts
                            If TestIntersExiste.Contains(Pt) = False Then
                                TestIntersExiste.Add(Pt)
                                Dim DistInters As Double = Dist(New Point2d(LineA.StartPoint.X, LineA.StartPoint.Y), New Point2d(Pt.X, Pt.Y))
                                'AllIntersection.Add(New PtsIntersection(Pt, DistPoly + DistInters))
                                AllIntersection.Add(DistPoly + DistInters, Pt)
                            End If
                        Next
                    Next
                Next

                'Addition de la longeur de la ligne en cours
                DistPoly += LineA.Length
            Next


            ''Chargement des dist dans une liste
            'Dim DistList As New List(Of Double)
            'For Each pt As PtsIntersection In AllIntersection
            '    DistList.Add(pt.Dist)
            'Next

            ''Tri de la liste
            'DistList.Sort()

            'Dim PtList As New List(Of Point3d)
            'For ri = 0 To DistList.Count - 1
            '    For Each pt As PtsIntersection In AllIntersection
            '        If DistList(ri) = pt.Dist Then
            '            PtList.Add(pt.Pts)
            '            Exit For
            '        End If
            '    Next
            'Next

            'Dim PolyID As ObjectId = PolyAxe.ObjectId
            'For Nb = 0 To PtList.Count - 1
            '    PolyID = BreakPolyline(PolyID, PtList(Nb))
            'Next

            Dim PolyID As ObjectId = PolyAxe.ObjectId
            For Each Nb In AllIntersection
                PolyID = BreakPolyline(PolyID, Nb.Value)
            Next

            Connect.Message(Ass.xTitle, "Analyse teminée avec succès!" & vbCrLf & " Nombre de segments : " & "PtList.Count" _
                             & "  /  Nombre d'intersection : " & CountInters & " (" & AllIntersection.Count & ")", False, 100, 100, "info")
            Connect.Message(Ass.xTitle, "", True, 100, 100)

        End If

    End Sub

    Public Sub ObjInterPoly() ' Objet à l'intérieur d'une polyligne 

        ' Boucle dans les poylignes sélectionnée
        Dim Objs As New Collection
        Objs = SelectObj(False, "Sélectionner les polyligne et les textes à analyser", True, True)
        Dim Connect As New Revo.connect
        Dim Ass As New Revo.RevoInfo
        Dim AcadDoc As Document = Application.DocumentManager.MdiActiveDocument

        If Objs.Count <> 0 Then

            'Tri des objets
            Dim PolysSurf As New Collection
            Dim LignesDiv As New Collection
            Dim Textes As New Collection
            For Each Obj As Entity In Objs
                If TypeName(Obj) Like "Polyline*" Then
                    Dim PolyTest As Polyline = Obj
                    If PolyTest.Closed = True Then
                        PolysSurf.Add(Item:=Obj, Key:=Obj.ObjectId.ToString)
                    Else
                        LignesDiv.Add(Item:=Obj, Key:=Obj.ObjectId.ToString)
                    End If
                ElseIf TypeName(Obj) Like "DBText" Then
                    Textes.Add(Item:=Obj, Key:=Obj.ObjectId.ToString)
                End If
            Next



            'Recherche des objets compris dans la surface
            Dim Listing As String = "" 'Contenu de l'analyse
            Dim LiveCount As Double = 0
            'Boucle dans les poly de surfaces
            For Each PolySurf As Polyline In PolysSurf
                LiveCount += 1
                Connect.Message(Ass.xTitle, "Analyse des surfaces en cours ... (" & Textes.Count & "/" & LignesDiv.Count & ")", False, LiveCount, PolysSurf.Count, "info")

                Listing += vbCrLf & FileIO.FileSystem.GetName(AcadDoc.Name) & ";" & PolySurf.Layer.ToString & ";" & Math.Round(PolySurf.Area, 3).ToString & ";"

                ' Texte compris à dans la polyligne surfacique
                Dim NumParc As String = ""
                For u = 1 To Textes.Count
                    If u > Textes.Count Then Exit For
                    If PtIsInsideLWP(Textes(u).Position, PolySurf.ObjectId) Then
                        If NumParc <> "" Then NumParc += "/"
                        NumParc += Textes(u).TextString
                        Textes.Remove(u)
                        u = u - 1
                    End If
                Next u
                Listing += NumParc & ";"


                ' Polyligne compris à dans la polyligne surfacique
                Dim LongLigneDiv As Double = 0
                For u = 1 To LignesDiv.Count
                    If u > LignesDiv.Count Then Exit For
                    Dim LigneDiv As Polyline = LignesDiv(u)
                    Dim InStart As Boolean = False
                    Dim InEnd As Boolean = False

                    'Test du point de départ
                    If PtIsInsideLWP(LigneDiv.GetPoint3dAt(0), PolySurf.ObjectId) Then
                        InStart = True
                    Else 'Sinon test si sur une ligne
                        ' Objet en intersection à dans poyligne surfacique
                        Dim LineA As Line = New Line(New Point3d(LigneDiv.GetPoint2dAt(0).X, LigneDiv.GetPoint2dAt(0).Y, 0), _
                                            New Point3d(LigneDiv.GetPoint2dAt(1).X, LigneDiv.GetPoint2dAt(1).Y, 0))
                        For i As Integer = 0 To PolySurf.NumberOfVertices - 1 'Boucle dans les sommes de la poyligne -----------------------
                            Dim lStart As Double = i
                            Dim lEnd As Double = i + 1
                            If lStart = PolySurf.NumberOfVertices - 1 Then lEnd = 0
                            Dim LineB As Line = New Line(New Point3d(PolySurf.GetPoint2dAt(lStart).X, PolySurf.GetPoint2dAt(lStart).Y, 0), _
                                                         New Point3d(PolySurf.GetPoint2dAt(lEnd).X, PolySurf.GetPoint2dAt(lEnd).Y, 0))
                            Dim IntersectPts As Point3dCollection = New Point3dCollection
                            LineA.IntersectWith(LineB, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)
                            If IntersectPts.Count <> 0 Then
                                If Math.Round(IntersectPts(0).X, 6) = Math.Round(LigneDiv.GetPoint3dAt(0).X, 6) And _
                                 Math.Round(IntersectPts(0).Y, 6) = Math.Round(LigneDiv.GetPoint3dAt(0).Y, 6) Then
                                    InStart = True
                                    Exit For
                                End If
                            End If

                        Next
                    End If

                    'Test du point de fin
                    If PtIsInsideLWP(LigneDiv.GetPoint3dAt(LigneDiv.NumberOfVertices - 1), PolySurf.ObjectId) Then
                        InEnd = True
                    Else 'Sinon test si sur une ligne
                        ' Objet en intersection à dans poyligne surfacique
                        Dim LineA As Line = New Line(New Point3d(LigneDiv.GetPoint2dAt(LigneDiv.NumberOfVertices - 2).X, LigneDiv.GetPoint2dAt(LigneDiv.NumberOfVertices - 2).Y, 0), _
                                            New Point3d(LigneDiv.GetPoint2dAt(LigneDiv.NumberOfVertices - 1).X, LigneDiv.GetPoint2dAt(LigneDiv.NumberOfVertices - 1).Y, 0))
                        For i As Integer = 0 To PolySurf.NumberOfVertices - 1 'Boucle dans les sommes de la poyligne -----------------------
                            Dim lStart As Double = i
                            Dim lEnd As Double = i + 1
                            If lStart = PolySurf.NumberOfVertices - 1 Then lEnd = 0
                            Dim LineB As Line = New Line(New Point3d(PolySurf.GetPoint2dAt(lStart).X, PolySurf.GetPoint2dAt(lStart).Y, 0), _
                                                         New Point3d(PolySurf.GetPoint2dAt(lEnd).X, PolySurf.GetPoint2dAt(lEnd).Y, 0))
                            Dim IntersectPts As Point3dCollection = New Point3dCollection
                            LineA.IntersectWith(LineB, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)
                            If IntersectPts.Count <> 0 Then
                                If Math.Round(IntersectPts(0).X, 6) = Math.Round(LigneDiv.GetPoint3dAt(LigneDiv.NumberOfVertices - 1).X, 6) And _
                                   Math.Round(IntersectPts(0).Y, 6) = Math.Round(LigneDiv.GetPoint3dAt(LigneDiv.NumberOfVertices - 1).Y, 6) Then
                                    InEnd = True
                                    Exit For
                                End If
                            End If

                        Next
                    End If


                    If InStart And InEnd Then
                        LongLigneDiv += LigneDiv.Length
                        LignesDiv.Remove(u)
                        u = u - 1
                    End If

                Next u

                Listing += Math.Round(LongLigneDiv, 3) & ";"

                If LignesDiv.Count = 0 Then Exit For

            Next

            ' Log avec :
            '    - Numéro de la parcelle (texte)
            '    - Longeur polyligne 
            '    - *dist min du Sommet le plus pret (+ x/y du sommet)
            '    - *Portée concernée (n° pyl -> n° pyl)
            '    - Area


            Dim SaveFileDialogExport As New System.Windows.Forms.SaveFileDialog
            Dim CheminFichier As String '= "c:\export-calque.txt"
            SaveFileDialogExport.Filter = "Fichier texte (*.csv)|*.csv"
            SaveFileDialogExport.Title = "Exporter des données"
            SaveFileDialogExport.FileName = ""
            CheminFichier = SaveFileDialogExport.ShowDialog()

            Dim Confirm
            Dim ListText() As String = Split(Listing, vbCrLf)
            Confirm = Revo.RevoFiler.EcritureFichier(SaveFileDialogExport.FileName, ListText, True)

            Connect.Message(Ass.xTitle, "Analyse teminée avec succès!" & vbCrLf & " Nombre de segments : " & PolysSurf.Count, False, 100, 100, "info")
            Connect.Message(Ass.xTitle, "", True, 100, 100)

        End If

    End Sub

End Class
