Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Module modProject
    ''' <summary>
    ''' Gets the value of the 'Echelle' entry from the custom properties held on the drawing
    ''' </summary>
    ''' <returns>The value in the 'Echelle' entry of the custom properties</returns>
    ''' <remarks></remarks>
    Public Function GetCurrentScale() As String

        Dim EchStr As String = "1000"
        Dim EchPSStr As String = "1"

        Try
            'Get the database for the current drawing
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'Get the custom properties collection
            Dim dictRecs As System.Collections.IDictionaryEnumerator = db.SummaryInfo.CustomProperties
            'Loop through the custom properties until we find the key we're looking for
            While dictRecs.MoveNext
                'If thie key is the 'Echelle' property
                If dictRecs.Key.ToString = "Echelle" Then
                    'Return its value
                    EchStr = dictRecs.Value.ToString
                End If

                If dictRecs.Key.ToString = "PS-Scale" Then
                    'Return its value
                    EchPSStr = dictRecs.Value.ToString
                End If

            End While
        Catch
        End Try

        Dim SumEch As Double = 1000
        SumEch = CDbl(EchStr) * CDbl(EchPSStr)
        'Return an empty string as the 'Echelle' key was not found
        Return SumEch.ToString 'Default value = 1:1000 ' include SQLite in Revo THA


    End Function

    Function GetScalePS() As String

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database

        'Chargement du facteur d'Annotation
        Dim EchPS As String = ""
        Try
            EchPS = GetFileProperty("PS-Scale")
        Catch
        End Try

        ' Recherche du facteur  "1:1000" (1/1000 ou  1/1)
        If EchPS = "" And IsNumeric(EchPS) = False Then 'Si pas de facteur
            EchPS = 1  'Par défaut
            Try
                Dim cm As ObjectContextManager = db.ObjectContextManager
                If cm IsNot Nothing Then  ' Now get the Annotation Scaling context collection
                    Dim occ As ObjectContextCollection = cm.GetContextCollection("ACDB_ANNOTATIONSCALES")  ' (named ACDB_ANNOTATIONSCALES_COLLECTION)
                    If occ IsNot Nothing Then
                        For Each occX As AnnotationScale In occ
                            If occX.Name = "1:1000" Then EchPS = occX.DrawingUnits
                        Next
                    End If
                End If
            Catch ex As System.Exception
            End Try
        End If

        Return EchPS
    End Function

    Function CheckExistQuadri() As Collection


        Dim entCol As New Collection

        Try

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor

            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim filterValues As TypedValue() = New TypedValue() {New TypedValue(CInt(DxfCode.Start), "INSERT"), New TypedValue(CInt(67), 1)}
                Dim filter As New SelectionFilter(filterValues)
                Dim foundBlocks As PromptSelectionResult = ed.SelectAll(filter)
                If foundBlocks.Status <> PromptStatus.OK Then
                    'Erreur
                Else

                    For Each id As ObjectId In foundBlocks.Value.GetObjectIds()
                        Dim br As BlockReference = DirectCast(tr.GetObject(id, OpenMode.ForRead), BlockReference)
                        Dim NameBL As String = ""
                        If br.IsDynamicBlock Then
                            Dim dbtrID As ObjectId = br.DynamicBlockTableRecord
                            Dim btr As BlockTableRecord = DirectCast(tr.GetObject(dbtrID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockTableRecord)
                            NameBL = btr.Name
                        Else
                            NameBL = br.Name
                        End If

                        If NameBL.ToUpper = "QUADRI_DYN" Then
                            'ed.WriteMessage(vbTab, br.Name, NameBL)
                            Dim newlm As LayoutManager = LayoutManager.Current
                            If newlm.CurrentLayout.ToUpper = LayoutName(br, tr).ToUpper Then
                                entCol.Add(br)
                            End If
                        End If
                    Next
                End If

            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

        If entCol.Count = 0 Then
            Return entCol
        Else
            Return entCol
        End If


    End Function


    ''' <summary>
    ''' Restores the layer states
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub RestoreDefaultLayerState()
        If RestoreLayerState("Revo_Defaut") = False Then 'JourCAD_Defaut 'Revo_Defaut include SQLite in Revo THA
            SwitchToLayersDefaultSettings()
        End If
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        ed.Regen()
    End Sub

    ''' <summary>
    ''' Set layer on/frozen/locked states based on the values in the layer_management table in the system.mdb database
    ''' </summary>
    ''' <remarks></remarks>
    Private Sub SwitchToLayersDefaultSettings()
        'Open the system.db3 database
        'OpenDatabase("system\system.mdb")
        'Open the Layer_Management table
        Dim Rvinfo As New Revo.RevoInfo

        Dim dt As System.Data.DataTable = GetTable("Layer_Management", dbName:=Rvinfo.DB3system) ' include SQLite in Revo THA
        If Not dt Is Nothing Then
            Dim strLayerName As String = ""
            'Thaw and activate layer 0
            FreezeLayer("0", False)
            ActivateLayer("0")
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Get the drawing layer table
                Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
                'Loop through the records
                For i As Integer = 0 To dt.Rows.Count - 1
                    strLayerName = dt.Rows(i)("JCLayer").ToString
                    'Check if the layer is in the layertable
                    If lt.Has(strLayerName) Then
                        'Get the layer table record
                        Dim ltr As LayerTableRecord = DirectCast(trans.GetObject(lt.Item(strLayerName), OpenMode.ForWrite), LayerTableRecord)
                        'Set the Off, Frozen and Locked states to the values from the database
                        ltr.IsOff = Not DirectCast(dt.Rows(i)("DefaultLayerOn"), Boolean)
                        ltr.IsFrozen = DirectCast(dt.Rows(i)("DefaultFreeze"), Boolean)
                        ltr.IsLocked = DirectCast(dt.Rows(i)("DefaultLock"), Boolean)
                    End If
                Next
                'Save the changes
                trans.Commit()
            End Using
        End If
    End Sub

    Public Function GetCurrentRotation() As Double

        Dim strRot As String = ""
        strRot = GetFileProperty("Rotation")

        If strRot = "" Then
            Return 100.0
        Else
            Return Val(strRot)
        End If

    End Function

    Public Function GetFileProperty(ByRef strProperty As String) As String

        'Get the database for the current drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Get the custom properties collection
        Dim dictRecs As System.Collections.IDictionaryEnumerator = db.SummaryInfo.CustomProperties
        'Loop through the custom properties until we find the key we're looking for
        While dictRecs.MoveNext
            'If thie key is the 'Echelle' property
            If dictRecs.Key.ToString = strProperty Then
                'Return its value
                Return dictRecs.Value.ToString
            End If
        End While

        'Return an empty string as the 'Echelle' key was not found
        Return ""

    End Function

    Public Function SetFileProperty(ByRef strProperty As String, strValue As String)

        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database
        Dim infoBuilder As New DatabaseSummaryInfoBuilder(db.SummaryInfo)
        Try
            infoBuilder.CustomPropertyTable.Add(strProperty, strValue)
        Catch ex As Exception
            infoBuilder.CustomPropertyTable.Item(strProperty) = strValue
        End Try
        db.SummaryInfo = infoBuilder.ToDatabaseSummaryInfo()

        Return True

        '#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        '        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
        '#Else 'Versio AutoCAD 2013 et +
        '            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
        '#End If

        '        Try 'Si l'info existe
        '            acDoc.SummaryInfo.AddCustomInfo(strProperty, strValue)
        '        Catch
        '            acDoc.SummaryInfo.SetCustomByKey(strProperty, strValue)
        '            'MsgBox("erreur d'écriture de propriété")
        '        End Try

        '        Return True

    End Function


End Module
