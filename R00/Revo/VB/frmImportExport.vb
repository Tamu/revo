Imports System.Windows.Forms
Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Public Class frmImportExport

    Public ActiveImport As Boolean = False ' Default : Exportation
    Dim SelectedOjs As New Collection
    Dim NumFormat As Double = -1
    Dim AutorizName As String = ""

    Private Sub frmImportExport_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Revo.RevoScript.ParamImportExportFrm.Clear()


        If ActiveImport = False Then 'Exportation
            Me.Text = "Exportation de données"
            LayoutSelect.Visible = True
            LayoutTheme.Visible = True
            LayoutComPlan.Visible = True
            LayoutDoublon.Visible = False


            'Sélection du fichier + Chargement des formats

            'Sélection des objets manuelle (défaut tout le fichiers)

            'Points du thèmes
            With Me.cboTheme
                .Items.Add("Tous") '0
                .Items.Add("MO/MUT/RAD : Tous") '1
                .Items.Add("MO : Tous") '2
                .Items.Add("MO : 01 PF") '3
                .Items.Add("MO : 02 CS") '4
                .Items.Add("MO : 03 OD") '5
                .Items.Add("MO : 06 BF") '6
                .Items.Add("MO : 09 COM") '7
                .Items.Add("MUT : Tous") '8
                .Items.Add("MUT : 01 PF") '9
                .Items.Add("MUT : 02 CS") '10
                .Items.Add("MUT : 03 OD") '11
                .Items.Add("MUT : 06 BF") '12
                .Items.Add("MUT : 09 COM") '13
                .Items.Add("IMPL") '14
                .Items.Add("TOPO") '15
                .Items.Add("RAD") '16
                .Items.Add("SOUT") '17


                '.Items.Add("SOUT : Eau potable") '17
                '.Items.Add("SOUT : E. P. (projet)") '18
                '.Items.Add("SOUT : Eaux claires") '19
                '.Items.Add("SOUT : E. C. (projet)") '20
                '.Items.Add("SOUT : Eaux unitaires") '21
                '.Items.Add("SOUT : Eaux usées") '22
                '.Items.Add("SOUT : E. U. (projet)") '23
                '.Items.Add("SOUT : Eclairage public") '24
                '.Items.Add("SOUT : Electricité") '25
                '.Items.Add("SOUT : Gaz") '26
                '.Items.Add("SOUT : Téléphone") '27
                '.Items.Add("SOUT : TV") '28
                .SelectedIndex = 0
            End With

            'Points des communes

            'Points des plans
            Dim ListPlans As List(Of String) = SelectPlans()
            For Each Plan As String In ListPlans
                CkComPlan.Items.Add(Plan.ToString)
            Next
           
            CkComPlan.CheckOnClick = True
            Me.Height = 308 'Normal : 349

        Else ' Importation
            Me.Text = "Importation de données"
            LayoutFile.Visible = False ' Fichier
            LayoutSelect.Visible = False
            LayoutTheme.Visible = True
            LayoutComPlan.Visible = False
            LayoutDoublon.Visible = False 'DESACTIVATION PROVISOIRE 06.10.2015 THA

            'Points du thèmes
            With Me.cboTheme
                .Items.Add("Automatique (code Nature)") '0
                .Items.Add("MO") '1
                .Items.Add("MUT") '2
                .SelectedIndex = 0
            End With

            Me.Height = 198 'Normal : 349

        End If

        'Détection des doublons


    End Sub

    Private Sub btnFileDialog_Click(sender As System.Object, e As System.EventArgs) Handles btnFileDialog.Click

        Dim FileDialog As New System.Windows.Forms.SaveFileDialog
        Dim RVinfo As New Revo.RevoInfo
        Dim XMLformat As String = RVinfo.XMLformatBase
        Dim NbreFormat As Double = 0
        Dim Filters As String = ""
        Dim ConfirmOK As Double

        'Importation des formats points : format.xml **********
        If File.Exists(XMLformat) Then 'Test de l'existance du fichier XML Revo
            Dim x As New RevoXML(XMLformat) 'lecture des extentions
            NbreFormat = x.countElements("/data-format/format")
            Dim VarExt As String = ""
            For i = 1 To NbreFormat
                VarExt = x.getElementValue("/data-format/format/extension", i)
                Filters += "|" & x.getElementValue("/data-format/format/name", i) & " (*." & VarExt & ")" & "|*." & VarExt
            Next
            If Strings.Left(Filters, 1) = "|" Then Filters = Mid(Filters, 2, Filters.Length - 1)
        End If

        FileDialog.Filter = Filters
        If ActiveImport = True Then FileDialog.Title = "Importation des données"
        If ActiveImport = False Then FileDialog.Title = "Exporter des données"
        FileDialog.FilterIndex = RVinfo.FormatExport
        FileDialog.FileName = ""
        ConfirmOK = FileDialog.ShowDialog()

        If ConfirmOK = 1 Then
            TBNameFile.Text = FileDialog.FileName
            NumFormat = FileDialog.FilterIndex

            'Chargement des block authoris :Load Block Name (LOAD format.xml
            Dim docXML As New System.Xml.XmlDocument
            docXML.Load(XMLformat)
            Dim NodeSymbols As System.Xml.XmlNodeList
            Dim Xroot As String = "/data-format/format[" & NumFormat & "]"
            NodeSymbols = docXML.SelectNodes(Xroot & "/topic/data/type/symbol/block")
            For Each NodeSymbol As System.Xml.XmlNode In NodeSymbols
                AutorizName += "[" & NodeSymbol.InnerText & "]"  '[PTS_BF_BORNE]
            Next

        End If

    End Sub

    Private Sub BtnSelectObj_Click(sender As System.Object, e As System.EventArgs) Handles BtnSelectObj.Click

        If NumFormat <> -1 Then


            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            SelectedOjs.Clear()


            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acOPrompt As PromptSelectionOptions = New PromptSelectionOptions()
                acOPrompt.MessageForAdding = "Sélectionner les points à exporter"
                '' Request for objects to be selected in the drawing area
                Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(acOPrompt)
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
                                If TypeName(acEnt) = "BlockReference" Then

                                    Try
                                        Dim Ident As String = ""
                                        Dim blkRef As BlockReference = DirectCast(acTrans.GetObject(acEnt.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockReference)
                                        Dim attCol As AttributeCollection = blkRef.AttributeCollection 'Get the block attributue collection
                                        Dim attRef As AttributeReference = Nothing
                                        Dim Titre As String = Me.Text
                                        If attCol IsNot Nothing Then
                                            For Each attId As ObjectId In attCol
                                                attRef = DirectCast(acTrans.GetObject(attId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeReference)
                                                Me.Text = "Recherche des attributs : " & attRef.TextString
                                                System.Windows.Forms.Application.DoEvents()
                                                If attRef.Tag.ToUpper = "IDENTDN" Then
                                                    Ident = Mid(attRef.TextString, 3, 10)
                                                    Exit For
                                                End If
                                            Next
                                        End If
                                        Me.Text = Titre

                                        If acEnt.BlockName Like "*Paper_Space*" Then
                                            'Ne pas charger les block Papier
                                        Else
                                            'Filtre Selon Nom de bloc authorisé dans le fichier Format.xml
                                            Dim BL As BlockReference = acEnt
                                            Dim AutorizeTxt As String = Replace(AutorizName, "REVOTHEME_", "")
                                            Dim BLnameTest As String = "[" & Replace(BL.Name, "MUT_", "") & "]"
                                            If InStr(AutorizeTxt.ToUpper, BLnameTest.ToUpper) <> 0 Then
                                                SelectedOjs.Add(acEnt, Ident & ":" & acEnt.ObjectId.ToString)
                                            End If
                                        End If

                                    Catch ex As Exception
                                        MsgBox(ex.Message)
                                    End Try
                                End If
                            End If
                        End If
                    Next
                End If
                '' Save the new object to the database
                acTrans.Commit()
                '' Dispose of the transaction
            End Using
            ' End selection of objects

            LblSelect.Text = "Sélection manuelle (" & SelectedOjs.Count & " objets)"

        Else
            MsgBox("Sélectionner d'abord un format", vbInformation + vbOKOnly, "Sélectionner un fichier")
        End If
    End Sub
    Function FiltreObjs(InputCollection As Collection, Filter As Double)

        Dim NewColl As New Collection
        Dim ListPlans As New List(Of String)
        For i = 0 To CkComPlan.CheckedItems.Count - 1
            Dim ComPlan() As String = Split(CkComPlan.CheckedItems(i).ToString, " : ")
            If ComPlan.Count = 2 Then
                ListPlans.Add(Format(CDbl(ComPlan(0)), "0000") & Format(CDbl(ComPlan(1)), "000000"))
            End If
        Next
        If ListPlans.Count = CkComPlan.Items.Count Then ListPlans.Clear()


        If cboTheme.SelectedIndex = 0 And ListPlans.Count = 0 Then
            'renvoi l'entier de la collection
            Me.LblSelect.Text = "Sélection manuelle (" & InputCollection.Count & " objets)"
            Return InputCollection

        Else 'filtre la collection

            For u = 1 To InputCollection.Count
                Dim Ent As Entity = InputCollection(u)
                Dim EntKey As String = (GetKey(InputCollection, u)).ToString

                Dim AddEntPlans As Boolean = False
                If ListPlans.Count = 0 Then AddEntPlans = True

                Try
                    If ListPlans.Count > 0 Then
                        Dim KeySplit() As String = Split(EntKey, ":")
                        For Each CheckPlan As String In ListPlans
                            If CheckPlan = KeySplit(0) Then
                                AddEntPlans = True ' OK pour ajout à la coll
                                Exit For
                            End If
                        Next


                    End If

                Catch
                    AddEntPlans = False
                End Try

                Dim AddEntLayer As Boolean = False
                Dim LA As String = Ent.Layer.ToString.ToUpper

                Select Case Filter
                    Case 0 ' Tous : 0
                        AddEntLayer = True ' ne filtre rien
                    Case 1 ' MO : Tous : 1
                        If LA Like "MO_*_PTS*" Or LA Like "MUT_*_PTS*" Or LA Like "RAD_*_PTS*" Or _
                              LA Like "MO_PFA*" Or LA Like "MUT_PFA*" Or LA Like "RAD_PFA*" Or _
                               LA Like "MO_PFP*" Or LA Like "MUT_PFP*" Or LA Like "RAD_PFP*" Then AddEntLayer = True
                    Case 2 ' MO : Tous : 2
                        If LA Like "MO_*_PTS*" Or LA Like "MO_PFA*" Or LA Like "MO_PFP*" Then AddEntLayer = True
                    Case 3 ' MO : 01 PF : 3 
                        If LA Like "MO_PFA*" Or LA Like "MO_PFP*" Then AddEntLayer = True
                    Case 4 ' MO : 02 CS : 4
                        If LA Like "MO_CS_PTS*" Then AddEntLayer = True
                    Case 5 ' MO : 03 OD : 5
                        If LA Like "MO_OD_PTS*" Then AddEntLayer = True
                    Case 6 ' MO : 06 BF : 6
                        If LA Like "MO_BF_PTS*" Then AddEntLayer = True
                    Case 7 ' MO : 09 COM : 7
                        If LA Like "MO_COM_PTS*" Then AddEntLayer = True
                    Case 8 ' MUT : Tous : 8
                        If LA Like "MUT_*_PTS*" Or LA Like "MUT_PFA*" Or LA Like "MUT_PFP*" Then AddEntLayer = True
                    Case 9 ' MUT : 01 PF : 9
                        If LA Like "MUT_PFA*" Or LA Like "MUT_PFP*" Then AddEntLayer = True
                    Case 10 ' MUT : 02 CS : 10
                        If LA Like "MUT_CS_PTS*" Then AddEntLayer = True
                    Case 11      ' MUT : 03 OD : 11
                        If LA Like "MUT_OD_PTS*" Then AddEntLayer = True
                    Case 12      ' MUT : 06 BF : 12
                        If LA Like "MUT_BF_PTS*" Then AddEntLayer = True
                    Case 13      ' MUT : 09 COM : 13
                        If LA Like "MUT_COM_PTS*" Then AddEntLayer = True
                    Case 14      ' IMPL : 14
                        If LA Like "IMPL_CONT_PTS*" Then AddEntLayer = True
                    Case 15      ' TOPO : 15
                        If LA Like "TOPO_PTS*" Then AddEntLayer = True
                    Case 16      ' RAD : 16
                        If LA Like "RAD_*PTS*" Or LA Like "RAD_PFA*" Or LA Like "RAD_PFP*" Then AddEntLayer = True
                    Case 17      ' SOUT : 17
                        If LA Like "SOUT_*" Then AddEntLayer = True


                        'Case 17      ' SOUT : Eau potable : 17
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 18      ' SOUT : E. P. (projet) : 18
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 19      ' SOUT : Eaux claires : 19
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 20      ' SOUT : E. C. (projet) : 20
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 21      ' SOUT : Eaux unitaires : 21
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 22      ' SOUT : Eaux usées : 22
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 23      ' SOUT : E. U. (projet) : 23
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 24      ' SOUT : Eclairage public : 24
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 25      ' SOUT : Electricité : 25
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 26      ' SOUT : Gaz : 26
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 27      ' SOUT : Téléphone : 27
                        '    If LA Like "xxx*" Then AddEntLayer = True
                        'Case 28      ' SOUT : TV : 28
                        '    If LA Like "xxx*" Then AddEntLayer = True
                End Select

                If AddEntPlans And AddEntLayer Then NewColl.Add(Ent, EntKey)
            Next

            Me.LblSelect.Text = "Sélection avec filtre (" & NewColl.Count & " objets)"
            Return NewColl
        End If

    End Function
    Private Function GetKey(Col As Collection, Index As Integer)
        Dim flg As Reflection.BindingFlags = Reflection.BindingFlags.Instance Or Reflection.BindingFlags.NonPublic
        Dim InternalList As Object = Col.GetType.GetMethod("InternalItemsList", flg).Invoke(Col, Nothing)
        Dim Item As Object = InternalList.GetType.GetProperty("Item", flg).GetValue(InternalList, {Index - 1})
        Dim Key As String = Item.GetType.GetField("m_Key", flg).GetValue(Item)
        Return Key
    End Function
    Private Sub cboTheme_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboTheme.SelectedIndexChanged
        FiltreObjs(SelectedOjs, cboTheme.SelectedIndex)
    End Sub
    Private Sub CkComPlan_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles CkComPlan.SelectedIndexChanged
        FiltreObjs(SelectedOjs, cboTheme.SelectedIndex)
    End Sub
    Private Sub btnNext_Click(sender As System.Object, e As System.EventArgs) Handles btnNext.Click


        'Ajout des objets en fonction des filtres
        Dim ParamFile As New Collection  'File Name,Filter Index, Obj Selected, Theme, Communes, Plans, Doublon

        Dim NewColl As Collection = FiltreObjs(SelectedOjs, cboTheme.SelectedIndex)

        If ActiveImport = True Then 'Importation

            'Reset
            ParamFile.Clear()
            Revo.RevoScript.ParamImportExportFrm.Clear()

            ' ParamFile   
            ParamFile.Add("") 'Sélection du fichier + Chargement des formats
            ParamFile.Add(-1) 'Filter Index,
            ParamFile.Add(Nothing) 'Sélection des objets manuelle (défaut tout le fichiers)
            ParamFile.Add(cboTheme.SelectedIndex) '4 Points du thèmes
            ParamFile.Add("") 'Points des communes/plans
            ParamFile.Add(cboxDoublons.CheckState) '6 Détection des doublons
            ParamFile.Add(Ck3D.CheckState) '7 Points 3D : OK
            Revo.RevoScript.ParamImportExportFrm = ParamFile
            Me.Close()

        Else 'Exportation

            'Check si ok
            Dim CheckobjsSelected As Boolean = False
            If ActiveImport = True Then
                CheckobjsSelected = True
            Else
                If NewColl.Count <> 0 Then CheckobjsSelected = True
            End If

            If TBNameFile.Text <> "" And CheckobjsSelected = True Then

                'Reset
                ParamFile.Clear()
                Revo.RevoScript.ParamImportExportFrm.Clear()

                Dim ListPlans As New List(Of String)
                For i = 0 To CkComPlan.CheckedItems.Count - 1
                    ListPlans.Add(CkComPlan.CheckedItems(i).ToString)
                Next

                ' ParamFile   
                ParamFile.Add(TBNameFile.Text) 'Sélection du fichier + Chargement des formats
                ParamFile.Add(NumFormat) 'Filter Index,
                ParamFile.Add(NewColl) 'Sélection des objets manuelle (défaut tout le fichiers)
                ParamFile.Add(cboTheme.SelectedIndex) 'index du Points du thèmes
                ParamFile.Add(ListPlans) 'Points des communes/plans
                ParamFile.Add(cboxDoublons.CheckState) 'Détection des doublons
                ParamFile.Add(Ck3D.CheckState) 'Points 3D : OK

                Revo.RevoScript.ParamImportExportFrm = ParamFile
                Me.Close()

            Else
                If NewColl.Count = 0 And ActiveImport = False Then
                    MsgBox("Aucun objet n'a été sélectionné" & vbCrLf & _
                           "Sélectionner des objets et contrôler les filtres", vbInformation + vbOKOnly, "Traitement des données")
                ElseIf Me.TBNameFile.Text = "" Then
                    MsgBox("Le nom du fichier et le format ne sont pas configuré", vbInformation + vbOKOnly, "Traitement des données")
                Else
                    MsgBox("Certain paramètre ne sont pas complétés", vbInformation + vbOKOnly, "Traitement des données")
                End If
            End If

        End If

    End Sub


    Function SelectPlans()
        Dim ListPlans As New List(Of String)

        Try


            'Entités administratives présentes dans le fichier
            Dim j As Object
            Dim strPlan As String = ""
            Dim strCom As String = ""

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

            'Sélection
            Dim objSelSet As Autodesk.AutoCAD.Interop.AcadSelectionSet
            Dim objAcadBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference

            'Restrictions pour la sélection
            Dim adDXFCode(2) As Short
            Dim adDXFGroup(2) As Object

            objSelSet = acDoc.SelectionSets.Add("JourCAD" & acDoc.SelectionSets.Count) 'Nom de la sélection
            adDXFCode(0) = 0
            adDXFGroup(0) = "INSERT" 'Type d'objet
            adDXFCode(1) = 2
            adDXFGroup(1) = "PLAN" 'Nom de bloc
            adDXFCode(2) = 8
            adDXFGroup(2) = "MO_RPL_PLAN" 'Calque

            'Effectue la sélection
            objSelSet.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , adDXFCode, adDXFGroup)

            'Boucle pour tous les éléments sélectionnés (polylignes)
            Dim strCode As String = ""
            Dim varAttributes As Object
            If objSelSet.Count <> 0 Then


                'Boucle pour chaque bloc de la couche "MO_RPL_PLAN"
                For Each objAcadBlock In objSelSet

                    '  MsgBox(TypeName(objAcadBlock))
                    'Dim BL As  = objAcadBlock

                    'Lecture des attributs
                    varAttributes = objAcadBlock.GetAttributes

                    For j = LBound(varAttributes) To UBound(varAttributes)
                        Select Case varAttributes(j).TagString
                            Case "IDENTDN"
                                strCom = Mid(varAttributes(j).TextString, 4, 3)
                            Case "NUMERO"
                                strPlan = varAttributes(j).TextString
                            Case "CODEPLAN"
                                strCode = varAttributes(j).TextString
                        End Select
                    Next

                    'Ajout de l'info à la ListView
                    'La fonction "décortique" une ligne "OBJE" Interlis => mise en forme bidon pour réutiliser la fctn
                    ListPlans.Add(strCom & " : " & strPlan)
                    'PlanToList("0 1 2 VD0" & strCom & "000000 " & strPlan & " 5 6 7 " & strCode & " 9", Me.LVPlans)

                Next objAcadBlock

            End If

            'Effacement de la sélection
            objSelSet.Delete()

        Catch

        End Try


        Return ListPlans
    End Function

  
End Class