Imports Autodesk.AutoCAD.DatabaseServices
Imports frms = System.Windows.Forms
Public Class frmPER
    Private Sub frmPer_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.AcceptButton = btnNext
        Me.CancelButton = btnCancel

        Dim Ass As New Revo.RevoInfo
        Dim strVal As String = Ass.PERFolder ' LireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "PERFolder")
        txtDestFolder.Text = strVal
    End Sub
    
    Private Sub btnSupprSel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSupprSel.Click
        'Check that there is something selected
        If lstLVObjets.SelectedIndex <> -1 Then
            'Remove the selected item
            lstLVObjets.Items.RemoveAt(lstLVObjets.SelectedIndex)
        End If
        'Disable the remove button if no objects in the list
        'btnSupprSel.Enabled = lstLVObjets.Items.Count > 0
        'And disable the Go button if nothing to process
        'btnNext.Enabled = btnSupprSel.Enabled

        If lstLVObjets.Items.Count > 0 Then
            btnSupprSel.Enabled = True
            btnNext.Enabled = btnSupprSel.Enabled
            btnNext.BackgroundImage = My.Resources.bouton_revo_app_active
            btnDigit.BackgroundImage = My.Resources.bouton_revo_app
        Else
            btnSupprSel.Enabled = False
            btnNext.Enabled = btnSupprSel.Enabled
            btnNext.BackgroundImage = My.Resources.bouton_revo_app
            btnDigit.BackgroundImage = My.Resources.bouton_revo_app_active
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        'Allow the user to select a folder
        Dim strFolder As String = BrowseForFolder("Sélectionner le dossier dans lequel enregistrer le fichier PER résultat:")
        'Check that the return value is not empty
        If Not strFolder.Equals(String.Empty) Then
            'Add Backslash to the selected folder path if it does not already end with one
            If Not strFolder.EndsWith("\") Then strFolder = strFolder & "\"
            'Set the destination folder text box value
            txtDestFolder.Text = strFolder

            Dim Connect As New Revo.connect
            Dim Ass As New Revo.RevoInfo
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "PERFolder".ToLower, txtDestFolder.Text)

        End If
    End Sub
    
    Private Sub btnDigit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDigit.Click
        Me.Hide()

        Dim doc As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ' Using trans As Transaction = db.TransactionManager.StartTransaction
        '      trans.Commit()
        ' End Using

        Try
           
         
            Dim blkId As ObjectId = Nothing
            'Ask the user to select the polyline
            Dim objPoly As Polyline = GetPolylineOnScreen("MO_BF,MUT_BF,MO_BF_DDP,MUT_BF_DDP,MO_CS_BAT,MUT_CS_BAT,MO_ODS_BATSOUT,MUT_ODS_BATSOUT", blkId)
            'check that they selected something
            If objPoly Is Nothing Then
                frms.MessageBox.Show("Aucun contour de bien-fonds ou de bâtiment sélectionné !", "Sélection incorrecte", _
                                     Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
            Else
                Dim PerItem As New PerListItem
                Dim strCom As String = ""
                Dim strNum As String = ""
                'Land
                If objPoly.Layer.Contains("_BF") Then
                    If blkId.IsNull Then
                        frms.MessageBox.Show("Aucun contour de bien-fonds ou de bâtiment sélectionné !", "Sélection incorrecte", _
                                              Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                    Else
                        If lstLVObjets.Items.Count = 0 Then
                            lstLVObjets.DisplayMember = "DisplayVal"
                            lstLVObjets.ValueMember = "PolyLayer"
                        End If
                        Dim edtItem As EDTListItem = GetInfoFromBlock(blkId)
                        With PerItem
                            .Com = edtItem.Com
                            .Num = edtItem.Num
                            .BlockID = edtItem.BlockID
                            .PolyLayer = objPoly.Layer
                            .PolyID = objPoly.ObjectId
                            .ID = "BF" & PerItem.ComNum
                            .DisplayVal = objPoly.Layer & " " & PerItem.ComNum
                        End With
                        AddItemToList(PerItem)
                    End If
                ElseIf objPoly.Layer.Contains("_CS_BAT") Then
                    Dim tv() As TypedValue = objPoly.XData.AsArray
                    strNum = tv(2).Value.ToString
                    'Number not found if new building
                    If strNum = "" Then
                        'Find the number of the building block inside the polyline
                        Dim blkBat As ObjectId = GetBATBlockByPolyline(objPoly.ObjectId, objPoly.Layer & "_NUM")
                        Dim strDesign As String = ""
                        If Not blkBat.IsNull Then
                            strNum = GetBlockAttribute(blkBat, "NUMERO")
                            strDesign = GetBlockAttribute(blkBat, "DESIGNATION")
                            AddRegAppTableRecord("REVO")
                            'Declare a new result buffer
                            Dim acResBuf As New ResultBuffer
                            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "REVO"))
                            'The ID is where the data is held
                            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "ID" & strNum))
                            'Add the data parts to the resultbuffer
                            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, strNum))
                            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, strDesign))
                            SetPolyXData(objPoly.ObjectId, acResBuf)
                        End If
                    End If
                    Dim strBF As String = GetBFNumFromBAT(objPoly.ObjectId, strCom)
                    PerItem = New PerListItem
                    With PerItem
                        .PolyID = objPoly.ObjectId
                        .PolyLayer = objPoly.Layer
                        .Com = strCom
                        .Num = strNum
                        .DisplayVal = objPoly.Layer & " " & PerItem.ComNum
                        .BF = strBF
                    End With
                    AddItemToList(PerItem)


                    'ElseIf objPoly.Layer.Contains("_ODS_BATS") Then
                    '    Dim tv() As TypedValue = objPoly.XData.AsArray
                    '    strNum = tv(2).Value.ToString
                    '    'Number not found if new building
                    '    If strNum = "" Then
                    '        'Find the number of the building block inside the polyline
                    '        Dim blkBat As ObjectId = GetBATBlockByPolyline(objPoly.ObjectId, objPoly.Layer & "_NUM")
                    '        Dim strDesign As String = ""
                    '        If Not blkBat.IsNull Then
                    '            strNum = GetBlockAttribute(blkBat, "NUMERO")
                    '            strDesign = GetBlockAttribute(blkBat, "DESIGNATION")
                    '            AddRegAppTableRecord("REVO")
                    '            'Declare a new result buffer
                    '            Dim acResBuf As New ResultBuffer
                    '            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "REVO"))
                    '            'The ID is where the data is held
                    '            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "ID" & strNum))
                    '            'Add the data parts to the resultbuffer
                    '            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, strNum))
                    '            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, strDesign))
                    '            SetPolyXData(objPoly.ObjectId, acResBuf)
                    '        End If
                    '    End If
                    '    Dim strBF As String = GetBFNumFromBAT(objPoly.ObjectId, strCom)
                    '    PerItem = New PerListItem
                    '    With PerItem
                    '        .PolyID = objPoly.ObjectId
                    '        .PolyLayer = objPoly.Layer
                    '        .Com = strCom
                    '        .Num = strNum
                    '        .DisplayVal = objPoly.Layer & " " & PerItem.ComNum
                    '        .BF = strBF
                    '    End With
                    '    AddItemToList(PerItem)
                End If
            End If

            If lstLVObjets.Items.Count > 0 Then
                btnSupprSel.Enabled = True
                btnNext.Enabled = btnSupprSel.Enabled
                btnNext.BackgroundImage = My.Resources.bouton_revo_app_active
                btnDigit.BackgroundImage = My.Resources.bouton_revo_app
            End If




        Catch 'ex As Exception
            MsgBox("Erreur de sélection", MsgBoxStyle.Critical, "PER : Sélection")
        End Try

        doc.LockDocument.Dispose()


        Me.Show()
    End Sub
    

    Private Sub AddItemToList(ByVal PerItem As PerListItem)
        Dim ls As List(Of PerListItem) = Nothing
        Try
            ls = lstLVObjets.Items.OfType(Of PerListItem).ToList
        Catch ex As Exception
            'Nothing in the list
        End Try
        If Not ls Is Nothing Then
            If ls.Where(Function(x) x.ID = PerItem.ID).ToList.Count > 0 Then Exit Sub
        End If
        lstLVObjets.Items.Add(PerItem)
        lstLVObjets.Refresh()
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        If txtDestFolder.Text = "" Then
            frms.MessageBox.Show("Vous devez d'abord définir le dossier de destination du fichier rapport !", "Saisie incomplète", _
                                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
        Else
            If Not System.IO.Directory.Exists(txtDestFolder.Text) Then
                frms.MessageBox.Show("Le dossier destination spécifié n'existe pas !", "Saisie incorrecte", _
                                      Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
            Else
                Dim Connect As New Revo.connect
                Dim Ass As New Revo.RevoInfo
                Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "PERFolder".ToLower, txtDestFolder.Text)
                'EcrireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "PERFolder", txtDestFolder.Text)
                Dim lst As New List(Of PerListItem)
                For i As Integer = 0 To lstLVObjets.Items.Count - 1
                    lst.Add(CType(lstLVObjets.Items(i), PerListItem))
                Next
                If lst.Count > 0 Then
                    Me.Hide()
                    CalculPER(lst)
                    Me.Close()
                End If
            End If
        End If
    End Sub

End Class