Imports Autodesk.AutoCAD.DatabaseServices
Imports frms = System.Windows.Forms
Imports System.IO
Public Class frmEDT

    Private Sub frmEDT_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.AcceptButton = btnNext
        Me.CancelButton = btnCancel

        Dim strVal As String
        'Destination directory report
        Dim Ass As New Revo.RevoInfo
        strVal = Ass.EDTFolder 'LireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "EDTFolder")
        Me.txtDestFolder.Text = strVal
        'RF Control Surface
        strVal = Ass.EDTCheckSurf 'LireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "EDTCheckSurf")
        Me.chkControleRF.Checked = Convert.ToBoolean(strVal = "1")
        'Ignore numbers Plans
        strVal = Ass.EDTIgnoreRPL ' LireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "EDTIgnoreRPL")
        Me.chkIgnorerRPL.Checked = Convert.ToBoolean(strVal = "1")
    End Sub

    Private Sub btnBrowse_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBrowse.Click
        'Allow the user to select a folder
        Dim strFolder As String = BrowseForFolder("Sélectionner le dossier dans lequel enregistrer le fichier EDT résultat:")
        'Check that the return value is not empty
        If Not strFolder.Equals(String.Empty) Then
            'Add Backslash to the selected folder path if it does not already end with one
            If Not strFolder.EndsWith("\") Then strFolder = strFolder & "\"
            'Set the destination folder text box value
            txtDestFolder.Text = strFolder

            'Write the registry values
            Dim Connect As New Revo.connect
            Dim Ass As New Revo.RevoInfo
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "EDTFolder".ToLower, Me.txtDestFolder.Text)
        End If
    End Sub

    Private Sub btnCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

    Private Sub btnSupprSel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSupprSel.Click
        'Check that there is something selected
        If lstLVObjets.SelectedIndex <> -1 Then
            'Remove the selected item
            lstLVObjets.Items.RemoveAt(lstLVObjets.SelectedIndex)
        End If
        'Disable the remove button if no objects in the list
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

    ''' <summary>
    ''' Asks the user to select either the polyline or the block to process
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    ''' <remarks></remarks>
    Private Sub btnDigit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDigit.Click
        Me.Hide()
        Dim blkID As ObjectId = Nothing
        'Call the function to ask the user to select on screen
        Dim objPoly As Polyline = GetPolylineOnScreen("MO_BF,MUT_BF", blkID)
        If objPoly Is Nothing Then
            frms.MessageBox.Show("Aucun contour de bien-fonds sélectionné !", "Sélection incorrecte", _
                                 Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
        Else
            If blkID.IsNull Then blkID = GetBFBlockByPolyline(objPoly, objPoly.Layer & "_IMMEUBLE")
            If blkID.IsNull Then
                frms.MessageBox.Show("Aucun contour de bien-fonds sélectionné !", "Sélection incorrecte", _
                                 Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
            Else

                If lstLVObjets.Items.Count = 0 Then
                    lstLVObjets.DisplayMember = "DisplayVal"
                    lstLVObjets.ValueMember = "PolyLayer"
                End If
                Dim EDTItem As EDTListItem = GetInfoFromBlock(blkID)
                EDTItem.PolyLayer = objPoly.Layer
                EDTItem.PolyID = objPoly.ObjectId
                EDTItem.ID = "BF" & EDTItem.ComNum
                EDTItem.DisplayVal = objPoly.Layer & " " & EDTItem.ComNum
                AddItemToList(EDTItem)
            End If
        End If

        If lstLVObjets.Items.Count > 0 Then
            btnSupprSel.Enabled = True
            btnNext.Enabled = btnSupprSel.Enabled
            btnNext.BackgroundImage = My.Resources.bouton_revo_app_active
            btnDigit.BackgroundImage = My.Resources.bouton_revo_app
        End If

        Me.Show()
    End Sub

    Private Sub btnNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNext.Click
        'Check that the destination folder for the report has been set
        If txtDestFolder.Text = "" Then
            'Not set so warn and exit
            frms.MessageBox.Show("Vous devez d'abord définir le dossier de destination du fichier rapport !", "Saisie incomplète", _
                                  Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
            Exit Sub
        Else
            'Check the the folder sepcified actually exists
            If Not Directory.Exists(txtDestFolder.Text) Then
                'Folder doesn;t exist so warn and exit
                frms.MessageBox.Show("Le dossier destination spécifié n'existe pas !", "Saisie incorrecte", _
                                      Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                Exit Sub
            End If
            'Write the registry values
            Dim Connect As New Revo.connect
            Dim Ass As New Revo.RevoInfo
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "EDTFolder".ToLower, Me.txtDestFolder.Text)
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "EDTCheckSurf".ToLower, Math.Abs(CInt(Me.chkControleRF.Checked)).ToString)
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "EDTIgnoreRPL".ToLower, Math.Abs(CInt(Me.chkIgnorerRPL.Checked)).ToString)
            'EcrireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "EDTFolder", Me.txtDestFolder.Text)
            'EcrireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "EDTCheckSurf", Math.Abs(CInt(Me.chkControleRF.Checked)).ToString)
            'EcrireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "EDTIgnoreRPL", Math.Abs(CInt(Me.chkIgnorerRPL.Checked)).ToString)

            'Store the settings in variables
            EDTFolder = Me.txtDestFolder.Text
            EDTCheckSurf = Me.chkControleRF.Checked.ToString
            EDTIgnoreRPL = Me.chkIgnorerRPL.Checked.ToString
            'Get the items from the list
            Dim lst As New List(Of EDTListItem)
            For i As Integer = 0 To lstLVObjets.Items.Count - 1
                lst.Add(CType(lstLVObjets.Items(i), EDTListItem))
            Next
            'Check that we've got some items to process
            If lst.Count > 0 Then
                'Hide the form
                Me.Hide()
                'And run the process on the list
                CalculEDT(lst, chkIgnorerRPL.Checked, chkControleRF.Checked)
            End If
        End If
    End Sub

    ''' <summary>
    ''' Adds the supplied item to the list if it is not already there
    ''' </summary>
    ''' <param name="EDTItem">The item to add to the list</param>
    ''' <remarks></remarks>
    Private Sub AddItemToList(ByVal EDTItem As EDTListItem)
        Dim ls As List(Of EDTListItem) = Nothing
        Try
            'Get the items already in the list
            ls = lstLVObjets.Items.OfType(Of EDTListItem).ToList
        Catch ex As Exception
            'Nothing in the list
        End Try
        If Not ls Is Nothing Then
            'If there are some items in the list then select any items with the ID of the item we're trying to add
            'If the count is greater than 0 then the item alreayd exists in the list so just exit
            If ls.Where(Function(x) x.ID = EDTItem.ID).ToList.Count > 0 Then Exit Sub
        End If
        'The list is empty or the item we're adding is not in the list so add it.
        lstLVObjets.Items.Add(EDTItem)
        lstLVObjets.Refresh()
    End Sub

    Private Sub btnXdata_Click(sender As System.Object, e As System.EventArgs) Handles btnXdata.Click
        Me.Hide()
        Dim MyCmd As New Revo.MyCommands
        MyCmd.RevoBatXdata()
        Me.Show()
    End Sub
End Class