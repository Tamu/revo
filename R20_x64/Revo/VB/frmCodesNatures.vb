Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic


Friend Class frmCodesNatures
    Inherits System.Windows.Forms.Form





    Private Sub cmdNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNext.Click

        'Teste si une nature est bien sélectionnée (et non un titre de domaine)
        Dim intNat As Short
        '  intNat = Val(Replace(VB.Left(tvNatures.SelectedNode.Text, 2), ".", ""))
        intNat = Val(Replace(tvNatures.SelectedNode.Name, "NAT", ""))
        If intNat > 0 Then

            Select Case Me.Tag
                Case "ImportGSI"
                    'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet frmImportPts.txtGSINat. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'frmImportPts.txtGSINat.Text = intNat
                Case "ImportXLS"
                    'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet frmImportPts.txtXLSNat. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'frmImportPts.txtXLSNat.Text = intNat
                Case "ImportLTop"
                    'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet frmImportPts.txtLTopNat. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    'frmImportPts.txtLTopNat.Text = intNat
                Case "InsertionPoint"
                    frmInsertionPts.SelectedNature = intNat
            End Select

            Me.Close()

        Else
            MsgBox("Sélectionner un code nature (1-99) dans l'arborescence !", MsgBoxStyle.Exclamation, "Nature de points")
        End If
    End Sub

    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click
        Me.Close()
    End Sub


    Private Sub tvNatures_AfterSelect(ByVal sender As System.Object, ByVal e As System.Windows.Forms.TreeViewEventArgs) Handles tvNatures.AfterSelect

    End Sub

    Private Sub tvNatures_MouseDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles tvNatures.MouseDoubleClick

        'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet Me.tvNatures.SelectedItem. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        '   If Val(Replace(VB.Left(Me.tvNatures.SelectedItem.Text, 2), ".", "")) > 0 Then Call cmdNext_Click(cmdNext, New System.EventArgs())

    End Sub

    Private Sub frmCodesNatures_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Me.AcceptButton = cmdNext
        Me.CancelButton = cmdCancel

    End Sub
End Class