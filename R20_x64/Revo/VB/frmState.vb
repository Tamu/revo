Public Class frmState

    Private Sub frmState_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.AcceptButton = BtnValid

        Dim ass As New Revo.RevoInfo
        Me.Icon = ass.Icon
    End Sub

    Private Sub State_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If e.CloseReason = 3 Then
            'If MsgBox("Souhaitez-vous vraiment fermer la fenêtre ?" _
            '            , MsgBoxStyle.Information + MsgBoxStyle.YesNo, "Stopper l'opération") = MsgBoxResult.No Then
            '           & vbCr & "Des données peuvents être perduent"
            'e.Cancel = True
            'Else
            ' Ferme la fenêtre   ("Stop toutes les opérations")
            'End If
        End If
    End Sub


    Private Sub BtnValid_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnValid.Click
        lbl_infos.Text = "..."
        ProgBar.Visible = True
        BtnValid.Visible = False
        BoxList.Visible = False
        Me.Close()
    End Sub
End Class
