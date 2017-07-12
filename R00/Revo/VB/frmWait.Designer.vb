<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWait
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmWait))
        Me.lblInfo1 = New System.Windows.Forms.Label()
        Me.lblInfo2 = New System.Windows.Forms.Label()
        Me.lblStatus1 = New System.Windows.Forms.Label()
        Me.progBar1 = New System.Windows.Forms.ProgressBar()
        Me.lblStatus2 = New System.Windows.Forms.Label()
        Me.progBar2 = New System.Windows.Forms.ProgressBar()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.lblSubTitle = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblInfo1
        '
        Me.lblInfo1.AutoSize = True
        Me.lblInfo1.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfo1.ForeColor = System.Drawing.Color.White
        Me.lblInfo1.Location = New System.Drawing.Point(12, 73)
        Me.lblInfo1.Name = "lblInfo1"
        Me.lblInfo1.Size = New System.Drawing.Size(56, 16)
        Me.lblInfo1.TabIndex = 1
        Me.lblInfo1.Text = "lblInfo1"
        '
        'lblInfo2
        '
        Me.lblInfo2.AutoSize = True
        Me.lblInfo2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfo2.ForeColor = System.Drawing.Color.White
        Me.lblInfo2.Location = New System.Drawing.Point(12, 98)
        Me.lblInfo2.Name = "lblInfo2"
        Me.lblInfo2.Size = New System.Drawing.Size(383, 14)
        Me.lblInfo2.TabIndex = 2
        Me.lblInfo2.Text = "Cette opération peut prendre quelques minutes, veuillez patienter..."
        '
        'lblStatus1
        '
        Me.lblStatus1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus1.ForeColor = System.Drawing.Color.White
        Me.lblStatus1.Location = New System.Drawing.Point(12, 130)
        Me.lblStatus1.Name = "lblStatus1"
        Me.lblStatus1.Size = New System.Drawing.Size(455, 14)
        Me.lblStatus1.TabIndex = 3
        Me.lblStatus1.Text = "0 ligne lue"
        Me.lblStatus1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'progBar1
        '
        Me.progBar1.Location = New System.Drawing.Point(15, 147)
        Me.progBar1.Name = "progBar1"
        Me.progBar1.Size = New System.Drawing.Size(440, 23)
        Me.progBar1.TabIndex = 4
        '
        'lblStatus2
        '
        Me.lblStatus2.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStatus2.ForeColor = System.Drawing.Color.White
        Me.lblStatus2.Location = New System.Drawing.Point(12, 182)
        Me.lblStatus2.Name = "lblStatus2"
        Me.lblStatus2.Size = New System.Drawing.Size(455, 14)
        Me.lblStatus2.TabIndex = 5
        Me.lblStatus2.Text = "0%"
        Me.lblStatus2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'progBar2
        '
        Me.progBar2.Location = New System.Drawing.Point(15, 201)
        Me.progBar2.Name = "progBar2"
        Me.progBar2.Size = New System.Drawing.Size(440, 23)
        Me.progBar2.TabIndex = 6
        '
        'btnCancel
        '
        Me.btnCancel.BackgroundImage = CType(resources.GetObject("btnCancel.BackgroundImage"), System.Drawing.Image)
        Me.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCancel.Location = New System.Drawing.Point(389, 9)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(66, 24)
        Me.btnCancel.TabIndex = 7
        Me.btnCancel.Tag = ""
        Me.btnCancel.Text = "Annuler "
        Me.btnCancel.UseVisualStyleBackColor = True
        Me.btnCancel.Visible = False
        '
        'lblSubTitle
        '
        Me.lblSubTitle.AutoSize = True
        Me.lblSubTitle.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubTitle.ForeColor = System.Drawing.Color.White
        Me.lblSubTitle.Location = New System.Drawing.Point(12, 32)
        Me.lblSubTitle.Name = "lblSubTitle"
        Me.lblSubTitle.Size = New System.Drawing.Size(63, 14)
        Me.lblSubTitle.TabIndex = 9
        Me.lblSubTitle.Text = "lblSubTitle"
        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Location = New System.Drawing.Point(12, 9)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(49, 16)
        Me.lblTitle.TabIndex = 8
        Me.lblTitle.Text = "lblTitle"
        '
        'frmWait
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DimGray
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(469, 240)
        Me.Controls.Add(Me.lblSubTitle)
        Me.Controls.Add(Me.lblTitle)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.progBar2)
        Me.Controls.Add(Me.lblStatus2)
        Me.Controls.Add(Me.progBar1)
        Me.Controls.Add(Me.lblStatus1)
        Me.Controls.Add(Me.lblInfo2)
        Me.Controls.Add(Me.lblInfo1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmWait"
        Me.Opacity = 0.85R
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Travail en cours ..."
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblInfo1 As System.Windows.Forms.Label
    Friend WithEvents lblInfo2 As System.Windows.Forms.Label
    Friend WithEvents lblStatus1 As System.Windows.Forms.Label
    Friend WithEvents progBar1 As System.Windows.Forms.ProgressBar
    Friend WithEvents lblStatus2 As System.Windows.Forms.Label
    Friend WithEvents progBar2 As System.Windows.Forms.ProgressBar
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents lblSubTitle As System.Windows.Forms.Label
    Friend WithEvents lblTitle As System.Windows.Forms.Label
End Class
