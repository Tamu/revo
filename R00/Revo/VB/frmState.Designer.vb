<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmState
    Inherits System.Windows.Forms.Form

    'Form remplace la méthode Dispose pour nettoyer la liste des composants.
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

    'Requise par le Concepteur Windows Form
    Private components As System.ComponentModel.IContainer

    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.  
    'Ne la modifiez pas à l'aide de l'éditeur de code.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmState))
        Me.ProgBar = New System.Windows.Forms.ProgressBar()
        Me.lbl_infos = New System.Windows.Forms.Label()
        Me.BoxList = New System.Windows.Forms.ComboBox()
        Me.BtnValid = New System.Windows.Forms.Button()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'ProgBar
        '
        Me.ProgBar.AccessibleDescription = ""
        Me.ProgBar.AccessibleName = ""
        Me.ProgBar.Location = New System.Drawing.Point(16, 55)
        Me.ProgBar.Name = "ProgBar"
        Me.ProgBar.Size = New System.Drawing.Size(556, 23)
        Me.ProgBar.Step = 1
        Me.ProgBar.TabIndex = 0
        Me.ProgBar.Tag = ""
        '
        'lbl_infos
        '
        Me.lbl_infos.AutoSize = True
        Me.lbl_infos.BackColor = System.Drawing.Color.Transparent
        Me.lbl_infos.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_infos.ForeColor = System.Drawing.Color.White
        Me.lbl_infos.Location = New System.Drawing.Point(12, 13)
        Me.lbl_infos.Name = "lbl_infos"
        Me.lbl_infos.Size = New System.Drawing.Size(166, 20)
        Me.lbl_infos.TabIndex = 1
        Me.lbl_infos.Text = "Traitement en cours ..."
        '
        'BoxList
        '
        Me.BoxList.FormattingEnabled = True
        Me.BoxList.Location = New System.Drawing.Point(16, 55)
        Me.BoxList.Name = "BoxList"
        Me.BoxList.Size = New System.Drawing.Size(483, 21)
        Me.BoxList.TabIndex = 2
        Me.BoxList.Visible = False
        '
        'BtnValid
        '
        Me.BtnValid.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.BtnValid.BackColor = System.Drawing.Color.LightGray
        Me.BtnValid.BackgroundImage = CType(resources.GetObject("BtnValid.BackgroundImage"), System.Drawing.Image)
        Me.BtnValid.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.BtnValid.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.BtnValid.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnValid.Location = New System.Drawing.Point(505, 53)
        Me.BtnValid.Name = "BtnValid"
        Me.BtnValid.Size = New System.Drawing.Size(67, 23)
        Me.BtnValid.TabIndex = 14
        Me.BtnValid.Text = "Valider"
        Me.BtnValid.UseVisualStyleBackColor = False
        Me.BtnValid.Visible = False
        '
        'btnSelect
        '
        Me.btnSelect.Anchor = System.Windows.Forms.AnchorStyles.Right
        Me.btnSelect.BackColor = System.Drawing.Color.LightGray
        Me.btnSelect.BackgroundImage = CType(resources.GetObject("btnSelect.BackgroundImage"), System.Drawing.Image)
        Me.btnSelect.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSelect.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnSelect.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSelect.Location = New System.Drawing.Point(434, 13)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(138, 23)
        Me.btnSelect.TabIndex = 15
        Me.btnSelect.Text = "Sélectionner des objets"
        Me.btnSelect.UseVisualStyleBackColor = False
        Me.btnSelect.Visible = False
        '
        'frmState
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gray
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.ClientSize = New System.Drawing.Size(584, 90)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.BtnValid)
        Me.Controls.Add(Me.BoxList)
        Me.Controls.Add(Me.lbl_infos)
        Me.Controls.Add(Me.ProgBar)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmState"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "REVO"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents ProgBar As System.Windows.Forms.ProgressBar
    Friend WithEvents lbl_infos As System.Windows.Forms.Label
    Friend WithEvents BoxList As System.Windows.Forms.ComboBox
    Friend WithEvents BtnValid As System.Windows.Forms.Button
    Friend WithEvents btnSelect As System.Windows.Forms.Button
End Class
