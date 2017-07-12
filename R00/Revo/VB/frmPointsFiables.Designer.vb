<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmPointsFiables
#Region "Code généré par le Concepteur Windows Form "
	<System.Diagnostics.DebuggerNonUserCode()> Public Sub New()
		MyBase.New()
		'Cet appel est requis par le Concepteur Windows Form.
		InitializeComponent()
	End Sub
	'Form remplace la méthode Dispose pour nettoyer la liste des composants.
	<System.Diagnostics.DebuggerNonUserCode()> Protected Overloads Overrides Sub Dispose(ByVal Disposing As Boolean)
		If Disposing Then
			If Not components Is Nothing Then
				components.Dispose()
			End If
		End If
		MyBase.Dispose(Disposing)
	End Sub
	'Requise par le Concepteur Windows Form
	Private components As System.ComponentModel.IContainer
	Public ToolTip1 As System.Windows.Forms.ToolTip
	Public WithEvents lblTitle As System.Windows.Forms.Label
	Public WithEvents lblSubtitle As System.Windows.Forms.Label
	Public WithEvents Frame1 As System.Windows.Forms.GroupBox
    Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Label4 As System.Windows.Forms.Label
	Public WithEvents chkPF As System.Windows.Forms.CheckBox
	Public WithEvents chkCS As System.Windows.Forms.CheckBox
	Public WithEvents chkOD As System.Windows.Forms.CheckBox
	Public WithEvents chkBF As System.Windows.Forms.CheckBox
	Public WithEvents chkCOM As System.Windows.Forms.CheckBox
	Public WithEvents Frame3 As System.Windows.Forms.GroupBox
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.
    'Ne la modifiez pas à l'aide de l'éditeur de code.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame1 = New System.Windows.Forms.GroupBox()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.lblSubtitle = New System.Windows.Forms.Label()
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Frame3 = New System.Windows.Forms.GroupBox()
        Me.chkPF = New System.Windows.Forms.CheckBox()
        Me.chkCS = New System.Windows.Forms.CheckBox()
        Me.chkOD = New System.Windows.Forms.CheckBox()
        Me.chkBF = New System.Windows.Forms.CheckBox()
        Me.chkCOM = New System.Windows.Forms.CheckBox()
        Me.cmdDeleteGrid = New System.Windows.Forms.Button()
        Me.Frame1.SuspendLayout()
        Me.Frame2.SuspendLayout()
        Me.Frame3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame1
        '
        Me.Frame1.BackColor = System.Drawing.Color.Transparent
        Me.Frame1.Controls.Add(Me.lblTitle)
        Me.Frame1.Controls.Add(Me.lblSubtitle)
        Me.Frame1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame1.Location = New System.Drawing.Point(-8, -8)
        Me.Frame1.Name = "Frame1"
        Me.Frame1.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame1.Size = New System.Drawing.Size(728, 88)
        Me.Frame1.TabIndex = 0
        Me.Frame1.TabStop = False
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Location = New System.Drawing.Point(26, 28)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTitle.Size = New System.Drawing.Size(344, 16)
        Me.lblTitle.TabIndex = 0
        Me.lblTitle.Text = "Mise en évidence des points fiables (en planimétrie)"
        '
        'lblSubtitle
        '
        Me.lblSubtitle.BackColor = System.Drawing.Color.Transparent
        Me.lblSubtitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSubtitle.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubtitle.ForeColor = System.Drawing.Color.White
        Me.lblSubtitle.Location = New System.Drawing.Point(26, 52)
        Me.lblSubtitle.Name = "lblSubtitle"
        Me.lblSubtitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSubtitle.Size = New System.Drawing.Size(200, 16)
        Me.lblSubtitle.TabIndex = 1
        Me.lblSubtitle.Text = "Définir les objets à traiter"
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.Color.Transparent
        Me.Frame2.Controls.Add(Me.cmdNext)
        Me.Frame2.Controls.Add(Me.btnCancel)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(-8, 312)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(448, 64)
        Me.Frame2.TabIndex = 1
        Me.Frame2.TabStop = False
        '
        'cmdNext
        '
        Me.cmdNext.BackColor = System.Drawing.Color.LightGray
        Me.cmdNext.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app_active
        Me.cmdNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdNext.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.cmdNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdNext.Location = New System.Drawing.Point(269, 24)
        Me.cmdNext.Margin = New System.Windows.Forms.Padding(3, 2, 3, 3)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(91, 25)
        Me.cmdNext.TabIndex = 17
        Me.cmdNext.Tag = "Démarrer le flux sélectionné"
        Me.cmdNext.Text = "Valider"
        Me.cmdNext.UseVisualStyleBackColor = False
        '
        'btnCancel
        '
        Me.btnCancel.BackColor = System.Drawing.Color.LightGray
        Me.btnCancel.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCancel.Location = New System.Drawing.Point(29, 24)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(3, 2, 3, 3)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(91, 25)
        Me.btnCancel.TabIndex = 16
        Me.btnCancel.Tag = "Démarrer le flux sélectionné"
        Me.btnCancel.Text = "Annuler"
        Me.btnCancel.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.BackColor = System.Drawing.Color.Transparent
        Me.Label4.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label4.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.ForeColor = System.Drawing.Color.White
        Me.Label4.Location = New System.Drawing.Point(16, 96)
        Me.Label4.Name = "Label4"
        Me.Label4.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label4.Size = New System.Drawing.Size(256, 16)
        Me.Label4.TabIndex = 2
        Me.Label4.Text = "Entités (types de points) à analyser:"
        '
        'Frame3
        '
        Me.Frame3.BackColor = System.Drawing.Color.Transparent
        Me.Frame3.Controls.Add(Me.chkPF)
        Me.Frame3.Controls.Add(Me.chkCS)
        Me.Frame3.Controls.Add(Me.chkOD)
        Me.Frame3.Controls.Add(Me.chkBF)
        Me.Frame3.Controls.Add(Me.chkCOM)
        Me.Frame3.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame3.Location = New System.Drawing.Point(22, 119)
        Me.Frame3.Name = "Frame3"
        Me.Frame3.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame3.Size = New System.Drawing.Size(330, 136)
        Me.Frame3.TabIndex = 3
        Me.Frame3.TabStop = False
        '
        'chkPF
        '
        Me.chkPF.BackColor = System.Drawing.Color.Transparent
        Me.chkPF.Checked = True
        Me.chkPF.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkPF.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkPF.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkPF.ForeColor = System.Drawing.Color.White
        Me.chkPF.Location = New System.Drawing.Point(10, 16)
        Me.chkPF.Name = "chkPF"
        Me.chkPF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkPF.Size = New System.Drawing.Size(246, 23)
        Me.chkPF.TabIndex = 0
        Me.chkPF.Text = "01 PF: PFP1/2/3"
        Me.chkPF.UseVisualStyleBackColor = False
        '
        'chkCS
        '
        Me.chkCS.BackColor = System.Drawing.Color.Transparent
        Me.chkCS.Checked = True
        Me.chkCS.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCS.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCS.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCS.ForeColor = System.Drawing.Color.White
        Me.chkCS.Location = New System.Drawing.Point(10, 40)
        Me.chkCS.Name = "chkCS"
        Me.chkCS.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCS.Size = New System.Drawing.Size(198, 23)
        Me.chkCS.TabIndex = 1
        Me.chkCS.Text = "02 CS: points particuliers"
        Me.chkCS.UseVisualStyleBackColor = False
        '
        'chkOD
        '
        Me.chkOD.BackColor = System.Drawing.Color.Transparent
        Me.chkOD.Checked = True
        Me.chkOD.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkOD.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkOD.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkOD.ForeColor = System.Drawing.Color.White
        Me.chkOD.Location = New System.Drawing.Point(10, 64)
        Me.chkOD.Name = "chkOD"
        Me.chkOD.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkOD.Size = New System.Drawing.Size(198, 23)
        Me.chkOD.TabIndex = 2
        Me.chkOD.Text = "03 OD: points particuliers"
        Me.chkOD.UseVisualStyleBackColor = False
        '
        'chkBF
        '
        Me.chkBF.BackColor = System.Drawing.Color.Transparent
        Me.chkBF.Checked = True
        Me.chkBF.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBF.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkBF.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkBF.ForeColor = System.Drawing.Color.White
        Me.chkBF.Location = New System.Drawing.Point(10, 88)
        Me.chkBF.Name = "chkBF"
        Me.chkBF.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkBF.Size = New System.Drawing.Size(198, 23)
        Me.chkBF.TabIndex = 3
        Me.chkBF.Text = "04 BF: points-limite"
        Me.chkBF.UseVisualStyleBackColor = False
        '
        'chkCOM
        '
        Me.chkCOM.BackColor = System.Drawing.Color.Transparent
        Me.chkCOM.Checked = True
        Me.chkCOM.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkCOM.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkCOM.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkCOM.ForeColor = System.Drawing.Color.White
        Me.chkCOM.Location = New System.Drawing.Point(10, 112)
        Me.chkCOM.Name = "chkCOM"
        Me.chkCOM.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkCOM.Size = New System.Drawing.Size(232, 23)
        Me.chkCOM.TabIndex = 4
        Me.chkCOM.Text = "09 COM: points-limite territoriaux"
        Me.chkCOM.UseVisualStyleBackColor = False
        '
        'cmdDeleteGrid
        '
        Me.cmdDeleteGrid.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.cmdDeleteGrid.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdDeleteGrid.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.cmdDeleteGrid.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdDeleteGrid.Image = Global.Revo.My.Resources.Resources.delete
        Me.cmdDeleteGrid.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDeleteGrid.Location = New System.Drawing.Point(104, 273)
        Me.cmdDeleteGrid.Name = "cmdDeleteGrid"
        Me.cmdDeleteGrid.Size = New System.Drawing.Size(174, 24)
        Me.cmdDeleteGrid.TabIndex = 6
        Me.cmdDeleteGrid.Text = " Effacer les objets existants"
        Me.cmdDeleteGrid.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdDeleteGrid.UseVisualStyleBackColor = True
        '
        'frmPointsFiables
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(374, 371)
        Me.Controls.Add(Me.cmdDeleteGrid)
        Me.Controls.Add(Me.Frame1)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Frame3)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPointsFiables"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Frame1.ResumeLayout(False)
        Me.Frame2.ResumeLayout(False)
        Me.Frame3.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents cmdNext As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents cmdDeleteGrid As System.Windows.Forms.Button
#End Region 
End Class