<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmInsertionPts
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
    Public WithEvents txtOri As System.Windows.Forms.TextBox
	Public WithEvents Label75 As System.Windows.Forms.Label
	Public WithEvents txtBlockID As System.Windows.Forms.TextBox
	Public WithEvents Frame2 As System.Windows.Forms.GroupBox
	Public WithEvents Label3 As System.Windows.Forms.Label
	Public WithEvents cboTheme As System.Windows.Forms.ComboBox
	Public WithEvents lblTheme As System.Windows.Forms.Label
	Public WithEvents cboType As System.Windows.Forms.ComboBox
	Public WithEvents lblType As System.Windows.Forms.Label
	Public WithEvents Label12 As System.Windows.Forms.Label
	Public WithEvents Label13 As System.Windows.Forms.Label
	Public WithEvents Label14 As System.Windows.Forms.Label
	Public WithEvents Label15 As System.Windows.Forms.Label
	Public WithEvents txtX As System.Windows.Forms.TextBox
	Public WithEvents txtY As System.Windows.Forms.TextBox
	Public WithEvents txtZ As System.Windows.Forms.TextBox
    Public WithEvents lblAttrVariable As System.Windows.Forms.Label
	Public WithEvents cboAttrVariable As System.Windows.Forms.ComboBox
	Public WithEvents lblExact As System.Windows.Forms.Label
	Public WithEvents cboExact As System.Windows.Forms.ComboBox
    Public WithEvents Label87 As System.Windows.Forms.Label
	Public WithEvents txtIdentDN As System.Windows.Forms.TextBox
	Public WithEvents Label85 As System.Windows.Forms.Label
	Public WithEvents txtPlan As System.Windows.Forms.TextBox
	Public WithEvents lblPlan As System.Windows.Forms.Label
	Public WithEvents txtCom As System.Windows.Forms.TextBox
	Public WithEvents Label74 As System.Windows.Forms.Label
	Public WithEvents txtNo As System.Windows.Forms.TextBox
	Public WithEvents lblCom As System.Windows.Forms.Label
	Public WithEvents cboSigne As System.Windows.Forms.ComboBox
	Public WithEvents lblSigne As System.Windows.Forms.Label
	Public WithEvents Label88 As System.Windows.Forms.Label
	Public WithEvents Label78 As System.Windows.Forms.Label
	Public WithEvents Label77 As System.Windows.Forms.Label
	Public WithEvents lblPrecAlt As System.Windows.Forms.Label
	Public WithEvents lblFiabAlt As System.Windows.Forms.Label
	Public WithEvents txtPrecPlan As System.Windows.Forms.TextBox
	Public WithEvents txtPrecAlt As System.Windows.Forms.TextBox
	Public WithEvents cboFiabPlan As System.Windows.Forms.ComboBox
	Public WithEvents cboFiabAlt As System.Windows.Forms.ComboBox
	Public WithEvents Label89 As System.Windows.Forms.Label
    Public WithEvents chkDigitSerie As System.Windows.Forms.CheckBox
	Public WithEvents Label90 As System.Windows.Forms.Label
	Public WithEvents cboCat As System.Windows.Forms.ComboBox
	Public WithEvents txtComment As System.Windows.Forms.TextBox
	Public WithEvents lblComment As System.Windows.Forms.Label
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.
    'Ne la modifiez pas à l'aide de l'éditeur de code.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmInsertionPts))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Frame2 = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.txtOri = New System.Windows.Forms.TextBox()
        Me.Label75 = New System.Windows.Forms.Label()
        Me.txtBlockID = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.cboTheme = New System.Windows.Forms.ComboBox()
        Me.lblTheme = New System.Windows.Forms.Label()
        Me.cboType = New System.Windows.Forms.ComboBox()
        Me.lblType = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.txtX = New System.Windows.Forms.TextBox()
        Me.txtY = New System.Windows.Forms.TextBox()
        Me.txtZ = New System.Windows.Forms.TextBox()
        Me.lblAttrVariable = New System.Windows.Forms.Label()
        Me.cboAttrVariable = New System.Windows.Forms.ComboBox()
        Me.lblExact = New System.Windows.Forms.Label()
        Me.cboExact = New System.Windows.Forms.ComboBox()
        Me.Label87 = New System.Windows.Forms.Label()
        Me.txtIdentDN = New System.Windows.Forms.TextBox()
        Me.Label85 = New System.Windows.Forms.Label()
        Me.txtPlan = New System.Windows.Forms.TextBox()
        Me.lblPlan = New System.Windows.Forms.Label()
        Me.txtCom = New System.Windows.Forms.TextBox()
        Me.Label74 = New System.Windows.Forms.Label()
        Me.txtNo = New System.Windows.Forms.TextBox()
        Me.lblCom = New System.Windows.Forms.Label()
        Me.cboSigne = New System.Windows.Forms.ComboBox()
        Me.lblSigne = New System.Windows.Forms.Label()
        Me.Label88 = New System.Windows.Forms.Label()
        Me.Label78 = New System.Windows.Forms.Label()
        Me.Label77 = New System.Windows.Forms.Label()
        Me.lblPrecAlt = New System.Windows.Forms.Label()
        Me.lblFiabAlt = New System.Windows.Forms.Label()
        Me.txtPrecPlan = New System.Windows.Forms.TextBox()
        Me.txtPrecAlt = New System.Windows.Forms.TextBox()
        Me.cboFiabPlan = New System.Windows.Forms.ComboBox()
        Me.cboFiabAlt = New System.Windows.Forms.ComboBox()
        Me.Label89 = New System.Windows.Forms.Label()
        Me.chkDigitSerie = New System.Windows.Forms.CheckBox()
        Me.Label90 = New System.Windows.Forms.Label()
        Me.cboCat = New System.Windows.Forms.ComboBox()
        Me.txtComment = New System.Windows.Forms.TextBox()
        Me.lblComment = New System.Windows.Forms.Label()
        Me.cmdDigit = New System.Windows.Forms.Button()
        Me.cmdListeNat = New System.Windows.Forms.Button()
        Me.lblSubtitle = New System.Windows.Forms.Label()
        Me.LVPlans = New System.Windows.Forms.ListView()
        Me.Commune = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Plan = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.NT = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Mensuration = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Alt3D = New System.Windows.Forms.RadioButton()
        Me.Alt2D = New System.Windows.Forms.RadioButton()
        Me.Frame2.SuspendLayout()
        Me.SuspendLayout()
        '
        'Frame2
        '
        Me.Frame2.BackColor = System.Drawing.Color.Transparent
        Me.Frame2.Controls.Add(Me.cmdCancel)
        Me.Frame2.Controls.Add(Me.cmdNext)
        Me.Frame2.Controls.Add(Me.txtOri)
        Me.Frame2.Controls.Add(Me.Label75)
        Me.Frame2.Controls.Add(Me.txtBlockID)
        Me.Frame2.ForeColor = System.Drawing.SystemColors.ControlText
        Me.Frame2.Location = New System.Drawing.Point(-8, 552)
        Me.Frame2.Name = "Frame2"
        Me.Frame2.Padding = New System.Windows.Forms.Padding(0)
        Me.Frame2.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Frame2.Size = New System.Drawing.Size(760, 64)
        Me.Frame2.TabIndex = 18
        Me.Frame2.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.cmdCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Location = New System.Drawing.Point(30, 20)
        Me.cmdCancel.Name = "cmdCancel"
        Me.cmdCancel.Size = New System.Drawing.Size(66, 24)
        Me.cmdCancel.TabIndex = 50
        Me.cmdCancel.Tag = ""
        Me.cmdCancel.Text = "Annuler"
        Me.cmdCancel.UseVisualStyleBackColor = True
        '
        'cmdNext
        '
        Me.cmdNext.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.cmdNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdNext.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.cmdNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdNext.Location = New System.Drawing.Point(675, 24)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(66, 24)
        Me.cmdNext.TabIndex = 0
        Me.cmdNext.Tag = ""
        Me.cmdNext.Text = "Valider"
        Me.cmdNext.UseVisualStyleBackColor = True
        '
        'txtOri
        '
        Me.txtOri.AcceptsReturn = True
        Me.txtOri.BackColor = System.Drawing.SystemColors.Window
        Me.txtOri.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtOri.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtOri.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtOri.Location = New System.Drawing.Point(506, 24)
        Me.txtOri.MaxLength = 0
        Me.txtOri.Name = "txtOri"
        Me.txtOri.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtOri.Size = New System.Drawing.Size(136, 22)
        Me.txtOri.TabIndex = 4
        Me.txtOri.Text = "100.0"
        Me.txtOri.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        Me.txtOri.Visible = True
        '
        'Label75
        '
        Me.Label75.BackColor = System.Drawing.Color.Transparent
        Me.Label75.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label75.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label75.ForeColor = System.Drawing.Color.White
        Me.Label75.Location = New System.Drawing.Point(322, 28)
        Me.Label75.Name = "Label75"
        Me.Label75.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label75.Size = New System.Drawing.Size(176, 17)
        Me.Label75.TabIndex = 3
        Me.Label75.Text = "Rotation du symbole [g]:"
        Me.Label75.Visible = True
        '
        'txtBlockID
        '
        Me.txtBlockID.AcceptsReturn = True
        Me.txtBlockID.BackColor = System.Drawing.SystemColors.Window
        Me.txtBlockID.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtBlockID.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtBlockID.Location = New System.Drawing.Point(130, 24)
        Me.txtBlockID.MaxLength = 0
        Me.txtBlockID.Name = "txtBlockID"
        Me.txtBlockID.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtBlockID.Size = New System.Drawing.Size(152, 21)
        Me.txtBlockID.TabIndex = 1
        Me.txtBlockID.Text = "txtBlockID (pour édition)"
        Me.txtBlockID.Visible = False
        '
        'Label3
        '
        Me.Label3.BackColor = System.Drawing.Color.Transparent
        Me.Label3.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label3.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.Label3.Location = New System.Drawing.Point(18, 57)
        Me.Label3.Name = "Label3"
        Me.Label3.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label3.Size = New System.Drawing.Size(136, 16)
        Me.Label3.TabIndex = 19
        Me.Label3.Text = "Type de point / entité:"
        '
        'cboTheme
        '
        Me.cboTheme.BackColor = System.Drawing.Color.White
        Me.cboTheme.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboTheme.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTheme.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboTheme.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboTheme.Location = New System.Drawing.Point(354, 89)
        Me.cboTheme.Name = "cboTheme"
        Me.cboTheme.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboTheme.Size = New System.Drawing.Size(136, 24)
        Me.cboTheme.TabIndex = 1
        '
        'lblTheme
        '
        Me.lblTheme.BackColor = System.Drawing.Color.Transparent
        Me.lblTheme.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTheme.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTheme.ForeColor = System.Drawing.Color.White
        Me.lblTheme.Location = New System.Drawing.Point(275, 93)
        Me.lblTheme.Name = "lblTheme"
        Me.lblTheme.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTheme.Size = New System.Drawing.Size(56, 16)
        Me.lblTheme.TabIndex = 20
        Me.lblTheme.Text = "Thème:"
        '
        'cboType
        '
        Me.cboType.BackColor = System.Drawing.Color.White
        Me.cboType.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboType.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboType.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboType.Location = New System.Drawing.Point(594, 89)
        Me.cboType.Name = "cboType"
        Me.cboType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboType.Size = New System.Drawing.Size(136, 24)
        Me.cboType.TabIndex = 2
        '
        'lblType
        '
        Me.lblType.BackColor = System.Drawing.Color.Transparent
        Me.lblType.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblType.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblType.ForeColor = System.Drawing.Color.White
        Me.lblType.Location = New System.Drawing.Point(514, 93)
        Me.lblType.Name = "lblType"
        Me.lblType.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblType.Size = New System.Drawing.Size(56, 16)
        Me.lblType.TabIndex = 21
        Me.lblType.Text = "Type:"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Transparent
        Me.Label12.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label12.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.Label12.Location = New System.Drawing.Point(18, 322)
        Me.Label12.Name = "Label12"
        Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label12.Size = New System.Drawing.Size(248, 16)
        Me.Label12.TabIndex = 22
        Me.Label12.Text = "Coordonnées [m]:"
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Transparent
        Me.Label13.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label13.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.ForeColor = System.Drawing.Color.White
        Me.Label13.Location = New System.Drawing.Point(18, 349)
        Me.Label13.Name = "Label13"
        Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label13.Size = New System.Drawing.Size(72, 19)
        Me.Label13.TabIndex = 31
        Me.Label13.Text = "Est (X):"
        '
        'Label14
        '
        Me.Label14.BackColor = System.Drawing.Color.Transparent
        Me.Label14.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label14.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label14.ForeColor = System.Drawing.Color.White
        Me.Label14.Location = New System.Drawing.Point(251, 350)
        Me.Label14.Name = "Label14"
        Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label14.Size = New System.Drawing.Size(55, 16)
        Me.Label14.TabIndex = 32
        Me.Label14.Text = "Nord (Y):"
        '
        'Label15
        '
        Me.Label15.BackColor = System.Drawing.Color.Transparent
        Me.Label15.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label15.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label15.ForeColor = System.Drawing.Color.White
        Me.Label15.Location = New System.Drawing.Point(519, 350)
        Me.Label15.Name = "Label15"
        Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label15.Size = New System.Drawing.Size(72, 16)
        Me.Label15.TabIndex = 33
        Me.Label15.Text = "Altitude (Z):"
        '
        'txtX
        '
        Me.txtX.AcceptsReturn = True
        Me.txtX.BackColor = System.Drawing.Color.White
        Me.txtX.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtX.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtX.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtX.Location = New System.Drawing.Point(91, 346)
        Me.txtX.MaxLength = 0
        Me.txtX.Name = "txtX"
        Me.txtX.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtX.Size = New System.Drawing.Size(136, 22)
        Me.txtX.TabIndex = 10
        Me.txtX.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtY
        '
        Me.txtY.AcceptsReturn = True
        Me.txtY.BackColor = System.Drawing.Color.White
        Me.txtY.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtY.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtY.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtY.Location = New System.Drawing.Point(314, 346)
        Me.txtY.MaxLength = 0
        Me.txtY.Name = "txtY"
        Me.txtY.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtY.Size = New System.Drawing.Size(136, 22)
        Me.txtY.TabIndex = 11
        Me.txtY.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtZ
        '
        Me.txtZ.AcceptsReturn = True
        Me.txtZ.BackColor = System.Drawing.Color.White
        Me.txtZ.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtZ.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtZ.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtZ.Location = New System.Drawing.Point(594, 346)
        Me.txtZ.MaxLength = 0
        Me.txtZ.Name = "txtZ"
        Me.txtZ.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtZ.Size = New System.Drawing.Size(136, 22)
        Me.txtZ.TabIndex = 12
        Me.txtZ.Tag = "Non spécifiée = 0 m"
        Me.txtZ.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblAttrVariable
        '
        Me.lblAttrVariable.BackColor = System.Drawing.Color.Transparent
        Me.lblAttrVariable.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblAttrVariable.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAttrVariable.ForeColor = System.Drawing.Color.White
        Me.lblAttrVariable.Location = New System.Drawing.Point(514, 126)
        Me.lblAttrVariable.Name = "lblAttrVariable"
        Me.lblAttrVariable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblAttrVariable.Size = New System.Drawing.Size(72, 20)
        Me.lblAttrVariable.TabIndex = 30
        Me.lblAttrVariable.Text = "Variable:"
        Me.lblAttrVariable.Visible = False
        '
        'cboAttrVariable
        '
        Me.cboAttrVariable.BackColor = System.Drawing.Color.White
        Me.cboAttrVariable.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboAttrVariable.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboAttrVariable.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboAttrVariable.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboAttrVariable.Location = New System.Drawing.Point(594, 121)
        Me.cboAttrVariable.Name = "cboAttrVariable"
        Me.cboAttrVariable.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboAttrVariable.Size = New System.Drawing.Size(136, 24)
        Me.cboAttrVariable.TabIndex = 5
        Me.cboAttrVariable.Visible = False
        '
        'lblExact
        '
        Me.lblExact.BackColor = System.Drawing.Color.Transparent
        Me.lblExact.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblExact.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblExact.ForeColor = System.Drawing.Color.White
        Me.lblExact.Location = New System.Drawing.Point(275, 126)
        Me.lblExact.Name = "lblExact"
        Me.lblExact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblExact.Size = New System.Drawing.Size(80, 16)
        Me.lblExact.TabIndex = 29
        Me.lblExact.Text = "Défini exact.:"
        '
        'cboExact
        '
        Me.cboExact.BackColor = System.Drawing.Color.White
        Me.cboExact.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboExact.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboExact.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboExact.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboExact.Location = New System.Drawing.Point(354, 122)
        Me.cboExact.Name = "cboExact"
        Me.cboExact.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboExact.Size = New System.Drawing.Size(136, 24)
        Me.cboExact.TabIndex = 4
        '
        'Label87
        '
        Me.Label87.BackColor = System.Drawing.Color.Transparent
        Me.Label87.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label87.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label87.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.Label87.Location = New System.Drawing.Point(18, 162)
        Me.Label87.Name = "Label87"
        Me.Label87.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label87.Size = New System.Drawing.Size(112, 16)
        Me.Label87.TabIndex = 23
        Me.Label87.Text = "Identification:"
        '
        'txtIdentDN
        '
        Me.txtIdentDN.AcceptsReturn = True
        Me.txtIdentDN.BackColor = System.Drawing.Color.White
        Me.txtIdentDN.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtIdentDN.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtIdentDN.Location = New System.Drawing.Point(91, 250)
        Me.txtIdentDN.MaxLength = 12
        Me.txtIdentDN.Name = "txtIdentDN"
        Me.txtIdentDN.ReadOnly = True
        Me.txtIdentDN.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtIdentDN.Size = New System.Drawing.Size(136, 22)
        Me.txtIdentDN.TabIndex = 8
        Me.txtIdentDN.TabStop = False
        Me.txtIdentDN.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label85
        '
        Me.Label85.BackColor = System.Drawing.Color.Transparent
        Me.Label85.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label85.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label85.ForeColor = System.Drawing.Color.White
        Me.Label85.Location = New System.Drawing.Point(18, 254)
        Me.Label85.Name = "Label85"
        Me.Label85.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label85.Size = New System.Drawing.Size(56, 16)
        Me.Label85.TabIndex = 24
        Me.Label85.Text = "IdentDN:"
        '
        'txtPlan
        '
        Me.txtPlan.AcceptsReturn = True
        Me.txtPlan.BackColor = System.Drawing.Color.White
        Me.txtPlan.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPlan.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPlan.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPlan.Location = New System.Drawing.Point(91, 218)
        Me.txtPlan.MaxLength = 4
        Me.txtPlan.Name = "txtPlan"
        Me.txtPlan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPlan.Size = New System.Drawing.Size(136, 22)
        Me.txtPlan.TabIndex = 7
        Me.txtPlan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblPlan
        '
        Me.lblPlan.BackColor = System.Drawing.Color.Transparent
        Me.lblPlan.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPlan.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlan.ForeColor = System.Drawing.Color.White
        Me.lblPlan.Location = New System.Drawing.Point(18, 222)
        Me.lblPlan.Name = "lblPlan"
        Me.lblPlan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPlan.Size = New System.Drawing.Size(56, 16)
        Me.lblPlan.TabIndex = 25
        Me.lblPlan.Text = "Plan:"
        '
        'txtCom
        '
        Me.txtCom.AcceptsReturn = True
        Me.txtCom.BackColor = System.Drawing.Color.White
        Me.txtCom.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtCom.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCom.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtCom.Location = New System.Drawing.Point(91, 186)
        Me.txtCom.MaxLength = 3
        Me.txtCom.Name = "txtCom"
        Me.txtCom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtCom.Size = New System.Drawing.Size(136, 22)
        Me.txtCom.TabIndex = 6
        Me.txtCom.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'Label74
        '
        Me.Label74.BackColor = System.Drawing.Color.Transparent
        Me.Label74.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label74.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label74.ForeColor = System.Drawing.Color.White
        Me.Label74.Location = New System.Drawing.Point(18, 285)
        Me.Label74.Name = "Label74"
        Me.Label74.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label74.Size = New System.Drawing.Size(64, 16)
        Me.Label74.TabIndex = 26
        Me.Label74.Text = "Numero:"
        '
        'txtNo
        '
        Me.txtNo.AcceptsReturn = True
        Me.txtNo.BackColor = System.Drawing.Color.White
        Me.txtNo.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtNo.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtNo.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtNo.Location = New System.Drawing.Point(91, 281)
        Me.txtNo.MaxLength = 12
        Me.txtNo.Name = "txtNo"
        Me.txtNo.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtNo.Size = New System.Drawing.Size(136, 22)
        Me.txtNo.TabIndex = 9
        Me.txtNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lblCom
        '
        Me.lblCom.BackColor = System.Drawing.Color.Transparent
        Me.lblCom.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblCom.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCom.ForeColor = System.Drawing.Color.White
        Me.lblCom.Location = New System.Drawing.Point(18, 190)
        Me.lblCom.Name = "lblCom"
        Me.lblCom.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblCom.Size = New System.Drawing.Size(80, 16)
        Me.lblCom.TabIndex = 27
        Me.lblCom.Text = "Commune:"
        '
        'cboSigne
        '
        Me.cboSigne.BackColor = System.Drawing.Color.White
        Me.cboSigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboSigne.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboSigne.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboSigne.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboSigne.Location = New System.Drawing.Point(91, 122)
        Me.cboSigne.Name = "cboSigne"
        Me.cboSigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboSigne.Size = New System.Drawing.Size(175, 24)
        Me.cboSigne.TabIndex = 3
        '
        'lblSigne
        '
        Me.lblSigne.BackColor = System.Drawing.Color.Transparent
        Me.lblSigne.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSigne.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSigne.ForeColor = System.Drawing.Color.White
        Me.lblSigne.Location = New System.Drawing.Point(18, 126)
        Me.lblSigne.Name = "lblSigne"
        Me.lblSigne.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSigne.Size = New System.Drawing.Size(56, 16)
        Me.lblSigne.TabIndex = 28
        Me.lblSigne.Text = "Signe:"
        '
        'Label88
        '
        Me.Label88.BackColor = System.Drawing.Color.Transparent
        Me.Label88.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label88.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label88.ForeColor = System.Drawing.SystemColors.ScrollBar
        Me.Label88.Location = New System.Drawing.Point(18, 418)
        Me.Label88.Name = "Label88"
        Me.Label88.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label88.Size = New System.Drawing.Size(120, 16)
        Me.Label88.TabIndex = 36
        Me.Label88.Text = "Attributs qualitatifs:"
        '
        'Label78
        '
        Me.Label78.BackColor = System.Drawing.Color.Transparent
        Me.Label78.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label78.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label78.ForeColor = System.Drawing.Color.White
        Me.Label78.Location = New System.Drawing.Point(274, 446)
        Me.Label78.Name = "Label78"
        Me.Label78.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label78.Size = New System.Drawing.Size(64, 16)
        Me.Label78.TabIndex = 37
        Me.Label78.Text = "PrecPlan:"
        '
        'Label77
        '
        Me.Label77.BackColor = System.Drawing.Color.Transparent
        Me.Label77.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label77.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label77.ForeColor = System.Drawing.Color.White
        Me.Label77.Location = New System.Drawing.Point(18, 445)
        Me.Label77.Name = "Label77"
        Me.Label77.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label77.Size = New System.Drawing.Size(64, 16)
        Me.Label77.TabIndex = 38
        Me.Label77.Text = "FiabPlan:"
        '
        'lblPrecAlt
        '
        Me.lblPrecAlt.BackColor = System.Drawing.Color.Transparent
        Me.lblPrecAlt.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblPrecAlt.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrecAlt.ForeColor = System.Drawing.Color.White
        Me.lblPrecAlt.Location = New System.Drawing.Point(274, 478)
        Me.lblPrecAlt.Name = "lblPrecAlt"
        Me.lblPrecAlt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblPrecAlt.Size = New System.Drawing.Size(64, 16)
        Me.lblPrecAlt.TabIndex = 39
        Me.lblPrecAlt.Text = "PrecAlt:"
        '
        'lblFiabAlt
        '
        Me.lblFiabAlt.BackColor = System.Drawing.Color.Transparent
        Me.lblFiabAlt.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFiabAlt.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFiabAlt.ForeColor = System.Drawing.Color.White
        Me.lblFiabAlt.Location = New System.Drawing.Point(18, 477)
        Me.lblFiabAlt.Name = "lblFiabAlt"
        Me.lblFiabAlt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblFiabAlt.Size = New System.Drawing.Size(64, 16)
        Me.lblFiabAlt.TabIndex = 40
        Me.lblFiabAlt.Text = "FiabAlt:"
        '
        'txtPrecPlan
        '
        Me.txtPrecPlan.AcceptsReturn = True
        Me.txtPrecPlan.BackColor = System.Drawing.Color.White
        Me.txtPrecPlan.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrecPlan.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrecPlan.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrecPlan.Location = New System.Drawing.Point(354, 442)
        Me.txtPrecPlan.MaxLength = 0
        Me.txtPrecPlan.Name = "txtPrecPlan"
        Me.txtPrecPlan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrecPlan.Size = New System.Drawing.Size(136, 22)
        Me.txtPrecPlan.TabIndex = 14
        Me.txtPrecPlan.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'txtPrecAlt
        '
        Me.txtPrecAlt.AcceptsReturn = True
        Me.txtPrecAlt.BackColor = System.Drawing.Color.White
        Me.txtPrecAlt.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtPrecAlt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPrecAlt.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtPrecAlt.Location = New System.Drawing.Point(354, 474)
        Me.txtPrecAlt.MaxLength = 0
        Me.txtPrecAlt.Name = "txtPrecAlt"
        Me.txtPrecAlt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtPrecAlt.Size = New System.Drawing.Size(136, 22)
        Me.txtPrecAlt.TabIndex = 16
        Me.txtPrecAlt.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'cboFiabPlan
        '
        Me.cboFiabPlan.BackColor = System.Drawing.Color.White
        Me.cboFiabPlan.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboFiabPlan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFiabPlan.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFiabPlan.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboFiabPlan.Location = New System.Drawing.Point(91, 442)
        Me.cboFiabPlan.Name = "cboFiabPlan"
        Me.cboFiabPlan.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboFiabPlan.Size = New System.Drawing.Size(136, 24)
        Me.cboFiabPlan.TabIndex = 13
        '
        'cboFiabAlt
        '
        Me.cboFiabAlt.BackColor = System.Drawing.Color.White
        Me.cboFiabAlt.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboFiabAlt.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboFiabAlt.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboFiabAlt.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboFiabAlt.Location = New System.Drawing.Point(91, 474)
        Me.cboFiabAlt.Name = "cboFiabAlt"
        Me.cboFiabAlt.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboFiabAlt.Size = New System.Drawing.Size(136, 24)
        Me.cboFiabAlt.TabIndex = 15
        '
        'Label89
        '
        Me.Label89.BackColor = System.Drawing.Color.Transparent
        Me.Label89.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label89.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label89.ForeColor = System.Drawing.Color.White
        Me.Label89.Location = New System.Drawing.Point(250, 162)
        Me.Label89.Name = "Label89"
        Me.Label89.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label89.Size = New System.Drawing.Size(280, 16)
        Me.Label89.TabIndex = 41
        Me.Label89.Text = "Entités administratives incluses dans le dessin:"
        '
        'chkDigitSerie
        '
        Me.chkDigitSerie.BackColor = System.Drawing.Color.Transparent
        Me.chkDigitSerie.Cursor = System.Windows.Forms.Cursors.Default
        Me.chkDigitSerie.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chkDigitSerie.ForeColor = System.Drawing.Color.White
        Me.chkDigitSerie.Location = New System.Drawing.Point(313, 376)
        Me.chkDigitSerie.Name = "chkDigitSerie"
        Me.chkDigitSerie.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.chkDigitSerie.Size = New System.Drawing.Size(149, 25)
        Me.chkDigitSerie.TabIndex = 43
        Me.chkDigitSerie.Text = "Digitalisation en série"
        Me.chkDigitSerie.UseVisualStyleBackColor = False
        '
        'Label90
        '
        Me.Label90.BackColor = System.Drawing.Color.Transparent
        Me.Label90.Cursor = System.Windows.Forms.Cursors.Default
        Me.Label90.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label90.ForeColor = System.Drawing.Color.White
        Me.Label90.Location = New System.Drawing.Point(18, 93)
        Me.Label90.Name = "Label90"
        Me.Label90.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label90.Size = New System.Drawing.Size(72, 16)
        Me.Label90.TabIndex = 44
        Me.Label90.Text = "Domaine:"
        '
        'cboCat
        '
        Me.cboCat.BackColor = System.Drawing.Color.White
        Me.cboCat.Cursor = System.Windows.Forms.Cursors.Default
        Me.cboCat.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboCat.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cboCat.ForeColor = System.Drawing.SystemColors.WindowText
        Me.cboCat.Location = New System.Drawing.Point(91, 89)
        Me.cboCat.Name = "cboCat"
        Me.cboCat.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.cboCat.Size = New System.Drawing.Size(175, 24)
        Me.cboCat.TabIndex = 0
        '
        'txtComment
        '
        Me.txtComment.AcceptsReturn = True
        Me.txtComment.BackColor = System.Drawing.Color.White
        Me.txtComment.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.txtComment.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtComment.ForeColor = System.Drawing.SystemColors.WindowText
        Me.txtComment.Location = New System.Drawing.Point(354, 121)
        Me.txtComment.MaxLength = 50
        Me.txtComment.Name = "txtComment"
        Me.txtComment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.txtComment.Size = New System.Drawing.Size(376, 22)
        Me.txtComment.TabIndex = 45
        '
        'lblComment
        '
        Me.lblComment.BackColor = System.Drawing.Color.Transparent
        Me.lblComment.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblComment.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblComment.ForeColor = System.Drawing.Color.White
        Me.lblComment.Location = New System.Drawing.Point(274, 126)
        Me.lblComment.Name = "lblComment"
        Me.lblComment.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblComment.Size = New System.Drawing.Size(88, 16)
        Me.lblComment.TabIndex = 46
        Me.lblComment.Text = "Commentaire:"
        '
        'cmdDigit
        '
        Me.cmdDigit.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app_active
        Me.cmdDigit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdDigit.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.cmdDigit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdDigit.Image = Global.Revo.My.Resources.Resources._select
        Me.cmdDigit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.cmdDigit.Location = New System.Drawing.Point(91, 376)
        Me.cmdDigit.Name = "cmdDigit"
        Me.cmdDigit.Size = New System.Drawing.Size(136, 24)
        Me.cmdDigit.TabIndex = 50
        Me.cmdDigit.Text = "Digitaliser à l'écran"
        Me.cmdDigit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.cmdDigit.UseVisualStyleBackColor = True
        '
        'cmdListeNat
        '
        Me.cmdListeNat.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.cmdListeNat.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdListeNat.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.cmdListeNat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdListeNat.Location = New System.Drawing.Point(21, 21)
        Me.cmdListeNat.Name = "cmdListeNat"
        Me.cmdListeNat.Size = New System.Drawing.Size(229, 24)
        Me.cmdListeNat.TabIndex = 51
        Me.cmdListeNat.Tag = ""
        Me.cmdListeNat.Text = "Choix selon liste des codes natures"
        Me.cmdListeNat.UseVisualStyleBackColor = True
        '
        'lblSubtitle
        '
        Me.lblSubtitle.AutoSize = True
        Me.lblSubtitle.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubtitle.ForeColor = System.Drawing.Color.White
        Me.lblSubtitle.Location = New System.Drawing.Point(274, 26)
        Me.lblSubtitle.Name = "lblSubtitle"
        Me.lblSubtitle.Size = New System.Drawing.Size(279, 14)
        Me.lblSubtitle.TabIndex = 52
        Me.lblSubtitle.Text = "Définir le type, la position et les attributs du point"
        '
        'LVPlans
        '
        Me.LVPlans.Activation = System.Windows.Forms.ItemActivation.OneClick
        Me.LVPlans.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.LVPlans.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.Commune, Me.Plan, Me.NT, Me.Mensuration})
        Me.LVPlans.Cursor = System.Windows.Forms.Cursors.Hand
        Me.LVPlans.FullRowSelect = True
        Me.LVPlans.GridLines = True
        Me.LVPlans.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.LVPlans.HideSelection = False
        Me.LVPlans.ImeMode = System.Windows.Forms.ImeMode.AlphaFull
        Me.LVPlans.Location = New System.Drawing.Point(250, 186)
        Me.LVPlans.MultiSelect = False
        Me.LVPlans.Name = "LVPlans"
        Me.LVPlans.Size = New System.Drawing.Size(481, 138)
        Me.LVPlans.TabIndex = 54
        Me.LVPlans.UseCompatibleStateImageBehavior = False
        Me.LVPlans.View = System.Windows.Forms.View.Details
        '
        'Commune
        '
        Me.Commune.Text = "Commune"
        Me.Commune.Width = 127
        '
        'Plan
        '
        Me.Plan.Text = "Plan"
        Me.Plan.Width = 92
        '
        'NT
        '
        Me.NT.Text = "NT"
        Me.NT.Width = 116
        '
        'Mensuration
        '
        Me.Mensuration.Text = "Mensuration"
        Me.Mensuration.Width = 144
        '
        'Alt3D
        '
        Me.Alt3D.AutoSize = True
        Me.Alt3D.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Alt3D.ForeColor = System.Drawing.Color.White
        Me.Alt3D.Location = New System.Drawing.Point(667, 374)
        Me.Alt3D.Name = "Alt3D"
        Me.Alt3D.Size = New System.Drawing.Size(67, 17)
        Me.Alt3D.TabIndex = 55
        Me.Alt3D.Tag = "L'objet enregistre l'altitude uniquement en Z et comme information"
        Me.Alt3D.Text = "Objet 3D"
        Me.Alt3D.UseVisualStyleBackColor = True
        '
        'Alt2D
        '
        Me.Alt2D.AutoSize = True
        Me.Alt2D.Checked = True
        Me.Alt2D.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.Alt2D.ForeColor = System.Drawing.Color.White
        Me.Alt2D.Location = New System.Drawing.Point(594, 374)
        Me.Alt2D.Name = "Alt2D"
        Me.Alt2D.Size = New System.Drawing.Size(67, 17)
        Me.Alt2D.TabIndex = 56
        Me.Alt2D.TabStop = True
        Me.Alt2D.Tag = "L'objet enregistre l'altitude uniquement en information"
        Me.Alt2D.Text = "Objet 2D"
        Me.Alt2D.UseVisualStyleBackColor = True
        '
        'frmInsertionPts
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(745, 612)
        Me.Controls.Add(Me.Alt2D)
        Me.Controls.Add(Me.Alt3D)
        Me.Controls.Add(Me.LVPlans)
        Me.Controls.Add(Me.lblSubtitle)
        Me.Controls.Add(Me.cmdListeNat)
        Me.Controls.Add(Me.cmdDigit)
        Me.Controls.Add(Me.Frame2)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.cboTheme)
        Me.Controls.Add(Me.lblTheme)
        Me.Controls.Add(Me.cboType)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.txtX)
        Me.Controls.Add(Me.txtY)
        Me.Controls.Add(Me.txtZ)
        Me.Controls.Add(Me.lblAttrVariable)
        Me.Controls.Add(Me.cboAttrVariable)
        Me.Controls.Add(Me.lblExact)
        Me.Controls.Add(Me.cboExact)
        Me.Controls.Add(Me.Label87)
        Me.Controls.Add(Me.txtIdentDN)
        Me.Controls.Add(Me.Label85)
        Me.Controls.Add(Me.txtPlan)
        Me.Controls.Add(Me.lblPlan)
        Me.Controls.Add(Me.txtCom)
        Me.Controls.Add(Me.Label74)
        Me.Controls.Add(Me.txtNo)
        Me.Controls.Add(Me.lblCom)
        Me.Controls.Add(Me.cboSigne)
        Me.Controls.Add(Me.lblSigne)
        Me.Controls.Add(Me.Label88)
        Me.Controls.Add(Me.Label78)
        Me.Controls.Add(Me.Label77)
        Me.Controls.Add(Me.lblPrecAlt)
        Me.Controls.Add(Me.lblFiabAlt)
        Me.Controls.Add(Me.txtPrecPlan)
        Me.Controls.Add(Me.txtPrecAlt)
        Me.Controls.Add(Me.cboFiabPlan)
        Me.Controls.Add(Me.cboFiabAlt)
        Me.Controls.Add(Me.Label89)
        Me.Controls.Add(Me.chkDigitSerie)
        Me.Controls.Add(Me.Label90)
        Me.Controls.Add(Me.cboCat)
        Me.Controls.Add(Me.txtComment)
        Me.Controls.Add(Me.lblComment)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 15)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmInsertionPts"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Insertion de point"
        Me.Frame2.ResumeLayout(False)
        Me.Frame2.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents cmdNext As System.Windows.Forms.Button
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdDigit As System.Windows.Forms.Button
    Friend WithEvents cmdListeNat As System.Windows.Forms.Button
    Friend WithEvents lblSubtitle As System.Windows.Forms.Label
    Friend WithEvents LVPlans As System.Windows.Forms.ListView
    Friend WithEvents Commune As System.Windows.Forms.ColumnHeader
    Friend WithEvents Plan As System.Windows.Forms.ColumnHeader
    Friend WithEvents NT As System.Windows.Forms.ColumnHeader
    Friend WithEvents Mensuration As System.Windows.Forms.ColumnHeader
    Friend WithEvents Alt3D As System.Windows.Forms.RadioButton
    Friend WithEvents Alt2D As System.Windows.Forms.RadioButton
#End Region
End Class