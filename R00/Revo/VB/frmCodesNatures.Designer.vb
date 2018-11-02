<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> Partial Class frmCodesNatures
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
    Public WithEvents imgList As System.Windows.Forms.Label
    'REMARQUE : la procédure suivante est requise par le Concepteur Windows Form
    'Elle peut être modifiée à l'aide du Concepteur Windows Form.
    'Ne la modifiez pas à l'aide de l'éditeur de code.
	<System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nœud1")
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nœud2")
        Dim TreeNode3 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nœud0", New System.Windows.Forms.TreeNode() {TreeNode1, TreeNode2})
        Dim TreeNode4 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nœud4")
        Dim TreeNode5 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nœud5")
        Dim TreeNode6 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Nœud3", New System.Windows.Forms.TreeNode() {TreeNode4, TreeNode5})
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmCodesNatures))
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.imgList = New System.Windows.Forms.Label()
        Me.lblSubtitle = New System.Windows.Forms.Label()
        Me.tvNatures = New System.Windows.Forms.TreeView()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.cmdCancel = New System.Windows.Forms.Button()
        Me.cmdNext = New System.Windows.Forms.Button()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.GroupBox1.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'imgList
        '
        Me.imgList.BackColor = System.Drawing.Color.Red
        Me.imgList.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.imgList.Cursor = System.Windows.Forms.Cursors.Default
        Me.imgList.ForeColor = System.Drawing.SystemColors.ControlText
        Me.imgList.Location = New System.Drawing.Point(250, 11)
        Me.imgList.Name = "imgList"
        Me.imgList.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.imgList.Size = New System.Drawing.Size(38, 38)
        Me.imgList.TabIndex = 2
        Me.imgList.Text = "imgList - ImageList"
        Me.imgList.TextAlign = System.Drawing.ContentAlignment.TopCenter
        Me.imgList.Visible = False
        '
        'lblSubtitle
        '
        Me.lblSubtitle.BackColor = System.Drawing.Color.Transparent
        Me.lblSubtitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblSubtitle.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSubtitle.ForeColor = System.Drawing.Color.White
        Me.lblSubtitle.Location = New System.Drawing.Point(3, 0)
        Me.lblSubtitle.Name = "lblSubtitle"
        Me.lblSubtitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblSubtitle.Size = New System.Drawing.Size(384, 16)
        Me.lblSubtitle.TabIndex = 4
        Me.lblSubtitle.Text = "Sélectionner le code nature à appliquer au(x) point(s) concerné(s)"
        '
        'tvNatures
        '
        Me.tvNatures.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tvNatures.Location = New System.Drawing.Point(3, 33)
        Me.tvNatures.Name = "tvNatures"
        TreeNode1.Name = "Nœud1"
        TreeNode1.Text = "Nœud1"
        TreeNode2.Name = "Nœud2"
        TreeNode2.Text = "Nœud2"
        TreeNode3.Name = "Nœud0"
        TreeNode3.Text = "Nœud0"
        TreeNode4.Name = "Nœud4"
        TreeNode4.Text = "Nœud4"
        TreeNode5.Name = "Nœud5"
        TreeNode5.Text = "Nœud5"
        TreeNode6.Name = "Nœud3"
        TreeNode6.Text = "Nœud3"
        Me.tvNatures.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode3, TreeNode6})
        Me.tvNatures.Size = New System.Drawing.Size(548, 388)
        Me.tvNatures.TabIndex = 5
        '
        'GroupBox1
        '
        Me.GroupBox1.BackColor = System.Drawing.Color.Transparent
        Me.GroupBox1.Controls.Add(Me.cmdCancel)
        Me.GroupBox1.Controls.Add(Me.imgList)
        Me.GroupBox1.Controls.Add(Me.cmdNext)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.ForeColor = System.Drawing.SystemColors.ControlText
        Me.GroupBox1.Location = New System.Drawing.Point(3, 427)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Padding = New System.Windows.Forms.Padding(0)
        Me.GroupBox1.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.GroupBox1.Size = New System.Drawing.Size(548, 49)
        Me.GroupBox1.TabIndex = 19
        Me.GroupBox1.TabStop = False
        '
        'cmdCancel
        '
        Me.cmdCancel.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.cmdCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.cmdCancel.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.cmdCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.cmdCancel.Location = New System.Drawing.Point(15, 17)
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
        Me.cmdNext.Location = New System.Drawing.Point(471, 17)
        Me.cmdNext.Name = "cmdNext"
        Me.cmdNext.Size = New System.Drawing.Size(66, 24)
        Me.cmdNext.TabIndex = 49
        Me.cmdNext.Tag = ""
        Me.cmdNext.Text = "Valider"
        Me.cmdNext.UseVisualStyleBackColor = True
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.tvNatures, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox1, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.lblSubtitle, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 55.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(554, 479)
        Me.TableLayoutPanel1.TabIndex = 20
        '
        'frmCodesNatures
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(554, 479)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Cursor = System.Windows.Forms.Cursors.Default
        Me.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Location = New System.Drawing.Point(3, 22)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmCodesNatures"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Liste des codes natures par domaine"
        Me.GroupBox1.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Public WithEvents lblSubtitle As System.Windows.Forms.Label
    Friend WithEvents tvNatures As System.Windows.Forms.TreeView
    Public WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents cmdCancel As System.Windows.Forms.Button
    Friend WithEvents cmdNext As System.Windows.Forms.Button
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
#End Region
End Class