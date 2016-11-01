<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmOptions
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
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle3 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmOptions))
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.GridParam = New System.Windows.Forms.DataGridView()
        Me.NameParam = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DossierCmd = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.VariableName = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.VariableType = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DefaultVar = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.PicLogo = New System.Windows.Forms.PictureBox()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnExportXML = New System.Windows.Forms.Button()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnEscape = New System.Windows.Forms.Button()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.btnNetwork = New System.Windows.Forms.Button()
        Me.btnImportXML = New System.Windows.Forms.Button()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ModifPath = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChargerParamDefautToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ChargerLeParamToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.btnLock = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        CType(Me.GridParam, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.GroupBox2.SuspendLayout()
        CType(Me.PicLogo, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.TableLayoutPanel1.SetColumnSpan(Me.GroupBox1, 2)
        Me.GroupBox1.Controls.Add(Me.GridParam)
        Me.GroupBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.ForeColor = System.Drawing.Color.White
        Me.GroupBox1.Location = New System.Drawing.Point(3, 68)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(624, 178)
        Me.GroupBox1.TabIndex = 0
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Environnement"
        '
        'GridParam
        '
        Me.GridParam.AllowUserToAddRows = False
        Me.GridParam.AllowUserToDeleteRows = False
        Me.GridParam.AllowUserToResizeRows = False
        Me.GridParam.BackgroundColor = System.Drawing.Color.Gray
        Me.GridParam.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.GridParam.CellBorderStyle = System.Windows.Forms.DataGridViewCellBorderStyle.None
        Me.GridParam.ColumnHeadersBorderStyle = System.Windows.Forms.DataGridViewHeaderBorderStyle.None
        Me.GridParam.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.GridParam.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.NameParam, Me.DossierCmd, Me.VariableName, Me.VariableType, Me.DefaultVar})
        Me.GridParam.Cursor = System.Windows.Forms.Cursors.Default
        DataGridViewCellStyle3.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle3.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle3.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle3.ForeColor = System.Drawing.Color.White
        DataGridViewCellStyle3.SelectionBackColor = System.Drawing.Color.FromArgb(CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer), CType(CType(224, Byte), Integer))
        DataGridViewCellStyle3.SelectionForeColor = System.Drawing.Color.Black
        DataGridViewCellStyle3.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.GridParam.DefaultCellStyle = DataGridViewCellStyle3
        Me.GridParam.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GridParam.GridColor = System.Drawing.Color.LightGray
        Me.GridParam.Location = New System.Drawing.Point(3, 17)
        Me.GridParam.MultiSelect = False
        Me.GridParam.Name = "GridParam"
        Me.GridParam.ReadOnly = True
        Me.GridParam.RowHeadersWidth = 4
        Me.GridParam.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.GridParam.RowTemplate.DefaultCellStyle.ForeColor = System.Drawing.Color.Black
        Me.GridParam.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.GridParam.Size = New System.Drawing.Size(618, 158)
        Me.GridParam.TabIndex = 2
        '
        'NameParam
        '
        DataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.NameParam.DefaultCellStyle = DataGridViewCellStyle1
        Me.NameParam.HeaderText = "Paramètre"
        Me.NameParam.Name = "NameParam"
        Me.NameParam.ReadOnly = True
        Me.NameParam.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'DossierCmd
        '
        Me.DossierCmd.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        Me.DossierCmd.DefaultCellStyle = DataGridViewCellStyle2
        Me.DossierCmd.HeaderText = "Dossier"
        Me.DossierCmd.Name = "DossierCmd"
        Me.DossierCmd.ReadOnly = True
        Me.DossierCmd.Resizable = System.Windows.Forms.DataGridViewTriState.[True]
        '
        'VariableName
        '
        Me.VariableName.HeaderText = "VariableName"
        Me.VariableName.Name = "VariableName"
        Me.VariableName.ReadOnly = True
        Me.VariableName.Visible = False
        '
        'VariableType
        '
        Me.VariableType.HeaderText = "VariableType"
        Me.VariableType.Name = "VariableType"
        Me.VariableType.ReadOnly = True
        Me.VariableType.Visible = False
        '
        'DefaultVar
        '
        Me.DefaultVar.HeaderText = "DefaultVar"
        Me.DefaultVar.Name = "DefaultVar"
        Me.DefaultVar.ReadOnly = True
        Me.DefaultVar.Visible = False
        '
        'GroupBox2
        '
        Me.GroupBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.GroupBox2.BackColor = System.Drawing.Color.White
        Me.GroupBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.TableLayoutPanel1.SetColumnSpan(Me.GroupBox2, 2)
        Me.GroupBox2.Controls.Add(Me.PicLogo)
        Me.GroupBox2.Controls.Add(Me.lblInfo)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(624, 59)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        '
        'PicLogo
        '
        Me.PicLogo.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PicLogo.Image = Global.Revo.My.Resources.Resources.ribbon_plug32
        Me.PicLogo.InitialImage = Nothing
        Me.PicLogo.Location = New System.Drawing.Point(579, 16)
        Me.PicLogo.Name = "PicLogo"
        Me.PicLogo.Size = New System.Drawing.Size(32, 32)
        Me.PicLogo.TabIndex = 5
        Me.PicLogo.TabStop = False
        '
        'lblInfo
        '
        Me.lblInfo.AutoSize = True
        Me.lblInfo.BackColor = System.Drawing.Color.Transparent
        Me.lblInfo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblInfo.Location = New System.Drawing.Point(17, 20)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(319, 26)
        Me.lblInfo.TabIndex = 4
        Me.lblInfo.Text = "Configuration de l’environnement et des paramètres personnalisés." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "Pour plus d'in" & _
    "formation, merci de consulter l'aide en ligne."
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.BackColor = System.Drawing.Color.Transparent
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox2, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.GroupBox1, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel2, 1, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel3, 0, 2)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 3
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 65.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 36.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(630, 285)
        Me.TableLayoutPanel1.TabIndex = 6
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 4
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 70.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 100.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 5.0!))
        Me.TableLayoutPanel2.Controls.Add(Me.btnExportXML, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnOK, 2, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.btnEscape, 1, 0)
        Me.TableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(318, 252)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(309, 30)
        Me.TableLayoutPanel2.TabIndex = 6
        '
        'btnExportXML
        '
        Me.btnExportXML.BackColor = System.Drawing.Color.Transparent
        Me.btnExportXML.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnExportXML.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnExportXML.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnExportXML.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnExportXML.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExportXML.ForeColor = System.Drawing.Color.Black
        Me.btnExportXML.Location = New System.Drawing.Point(3, 3)
        Me.btnExportXML.Name = "btnExportXML"
        Me.btnExportXML.Size = New System.Drawing.Size(100, 24)
        Me.btnExportXML.TabIndex = 8
        Me.btnExportXML.Text = "Exporter config."
        Me.btnExportXML.UseVisualStyleBackColor = False
        '
        'btnOK
        '
        Me.btnOK.BackColor = System.Drawing.Color.Transparent
        Me.btnOK.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnOK.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnOK.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnOK.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnOK.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnOK.ForeColor = System.Drawing.Color.Black
        Me.btnOK.Location = New System.Drawing.Point(221, 3)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 24)
        Me.btnOK.TabIndex = 9
        Me.btnOK.Text = "Appliquer"
        Me.btnOK.UseVisualStyleBackColor = False
        '
        'btnEscape
        '
        Me.btnEscape.BackColor = System.Drawing.Color.Transparent
        Me.btnEscape.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnEscape.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEscape.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnEscape.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnEscape.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnEscape.ForeColor = System.Drawing.Color.Black
        Me.btnEscape.Location = New System.Drawing.Point(137, 3)
        Me.btnEscape.Name = "btnEscape"
        Me.btnEscape.Size = New System.Drawing.Size(64, 24)
        Me.btnEscape.TabIndex = 10
        Me.btnEscape.Text = "Annuler"
        Me.btnEscape.UseVisualStyleBackColor = False
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 3
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 154.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 125.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.btnNetwork, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.btnImportXML, 2, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.btnLock, 1, 0)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 252)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(309, 30)
        Me.TableLayoutPanel3.TabIndex = 7
        '
        'btnNetwork
        '
        Me.btnNetwork.BackColor = System.Drawing.Color.Transparent
        Me.btnNetwork.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app_active
        Me.btnNetwork.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnNetwork.Dock = System.Windows.Forms.DockStyle.Left
        Me.btnNetwork.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnNetwork.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNetwork.ForeColor = System.Drawing.Color.Black
        Me.btnNetwork.Location = New System.Drawing.Point(3, 3)
        Me.btnNetwork.Name = "btnNetwork"
        Me.btnNetwork.Size = New System.Drawing.Size(147, 24)
        Me.btnNetwork.TabIndex = 4
        Me.btnNetwork.Text = "Créer un espace partagé"
        Me.btnNetwork.UseVisualStyleBackColor = False
        '
        'btnImportXML
        '
        Me.btnImportXML.BackColor = System.Drawing.Color.Transparent
        Me.btnImportXML.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnImportXML.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnImportXML.Dock = System.Windows.Forms.DockStyle.Right
        Me.btnImportXML.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnImportXML.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnImportXML.ForeColor = System.Drawing.Color.Black
        Me.btnImportXML.Location = New System.Drawing.Point(206, 3)
        Me.btnImportXML.Name = "btnImportXML"
        Me.btnImportXML.Size = New System.Drawing.Size(100, 24)
        Me.btnImportXML.TabIndex = 8
        Me.btnImportXML.Text = "Importer config."
        Me.btnImportXML.UseVisualStyleBackColor = False
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ModifPath, Me.ChargerParamDefautToolStripMenuItem, Me.ChargerLeParamToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(232, 70)
        '
        'ModifPath
        '
        Me.ModifPath.Name = "ModifPath"
        Me.ModifPath.Size = New System.Drawing.Size(231, 22)
        Me.ModifPath.Text = "Modifier le chemin d'accès"
        '
        'ChargerParamDefautToolStripMenuItem
        '
        Me.ChargerParamDefautToolStripMenuItem.Name = "ChargerParamDefautToolStripMenuItem"
        Me.ChargerParamDefautToolStripMenuItem.Size = New System.Drawing.Size(231, 22)
        Me.ChargerParamDefautToolStripMenuItem.Text = "Charger paramètre par défaut"
        '
        'ChargerLeParamToolStripMenuItem
        '
        Me.ChargerLeParamToolStripMenuItem.Name = "ChargerLeParamToolStripMenuItem"
        Me.ChargerLeParamToolStripMenuItem.Size = New System.Drawing.Size(231, 22)
        Me.ChargerLeParamToolStripMenuItem.Text = "Charger paramètre Autocad"
        Me.ChargerLeParamToolStripMenuItem.Visible = False
        '
        'btnLock
        '
        Me.btnLock.BackgroundImage = Global.Revo.My.Resources.Resources.lock
        Me.btnLock.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnLock.FlatAppearance.BorderSize = 0
        Me.btnLock.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnLock.Location = New System.Drawing.Point(157, 3)
        Me.btnLock.Name = "btnLock"
        Me.btnLock.Size = New System.Drawing.Size(24, 23)
        Me.btnLock.TabIndex = 9
        Me.btnLock.UseVisualStyleBackColor = True
        Me.btnLock.Visible = False
        '
        'frmOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(630, 285)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.SizableToolWindow
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmOptions"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Plugin Options"
        Me.GroupBox1.ResumeLayout(False)
        CType(Me.GridParam, System.ComponentModel.ISupportInitialize).EndInit()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        CType(Me.PicLogo, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents PicLogo As System.Windows.Forms.PictureBox
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents GridParam As System.Windows.Forms.DataGridView
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnNetwork As System.Windows.Forms.Button
    Friend WithEvents btnImportXML As System.Windows.Forms.Button
    Friend WithEvents btnEscape As System.Windows.Forms.Button
    Friend WithEvents btnOK As System.Windows.Forms.Button
    Friend WithEvents btnExportXML As System.Windows.Forms.Button
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents ContextMenuStrip1 As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents ModifPath As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FolderBrowserDialog1 As System.Windows.Forms.FolderBrowserDialog
    Friend WithEvents ChargerLeParamToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ChargerParamDefautToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents NameParam As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DossierCmd As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VariableName As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents VariableType As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DefaultVar As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents SaveFileDialog1 As System.Windows.Forms.SaveFileDialog
    Friend WithEvents btnLock As System.Windows.Forms.Button
End Class
