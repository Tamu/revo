<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmImportExport
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmImportExport))
        Me.LayoutComPlan = New System.Windows.Forms.TableLayoutPanel()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.CkComPlan = New System.Windows.Forms.CheckedListBox()
        Me.LayoutTheme = New System.Windows.Forms.TableLayoutPanel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.cboTheme = New System.Windows.Forms.ComboBox()
        Me.LayoutSelect = New System.Windows.Forms.TableLayoutPanel()
        Me.BtnSelectObj = New System.Windows.Forms.Button()
        Me.LblSelect = New System.Windows.Forms.Label()
        Me.LayoutDoublon = New System.Windows.Forms.TableLayoutPanel()
        Me.cboxDoublons = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lblInfo = New System.Windows.Forms.Label()
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.LayoutFile = New System.Windows.Forms.TableLayoutPanel()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.btnFileDialog = New System.Windows.Forms.Button()
        Me.TBNameFile = New System.Windows.Forms.TextBox()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.Ck3D = New System.Windows.Forms.CheckBox()
        Me.LayoutForm = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.LayoutComPlan.SuspendLayout()
        Me.LayoutTheme.SuspendLayout()
        Me.LayoutSelect.SuspendLayout()
        Me.LayoutDoublon.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.LayoutFile.SuspendLayout()
        Me.LayoutForm.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'LayoutComPlan
        '
        Me.LayoutComPlan.ColumnCount = 2
        Me.LayoutComPlan.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 130.0!))
        Me.LayoutComPlan.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.LayoutComPlan.Controls.Add(Me.Label2, 0, 0)
        Me.LayoutComPlan.Controls.Add(Me.CkComPlan, 1, 0)
        Me.LayoutComPlan.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutComPlan.Location = New System.Drawing.Point(3, 147)
        Me.LayoutComPlan.Name = "LayoutComPlan"
        Me.LayoutComPlan.RowCount = 1
        Me.LayoutComPlan.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.LayoutComPlan.Size = New System.Drawing.Size(397, 84)
        Me.LayoutComPlan.TabIndex = 9
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label2.AutoSize = True
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Location = New System.Drawing.Point(3, 29)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(122, 26)
        Me.Label2.TabIndex = 5
        Me.Label2.Tag = "Séparer par des virgules"
        Me.Label2.Text = "Points des communes et plans sélectionnés"
        '
        'CkComPlan
        '
        Me.CkComPlan.Dock = System.Windows.Forms.DockStyle.Fill
        Me.CkComPlan.FormattingEnabled = True
        Me.CkComPlan.Location = New System.Drawing.Point(133, 3)
        Me.CkComPlan.Name = "CkComPlan"
        Me.CkComPlan.Size = New System.Drawing.Size(261, 78)
        Me.CkComPlan.TabIndex = 6
        '
        'LayoutTheme
        '
        Me.LayoutTheme.ColumnCount = 2
        Me.LayoutTheme.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 130.0!))
        Me.LayoutTheme.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.LayoutTheme.Controls.Add(Me.Label1, 0, 0)
        Me.LayoutTheme.Controls.Add(Me.cboTheme, 1, 0)
        Me.LayoutTheme.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutTheme.Location = New System.Drawing.Point(3, 117)
        Me.LayoutTheme.Name = "LayoutTheme"
        Me.LayoutTheme.RowCount = 1
        Me.LayoutTheme.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.LayoutTheme.Size = New System.Drawing.Size(397, 24)
        Me.LayoutTheme.TabIndex = 8
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label1.AutoSize = True
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(3, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(106, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Tag = "Tous les points d'un thème"
        Me.Label1.Text = "Filtre/cible par thème"
        '
        'cboTheme
        '
        Me.cboTheme.Dock = System.Windows.Forms.DockStyle.Fill
        Me.cboTheme.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cboTheme.FormattingEnabled = True
        Me.cboTheme.Location = New System.Drawing.Point(133, 3)
        Me.cboTheme.Name = "cboTheme"
        Me.cboTheme.Size = New System.Drawing.Size(261, 21)
        Me.cboTheme.TabIndex = 6
        Me.cboTheme.Tag = "Tous les points d'un thème"
        '
        'LayoutSelect
        '
        Me.LayoutSelect.ColumnCount = 2
        Me.LayoutSelect.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.LayoutSelect.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 40.0!))
        Me.LayoutSelect.Controls.Add(Me.BtnSelectObj, 1, 0)
        Me.LayoutSelect.Controls.Add(Me.LblSelect, 0, 0)
        Me.LayoutSelect.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutSelect.Location = New System.Drawing.Point(3, 87)
        Me.LayoutSelect.Name = "LayoutSelect"
        Me.LayoutSelect.RowCount = 1
        Me.LayoutSelect.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.LayoutSelect.Size = New System.Drawing.Size(397, 24)
        Me.LayoutSelect.TabIndex = 7
        '
        'BtnSelectObj
        '
        Me.BtnSelectObj.BackColor = System.Drawing.Color.Transparent
        Me.BtnSelectObj.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.BtnSelectObj.Dock = System.Windows.Forms.DockStyle.Fill
        Me.BtnSelectObj.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.BtnSelectObj.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.BtnSelectObj.ForeColor = System.Drawing.Color.Black
        Me.BtnSelectObj.Image = Global.Revo.My.Resources.Resources._select
        Me.BtnSelectObj.Location = New System.Drawing.Point(360, 3)
        Me.BtnSelectObj.Name = "BtnSelectObj"
        Me.BtnSelectObj.Size = New System.Drawing.Size(34, 18)
        Me.BtnSelectObj.TabIndex = 7
        Me.BtnSelectObj.Tag = "Par défaut : tout le dessin"
        Me.BtnSelectObj.UseVisualStyleBackColor = False
        '
        'LblSelect
        '
        Me.LblSelect.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.LblSelect.AutoSize = True
        Me.LblSelect.ForeColor = System.Drawing.Color.White
        Me.LblSelect.Location = New System.Drawing.Point(3, 5)
        Me.LblSelect.Name = "LblSelect"
        Me.LblSelect.Size = New System.Drawing.Size(137, 13)
        Me.LblSelect.TabIndex = 8
        Me.LblSelect.Tag = "Par défaut : tout le dessin"
        Me.LblSelect.Text = "Sélection manuelle (0 objet)"
        '
        'LayoutDoublon
        '
        Me.LayoutDoublon.ColumnCount = 1
        Me.LayoutDoublon.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.LayoutDoublon.Controls.Add(Me.cboxDoublons, 0, 0)
        Me.LayoutDoublon.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutDoublon.Location = New System.Drawing.Point(3, 237)
        Me.LayoutDoublon.Name = "LayoutDoublon"
        Me.LayoutDoublon.RowCount = 1
        Me.LayoutDoublon.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.LayoutDoublon.Size = New System.Drawing.Size(397, 34)
        Me.LayoutDoublon.TabIndex = 12
        '
        'cboxDoublons
        '
        Me.cboxDoublons.AutoSize = True
        Me.cboxDoublons.ForeColor = System.Drawing.Color.White
        Me.cboxDoublons.Location = New System.Drawing.Point(6, 6)
        Me.cboxDoublons.Margin = New System.Windows.Forms.Padding(6, 6, 3, 3)
        Me.cboxDoublons.Name = "cboxDoublons"
        Me.cboxDoublons.Size = New System.Drawing.Size(213, 17)
        Me.cboxDoublons.TabIndex = 6
        Me.cboxDoublons.Tag = "Recherche selon les attributs ""Numero"" et ""IdentDN"""
        Me.cboxDoublons.Text = "Détection des doublons (même numéro)"
        Me.cboxDoublons.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.None
        Me.GroupBox2.BackColor = System.Drawing.Color.White
        Me.GroupBox2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None
        Me.GroupBox2.Controls.Add(Me.lblInfo)
        Me.GroupBox2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.GroupBox2.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.GroupBox2.Location = New System.Drawing.Point(3, 3)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(397, 44)
        Me.GroupBox2.TabIndex = 5
        Me.GroupBox2.TabStop = False
        '
        'lblInfo
        '
        Me.lblInfo.AutoSize = True
        Me.lblInfo.BackColor = System.Drawing.Color.Transparent
        Me.lblInfo.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.lblInfo.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblInfo.Location = New System.Drawing.Point(17, 17)
        Me.lblInfo.Name = "lblInfo"
        Me.lblInfo.Size = New System.Drawing.Size(199, 14)
        Me.lblInfo.TabIndex = 4
        Me.lblInfo.Text = "Sélectionner le format et les paramètres"
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialog1"
        '
        'LayoutFile
        '
        Me.LayoutFile.ColumnCount = 3
        Me.LayoutFile.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 130.0!))
        Me.LayoutFile.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.LayoutFile.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 40.0!))
        Me.LayoutFile.Controls.Add(Me.Label3, 0, 0)
        Me.LayoutFile.Controls.Add(Me.btnFileDialog, 2, 0)
        Me.LayoutFile.Controls.Add(Me.TBNameFile, 1, 0)
        Me.LayoutFile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutFile.Location = New System.Drawing.Point(3, 53)
        Me.LayoutFile.Name = "LayoutFile"
        Me.LayoutFile.RowCount = 1
        Me.LayoutFile.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 30.0!))
        Me.LayoutFile.Size = New System.Drawing.Size(397, 28)
        Me.LayoutFile.TabIndex = 0
        '
        'Label3
        '
        Me.Label3.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label3.AutoSize = True
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Location = New System.Drawing.Point(3, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(38, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Tag = "Sélection du fichier + format"
        Me.Label3.Text = "Fichier"
        '
        'btnFileDialog
        '
        Me.btnFileDialog.BackColor = System.Drawing.Color.Transparent
        Me.btnFileDialog.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnFileDialog.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnFileDialog.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnFileDialog.ForeColor = System.Drawing.Color.Black
        Me.btnFileDialog.Image = Global.Revo.My.Resources.Resources.open
        Me.btnFileDialog.Location = New System.Drawing.Point(360, 3)
        Me.btnFileDialog.Name = "btnFileDialog"
        Me.btnFileDialog.Size = New System.Drawing.Size(34, 24)
        Me.btnFileDialog.TabIndex = 0
        Me.btnFileDialog.Tag = "Sélectionner un format"
        Me.btnFileDialog.UseVisualStyleBackColor = False
        '
        'TBNameFile
        '
        Me.TBNameFile.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TBNameFile.Location = New System.Drawing.Point(133, 3)
        Me.TBNameFile.Name = "TBNameFile"
        Me.TBNameFile.Size = New System.Drawing.Size(221, 20)
        Me.TBNameFile.TabIndex = 2
        Me.TBNameFile.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'btnNext
        '
        Me.btnNext.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnNext.BackColor = System.Drawing.Color.Transparent
        Me.btnNext.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnNext.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNext.ForeColor = System.Drawing.Color.Black
        Me.btnNext.Location = New System.Drawing.Point(307, 5)
        Me.btnNext.Margin = New System.Windows.Forms.Padding(3, 3, 3, 5)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(87, 25)
        Me.btnNext.TabIndex = 7
        Me.btnNext.Tag = "Identification et mise à jour de la licence"
        Me.btnNext.Text = "Valider"
        Me.btnNext.UseVisualStyleBackColor = False
        '
        'Ck3D
        '
        Me.Ck3D.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Ck3D.AutoSize = True
        Me.Ck3D.Checked = False
        'Me.Ck3D.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Ck3D.ForeColor = System.Drawing.Color.White
        Me.Ck3D.Location = New System.Drawing.Point(6, 10)
        Me.Ck3D.Margin = New System.Windows.Forms.Padding(6, 6, 3, 8)
        Me.Ck3D.Name = "Ck3D"
        Me.Ck3D.Size = New System.Drawing.Size(229, 17)
        Me.Ck3D.TabIndex = 7
        Me.Ck3D.Tag = "La coordonnée Z est prioritaire sur l'attribut"
        Me.Ck3D.Text = "Charger les points en 3D (coord. Z du bloc)"
        Me.Ck3D.UseVisualStyleBackColor = True
        '
        'LayoutForm
        '
        Me.LayoutForm.BackColor = System.Drawing.Color.Transparent
        Me.LayoutForm.ColumnCount = 1
        Me.LayoutForm.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.LayoutForm.Controls.Add(Me.LayoutFile, 0, 1)
        Me.LayoutForm.Controls.Add(Me.GroupBox2, 0, 0)
        Me.LayoutForm.Controls.Add(Me.TableLayoutPanel3, 0, 6)
        Me.LayoutForm.Controls.Add(Me.LayoutDoublon, 0, 5)
        Me.LayoutForm.Controls.Add(Me.LayoutComPlan, 0, 4)
        Me.LayoutForm.Controls.Add(Me.LayoutTheme, 0, 3)
        Me.LayoutForm.Controls.Add(Me.LayoutSelect, 0, 2)
        Me.LayoutForm.Dock = System.Windows.Forms.DockStyle.Fill
        Me.LayoutForm.Location = New System.Drawing.Point(0, 0)
        Me.LayoutForm.Name = "LayoutForm"
        Me.LayoutForm.RowCount = 7
        Me.LayoutForm.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.LayoutForm.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.LayoutForm.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.LayoutForm.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.LayoutForm.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.LayoutForm.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.LayoutForm.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.LayoutForm.Size = New System.Drawing.Size(403, 315)
        Me.LayoutForm.TabIndex = 8
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.ColumnCount = 2
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70.0!))
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.Ck3D, 0, 0)
        Me.TableLayoutPanel3.Controls.Add(Me.btnNext, 1, 0)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 277)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 1
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(397, 35)
        Me.TableLayoutPanel3.TabIndex = 0
        '
        'frmImportExport
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(403, 315)
        Me.Controls.Add(Me.LayoutForm)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmImportExport"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Importation/Exportation de données"
        Me.LayoutComPlan.ResumeLayout(False)
        Me.LayoutComPlan.PerformLayout()
        Me.LayoutTheme.ResumeLayout(False)
        Me.LayoutTheme.PerformLayout()
        Me.LayoutSelect.ResumeLayout(False)
        Me.LayoutSelect.PerformLayout()
        Me.LayoutDoublon.ResumeLayout(False)
        Me.LayoutDoublon.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.LayoutFile.ResumeLayout(False)
        Me.LayoutFile.PerformLayout()
        Me.LayoutForm.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents lblInfo As System.Windows.Forms.Label
    Friend WithEvents OpenFileDialog1 As System.Windows.Forms.OpenFileDialog
    Friend WithEvents LayoutFile As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents btnFileDialog As System.Windows.Forms.Button
    Friend WithEvents TBNameFile As System.Windows.Forms.TextBox
    Friend WithEvents LayoutComPlan As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LayoutTheme As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents cboTheme As System.Windows.Forms.ComboBox
    Friend WithEvents LayoutSelect As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents BtnSelectObj As System.Windows.Forms.Button
    Friend WithEvents LblSelect As System.Windows.Forms.Label
    Friend WithEvents LayoutDoublon As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents cboxDoublons As System.Windows.Forms.CheckBox
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents CkComPlan As System.Windows.Forms.CheckedListBox
    Friend WithEvents Ck3D As System.Windows.Forms.CheckBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents LayoutForm As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
End Class
