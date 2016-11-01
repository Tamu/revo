<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmPER
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmPER))
        Me.TableLayout = New System.Windows.Forms.TableLayoutPanel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.lblDestFolder = New System.Windows.Forms.Label()
        Me.btnBrowse = New System.Windows.Forms.Button()
        Me.txtDestFolder = New System.Windows.Forms.TextBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.lstLVObjets = New System.Windows.Forms.ListBox()
        Me.btnNext = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnSupprSel = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnDigit = New System.Windows.Forms.Button()
        Me.TableLayout.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayout
        '
        Me.TableLayout.BackColor = System.Drawing.Color.Transparent
        Me.TableLayout.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayout.ColumnCount = 1
        Me.TableLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayout.Controls.Add(Me.Panel2, 0, 1)
        Me.TableLayout.Controls.Add(Me.Panel1, 0, 0)
        Me.TableLayout.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayout.Location = New System.Drawing.Point(0, 0)
        Me.TableLayout.Margin = New System.Windows.Forms.Padding(1)
        Me.TableLayout.Name = "TableLayout"
        Me.TableLayout.Padding = New System.Windows.Forms.Padding(2)
        Me.TableLayout.RowCount = 2
        Me.TableLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 240.0!))
        Me.TableLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayout.Size = New System.Drawing.Size(409, 320)
        Me.TableLayout.TabIndex = 15
        '
        'Panel2
        '
        Me.Panel2.Controls.Add(Me.lblDestFolder)
        Me.Panel2.Controls.Add(Me.btnBrowse)
        Me.Panel2.Controls.Add(Me.txtDestFolder)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel2.Location = New System.Drawing.Point(6, 247)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(397, 67)
        Me.Panel2.TabIndex = 1
        '
        'lblDestFolder
        '
        Me.lblDestFolder.AutoSize = True
        Me.lblDestFolder.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDestFolder.ForeColor = System.Drawing.Color.White
        Me.lblDestFolder.Location = New System.Drawing.Point(6, 11)
        Me.lblDestFolder.Name = "lblDestFolder"
        Me.lblDestFolder.Size = New System.Drawing.Size(201, 14)
        Me.lblDestFolder.TabIndex = 3
        Me.lblDestFolder.Text = "Dossier de destination du fichier résultat"
        '
        'btnBrowse
        '
        Me.btnBrowse.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnBrowse.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btnBrowse.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnBrowse.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBrowse.Image = Global.Revo.My.Resources.Resources.open
        Me.btnBrowse.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnBrowse.Location = New System.Drawing.Point(316, 28)
        Me.btnBrowse.Name = "btnBrowse"
        Me.btnBrowse.Size = New System.Drawing.Size(77, 24)
        Me.btnBrowse.TabIndex = 1
        Me.btnBrowse.Text = "Parcourir"
        Me.btnBrowse.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnBrowse.UseVisualStyleBackColor = True
        '
        'txtDestFolder
        '
        Me.txtDestFolder.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDestFolder.Location = New System.Drawing.Point(9, 29)
        Me.txtDestFolder.Name = "txtDestFolder"
        Me.txtDestFolder.Size = New System.Drawing.Size(303, 21)
        Me.txtDestFolder.TabIndex = 4
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.lstLVObjets)
        Me.Panel1.Controls.Add(Me.btnNext)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.btnSupprSel)
        Me.Panel1.Controls.Add(Me.btnCancel)
        Me.Panel1.Controls.Add(Me.btnDigit)
        Me.Panel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel1.Location = New System.Drawing.Point(6, 6)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(397, 234)
        Me.Panel1.TabIndex = 0
        '
        'lstLVObjets
        '
        Me.lstLVObjets.FormattingEnabled = True
        Me.lstLVObjets.Location = New System.Drawing.Point(9, 59)
        Me.lstLVObjets.Name = "lstLVObjets"
        Me.lstLVObjets.Size = New System.Drawing.Size(234, 160)
        Me.lstLVObjets.TabIndex = 2
        '
        'btnNext
        '
        Me.btnNext.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnNext.Enabled = False
        Me.btnNext.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnNext.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnNext.Location = New System.Drawing.Point(327, 195)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(66, 24)
        Me.btnNext.TabIndex = 2
        Me.btnNext.Tag = ""
        Me.btnNext.Text = "Valider"
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Location = New System.Drawing.Point(6, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(337, 28)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Sélectionner les objets à analyser (bien-fonds et bâtiments) " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "via la Sélection à" & _
    " l'écran et définir les options de traitement."
        '
        'btnSupprSel
        '
        Me.btnSupprSel.BackgroundImage = CType(resources.GetObject("btnSupprSel.BackgroundImage"), System.Drawing.Image)
        Me.btnSupprSel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnSupprSel.Enabled = False
        Me.btnSupprSel.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnSupprSel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSupprSel.Image = Global.Revo.My.Resources.Resources.delete
        Me.btnSupprSel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnSupprSel.Location = New System.Drawing.Point(257, 89)
        Me.btnSupprSel.Name = "btnSupprSel"
        Me.btnSupprSel.Size = New System.Drawing.Size(136, 24)
        Me.btnSupprSel.TabIndex = 8
        Me.btnSupprSel.Text = "Supprimer de la liste   "
        Me.btnSupprSel.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnSupprSel.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.BackgroundImage = CType(resources.GetObject("btnCancel.BackgroundImage"), System.Drawing.Image)
        Me.btnCancel.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnCancel.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnCancel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnCancel.Location = New System.Drawing.Point(257, 195)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(66, 24)
        Me.btnCancel.TabIndex = 3
        Me.btnCancel.Tag = ""
        Me.btnCancel.Text = "Annuler "
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnDigit
        '
        Me.btnDigit.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app_active
        Me.btnDigit.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnDigit.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnDigit.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDigit.Image = Global.Revo.My.Resources.Resources._select
        Me.btnDigit.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btnDigit.Location = New System.Drawing.Point(257, 59)
        Me.btnDigit.Name = "btnDigit"
        Me.btnDigit.Size = New System.Drawing.Size(136, 24)
        Me.btnDigit.TabIndex = 0
        Me.btnDigit.Text = "Sélectionner à l'écran"
        Me.btnDigit.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btnDigit.UseVisualStyleBackColor = True
        '
        'frmPER
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(409, 320)
        Me.Controls.Add(Me.TableLayout)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmPER"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Contour périmétrique"
        Me.TableLayout.ResumeLayout(False)
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayout As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents lblDestFolder As System.Windows.Forms.Label
    Friend WithEvents btnBrowse As System.Windows.Forms.Button
    Friend WithEvents txtDestFolder As System.Windows.Forms.TextBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents lstLVObjets As System.Windows.Forms.ListBox
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSupprSel As System.Windows.Forms.Button
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnDigit As System.Windows.Forms.Button
End Class
