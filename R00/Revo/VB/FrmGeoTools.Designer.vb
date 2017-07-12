<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmGeoTools
    Inherits System.Windows.Forms.UserControl

    'UserControl remplace la méthode Dispose pour nettoyer la liste des composants.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel3 = New System.Windows.Forms.TableLayoutPanel()
        Me.ListFind = New System.Windows.Forms.ListBox()
        Me.TableLayoutPanel5 = New System.Windows.Forms.TableLayoutPanel()
        Me.lblFindBF = New System.Windows.Forms.Label()
        Me.btnBFfound = New System.Windows.Forms.Button()
        Me.txtBFfound = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel4 = New System.Windows.Forms.TableLayoutPanel()
        Me.AnalyseBF = New System.Windows.Forms.Button()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.TableLayoutPanel2 = New System.Windows.Forms.TableLayoutPanel()
        Me.XdataBat = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel4 = New System.Windows.Forms.Panel()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.TableLayoutPanel3.SuspendLayout()
        Me.TableLayoutPanel5.SuspendLayout()
        Me.TableLayoutPanel4.SuspendLayout()
        Me.TableLayoutPanel2.SuspendLayout()
        Me.Panel4.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.BackColor = System.Drawing.Color.Transparent
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.TableLayoutPanel3, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.Panel4, 0, 0)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(0, 0)
        Me.TableLayoutPanel1.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 2
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 80.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 300.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(290, 400)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'TableLayoutPanel3
        '
        Me.TableLayoutPanel3.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayoutPanel3.ColumnCount = 1
        Me.TableLayoutPanel3.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel3.Controls.Add(Me.ListFind, 0, 3)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel5, 0, 2)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel4, 0, 1)
        Me.TableLayoutPanel3.Controls.Add(Me.TableLayoutPanel2, 0, 0)
        Me.TableLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel3.Location = New System.Drawing.Point(3, 83)
        Me.TableLayoutPanel3.Name = "TableLayoutPanel3"
        Me.TableLayoutPanel3.RowCount = 4
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20.0!))
        Me.TableLayoutPanel3.Size = New System.Drawing.Size(284, 314)
        Me.TableLayoutPanel3.TabIndex = 63
        '
        'ListFind
        '
        Me.ListFind.BackColor = System.Drawing.Color.White
        Me.ListFind.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListFind.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ListFind.FormattingEnabled = True
        Me.ListFind.ItemHeight = 18
        Me.ListFind.Location = New System.Drawing.Point(4, 157)
        Me.ListFind.Name = "ListFind"
        Me.ListFind.Size = New System.Drawing.Size(276, 153)
        Me.ListFind.TabIndex = 64
        '
        'TableLayoutPanel5
        '
        Me.TableLayoutPanel5.ColumnCount = 3
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 75.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 35.0!))
        Me.TableLayoutPanel5.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel5.Controls.Add(Me.lblFindBF, 2, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.btnBFfound, 1, 0)
        Me.TableLayoutPanel5.Controls.Add(Me.txtBFfound, 0, 0)
        Me.TableLayoutPanel5.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel5.Location = New System.Drawing.Point(4, 106)
        Me.TableLayoutPanel5.Name = "TableLayoutPanel5"
        Me.TableLayoutPanel5.RowCount = 1
        Me.TableLayoutPanel5.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel5.Size = New System.Drawing.Size(276, 44)
        Me.TableLayoutPanel5.TabIndex = 63
        '
        'lblFindBF
        '
        Me.lblFindBF.AutoSize = True
        Me.lblFindBF.Location = New System.Drawing.Point(112, 8)
        Me.lblFindBF.Margin = New System.Windows.Forms.Padding(2, 8, 0, 0)
        Me.lblFindBF.Name = "lblFindBF"
        Me.lblFindBF.Size = New System.Drawing.Size(137, 26)
        Me.lblFindBF.TabIndex = 57
        Me.lblFindBF.Text = "Recherche de parcelle (BF)" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(Espace pour lancer)"
        '
        'btnBFfound
        '
        Me.btnBFfound.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btnBFfound.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnBFfound.Dock = System.Windows.Forms.DockStyle.Top
        Me.btnBFfound.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btnBFfound.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnBFfound.Location = New System.Drawing.Point(76, 8)
        Me.btnBFfound.Margin = New System.Windows.Forms.Padding(1, 8, 0, 0)
        Me.btnBFfound.Name = "btnBFfound"
        Me.btnBFfound.Size = New System.Drawing.Size(34, 28)
        Me.btnBFfound.TabIndex = 0
        Me.btnBFfound.Tag = ""
        Me.btnBFfound.Text = "GO"
        Me.btnBFfound.UseVisualStyleBackColor = True
        '
        'txtBFfound
        '
        Me.txtBFfound.Dock = System.Windows.Forms.DockStyle.Top
        Me.txtBFfound.Location = New System.Drawing.Point(3, 12)
        Me.txtBFfound.Margin = New System.Windows.Forms.Padding(3, 12, 3, 3)
        Me.txtBFfound.Name = "txtBFfound"
        Me.txtBFfound.Size = New System.Drawing.Size(69, 20)
        Me.txtBFfound.TabIndex = 1
        '
        'TableLayoutPanel4
        '
        Me.TableLayoutPanel4.ColumnCount = 2
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 75.0!))
        Me.TableLayoutPanel4.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel4.Controls.Add(Me.AnalyseBF, 0, 0)
        Me.TableLayoutPanel4.Controls.Add(Me.Label5, 1, 0)
        Me.TableLayoutPanel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel4.Location = New System.Drawing.Point(4, 55)
        Me.TableLayoutPanel4.Name = "TableLayoutPanel4"
        Me.TableLayoutPanel4.RowCount = 1
        Me.TableLayoutPanel4.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel4.Size = New System.Drawing.Size(276, 44)
        Me.TableLayoutPanel4.TabIndex = 62
        '
        'AnalyseBF
        '
        Me.AnalyseBF.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.AnalyseBF.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.AnalyseBF.Dock = System.Windows.Forms.DockStyle.Top
        Me.AnalyseBF.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.AnalyseBF.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.AnalyseBF.Image = Global.Revo.My.Resources.Resources.MOvaliderMut
        Me.AnalyseBF.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.AnalyseBF.Location = New System.Drawing.Point(5, 8)
        Me.AnalyseBF.Margin = New System.Windows.Forms.Padding(5, 8, 0, 0)
        Me.AnalyseBF.Name = "AnalyseBF"
        Me.AnalyseBF.Size = New System.Drawing.Size(70, 28)
        Me.AnalyseBF.TabIndex = 3
        Me.AnalyseBF.Tag = ""
        Me.AnalyseBF.Text = "      Lancer"
        Me.AnalyseBF.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(77, 8)
        Me.Label5.Margin = New System.Windows.Forms.Padding(2, 8, 0, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(153, 26)
        Me.Label5.TabIndex = 57
        Me.Label5.Text = "Analyse des polylignes fermées" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "(BF, CS, OD : surfacique)"
        '
        'TableLayoutPanel2
        '
        Me.TableLayoutPanel2.ColumnCount = 2
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Absolute, 75.0!))
        Me.TableLayoutPanel2.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle())
        Me.TableLayoutPanel2.Controls.Add(Me.XdataBat, 0, 0)
        Me.TableLayoutPanel2.Controls.Add(Me.Label2, 1, 0)
        Me.TableLayoutPanel2.Location = New System.Drawing.Point(4, 4)
        Me.TableLayoutPanel2.Name = "TableLayoutPanel2"
        Me.TableLayoutPanel2.RowCount = 1
        Me.TableLayoutPanel2.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 50.0!))
        Me.TableLayoutPanel2.Size = New System.Drawing.Size(276, 44)
        Me.TableLayoutPanel2.TabIndex = 61
        '
        'XdataBat
        '
        Me.XdataBat.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.XdataBat.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.XdataBat.Dock = System.Windows.Forms.DockStyle.Top
        Me.XdataBat.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.XdataBat.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.XdataBat.Image = Global.Revo.My.Resources.Resources._select
        Me.XdataBat.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.XdataBat.Location = New System.Drawing.Point(5, 8)
        Me.XdataBat.Margin = New System.Windows.Forms.Padding(5, 8, 0, 0)
        Me.XdataBat.Name = "XdataBat"
        Me.XdataBat.Size = New System.Drawing.Size(70, 28)
        Me.XdataBat.TabIndex = 2
        Me.XdataBat.Tag = ""
        Me.XdataBat.Text = "      Lancer"
        Me.XdataBat.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(77, 8)
        Me.Label2.Margin = New System.Windows.Forms.Padding(2, 8, 0, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(164, 26)
        Me.Label2.TabIndex = 57
        Me.Label2.Text = "Mise à jour de la relations entre" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "le n° du bâtiment et son périmètre"
        '
        'Panel4
        '
        Me.Panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Panel4.Controls.Add(Me.Label1)
        Me.Panel4.Controls.Add(Me.lblTitle)
        Me.Panel4.Dock = System.Windows.Forms.DockStyle.Fill
        Me.Panel4.Location = New System.Drawing.Point(3, 5)
        Me.Panel4.Margin = New System.Windows.Forms.Padding(3, 5, 3, 3)
        Me.Panel4.Name = "Panel4"
        Me.Panel4.Size = New System.Drawing.Size(284, 72)
        Me.Panel4.TabIndex = 4
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(8, 23)
        Me.Label1.Margin = New System.Windows.Forms.Padding(10, 11, 10, 11)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(163, 26)
        Me.Label1.TabIndex = 33
        Me.Label1.Text = "Des outils pour la mensuration et " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "d'analyse geospatiale"
        '
        'lblTitle
        '
        Me.lblTitle.BackColor = System.Drawing.Color.Transparent
        Me.lblTitle.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblTitle.Font = New System.Drawing.Font("Tahoma", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTitle.ForeColor = System.Drawing.SystemColors.ControlText
        Me.lblTitle.Location = New System.Drawing.Point(8, 6)
        Me.lblTitle.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.lblTitle.Size = New System.Drawing.Size(72, 20)
        Me.lblTitle.TabIndex = 34
        Me.lblTitle.Text = "GeoTools"
        '
        'FrmGeoTools
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoScroll = True
        Me.AutoScrollMargin = New System.Drawing.Size(0, 1)
        Me.AutoScrollMinSize = New System.Drawing.Size(0, 300)
        Me.AutoSize = True
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.Margin = New System.Windows.Forms.Padding(6, 7, 6, 7)
        Me.MinimumSize = New System.Drawing.Size(290, 0)
        Me.Name = "FrmGeoTools"
        Me.Size = New System.Drawing.Size(290, 400)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.TableLayoutPanel3.ResumeLayout(False)
        Me.TableLayoutPanel5.ResumeLayout(False)
        Me.TableLayoutPanel5.PerformLayout()
        Me.TableLayoutPanel4.ResumeLayout(False)
        Me.TableLayoutPanel4.PerformLayout()
        Me.TableLayoutPanel2.ResumeLayout(False)
        Me.TableLayoutPanel2.PerformLayout()
        Me.Panel4.ResumeLayout(False)
        Me.Panel4.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Private m_Host As Autodesk.AutoCAD.Windows.PaletteSet

    Public Sub New(ByRef Host As Autodesk.AutoCAD.Windows.PaletteSet)

        ' Cet appel est requis par le concepteur.
        InitializeComponent()

        ' Ajoutez une initialisation quelconque après l'appel InitializeComponent().
        m_Host = Host

        LoadVariable()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents Panel4 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Public WithEvents lblTitle As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel3 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel2 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents XdataBat As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel4 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents AnalyseBF As System.Windows.Forms.Button
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents TableLayoutPanel5 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents lblFindBF As System.Windows.Forms.Label
    Friend WithEvents btnBFfound As System.Windows.Forms.Button
    Friend WithEvents txtBFfound As System.Windows.Forms.TextBox
    Friend WithEvents ListFind As System.Windows.Forms.ListBox
End Class
