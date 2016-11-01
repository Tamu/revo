<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmDivSurf
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
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmDivSurf))
        Me.TableLayout = New System.Windows.Forms.TableLayoutPanel()
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.pan_enregistrement = New System.Windows.Forms.Panel()
        Me.btn_calcul = New System.Windows.Forms.Button()
        Me.CBox_enregistrement = New System.Windows.Forms.CheckBox()
        Me.lbl_enregistrement = New System.Windows.Forms.Label()
        Me.btn_enregistrement = New System.Windows.Forms.Button()
        Me.tbox_enregistrement = New System.Windows.Forms.TextBox()
        Me.pan_init = New System.Windows.Forms.Panel()
        Me.pan_methode = New System.Windows.Forms.Panel()
        Me.pan_orientation_pt = New System.Windows.Forms.Panel()
        Me.lbl_orientation_pt = New System.Windows.Forms.Label()
        Me.bnt_orientation_pt = New System.Windows.Forms.Button()
        Me.pan_orientation_dir = New System.Windows.Forms.Panel()
        Me.rbnt_dir_perp = New System.Windows.Forms.RadioButton()
        Me.rbnt_dir_par = New System.Windows.Forms.RadioButton()
        Me.lbl_orientation_dir = New System.Windows.Forms.Label()
        Me.bnt_orientation_dir = New System.Windows.Forms.Button()
        Me.rbtn_methode_pt_fixe = New System.Windows.Forms.RadioButton()
        Me.rbtn_methode_direction = New System.Windows.Forms.RadioButton()
        Me.lbl_methode = New System.Windows.Forms.Label()
        Me.pan_selection_parcelle = New System.Windows.Forms.Panel()
        Me.lbl_selection_parcelle = New System.Windows.Forms.Label()
        Me.btn_selection_parcelle = New System.Windows.Forms.Button()
        Me.pan_N_parcelle = New System.Windows.Forms.Panel()
        Me.tbox_N_parcelle = New System.Windows.Forms.TextBox()
        Me.lbl_N_parcelle_1 = New System.Windows.Forms.Label()
        Me.pan_calcul_surf = New System.Windows.Forms.Panel()
        Me.lbl_surf_reste = New System.Windows.Forms.Label()
        Me.DataGridView_suface = New System.Windows.Forms.DataGridView()
        Me.Col_surface = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.rbtn_calcul_surface_choisie = New System.Windows.Forms.RadioButton()
        Me.rbtn_calcul_surface_egale = New System.Windows.Forms.RadioButton()
        Me.lbl_calcul_surf_1 = New System.Windows.Forms.Label()
        Me.ToolTip_selection_orientation = New System.Windows.Forms.ToolTip(Me.components)
        Me.chkLayerMO = New System.Windows.Forms.CheckBox()
        Me.TableLayout.SuspendLayout()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.pan_enregistrement.SuspendLayout()
        Me.pan_init.SuspendLayout()
        Me.pan_methode.SuspendLayout()
        Me.pan_orientation_pt.SuspendLayout()
        Me.pan_orientation_dir.SuspendLayout()
        Me.pan_selection_parcelle.SuspendLayout()
        Me.pan_N_parcelle.SuspendLayout()
        Me.pan_calcul_surf.SuspendLayout()
        CType(Me.DataGridView_suface, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'TableLayout
        '
        Me.TableLayout.BackColor = System.Drawing.Color.Transparent
        Me.TableLayout.CellBorderStyle = System.Windows.Forms.TableLayoutPanelCellBorderStyle.[Single]
        Me.TableLayout.ColumnCount = 1
        Me.TableLayout.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayout.Controls.Add(Me.TableLayoutPanel1, 0, 0)
        Me.TableLayout.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayout.Location = New System.Drawing.Point(0, 0)
        Me.TableLayout.Margin = New System.Windows.Forms.Padding(1)
        Me.TableLayout.Name = "TableLayout"
        Me.TableLayout.Padding = New System.Windows.Forms.Padding(2)
        Me.TableLayout.RowCount = 1
        Me.TableLayout.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 945.0!))
        Me.TableLayout.Size = New System.Drawing.Size(823, 486)
        Me.TableLayout.TabIndex = 14
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.ColumnCount = 1
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.pan_enregistrement, 0, 4)
        Me.TableLayoutPanel1.Controls.Add(Me.pan_init, 0, 1)
        Me.TableLayoutPanel1.Controls.Add(Me.pan_N_parcelle, 0, 2)
        Me.TableLayoutPanel1.Controls.Add(Me.pan_calcul_surf, 0, 3)
        Me.TableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(6, 6)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 5
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle())
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 207.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 60.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 150.0!))
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 295.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(811, 939)
        Me.TableLayoutPanel1.TabIndex = 2
        '
        'pan_enregistrement
        '
        Me.pan_enregistrement.BackColor = System.Drawing.Color.Transparent
        Me.pan_enregistrement.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan_enregistrement.Controls.Add(Me.chkLayerMO)
        Me.pan_enregistrement.Controls.Add(Me.btn_calcul)
        Me.pan_enregistrement.Controls.Add(Me.CBox_enregistrement)
        Me.pan_enregistrement.Controls.Add(Me.lbl_enregistrement)
        Me.pan_enregistrement.Controls.Add(Me.btn_enregistrement)
        Me.pan_enregistrement.Controls.Add(Me.tbox_enregistrement)
        Me.pan_enregistrement.Dock = System.Windows.Forms.DockStyle.Top
        Me.pan_enregistrement.Location = New System.Drawing.Point(3, 420)
        Me.pan_enregistrement.Name = "pan_enregistrement"
        Me.pan_enregistrement.Size = New System.Drawing.Size(805, 49)
        Me.pan_enregistrement.TabIndex = 15
        '
        'btn_calcul
        '
        Me.btn_calcul.BackColor = System.Drawing.Color.LightGray
        Me.btn_calcul.BackgroundImage = CType(resources.GetObject("btn_calcul.BackgroundImage"), System.Drawing.Image)
        Me.btn_calcul.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btn_calcul.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btn_calcul.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_calcul.Location = New System.Drawing.Point(707, 11)
        Me.btn_calcul.Margin = New System.Windows.Forms.Padding(3, 2, 3, 3)
        Me.btn_calcul.Name = "btn_calcul"
        Me.btn_calcul.Size = New System.Drawing.Size(91, 25)
        Me.btn_calcul.TabIndex = 15
        Me.btn_calcul.Tag = "Démarrer le flux sélectionné"
        Me.btn_calcul.Text = "Calcul"
        Me.btn_calcul.UseVisualStyleBackColor = False
        '
        'CBox_enregistrement
        '
        Me.CBox_enregistrement.AutoSize = True
        Me.CBox_enregistrement.ForeColor = System.Drawing.Color.White
        Me.CBox_enregistrement.Location = New System.Drawing.Point(9, 57)
        Me.CBox_enregistrement.Name = "CBox_enregistrement"
        Me.CBox_enregistrement.Size = New System.Drawing.Size(148, 17)
        Me.CBox_enregistrement.TabIndex = 5
        Me.CBox_enregistrement.Text = "Enregistrement du résultat"
        Me.CBox_enregistrement.UseVisualStyleBackColor = True
        Me.CBox_enregistrement.Visible = False
        '
        'lbl_enregistrement
        '
        Me.lbl_enregistrement.AutoSize = True
        Me.lbl_enregistrement.Font = New System.Drawing.Font("Arial", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_enregistrement.ForeColor = System.Drawing.Color.White
        Me.lbl_enregistrement.Location = New System.Drawing.Point(6, 11)
        Me.lbl_enregistrement.Name = "lbl_enregistrement"
        Me.lbl_enregistrement.Size = New System.Drawing.Size(201, 14)
        Me.lbl_enregistrement.TabIndex = 3
        Me.lbl_enregistrement.Text = "Dossier de destination du fichier résultat"
        Me.lbl_enregistrement.Visible = False
        '
        'btn_enregistrement
        '
        Me.btn_enregistrement.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app
        Me.btn_enregistrement.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center
        Me.btn_enregistrement.Enabled = False
        Me.btn_enregistrement.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btn_enregistrement.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_enregistrement.Image = Global.Revo.My.Resources.Resources.open
        Me.btn_enregistrement.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_enregistrement.Location = New System.Drawing.Point(316, 28)
        Me.btn_enregistrement.Name = "btn_enregistrement"
        Me.btn_enregistrement.Size = New System.Drawing.Size(77, 24)
        Me.btn_enregistrement.TabIndex = 1
        Me.btn_enregistrement.Text = "Parcourir"
        Me.btn_enregistrement.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_enregistrement.UseVisualStyleBackColor = True
        Me.btn_enregistrement.Visible = False
        '
        'tbox_enregistrement
        '
        Me.tbox_enregistrement.Enabled = False
        Me.tbox_enregistrement.Font = New System.Drawing.Font("Microsoft Sans Serif", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.tbox_enregistrement.Location = New System.Drawing.Point(9, 29)
        Me.tbox_enregistrement.Multiline = True
        Me.tbox_enregistrement.Name = "tbox_enregistrement"
        Me.tbox_enregistrement.Size = New System.Drawing.Size(303, 20)
        Me.tbox_enregistrement.TabIndex = 4
        Me.tbox_enregistrement.Visible = False
        '
        'pan_init
        '
        Me.pan_init.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan_init.Controls.Add(Me.pan_methode)
        Me.pan_init.Controls.Add(Me.pan_selection_parcelle)
        Me.pan_init.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pan_init.Location = New System.Drawing.Point(3, 3)
        Me.pan_init.Name = "pan_init"
        Me.pan_init.Size = New System.Drawing.Size(805, 201)
        Me.pan_init.TabIndex = 11
        '
        'pan_methode
        '
        Me.pan_methode.Controls.Add(Me.pan_orientation_pt)
        Me.pan_methode.Controls.Add(Me.pan_orientation_dir)
        Me.pan_methode.Controls.Add(Me.rbtn_methode_pt_fixe)
        Me.pan_methode.Controls.Add(Me.rbtn_methode_direction)
        Me.pan_methode.Controls.Add(Me.lbl_methode)
        Me.pan_methode.Location = New System.Drawing.Point(9, 77)
        Me.pan_methode.Name = "pan_methode"
        Me.pan_methode.Size = New System.Drawing.Size(837, 106)
        Me.pan_methode.TabIndex = 8
        '
        'pan_orientation_pt
        '
        Me.pan_orientation_pt.Controls.Add(Me.lbl_orientation_pt)
        Me.pan_orientation_pt.Controls.Add(Me.bnt_orientation_pt)
        Me.pan_orientation_pt.Location = New System.Drawing.Point(405, 66)
        Me.pan_orientation_pt.Name = "pan_orientation_pt"
        Me.pan_orientation_pt.Size = New System.Drawing.Size(388, 40)
        Me.pan_orientation_pt.TabIndex = 12
        Me.pan_orientation_pt.Visible = False
        '
        'lbl_orientation_pt
        '
        Me.lbl_orientation_pt.AutoSize = True
        Me.lbl_orientation_pt.ForeColor = System.Drawing.Color.White
        Me.lbl_orientation_pt.Location = New System.Drawing.Point(144, 17)
        Me.lbl_orientation_pt.Name = "lbl_orientation_pt"
        Me.lbl_orientation_pt.Size = New System.Drawing.Size(227, 13)
        Me.lbl_orientation_pt.TabIndex = 6
        Me.lbl_orientation_pt.Text = "Sélection du point fixe et du sens de la division"
        '
        'bnt_orientation_pt
        '
        Me.bnt_orientation_pt.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app_active
        Me.bnt_orientation_pt.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.bnt_orientation_pt.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.bnt_orientation_pt.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.bnt_orientation_pt.Image = Global.Revo.My.Resources.Resources._select
        Me.bnt_orientation_pt.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bnt_orientation_pt.Location = New System.Drawing.Point(3, 11)
        Me.bnt_orientation_pt.Name = "bnt_orientation_pt"
        Me.bnt_orientation_pt.Size = New System.Drawing.Size(136, 24)
        Me.bnt_orientation_pt.TabIndex = 5
        Me.bnt_orientation_pt.Text = "Sélectionner à l'écran"
        Me.bnt_orientation_pt.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.bnt_orientation_pt.UseVisualStyleBackColor = True
        '
        'pan_orientation_dir
        '
        Me.pan_orientation_dir.Controls.Add(Me.rbnt_dir_perp)
        Me.pan_orientation_dir.Controls.Add(Me.rbnt_dir_par)
        Me.pan_orientation_dir.Controls.Add(Me.lbl_orientation_dir)
        Me.pan_orientation_dir.Controls.Add(Me.bnt_orientation_dir)
        Me.pan_orientation_dir.Location = New System.Drawing.Point(405, 9)
        Me.pan_orientation_dir.Name = "pan_orientation_dir"
        Me.pan_orientation_dir.Size = New System.Drawing.Size(388, 61)
        Me.pan_orientation_dir.TabIndex = 11
        Me.pan_orientation_dir.Visible = False
        '
        'rbnt_dir_perp
        '
        Me.rbnt_dir_perp.AutoSize = True
        Me.rbnt_dir_perp.ForeColor = System.Drawing.Color.White
        Me.rbnt_dir_perp.Location = New System.Drawing.Point(174, 8)
        Me.rbnt_dir_perp.Name = "rbnt_dir_perp"
        Me.rbnt_dir_perp.Size = New System.Drawing.Size(135, 17)
        Me.rbnt_dir_perp.TabIndex = 6
        Me.rbnt_dir_perp.Text = "division perpendiculaire"
        Me.rbnt_dir_perp.UseVisualStyleBackColor = True
        '
        'rbnt_dir_par
        '
        Me.rbnt_dir_par.AutoSize = True
        Me.rbnt_dir_par.ForeColor = System.Drawing.Color.White
        Me.rbnt_dir_par.Location = New System.Drawing.Point(20, 8)
        Me.rbnt_dir_par.Name = "rbnt_dir_par"
        Me.rbnt_dir_par.Size = New System.Drawing.Size(102, 17)
        Me.rbnt_dir_par.TabIndex = 5
        Me.rbnt_dir_par.Text = "division parallèle"
        Me.rbnt_dir_par.UseVisualStyleBackColor = True
        '
        'lbl_orientation_dir
        '
        Me.lbl_orientation_dir.AutoSize = True
        Me.lbl_orientation_dir.ForeColor = System.Drawing.Color.White
        Me.lbl_orientation_dir.Location = New System.Drawing.Point(146, 37)
        Me.lbl_orientation_dir.Name = "lbl_orientation_dir"
        Me.lbl_orientation_dir.Size = New System.Drawing.Size(238, 13)
        Me.lbl_orientation_dir.TabIndex = 2
        Me.lbl_orientation_dir.Text = "Sélection de l'orientation et du sens de la division"
        '
        'bnt_orientation_dir
        '
        Me.bnt_orientation_dir.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app_active
        Me.bnt_orientation_dir.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.bnt_orientation_dir.Enabled = False
        Me.bnt_orientation_dir.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.bnt_orientation_dir.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.bnt_orientation_dir.Image = Global.Revo.My.Resources.Resources._select
        Me.bnt_orientation_dir.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.bnt_orientation_dir.Location = New System.Drawing.Point(3, 31)
        Me.bnt_orientation_dir.Name = "bnt_orientation_dir"
        Me.bnt_orientation_dir.Size = New System.Drawing.Size(136, 24)
        Me.bnt_orientation_dir.TabIndex = 1
        Me.bnt_orientation_dir.Text = "Sélectionner à l'écran"
        Me.bnt_orientation_dir.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.bnt_orientation_dir.UseVisualStyleBackColor = True
        '
        'rbtn_methode_pt_fixe
        '
        Me.rbtn_methode_pt_fixe.AutoSize = True
        Me.rbtn_methode_pt_fixe.Enabled = False
        Me.rbtn_methode_pt_fixe.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.rbtn_methode_pt_fixe.Location = New System.Drawing.Point(225, 77)
        Me.rbtn_methode_pt_fixe.Name = "rbtn_methode_pt_fixe"
        Me.rbtn_methode_pt_fixe.Size = New System.Drawing.Size(161, 17)
        Me.rbtn_methode_pt_fixe.TabIndex = 10
        Me.rbtn_methode_pt_fixe.Text = "Division passant par un point"
        Me.rbtn_methode_pt_fixe.UseVisualStyleBackColor = True
        '
        'rbtn_methode_direction
        '
        Me.rbtn_methode_direction.AutoSize = True
        Me.rbtn_methode_direction.Enabled = False
        Me.rbtn_methode_direction.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.rbtn_methode_direction.Location = New System.Drawing.Point(225, 17)
        Me.rbtn_methode_direction.Name = "rbtn_methode_direction"
        Me.rbtn_methode_direction.Size = New System.Drawing.Size(163, 17)
        Me.rbtn_methode_direction.TabIndex = 8
        Me.rbtn_methode_direction.Text = "Division suivant une direction"
        Me.rbtn_methode_direction.UseVisualStyleBackColor = True
        '
        'lbl_methode
        '
        Me.lbl_methode.AutoSize = True
        Me.lbl_methode.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_methode.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.lbl_methode.Location = New System.Drawing.Point(19, 48)
        Me.lbl_methode.Name = "lbl_methode"
        Me.lbl_methode.Size = New System.Drawing.Size(201, 14)
        Me.lbl_methode.TabIndex = 7
        Me.lbl_methode.Text = "Sélectionner la méthode de division"
        '
        'pan_selection_parcelle
        '
        Me.pan_selection_parcelle.Controls.Add(Me.lbl_selection_parcelle)
        Me.pan_selection_parcelle.Controls.Add(Me.btn_selection_parcelle)
        Me.pan_selection_parcelle.Location = New System.Drawing.Point(9, 18)
        Me.pan_selection_parcelle.Name = "pan_selection_parcelle"
        Me.pan_selection_parcelle.Size = New System.Drawing.Size(538, 42)
        Me.pan_selection_parcelle.TabIndex = 7
        '
        'lbl_selection_parcelle
        '
        Me.lbl_selection_parcelle.AutoSize = True
        Me.lbl_selection_parcelle.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_selection_parcelle.ForeColor = System.Drawing.Color.White
        Me.lbl_selection_parcelle.Location = New System.Drawing.Point(19, 19)
        Me.lbl_selection_parcelle.Name = "lbl_selection_parcelle"
        Me.lbl_selection_parcelle.Size = New System.Drawing.Size(196, 14)
        Me.lbl_selection_parcelle.TabIndex = 1
        Me.lbl_selection_parcelle.Text = "Sélectionner le bien-fonds à diviser"
        '
        'btn_selection_parcelle
        '
        Me.btn_selection_parcelle.BackgroundImage = Global.Revo.My.Resources.Resources.bouton_revo_app_active
        Me.btn_selection_parcelle.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btn_selection_parcelle.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.btn_selection_parcelle.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btn_selection_parcelle.Image = Global.Revo.My.Resources.Resources._select
        Me.btn_selection_parcelle.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.btn_selection_parcelle.Location = New System.Drawing.Point(227, 14)
        Me.btn_selection_parcelle.Name = "btn_selection_parcelle"
        Me.btn_selection_parcelle.Size = New System.Drawing.Size(136, 24)
        Me.btn_selection_parcelle.TabIndex = 0
        Me.btn_selection_parcelle.Text = "Sélectionner à l'écran"
        Me.btn_selection_parcelle.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.btn_selection_parcelle.UseVisualStyleBackColor = True
        '
        'pan_N_parcelle
        '
        Me.pan_N_parcelle.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan_N_parcelle.Controls.Add(Me.tbox_N_parcelle)
        Me.pan_N_parcelle.Controls.Add(Me.lbl_N_parcelle_1)
        Me.pan_N_parcelle.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pan_N_parcelle.Location = New System.Drawing.Point(3, 210)
        Me.pan_N_parcelle.Name = "pan_N_parcelle"
        Me.pan_N_parcelle.Size = New System.Drawing.Size(805, 54)
        Me.pan_N_parcelle.TabIndex = 12
        '
        'tbox_N_parcelle
        '
        Me.tbox_N_parcelle.Enabled = False
        Me.tbox_N_parcelle.Location = New System.Drawing.Point(258, 14)
        Me.tbox_N_parcelle.Name = "tbox_N_parcelle"
        Me.tbox_N_parcelle.Size = New System.Drawing.Size(33, 20)
        Me.tbox_N_parcelle.TabIndex = 4
        '
        'lbl_N_parcelle_1
        '
        Me.lbl_N_parcelle_1.AutoSize = True
        Me.lbl_N_parcelle_1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_N_parcelle_1.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.lbl_N_parcelle_1.Location = New System.Drawing.Point(33, 16)
        Me.lbl_N_parcelle_1.Name = "lbl_N_parcelle_1"
        Me.lbl_N_parcelle_1.Size = New System.Drawing.Size(163, 14)
        Me.lbl_N_parcelle_1.TabIndex = 3
        Me.lbl_N_parcelle_1.Text = "Nombre de parcelle à créer :"
        '
        'pan_calcul_surf
        '
        Me.pan_calcul_surf.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.pan_calcul_surf.Controls.Add(Me.lbl_surf_reste)
        Me.pan_calcul_surf.Controls.Add(Me.DataGridView_suface)
        Me.pan_calcul_surf.Controls.Add(Me.rbtn_calcul_surface_choisie)
        Me.pan_calcul_surf.Controls.Add(Me.rbtn_calcul_surface_egale)
        Me.pan_calcul_surf.Controls.Add(Me.lbl_calcul_surf_1)
        Me.pan_calcul_surf.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pan_calcul_surf.Location = New System.Drawing.Point(3, 270)
        Me.pan_calcul_surf.Name = "pan_calcul_surf"
        Me.pan_calcul_surf.Size = New System.Drawing.Size(805, 144)
        Me.pan_calcul_surf.TabIndex = 13
        '
        'lbl_surf_reste
        '
        Me.lbl_surf_reste.AutoSize = True
        Me.lbl_surf_reste.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_surf_reste.ForeColor = System.Drawing.Color.White
        Me.lbl_surf_reste.Location = New System.Drawing.Point(288, 102)
        Me.lbl_surf_reste.Name = "lbl_surf_reste"
        Me.lbl_surf_reste.Size = New System.Drawing.Size(180, 14)
        Me.lbl_surf_reste.TabIndex = 13
        Me.lbl_surf_reste.Text = "Surface de la dernière parcelle :"
        Me.lbl_surf_reste.Visible = False
        '
        'DataGridView_suface
        '
        Me.DataGridView_suface.AllowUserToAddRows = False
        Me.DataGridView_suface.AllowUserToDeleteRows = False
        Me.DataGridView_suface.BackgroundColor = System.Drawing.Color.DimGray
        Me.DataGridView_suface.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGridView_suface.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView_suface.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.Col_surface})
        Me.DataGridView_suface.Location = New System.Drawing.Point(622, 3)
        Me.DataGridView_suface.Name = "DataGridView_suface"
        Me.DataGridView_suface.RowHeadersVisible = False
        Me.DataGridView_suface.Size = New System.Drawing.Size(101, 136)
        Me.DataGridView_suface.TabIndex = 12
        Me.DataGridView_suface.Visible = False
        '
        'Col_surface
        '
        Me.Col_surface.HeaderText = "Surface(s)"
        Me.Col_surface.Name = "Col_surface"
        '
        'rbtn_calcul_surface_choisie
        '
        Me.rbtn_calcul_surface_choisie.AutoSize = True
        Me.rbtn_calcul_surface_choisie.Enabled = False
        Me.rbtn_calcul_surface_choisie.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.rbtn_calcul_surface_choisie.Location = New System.Drawing.Point(291, 59)
        Me.rbtn_calcul_surface_choisie.Name = "rbtn_calcul_surface_choisie"
        Me.rbtn_calcul_surface_choisie.Size = New System.Drawing.Size(183, 17)
        Me.rbtn_calcul_surface_choisie.TabIndex = 11
        Me.rbtn_calcul_surface_choisie.Text = "Surfaces imposées par l'utilisateur"
        Me.rbtn_calcul_surface_choisie.UseVisualStyleBackColor = True
        '
        'rbtn_calcul_surface_egale
        '
        Me.rbtn_calcul_surface_egale.AutoSize = True
        Me.rbtn_calcul_surface_egale.Enabled = False
        Me.rbtn_calcul_surface_egale.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.rbtn_calcul_surface_egale.Location = New System.Drawing.Point(291, 22)
        Me.rbtn_calcul_surface_egale.Name = "rbtn_calcul_surface_egale"
        Me.rbtn_calcul_surface_egale.Size = New System.Drawing.Size(101, 17)
        Me.rbtn_calcul_surface_egale.TabIndex = 10
        Me.rbtn_calcul_surface_egale.Text = "Surfaces égales"
        Me.rbtn_calcul_surface_egale.UseVisualStyleBackColor = True
        '
        'lbl_calcul_surf_1
        '
        Me.lbl_calcul_surf_1.AutoSize = True
        Me.lbl_calcul_surf_1.Font = New System.Drawing.Font("Tahoma", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_calcul_surf_1.ForeColor = System.Drawing.SystemColors.AppWorkspace
        Me.lbl_calcul_surf_1.Location = New System.Drawing.Point(33, 25)
        Me.lbl_calcul_surf_1.Name = "lbl_calcul_surf_1"
        Me.lbl_calcul_surf_1.Size = New System.Drawing.Size(227, 14)
        Me.lbl_calcul_surf_1.TabIndex = 4
        Me.lbl_calcul_surf_1.Text = "Sélectionner le mode calcul des surfaces"
        '
        'ToolTip_selection_orientation
        '
        Me.ToolTip_selection_orientation.Tag = "Vous allez sélectionnerdeux points donnant l'orientation de la division. Ensuite " & _
    "vous pourrez choisir de quel côté la première division doit être faite"
        '
        'chkLayerMO
        '
        Me.chkLayerMO.AutoSize = True
        Me.chkLayerMO.ForeColor = System.Drawing.Color.White
        Me.chkLayerMO.Location = New System.Drawing.Point(545, 15)
        Me.chkLayerMO.Name = "chkLayerMO"
        Me.chkLayerMO.Size = New System.Drawing.Size(148, 17)
        Me.chkLayerMO.TabIndex = 16
        Me.chkLayerMO.Text = "Activer le calque mutation"
        Me.chkLayerMO.UseVisualStyleBackColor = True
        '
        'frmDivSurf
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.DimGray
        Me.ClientSize = New System.Drawing.Size(823, 486)
        Me.Controls.Add(Me.TableLayout)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmDivSurf"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Division de parcelle (version bêta)"
        Me.TableLayout.ResumeLayout(False)
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.pan_enregistrement.ResumeLayout(False)
        Me.pan_enregistrement.PerformLayout()
        Me.pan_init.ResumeLayout(False)
        Me.pan_methode.ResumeLayout(False)
        Me.pan_methode.PerformLayout()
        Me.pan_orientation_pt.ResumeLayout(False)
        Me.pan_orientation_pt.PerformLayout()
        Me.pan_orientation_dir.ResumeLayout(False)
        Me.pan_orientation_dir.PerformLayout()
        Me.pan_selection_parcelle.ResumeLayout(False)
        Me.pan_selection_parcelle.PerformLayout()
        Me.pan_N_parcelle.ResumeLayout(False)
        Me.pan_N_parcelle.PerformLayout()
        Me.pan_calcul_surf.ResumeLayout(False)
        Me.pan_calcul_surf.PerformLayout()
        CType(Me.DataGridView_suface, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents TableLayout As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents pan_init As System.Windows.Forms.Panel
    Friend WithEvents pan_N_parcelle As System.Windows.Forms.Panel
    Friend WithEvents ToolTip_selection_orientation As System.Windows.Forms.ToolTip
    Friend WithEvents tbox_N_parcelle As System.Windows.Forms.TextBox
    Friend WithEvents lbl_N_parcelle_1 As System.Windows.Forms.Label
    Friend WithEvents pan_methode As System.Windows.Forms.Panel
    Friend WithEvents pan_orientation_dir As System.Windows.Forms.Panel
    Friend WithEvents lbl_orientation_dir As System.Windows.Forms.Label
    Friend WithEvents bnt_orientation_dir As System.Windows.Forms.Button
    Friend WithEvents rbtn_methode_pt_fixe As System.Windows.Forms.RadioButton
    Friend WithEvents rbtn_methode_direction As System.Windows.Forms.RadioButton
    Friend WithEvents lbl_methode As System.Windows.Forms.Label
    Friend WithEvents pan_selection_parcelle As System.Windows.Forms.Panel
    Friend WithEvents lbl_selection_parcelle As System.Windows.Forms.Label
    Friend WithEvents btn_selection_parcelle As System.Windows.Forms.Button
    Friend WithEvents rbnt_dir_perp As System.Windows.Forms.RadioButton
    Friend WithEvents rbnt_dir_par As System.Windows.Forms.RadioButton
    Friend WithEvents pan_orientation_pt As System.Windows.Forms.Panel
    Friend WithEvents lbl_orientation_pt As System.Windows.Forms.Label
    Friend WithEvents bnt_orientation_pt As System.Windows.Forms.Button
    Friend WithEvents pan_calcul_surf As System.Windows.Forms.Panel
    Friend WithEvents DataGridView_suface As System.Windows.Forms.DataGridView
    Friend WithEvents rbtn_calcul_surface_choisie As System.Windows.Forms.RadioButton
    Friend WithEvents rbtn_calcul_surface_egale As System.Windows.Forms.RadioButton
    Friend WithEvents lbl_calcul_surf_1 As System.Windows.Forms.Label
    Friend WithEvents pan_enregistrement As System.Windows.Forms.Panel
    Friend WithEvents CBox_enregistrement As System.Windows.Forms.CheckBox
    Friend WithEvents lbl_enregistrement As System.Windows.Forms.Label
    Friend WithEvents btn_enregistrement As System.Windows.Forms.Button
    Friend WithEvents tbox_enregistrement As System.Windows.Forms.TextBox
    Friend WithEvents lbl_surf_reste As System.Windows.Forms.Label
    Friend WithEvents Col_surface As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents btn_calcul As System.Windows.Forms.Button
    Friend WithEvents chkLayerMO As System.Windows.Forms.CheckBox
End Class
