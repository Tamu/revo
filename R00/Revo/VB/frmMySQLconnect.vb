
Imports System
Imports System.Data
Imports System.Windows.Forms
Imports MySql.Data.MySqlClient


Public Class MySqlConnect

    Inherits System.Windows.Forms.Form
    Dim conn As MySqlConnection
    Dim data As DataTable
    Dim da As MySqlDataAdapter
    Dim cb As MySqlCommandBuilder



#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call

    End Sub

    'Form overrides dispose to clean up the component list.
    Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
        If disposing Then
            If Not (components Is Nothing) Then
                components.Dispose()
            End If
        End If
        MyBase.Dispose(disposing)
    End Sub

    
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MySqlConnect))
        Me.databaseList = New System.Windows.Forms.ComboBox()
        Me.label5 = New System.Windows.Forms.Label()
        Me.dataGrid = New System.Windows.Forms.DataGrid()
        Me.tables = New System.Windows.Forms.ComboBox()
        Me.password = New System.Windows.Forms.TextBox()
        Me.label3 = New System.Windows.Forms.Label()
        Me.userid = New System.Windows.Forms.TextBox()
        Me.label2 = New System.Windows.Forms.Label()
        Me.server = New System.Windows.Forms.TextBox()
        Me.label1 = New System.Windows.Forms.Label()
        Me.label4 = New System.Windows.Forms.Label()
        Me.connectBtn = New System.Windows.Forms.Button()
        Me.updateBtn = New System.Windows.Forms.Button()
        Me.Port = New System.Windows.Forms.TextBox()
        CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'databaseList
        '
        Me.databaseList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.databaseList.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.databaseList.Location = New System.Drawing.Point(80, 76)
        Me.databaseList.Name = "databaseList"
        Me.databaseList.Size = New System.Drawing.Size(296, 24)
        Me.databaseList.TabIndex = 24
        '
        'label5
        '
        Me.label5.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.label5.ForeColor = System.Drawing.Color.White
        Me.label5.Location = New System.Drawing.Point(8, 80)
        Me.label5.Name = "label5"
        Me.label5.Size = New System.Drawing.Size(64, 16)
        Me.label5.TabIndex = 23
        Me.label5.Text = "Databases"
        '
        'dataGrid
        '
        Me.dataGrid.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dataGrid.DataMember = ""
        Me.dataGrid.HeaderForeColor = System.Drawing.SystemColors.ControlText
        Me.dataGrid.Location = New System.Drawing.Point(8, 136)
        Me.dataGrid.Name = "dataGrid"
        Me.dataGrid.Size = New System.Drawing.Size(450, 384)
        Me.dataGrid.TabIndex = 21
        '
        'tables
        '
        Me.tables.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.tables.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.tables.Location = New System.Drawing.Point(80, 104)
        Me.tables.Name = "tables"
        Me.tables.Size = New System.Drawing.Size(296, 24)
        Me.tables.TabIndex = 20
        '
        'password
        '
        Me.password.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.password.Location = New System.Drawing.Point(264, 37)
        Me.password.Name = "password"
        Me.password.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.password.Size = New System.Drawing.Size(112, 22)
        Me.password.TabIndex = 18
        '
        'label3
        '
        Me.label3.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.label3.ForeColor = System.Drawing.Color.White
        Me.label3.Location = New System.Drawing.Point(192, 39)
        Me.label3.Name = "label3"
        Me.label3.Size = New System.Drawing.Size(68, 18)
        Me.label3.TabIndex = 17
        Me.label3.Text = "Password"
        '
        'userid
        '
        Me.userid.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.userid.Location = New System.Drawing.Point(56, 37)
        Me.userid.Name = "userid"
        Me.userid.Size = New System.Drawing.Size(120, 22)
        Me.userid.TabIndex = 16
        '
        'label2
        '
        Me.label2.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.label2.ForeColor = System.Drawing.Color.White
        Me.label2.Location = New System.Drawing.Point(8, 39)
        Me.label2.Name = "label2"
        Me.label2.Size = New System.Drawing.Size(42, 16)
        Me.label2.TabIndex = 15
        Me.label2.Text = "User"
        '
        'server
        '
        Me.server.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.server.Location = New System.Drawing.Point(56, 8)
        Me.server.Name = "server"
        Me.server.Size = New System.Drawing.Size(202, 22)
        Me.server.TabIndex = 14
        '
        'label1
        '
        Me.label1.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.label1.ForeColor = System.Drawing.Color.White
        Me.label1.Location = New System.Drawing.Point(8, 10)
        Me.label1.Name = "label1"
        Me.label1.Size = New System.Drawing.Size(48, 16)
        Me.label1.TabIndex = 12
        Me.label1.Text = "Server"
        '
        'label4
        '
        Me.label4.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.label4.ForeColor = System.Drawing.Color.White
        Me.label4.Location = New System.Drawing.Point(8, 107)
        Me.label4.Name = "label4"
        Me.label4.Size = New System.Drawing.Size(64, 16)
        Me.label4.TabIndex = 13
        Me.label4.Text = "Tables"
        '
        'connectBtn
        '
        Me.connectBtn.BackColor = System.Drawing.Color.LightGray
        Me.connectBtn.BackgroundImage = CType(resources.GetObject("connectBtn.BackgroundImage"), System.Drawing.Image)
        Me.connectBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.connectBtn.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.connectBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.connectBtn.Location = New System.Drawing.Point(391, 8)
        Me.connectBtn.Name = "connectBtn"
        Me.connectBtn.Size = New System.Drawing.Size(67, 23)
        Me.connectBtn.TabIndex = 25
        Me.connectBtn.Text = "Connect"
        Me.connectBtn.UseVisualStyleBackColor = False
        '
        'updateBtn
        '
        Me.updateBtn.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.updateBtn.BackColor = System.Drawing.Color.LightGray
        Me.updateBtn.BackgroundImage = CType(resources.GetObject("updateBtn.BackgroundImage"), System.Drawing.Image)
        Me.updateBtn.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.updateBtn.FlatAppearance.BorderColor = System.Drawing.Color.Gray
        Me.updateBtn.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.updateBtn.Location = New System.Drawing.Point(391, 104)
        Me.updateBtn.Name = "updateBtn"
        Me.updateBtn.Size = New System.Drawing.Size(67, 23)
        Me.updateBtn.TabIndex = 26
        Me.updateBtn.Text = "Update"
        Me.updateBtn.UseVisualStyleBackColor = False
        '
        'Port
        '
        Me.Port.Font = New System.Drawing.Font("Arial", 9.75!)
        Me.Port.Location = New System.Drawing.Point(264, 7)
        Me.Port.Name = "Port"
        Me.Port.Size = New System.Drawing.Size(112, 22)
        Me.Port.TabIndex = 27
        Me.Port.Text = "3306"
        '
        'MySqlConnect
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.BackColor = System.Drawing.Color.Gray
        Me.ClientSize = New System.Drawing.Size(466, 525)
        Me.Controls.Add(Me.Port)
        Me.Controls.Add(Me.updateBtn)
        Me.Controls.Add(Me.connectBtn)
        Me.Controls.Add(Me.databaseList)
        Me.Controls.Add(Me.label5)
        Me.Controls.Add(Me.dataGrid)
        Me.Controls.Add(Me.tables)
        Me.Controls.Add(Me.password)
        Me.Controls.Add(Me.label3)
        Me.Controls.Add(Me.userid)
        Me.Controls.Add(Me.label2)
        Me.Controls.Add(Me.server)
        Me.Controls.Add(Me.label1)
        Me.Controls.Add(Me.label4)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MySqlConnect"
        Me.Text = "MySQL"
        CType(Me.dataGrid, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

#End Region

    Private Sub MySqlConnect_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim ass As New Revo.RevoInfo
        Me.Icon = Ass.Icon
    End Sub
    Private Sub connectBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles connectBtn.Click
        If Not conn Is Nothing Then conn.Close()

        Dim PortDefault As String = 3306
        If Port.Text = "" Or IsNumeric(Port.Text) = False Then Port.Text = PortDefault

        Dim connStr As String
        connStr = String.Format("server={0}; Port={1}; user id={2}; password={3}; database=mysql; pooling=false", _
    server.Text, Port.Text, userid.Text, password.Text)

        Try
            conn = New MySqlConnection(connStr)
            conn.Open()

            GetDatabases()
        Catch ex As MySqlException
            MessageBox.Show("Error connecting to the server: " + ex.Message)
        End Try
    End Sub

    Private Sub GetDatabases()
        Dim reader As MySqlDataReader
        reader = Nothing

        Dim cmd As New MySqlCommand("SHOW DATABASES", conn)
        Try
            reader = cmd.ExecuteReader()
            databaseList.Items.Clear()

            While (reader.Read())
                databaseList.Items.Add(reader.GetString(0))
            End While
        Catch ex As MySqlException
            MessageBox.Show("Failed to populate database list: " + ex.Message)
        Finally
            If Not reader Is Nothing Then reader.Close()
        End Try

    End Sub

    Private Sub databaseList_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles databaseList.SelectedIndexChanged
        Dim reader As MySqlDataReader = Nothing

        conn.ChangeDatabase(databaseList.SelectedItem.ToString())

        Dim cmd As New MySqlCommand("SHOW TABLES", conn)

        Try
            reader = cmd.ExecuteReader()
            tables.Items.Clear()

            While (reader.Read())
                tables.Items.Add(reader.GetString(0))
            End While

        Catch ex As MySqlException
            MessageBox.Show("Failed to populate table list: " + ex.Message)
        Finally
            If Not reader Is Nothing Then reader.Close()
        End Try
    End Sub

    Private Sub tables_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles tables.SelectedIndexChanged
        data = New DataTable

        da = New MySqlDataAdapter("SELECT * FROM " + tables.SelectedItem.ToString(), conn)
        cb = New MySqlCommandBuilder(da)
        Try
            da.Fill(data)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        dataGrid.DataSource = data
    End Sub

    
   
    Private Sub updateBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles updateBtn.Click
        Dim changes As DataTable = data.GetChanges()
        da.Update(changes)
        data.AcceptChanges()
    End Sub

    
 
End Class