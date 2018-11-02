Imports System.Data
Imports frms = System.Windows.Forms
Module modDBFunctions
    'Public objDb As dao.Database
    'Public rs As dao.Recordset
    'Public ds As System.Data.DataSet
    'Public dbCon As OleDb.OleDbConnection = Nothing
    Public MainDS As System.Data.DataSet = Nothing
    Public da As OleDb.OleDbDataAdapter = Nothing
    Dim builder As OleDb.OleDbCommandBuilder = Nothing
    Public dbCon As SQLite.SQLiteConnection = Nothing
    Public Const DBPwd As String = "chien-enrage-79"
    Dim Ass As New Revo.RevoInfo ' include SQLite in Revo THA
    Dim DB3interlis As String = Ass.DB3interlis



    ''' <summary>
    ''' Deletes the data from all of the tables in the dataset except for the tables whose names begin with "_"
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearTables()
        If dbCon Is Nothing Then
            dbCon = SQLDBCon(Nothing)
        End If
        If Not dbCon Is Nothing Then
            Dim cmd As SQLite.SQLiteCommand = dbCon.CreateCommand()
            cmd.Transaction = dbCon.BeginTransaction
            Try
                Using tbl As System.Data.DataTable = dbCon.GetSchema("Tables")
                    For Each r As DataRow In tbl.Rows
                        If Not r(2).ToString.StartsWith("_") Then
                             cmd.CommandText = "DELETE FROM " & r(2).ToString
                            cmd.ExecuteNonQuery()
                        End If
                    Next
                End Using
                cmd.Transaction.Commit()
            Catch ex As Exception
                frms.MessageBox.Show(ex.Message, "Clear Tables", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Finally
                ClearSQLCommand(cmd)
            End Try
        End If

    End Sub


    ''' <summary>
    ''' Creates the connection to the SQLite database
    ''' </summary>
    ''' <param name="dbName">The name of the database to connect to</param>
    ''' <returns>The SQLite connection</returns>
    ''' <remarks></remarks>
    Public Function SQLDBCon(Optional ByVal dbName As String = Nothing) As SQLite.SQLiteConnection ' include SQLite in Revo THA

        Dim DBPath As String = Ass.SystemPath ' include SQLite in Revo THA
        If dbName Is Nothing Then dbName = DB3interlis '"Interlis.db3" ' include SQLite in Revo THA

        Dim SQLCon As New SQLite.SQLiteConnection
        Try
            SQLCon.ConnectionString = String.Format("Data Source = {0}{1}; Version=3; Password={2};", DBPath, dbName, DBPwd) ' include SQLite in Revo THA
            Dim cmd As New SQLite.SQLiteCommand
            SQLCon.Open()
            cmd.Connection = SQLCon
            cmd.CommandText = "PRAGMA temp_store = 2;"
            cmd.ExecuteNonQuery()
            cmd.CommandText = "PRAGMA cache_size = 10000;"
            cmd.ExecuteNonQuery()
            Return SQLCon
        Catch ex As Exception
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Gets the data from the database for the desired table
    ''' </summary>
    ''' <param name="TableName">The table name to get the data from</param>
    ''' <param name="SQL">Optionally the SQL string to use to select the data</param>
    ''' <param name="dbName">Optionally the name of the database to get the data from</param>
    ''' <returns>A datatable of the data</returns>
    ''' <remarks></remarks>
    Public Function GetTable(ByVal TableName As String, Optional ByVal SQL As String = Nothing, Optional ByVal dbName As String = Nothing) As System.Data.DataTable
        'Check if the Main dataset exists
        If MainDS Is Nothing Then MainDS = New System.Data.DataSet
        'Declare the return datatable
        Dim retDt As System.Data.DataTable = Nothing
        Try
            'Try to get the table from the main dataset first
            retDt = MainDS.Tables(TableName)
        Catch ex As Exception
            ' MsgBox("ModDBFunctions / GetTable : " & ex.Message)
        End Try
        'If the table was not in the main dataset 
        If retDt Is Nothing Then
            'Then get the data from the database
            retDt = GetRecordset(TableName, SQL, dbName)
        End If
        Return retDt
    End Function

    ''' <summary>
    ''' Gets the data from the database for the desired table
    ''' </summary>
    ''' <param name="TableName">The table name to get the data from</param>
    ''' <param name="SQL">Optionally the SQL string to use to select the data</param>
    ''' <param name="dbName">Optionally the name of the database to get the data from</param>
    ''' <returns>A datatable of the data</returns>
    ''' <remarks></remarks>
    Public Function GetQueryResult(ByVal TableName As String, Optional ByVal SQL As String = Nothing, Optional ByVal dbName As String = Nothing) As System.Data.DataTable
        Dim retDt As System.Data.DataTable = GetRecordset(TableName, SQL, dbName)
        Return retDt
    End Function

    ''' <summary>
    ''' Gets the desired recordset data from the SQLite database
    ''' </summary>
    ''' <param name="TableName">The name of the table to get the data from</param>
    ''' <param name="SQL">Optional SQL query string to use to select the data (defaults to SELECT * FROM Tablename)</param>
    ''' <param name="dbName">Optional database name to get the data from</param>
    ''' <returns>A datatable of the data</returns>
    ''' <remarks></remarks>
    Public Function GetRecordset(ByVal TableName As String, Optional ByVal SQL As String = Nothing, Optional ByVal dbName As String = Nothing) As System.Data.DataTable
        Dim dSet As New System.Data.DataSet 'Temp dataset

        'Check the state of the connection
        If dbCon Is Nothing Or Not dbName Is Nothing Then
            'Get the connection if it is not already connected 
            dbCon = SQLDBCon(dbName)
        End If
        Try
            'Check that we've got a connection now
            If Not dbCon Is Nothing Then
                'If the SQL string is nothing then set it to the default
                If SQL Is Nothing Then SQL = String.Format("SELECT * FROM {0}", TableName)
                'Create the dataadapter
                Dim da As SQLite.SQLiteDataAdapter = New SQLite.SQLiteDataAdapter(SQL, dbCon)
                'Fill the dataset
                da.Fill(dSet, TableName)
            End If
            'We only want to add this table to the main dataset if it's one of the Interlis tables
            'So check first that dbname is nothing (i.e. we're connecting to the default Interlis database)
            If dbName Is Nothing Then
                'Copy the returned data into the main dataset
                MainDS.Tables.Add(dSet.Tables(0).Copy)
                'Return that datatable
                Return MainDS.Tables(TableName)
            Else
                'dbName is not nothing we we just want to return the datatable
                Return dSet.Tables(0)
            End If
        Catch ex As Exception
            conn.Message(Ass.xTitle, "Erreur de lecture de la BD :" & vbCrLf & ex.Message, False, 50, 100) ' include SQLite in Revo THA)
            ' frms.MessageBox.Show(ex.Message)
            Return Nothing
        End Try
    End Function

    Public Function GetRecordsetUsingWrapper(ByVal TableName As String, Optional ByVal SQL As String = Nothing, Optional ByVal dbName As String = Nothing) As System.Data.DataTable
        Try
            ClearDBPassword()
            Dim db As New SQLiteWrapper.SQLiteBase
            'Dim DBPath As String = Ass.SystemPath ' include SQLite in Revo THA
            Dim strPath As String = System.IO.Path.Combine(Ass.SystemPath, dbName) ' String.Format("{0}{1}", DBPath, dbName)
            'Dim strPath As String = String.Format("{0}{1}", DBPath, dbName)
            db.OpenDatabase(strPath)
            Dim dt As System.Data.DataTable = db.ExecuteQuery(SQL)
            Return dt
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message)
            Return Nothing
        Finally
            SetDBPassword()
        End Try

    End Function

    ''' <summary>
    ''' Writes the data in the dataset to the SQLite database
    ''' </summary>
    ''' <returns>True if sucessful</returns>
    ''' <remarks></remarks>
    Public Function UpdateData() As Boolean
        If dbCon Is Nothing Then
            dbCon = SQLDBCon(Nothing)
        End If
        Try
            If Not dbCon Is Nothing Then
                Using t As SQLite.SQLiteTransaction = dbCon.BeginTransaction()
                    For Each dt As DataTable In MainDS.Tables
                        Dim da As SQLite.SQLiteDataAdapter = New SQLite.SQLiteDataAdapter(String.Format("SELECT * FROM {0}", dt.TableName), dbCon)
                        Dim sb As SQLite.SQLiteCommandBuilder = New SQLite.SQLiteCommandBuilder(da)
                        da.Update(MainDS, dt.TableName.ToString)
                    Next
                    t.Commit()
                End Using
            End If
            Return True
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "Update Data", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Clear the SQLite command
    ''' </summary>
    ''' <param name="cmd">The SQLite command to clear</param>
    ''' <remarks></remarks>
    Private Sub ClearSQLCommand(ByVal cmd As SQLite.SQLiteCommand)
        If Not cmd Is Nothing Then
            Try
                cmd.Dispose()
            Catch ex As Exception
            End Try
        End If
    End Sub

    ''' <summary>
    ''' Reads the queries from the XML file
    ''' </summary>
    ''' <returns>A datatable of the queries to run</returns>
    ''' <remarks></remarks>
    Public Function GetQueries() As System.Data.DataTable
        Dim dsQs As New System.Data.DataSet
        Dim DBPath As String = Ass.SystemPath ' include SQLite in Revo THA

        Try
            dsQs.ReadXml(String.Format("{0}\Queries.xml", DBPath)) ' include SQLite in Revo THA
            Return dsQs.Tables(0)
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "Get Queries", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Used to set the password on the SQLite database
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub SetDBPassword()
        Dim con As New SQLite.SQLiteConnection
        Dim DBPath As String = Ass.SystemPath ' include SQLite in Revo THA

        'con.ConnectionString = String.Format("Data Source = {0}\" & Ass.DB3interlis, DBPath) ' include SQLite in Revo THA
        con.ConnectionString = String.Format("Data Source = {0}" & Ass.DB3system, DBPath)

        con.Open()
        con.ChangePassword(DBPwd)
        con.Dispose()
    End Sub

    ''' <summary>
    ''' Used to set the password on the SQLite database
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ClearDBPassword()
        Dim con As New SQLite.SQLiteConnection

        Dim DBPath As String = Ass.SystemPath ' include SQLite in Revo THA
        'con.ConnectionString = String.Format("Data Source = {0}\" & Ass.DB3interlis & "; PASSWORD={1}", DBPath, DBPwd) ' include SQLite in Revo THA
        con.ConnectionString = String.Format("Data Source = {0}" & Ass.DB3system & "; PASSWORD={1}", DBPath, DBPwd)

        con.Open()
        con.ChangePassword("")
        con.Dispose()
    End Sub


    ''' <summary>
    ''' Writes the data in the supplied table to the database
    ''' </summary>
    ''' <param name="dt">The datatable containing the data to write</param>
    ''' <param name="dbName">The name of the database to update</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function UpdateTable(ByVal dt As DataTable, ByVal dbName As String) As Boolean
        'Connect to the database
        dbCon = SQLDBCon(dbName)
        Try
            If Not dbCon Is Nothing Then
                'Update the table with the data from the supplied datatable
                Using t As SQLite.SQLiteTransaction = dbCon.BeginTransaction
                    Dim da As New SQLite.SQLiteDataAdapter(String.Format("SELECT * FROM {0}", dt.TableName), dbCon)
                    Dim cb As New SQLite.SQLiteCommandBuilder(da)
                    da.Update(dt)
                    t.Commit()
                End Using
            End If
            'Close the connection to the database
            dbCon.Close()
            dbCon = Nothing
            Return True
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "Update Table", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Deletes all of the data from the table name supplied
    ''' </summary>
    ''' <param name="TableName">The name of the table to delete the data from</param>
    ''' <remarks></remarks>
    Public Sub ClearTable(ByVal TableName As String)
        'Connect to the System Database
        Dim con As SQLite.SQLiteConnection = SQLDBCon(Ass.DB3system)
        If Not con Is Nothing Then
            Try
                Dim cmd As SQLite.SQLiteCommand = con.CreateCommand
                cmd.CommandText = String.Format("DELETE FROM {0}", TableName)
                cmd.ExecuteNonQuery()
            Catch ex As Exception
                frms.MessageBox.Show(ex.Message, "Clear Table", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Finally
                con.Close()
                con.Dispose()
            End Try
        End If
    End Sub


    Public Sub CloseDB() ' Deconnexion BD
        Try
            If dbCon IsNot Nothing Then 'Connexion de BD existante 
                dbCon.Close()
                dbCon.Dispose()
            End If
        Catch
        End Try
    End Sub
End Module
