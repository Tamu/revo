
Imports System
Imports System.Data
Imports System.Windows.Forms
Imports MySql.Data.MySqlClient ' Importation de la classe MySQL

Imports Autodesk.AutoCAD.ApplicationServices

Namespace Revo

    Public Class RevoMySQL

        Inherits System.Windows.Forms.Form
        Dim connBD As MySqlConnection
        Dim data As DataTable
        Dim da As MySqlDataAdapter
        Dim cb As MySqlCommandBuilder
        'creer un dataset
        Dim dts As DataSet
         'creer une instance de ma class perso pour les requete
        Dim db As New cl_mysql


        Private Function GetMySQLDB()
            Dim CollxBD As New Collection
            Dim reader As MySqlDataReader
            reader = Nothing

            Dim cmd As New MySqlCommand("SHOW DATABASES", connBD)
            Try
                reader = cmd.ExecuteReader()
                'databaseList.Items.Clear()
                CollxBD.Clear()

                While (reader.Read())
                    'databaseList.Items.Add(reader.GetString(0))
                    CollxBD.Add(reader.GetString(0))
                End While
            Catch ex As MySqlException
                MessageBox.Show("Failed to populate database list: " + ex.Message)
            Finally
                If Not reader Is Nothing Then reader.Close()
            End Try

            Return CollxBD
        End Function

        Public Function MySQLquery(ByVal Collprop)
            Dim RevoConnect As New connect
            'Dim Colldata As New Collection
            '      [[Propriété]] MySQL
            Dim ReqSQL As String = "" '    # SQL = "SELECT date(from_unixtime(ma_elem_techniques.MATDateLevee))  FROM ma_elem_techniques, ma_mandate WHERE ma_mandate.MADCode = 'CP07230' and ma_mandate.MADId = ma_elem_techniques.MATMandate"
            Dim URLsrv As String = "" '    # SQLurl = ""
            Dim Portsrv As String = "3306" ' Dans le SQLurl avec 00.00.00.00:8889
            Dim User As String = ""   '    # SQLuser = ""
            Dim Pass As String = ""   '    # SQLpass = ""
            Dim DBname As String = ""     '    # SQLdb = ""
            For Each Prop() As String In Collprop
                If Prop(0).ToLower = "sql" Then 'SQL = "SELECT ma_table...
                    ReqSQL = Prop(1)
                ElseIf Prop(0).ToLower = "sqlurl" Then 'SQLurl = ""
                    URLsrv = Prop(1)
                    If InStr(URLsrv, ":") <> 0 Then 'si port trouvé
                        Dim SplitURL() As String = Split(URLsrv, ":")
                        If SplitURL.Length = 2 Then
                            URLsrv = SplitURL(0)
                            If IsNumeric(SplitURL(1)) Then Portsrv = SplitURL(1)
                        End If
                    End If
                ElseIf Prop(0).ToLower = "sqluser" Then 'SQLuser = ""
                    User = Prop(1)
                ElseIf Prop(0).ToLower = "sqlpass" Then  'SQLpass = ""
                    Pass = Prop(1)
                ElseIf Prop(0).ToLower = "sqldb" Then 'SQLdb = ""
                    DBname = Prop(1)
                End If
            Next

            If ReqSQL <> "" And URLsrv <> "" And User <> "" And Pass <> "" And DBname <> "" Then
                Try
                    db.setvar(URLsrv, portsrv, User, Pass, DBname)
                    dts = db.mysqlquery(ReqSQL)
                    'Renvoi : dts.Tables(0)
                    Return dts
                Catch ex As Exception
                    RevoConnect.RevoLog(RevoConnect.DateLog & "MySQL" & vbTab & False & vbTab & "Error MySQL Host: " & ex.Message)
                End Try

            Else 'Element de connexion manquant
                MsgBox("Erreur de connexion à la base de donnée", MsgBoxStyle.Critical, "Connection MySQL")
                RevoConnect.RevoLog(RevoConnect.DateLog & "MySQL" & vbTab & False & vbTab & "Error MySQL URL/Pass:" & URLsrv)
            End If

            Return dts
            
        End Function

    End Class


    Public Class cl_mysql

        Dim HostConn As New MySqlConnection 'Pour une connexion a base de données MySQL
        Dim dta As MySqlDataAdapter 'Data adapter
        Dim dts As New DataSet 'Dataset
        Dim requete As String 'Chaine ou sera stocker les requetes
        Dim serv, port, user, pass, database, table As String


        'fonction d'enregistrement des variables user , pass ,serveur et eventuellement database 
        Function setvar(ByVal servv, ByVal portv, ByVal userv, ByVal passv, ByVal databasev)
            'on ferme la connection si jamais elle existe
            HostConn.Close()
            serv = servv
            port = portv
            user = userv
            pass = passv
            database = databasev
            'on attribut les variables de la nouvelle connection
            HostConn.ConnectionString = "server=" + serv + ";" _
                                 & "Port=" + port + ";" _
                                 & "user id=" + user + ";" _
                                 & "password=" + pass + ";" _
                                 & "database=" + database + ""
            Return 1
        End Function
        'charger la table
        Function loaddb(ByVal table As String)
            'on clear le dts si jammais il y a quelque chose
            dts.Clear()
            Try
                requete = "SELECT * FROM " + table + ""
                'on ouvre la connection
                HostConn.Open()
                'on execute la requete
                dta = New MySqlDataAdapter(requete, HostConn)
                'on rempli le dts avec la table demandée
                dta.Fill(dts, table)
                'on retourne le dts
                Return dts
                'si il y a une erreur
            Catch myerror As MySqlException
                'afficher le msg d'erreur
                MessageBox.Show("Error Connecting to Database: " & myerror.Message)
                Return 0
            Catch
                Return 0
            End Try

        End Function
        'fonction pour enregistrer la base modifier ( on envoi en paramettre la table en cours )
        Function savedb(ByVal data As DataSet, ByVal table As String)
            Try
                'commande de mise a jour de la base
                Dim MyCommBuild As New MySqlCommandBuilder(dta)
                dta.Update(data, table)
                Return 1
            Catch err As Exception
                'si erreur on affiche l'erreur et on charge la table de la base
                MsgBox(err.Message, MsgBoxStyle.Exclamation, "error")
                dts = New DataSet
                requete = "SELECT * FROM " + table + ""

                dta = New MySqlDataAdapter(requete, HostConn)
                dta.Fill(dts, table)
                Return dts
                Exit Function
            End Try
        End Function
        'fonction pour requete specifique
        Function mysqlquery(ByVal query As String)
            dts.Clear()
            Try
                requete = query
                HostConn.Open()
                'on execute la requete
                dta = New MySqlDataAdapter(requete, HostConn)
                dta.Fill(dts)
                Return dts
                'on renvoi la table
            Catch myerror As MySqlException

                MessageBox.Show("Error Connecting to Database: " & myerror.Message)
                Return 0
            End Try
        End Function
    End Class


End Namespace
