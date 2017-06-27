
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput

Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common

Imports System.Runtime.InteropServices

Imports Autodesk.AutoCAD.Colors
Imports frms = System.Windows.Forms
Imports Autodesk.AutoCAD
Imports Microsoft.VisualBasic
Imports System.IO
'Imports swisstopo


Namespace Revo

    Public Class RevoScript
        Dim Connect As New Revo.connect
        Dim Ass As New Revo.RevoInfo
        Dim WebForm As New frmWebbrowser 'Lancement du Nav Web
        'Dim Webbrowser2 As New System.Windows.Forms.WebBrowser 'Lancement du Nav Web
        'Dim WebbrowserState As String = "" 'Lancement du Nav Web

        Public CollAlignedDimension As New Collection
        Public CollArc As New Collection
        Public CollArcDimension As New Collection
        Public CollBlockReference As New Collection
        Public CollCircle As New Collection
        Public CollDBPoint As New Collection
        Public CollDBText As New Collection
        Public CollDiametricDimension As New Collection
        Public CollEllipse As New Collection
        Public CollOthers As New Collection
        Public CollAttribute As New Collection
        Public CollHatch As New Collection
        Public CollHelix As New Collection
        Public CollLineAngularDimension2 As New Collection
        Public CollLine As New Collection
        Public CollMText As New Collection
        Public CollPolyline3d As New Collection
        Public CollPolyline2d As New Collection
        Public CollPolyline As New Collection
        Public CollRadialDimension As New Collection
        Public CollRay As New Collection
        Public CollRotatedDimension As New Collection
        Public CollSpline As New Collection
        Public CollXline As New Collection
        ' Référence externe

        Public CollVar As New Collection 'Stockage de toutes les variables #Var
        Public CollTableHTML As New Collection 'Stockage des tableaux HTML

        Public CollLoad As Boolean = False
        Dim TimeCheck As New Timers.Timer
        Public Shared ParamImportExportFrm As New Collection

        Public Shared Cmds As New List(Of String)
        Private Shared Function Compare1(a As KeyValuePair(Of String, Integer), b As KeyValuePair(Of String, Integer)) As Integer
            Return a.Value.CompareTo(b.Value)
        End Function


        Public Function StartScript(ByVal FlowFiles As List(Of String), ByVal ImportFiles As List(Of String), Optional ByVal LicR As String = "")
            Dim i As Double = 0
            Dim z As Double = 0
            Dim ParamVar As New List(Of String)
            Dim ImportData As New List(Of String)
            Dim TestCmd As String = ""

            Cmds.Clear() 'Vider le tableau des commandes
            CollLoad = False 'Chargement des objets = null
            'Vide le tableau des objets
            CollAlignedDimension.Clear() : CollArc.Clear() : CollArcDimension.Clear() : CollBlockReference.Clear() : CollCircle.Clear() : CollDBPoint.Clear() : CollDBText.Clear() : CollDiametricDimension.Clear() : CollEllipse.Clear() : CollAttribute.Clear() : CollHatch.Clear() : CollHelix.Clear() : CollLineAngularDimension2.Clear() : CollLine.Clear() : CollMText.Clear() : CollPolyline3d.Clear() : CollPolyline2d.Clear() : CollPolyline.Clear() : CollRadialDimension.Clear() : CollRay.Clear() : CollRotatedDimension.Clear() : CollSpline.Clear() : CollXline.Clear()
            CollOthers.Clear()

            CollTableHTML.Clear()
            CollVar.Clear() 'Vider les #Var

            Connect.Message(Ass.xTitle, "Traitement en cours", False, 10, 100)

            'Load all Script in a List
            Connect.Message(Ass.xTitle, "Chargement des actions", False, 20, 100)
            For Each FlowFile As String In FlowFiles 'lit tous les fichiers commande + fichiers liés
                ParamVar = ReadScript(FlowFile, LicR)
            Next



            'Vide le tableau des objets
            CollAlignedDimension.Clear() : CollArc.Clear() : CollArcDimension.Clear() : CollBlockReference.Clear() : CollCircle.Clear() : CollDBPoint.Clear() : CollDBText.Clear() : CollDiametricDimension.Clear() : CollEllipse.Clear() : CollAttribute.Clear() : CollHatch.Clear() : CollHelix.Clear() : CollLineAngularDimension2.Clear() : CollLine.Clear() : CollMText.Clear() : CollPolyline3d.Clear() : CollPolyline2d.Clear() : CollPolyline.Clear() : CollRadialDimension.Clear() : CollRay.Clear() : CollRotatedDimension.Clear() : CollSpline.Clear() : CollXline.Clear()
            CollOthers.Clear()
            z = 0 : CollLoad = False 'Chargement des objets = null

            'Initialise le DWG/DXF and start cmds
            If Application.DocumentManager.Count <> 0 Then

                'Open file for Appli script
                'If Revo.RevoFiler.ActiveDraw(FileData) Then

                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                Connect.Message(Ass.xTitle, "Traitement du fichier: " & vbCr & Right(acDoc.Name, 40), False, 1, 100)

                'Boucle sur les commands ****************
                TestCmd = Cmd(Cmds, "", acDoc.Name, TestCmd, ImportFiles)

            End If


            'Fermeture du web
            Try
                WebForm.Close()
                WebForm.Dispose()
            Catch
                MsgBox("Erreur de clôture des fenêtres")
            End Try

            'Ouverture de l'historique
            If Connect.ActLog = True Then RevoFiler.OpenExe(Ass.LogPath)

            'Stop la forms des states
            If TestCmd <> "STOP" Then
                Connect.Message(Ass.xTitle, "Opération terminée avec succès", False, 100, 100, "info")
                Connect.Message(Ass.xTitle, "", True, 100, 100)
            Else
                Connect.Message(Ass.xTitle, "", True, 100, 100)
                Return False
            End If

            Return True '"ok"
        End Function


        '
        'Ancienne commande

        'If InfoCmd(0) = "#ECRASERBLOCK" Then cmdEcraserBlock(CmdLine)
        'If InfoCmd(0) = "#UPDATEEXCEL" Then cmdUpdateExcel(CmdLine)

        'If InfoCmd(0) = "#PROPOBJ" Then cmdPropOBJ(CmdLine)

        'If InfoCmd(0) = "#RENOBL" Then cmdrenoBL(CmdLine)
        'If InfoCmd(0) = "#RENOPRE" Then cmdRenoPRE(CmdLine)
        'If InfoCmd(0) = "#TRIPARCA" Then cmdTriParCA(CmdLine)

        'If InfoCmd(0) = "#UPDYNBLOCK" Then cmdUpDynBlock(CmdLine)

        'If InfoCmd(0) = "#NEWPRE" Then cmdNewPRE(CmdLine)
        'If InfoCmd(0) = "#NEWBL" Then cmdNewBL(CmdLine)
        'If InfoCmd(0) = "#DEFBL" Then cmdDefBL(CmdLine)
        'If InfoCmd(0) = "#ATTRIB" Then cmdAttrib(CmdLine)
        'If InfoCmd(0) = "#NEWHACH" Then cmdNewHACH(CmdLine)
        'If InfoCmd(0) = "#NEWXREF" Then cmdNewXref(CmdLine)
        'If InfoCmd(0) = "#PLOT" Then cmdPlot(CmdLine)
        'If InfoCmd(0) = "#FINDVAR" Then cmdFindVar(CmdLine)

        'Fonction Supprimée
        'If InfoCmd(0) = "#PURGETOUT" Then cmdPurgeTout(CmdLine)


        ''' <summary>
        ''' Revo Command: Cmd
        ''' </summary>
        ''' <param name="Cmds">Commande line</param>
        ''' <remarks>This fonction start all Revo Command</remarks>
        Public Function Cmd(ByVal Cmds As List(Of String), ByVal LoopVarName As String, ByVal DocName As String, ByVal TestCmd As String, ByVal ImportFiles As List(Of String))  ' Remarks
            'Check and start command
            Dim InfoCmd() As String
            Dim CollLoop As New Collection
            Dim z As Double = 0
            Dim NbreData As Double = 0
            Dim RevoVar As New ScriptVar("", Nothing, "", Nothing, Nothing, Nothing)
            Dim ActiveLoop As Boolean = False

            If LoopVarName <> "" Then
                'Boucle dans les variables existantes le chargement
                For Each coll As ScriptVar In CollVar
                    If LoopVarName.ToUpper = coll.Name.ToUpper Then
                        RevoVar = coll
                        Exit For
                    End If
                Next
            Else
                RevoVar = New ScriptVar("", Nothing, "", Nothing, Nothing, Nothing)
            End If

            'Boucle dans les var : LOOP
            If RevoVar.Data IsNot Nothing Then NbreData = RevoVar.Data.Count - 1
            For i As Double = 0 To NbreData
                Try

                    'Si fichier activé execute les commandes
                    For Each CmdLine In Cmds ' Boucle sur commande
                        z = z + 1
                        ' frms.Application.DoEvents()
                        Connect.Message("Action n° " & z & " en cours de traitement ...", "Traitement du fichier: ..." & vbCr & Right(DocName, 40), False, z, Cmds.Count)

                        'REMPLACE CmdLine par ListVars ******* A PAUFINER
                        If RevoVar.Name <> "" Then
                            If RevoVar.CoordX.Count - 1 <= i Then
                                CmdLine = Replace(CmdLine, RevoVar.Name & ".X", RevoVar.CoordX(i))
                                CmdLine = Replace(CmdLine, RevoVar.Name & ".Y", RevoVar.CoordY(i))
                            End If
                            If RevoVar.Data.Count - 1 <= i Then
                                CmdLine = Replace(CmdLine, LoopVarName, RevoVar.Data(i))
                            End If
                        End If

                        InfoCmd = Split(CmdLine, ";")
                        InfoCmd(0) = InfoCmd(0).ToUpper

                        ' --- Uniquement sur les fichiers DWG + DXF -----------------------------------
                        If Right(DocName.ToUpper, 4) = ".DWG" Or Right(DocName.ToUpper, 4) = ".DXF" Or Right(DocName.ToUpper, 4) = ".DWT" Then

                            'Fonction sur le script lui même

                            If InfoCmd(0) = "#LOOP" Then
                                Dim LoopName() As String = SplitCmd(CmdLine, 1)
                                LoopName(1) = Trim(LoopName(1))
                                If LoopName(1) <> "" Then

                                    'Commence une boucle (ajout de la boucle)
                                    If CollLoop.Contains(LoopName(1).ToUpper) = False Then
                                        CollLoop.Add(New List(Of String), LoopName(1).ToUpper) 'Add to coll
                                        ActiveLoop = True

                                    Else 'Cherche la fin de la boucle et lance la boucle
                                        TestCmd = Cmd(CollLoop.Item(LoopName(1).ToUpper), LoopName(1), DocName, TestCmd, ImportFiles)
                                        CollLoop.Remove(LoopName(1).ToUpper) 'Remove to coll
                                        ActiveLoop = False
                                    End If
                                End If
                            Else
                                For Each CmdsLoop As List(Of String) In CollLoop
                                    CmdsLoop.Add(CmdLine)
                                Next
                            End If

                            If ActiveLoop = False Then 'Active

                                '#FILE => Commande traité au niveau du module lors du chargement de dessin (LoopDWG)
                                If InfoCmd(0) = "#FILE" Then conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower & "/fileaction", "action", CmdLine)

                                'Fonction sur un fichier DWG
                                If InfoCmd(0) = "#VAR" Then Connect.RevoLog(cmdVar(CmdLine, LoopVarName))
                                If InfoCmd(0) = "#CMD" Then Connect.RevoLog(cmdCmd(CmdLine))
                                If InfoCmd(0) = "#PS" Then Connect.RevoLog(cmdPS(CmdLine))
                                If InfoCmd(0) = "#LA" Then Connect.RevoLog(cmdLA(CmdLine))
                                If InfoCmd(0) = "#BL" Then Connect.RevoLog(cmdBL(CmdLine))
                                If InfoCmd(0) = "#HA" Then Connect.RevoLog(cmdHA(CmdLine))
                                If InfoCmd(0) = "#OBJ" Then Connect.RevoLog(cmdOBJ(CmdLine))
                                If InfoCmd(0) = "#STY" Then Connect.RevoLog(cmdSTY(CmdLine))
                                If InfoCmd(0) = "#PURGE" Then Connect.RevoLog(cmdPurge(CmdLine))
                                If InfoCmd(0) = "#ZOOM" Then Connect.RevoLog(cmdZoom(CmdLine))
                                If InfoCmd(0) = "#EXPORT" Then Connect.RevoLog(cmdExport(CmdLine))
                                If InfoCmd(0) = "#IMPORT" Then
                                    TestCmd = cmdImport(CmdLine, ImportFiles)
                                    Connect.RevoLog(TestCmd)
                                    If TestCmd.ToUpper = "STOP" Then
                                        Return TestCmd
                                    End If
                                End If

                            End If

                        End If

                        ' --- Sur tous type de fichiers DWG + DXF + PTS + ITF  ----------------------
                        'If ActiveLoop = False Then 'Active 
                        'If InfoCmd(0) = "#IMPORT" Then
                        'TestCmd = cmdImport(CmdLine, ImportData)
                        'If TestCmd <> "IMPORT-ITF" Then
                        'Connect.RevoLog(TestCmd)
                        'Else
                        'Connect.Message(Ass.xTitle, "Importation en cours ...", True, 70, 100, "info")
                        'End If

                        'End If
                        'End If

                    Next 'Fin de boucle dans les cmds

                Catch 'ex As Exception
                    ' MsgBox("Erreur dans la boucle cmds dans Cmd") ' & ex.Message)
                End Try
            Next 'Fin de boucle dans les ListVar

            Return TestCmd
        End Function


        Public Shared Function ReadScript(ByVal FileName As String, ByVal LicRevo As String)
            Dim FichierALire As String = FileName
            Dim Script As New Revo.RevoScript
            Dim Connect As New Revo.connect
            'Dim Test As String = ""
            Dim Param As New List(Of String)
            Dim ParamImport As Boolean = False
            Param.Add("OK")
            Param.Add(False)

            'Verification de l'existance du FichierALire
            If System.IO.File.Exists(FichierALire) Then
                Try
                    Dim sr As StreamReader = New StreamReader(FichierALire, System.Text.Encoding.Default)
                    Dim ligne As String
                    '--- Traitement du fichier ligne par ligne
                    While Not sr.EndOfStream()
                        ligne = sr.ReadLine()

                        If Mid(ligne, 1, 1) = "#" Then
                            If Mid(ligne.ToUpper, 1, 12) = "#LOADSCRIPT;" Then  'implémentation futur: Si variable traite dans une 2ème étapes
                                'Load script
                                Dim CmdsLigne() As String
                                CmdsLigne = Split(ligne, ";")
                                ReadScript(CmdsLigne(1), LicRevo)
                            Else
                                Param = Script.TestScriptVar(ligne, LicRevo)  'Test s'il faut ajouter la commande
                                If Param(0) = True Then Cmds.Add(Replace(ligne, "]];", "]]")) 'Fusionne Param + Valeur
                                If Param(1) = True Then ParamImport = True

                            End If
                        Else
                        End If

                    End While
                    '--- Referme StreamReader
                    sr.Close()
                    Param(0) = "OK"
                    Param(1) = ParamImport
                Catch ex As Exception
                    'Traitement de l'exception sinon :
                    Throw ex
                    Connect.RevoLog(Connect.DateLog & "Read Script" & vbTab & False & vbTab & "Erreur lecture: " & ex.Message & vbTab & FileName)
                    Param(0) = "ERR"
                End Try
            Else
                Connect.RevoLog(Connect.DateLog & "Read Script" & vbTab & False & vbTab & "Erreur lecture: " & "fichier inexistant" & vbTab & FileName)
                Param(0) = "ERR"
            End If

            Return Param
        End Function


        ''' <summary>
        ''' Test Variable and licence Revo
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <param name="LicRevo">Name of licence</param>
        ''' <remarks></remarks>
        Public Function TestScriptVar(ByVal Cmd As String, ByVal LicRevo As String) As List(Of String)
            Dim TestCmd() As String = SplitCmd(Cmd, 2)
            Dim Param As New List(Of String)

            'Tableau Param: 0 = Commande valide (True/False)
            '               1 = Commande import (True/False)

            Param.Add(False) '0 Commande non-valide

            'Commande Import (si NomFichier vide import uniquement dans le fichier courant)
            If TestCmd(0).ToUpper = "#IMPORT" And Trim(TestCmd(2)) = "" Then
                Param.Add(True) '1: Commande import (True)
            Else
                Param.Add(False) '1: Commande import (False)
            End If

            'Gestion des variables
            'Test sur variables

            'Gestion de commande spéciales
            If TestCmd(0).ToUpper = "#TITLE" Or TestCmd(0).ToUpper = "#DESCRIPTION" Or TestCmd(0).ToUpper = "#GROUPE" Then
                Param(0) = False
            Else
                Param(0) = True
            End If

            Return Param
        End Function
        ''' <summary>
        ''' Split Revo commande in the Table
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <param name="NbreMinDe0">Minimum param for the commande (count only param)</param>
        ''' <remarks></remarks>
        Public Function SplitCmd(ByVal Cmd As String, Optional ByVal NbreMinDe0 As Double = 0)

            Dim Cmdinfo() As String
            Cmdinfo = Split(Cmd, ";")

            If NbreMinDe0 <> 0 Then
                If Cmdinfo.GetUpperBound(0) < NbreMinDe0 Then
                    Cmd = Cmd & Replace(Space(NbreMinDe0 - Cmdinfo.GetUpperBound(0)), " ", ";")
                    Cmdinfo = Split(Cmd, ";")
                End If
            End If

            Return Cmdinfo

        End Function


        ''' <summary>
        ''' Revo Command: #Var
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks>Chargement de données externe et stockée dans une variable.</remarks>
        Public Function cmdVar(ByVal Cmd As String, ByVal LoopVarName As String)  ' Remarks
            '#Var;/Name/;Type;[[PropA]]Val;[[PropB]]Val;
            'Chargement de données externe et stockée dans une variable.
            '
            '    Balise
            '    0  Name : ""
            '    1  Type : Acad / MySQL / http / DataReader
            '    *  [[Prop…]]Val…
            '
            '      [[Propriété]] Acad
            '    * ToDay = dd.mm.yyyy
            '    * Folder = nbre depuis la racine
            '    * Case = Lower / Upper /
            '    * Replace = "abc"||"xyz"
            '    * Mid = Start||Lenght
            '
            '      [[Propriété]] MySQL
            '    # SQL = "SELECT ma_table./champs/ FROM ma_table WHERE ma_table.MonChamps = /NumAffaire/"
            '    # SQLurl = ""
            '    # SQLuser = ""
            '    # SQLpass = ""

            Dim VarName As String = ""
            Dim VarData As New List(Of String)
            Dim VarType As String = ""
            Dim CoordX As New List(Of String)
            Dim CoordY As New List(Of String)
            Dim Collprop As New Collection
            'Dim InitVar As Boolean = False

            Dim Cmdinfo() As String
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim curDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim curDWG As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Cmdinfo = SplitCmd(Cmd, 2) '2 = min de paramètre obligatoire

            If Cmdinfo(1) <> "" Then 'si la ligne n'est pas vide
                'Chargement des propriétés
                VarName = Cmdinfo(1) 'Name variable
                VarType = Cmdinfo(2) 'Type variable  

                For i = 3 To Cmdinfo.Count - 1 'Boucle prop
                    If InStr(Cmdinfo(i), "[[") <> 0 And InStr(Cmdinfo(i), "]]") <> 0 Then
                        Cmdinfo(i) = Replace(Replace(Cmdinfo(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                        'Dim TabPropName As String = ""
                        'Dim TabPropData As New List(Of String)
                        Dim TabProp() As String
                        TabProp = SplitCmd(Cmdinfo(i), 1)
                        TabProp(0) = TabProp(0).ToLower

                        'Boucle dans les variables existantes pour le remplacement
                        For Each coll As ScriptVar In CollVar
                            TabProp(1) = Replace(TabProp(1), "¦", ";") 'remplacement des ;
                            If InStr(TabProp(1), coll.Name) <> 0 Then 'Recherche s'il faut remplacer une var
                                Dim DataMultiline As String = ""
                                For Each DataX As String In coll.Data ' Transforme la list en une variable multiline
                                    If DataMultiline <> "" Then DataMultiline += vbCrLf
                                    DataMultiline += DataX
                                Next

                                TabProp(1) = Replace(TabProp(1), coll.Name, DataMultiline) ' remplacement des /var/ dans toute les lignes

                            End If
                            If coll.Name.ToLower = VarName.ToLower And LoopVarName = "" Then
                                VarData = coll.Data 'Chargement d'existant
                            End If

                        Next

                        If TabProp(0) <> "" Then ' x Propriété
                            Collprop.Add(TabProp)
                        End If
                    End If
                Next


                'Chargement des variables  :       [[Propriété]] MySQL 
                If VarType.ToLower = "mysql" Then
                    Dim BDdata As New Collection
                    Dim rMySQL As New RevoMySQL
                    Dim dts As DataSet
                    dts = rMySQL.MySQLquery(Collprop)
                    Try
                        For Each row As DataRow In dts.Tables(0).Rows
                            VarData.Add(row(0).ToString())
                            ' FichiersData.Add(IO.Path.Combine(Fenetre.GridData.Rows(i).Cells(2).Value, Fenetre.GridData.Rows(i).Cells(1).Value))
                        Next
                    Catch
                    End Try
                End If

                'Chargement des variables  :       [[Propriété]] DataReader 
                If VarType.ToLower = "datareader" Then VarData = VarDataReader(Collprop, VarData, VarName, curDwg, LoopVarName)

                'Chargement des variables  :       [[Propriété]] DataReader 
                If VarType.ToLower = "xls" Then VarData = VarXLS(Collprop, VarData, VarName, curDwg, LoopVarName)

                'Chargement des variables  :       [[Propriété]] Acad 
                If VarType.ToLower = "acad" Then
                    VarData = VarAcad(Collprop, VarData, VarName, curDwg, LoopVarName)
                    CoordX.AddRange(Split(VarData(VarData.Count - 2), "|"))
                    CoordY.AddRange(Split(VarData(VarData.Count - 1), "|"))
                    VarData.RemoveAt(VarData.Count - 1)
                    VarData.RemoveAt(VarData.Count - 1)
                End If


                'Chargement des variables  :       [[Propriété]] HTTP 
                If VarType.ToLower = "http" Then VarData = VarHttp(Collprop, VarData, curDwg, LoopVarName)


                ' Recherche si Var existante dans la collection
                Dim TestWrite As Boolean = True
                For Each coll As ScriptVar In CollVar
                    If VarName.ToLower = coll.Name.ToLower Then
                        coll.Data = VarData
                        TestWrite = False
                    End If
                Next ' Si pas de Var complète la collection

                If TestWrite Then CollVar.Add(New ScriptVar(VarName, VarData, VarType, Collprop, CoordX, CoordY))

                ' MsgBox(VarName & " = " & VarData)

            Else 'ignore the command
                Return Connect.DateLog & "Cmd Var" & vbTab & False & vbTab & "Erreur null string: " & vbTab & curDwg.Name
            End If

            Return Connect.DateLog & "Cmd Var" & vbTab & True & vbTab & "Name: " & Cmdinfo(1) & vbTab & curDwg.Name

        End Function
        Private Function VarDataReader(ByVal Collprop As Collection, ByVal VarData As List(Of String), ByVal VarName As String, ByVal CurDWG As AcadDocument, ByVal LoopVarName As String) As List(Of String)

            '      [[Propriété]] MySQL
            Dim ReqSQL As String = "" '    # SQL = "SELECT date(from_unixtime(ma_elem_techniques.MATDateLevee))  FROM ma_elem_techniques, ma_mandate WHERE ma_mandate.MADCode = 'CP07230' and ma_mandate.MADId = ma_elem_techniques.MATMandate"
            Dim URLsrv As String = "" '    # SQLurl = ""
            Dim User As String = ""   '    # SQLuser = ""
            Dim Pass As String = ""   '    # SQLpass = ""
            Dim Workgroup As String = "" ' # SQLwg = "" sécu *.mdw
            Dim DBname As String = ""     '    # SQLdb = ""
            For Each Prop() As String In Collprop
                If Prop(0).ToLower = "sql" Then 'SQL = "SELECT ma_table...
                    ReqSQL = Prop(1)
                ElseIf Prop(0).ToLower = "sqlurl" Then 'SQLurl = ""
                    URLsrv = Prop(1)
                ElseIf Prop(0).ToLower = "sqluser" Then 'SQLuser = ""
                    User = Prop(1)
                ElseIf Prop(0).ToLower = "sqlpass" Then  'SQLpass = ""
                    Pass = Prop(1)
                ElseIf Prop(0).ToLower = "sqldb" Then 'SQLdb = ""
                    DBname = Prop(1)
                ElseIf Prop(0).ToLower = "sqlwg" Then 'SQLdb = ""
                    Workgroup = Prop(1)
                End If
            Next

            URLsrv = Path.Combine(URLsrv, DBname)
            Dim ListCoord As New List(Of String) ' Lecture en boucle

            'Password=' le mot de passe , 'User ID=' Admin
            Try
                'cnn.Open "Provider=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\nwind.mdb;Jet OLEDB:Database Password=DBPassword"
                '  Dim MyConnexion As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection("Provider=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & _
                '             URLsrv & ";Jet OLEDB:Database Password=" & Pass & ";") '& ";User ID=" & User & ";Password=" & Pass & ";")  '"C:\consultation.mdb""

                ' Dim MyConnexion As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection("Provider=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & _
                '                           URLsrv & ";Uid=" & User & ";Pwd=" & Pass & ";")  '"C:\consultation.mdb""


                '   Dim MyConnexion As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & _
                '                          URLsrv & ";User Id=admin;Password=;")


                Dim MyConnexion As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection("Provider=MSDataShape;Data Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & _
                                         URLsrv & ";Jet OLEDB:System Database=" & Workgroup & ";" & "User ID=" & User & ";Password=" & Pass & ";")  '"C:\consultation.mdb""




                Dim Mycommand As System.Data.OleDb.OleDbCommand = MyConnexion.CreateCommand()
                Mycommand.CommandText = ReqSQL '"SELECT NOM FROM QUESTIONS"
                MyConnexion.Open() 'Ouverture de la DB
                'Dim myReader As System.Data.OleDb.OleDbDataReader = Mycommand.ExecuteReader()
                Dim myReader As System.Data.OleDb.OleDbDataReader = Mycommand.ExecuteReader(CommandBehavior.CloseConnection)
                Do While myReader.Read() 'Boucle dans les valeurs
                    ListCoord.Add(myReader.GetValue(0))
                Loop
                'myReader.Close() ' Si fermeture manuel => sans .CloseConnection
                'MyConnexion.Close()

            Catch ex As System.Data.OleDb.OleDbException 'ex As Exception
                MsgBox(ex.Message) 'http://support.microsoft.com/kb/823913/fr
            Catch ex2 As System.Data.SqlClient.SqlException 'ex As Exception
                MsgBox(ex2.Message)
            Catch
                MsgBox("Une erreur a été trouvée dans la lecture de la BD :" & DBname)
            End Try

            Return ListCoord

            'Modif pour du SQLserveur : http://plasserre.developpez.com/cours/vb-net/?page=bases-donnees2#XVII-D-2
            'Imports System.Data.SqlClient
            'Dim MyConnexion As SqlConnection = New SqlConnection("Data Source=localhost;" & _
            '"Integrated Security=SSPI;Initial Catalog=northwind")
            'Dim Mycommand As SqlCommand = MyConnexion.CreateCommand()
            'Dim myReader As SqlDataReader = Mycommand.ExecuteReader()

        End Function

        Private Function VarXLS(ByVal Collprop As Collection, ByVal VarData As List(Of String), ByVal VarName As String, ByVal CurDWG As AcadDocument, ByVal LoopVarName As String) As List(Of String)
            '#Var;/Name/;XLS;[[Url]];..rechercheExcel.xls;[[Sheet]];Feuil1;[[FindV]];*txt*||A1:D31||2;

            '    [[Propriété]] XLS
            Dim Urlfile As String = "" ' fichier.xls
            Dim NomFeuille As String = "" ' Feuil1
            Dim FindV() As String  ' *txt*||A5:D31||2 (fonctionne comme RechercheV) 
            Dim ListData As New List(Of String) ' Lecture en boucle
            Dim FindTXT As String = ""
            Dim Zone As String = "A1:B1000"
            Dim ColData As Integer = 1 'Commence par 0 (1 = 2eme col)

            For Each Prop() As String In Collprop
                If Prop(0).ToLower = "url" Then 'Urlfile
                    Urlfile = Prop(1)
                ElseIf Prop(0).ToLower = "findv" Then 'SQLurl = ""
                    FindV = Split(Prop(1), "||")
                    FindTXT = FindV(0)
                    Zone = FindV(1)
                    If IsNumeric(FindV(2)) Then ColData = CInt(FindV(2))
                ElseIf Prop(0).ToLower = "sheet" Then 'SQLurl = ""
                    NomFeuille = Prop(1) & "$" 'Ajout du $ pour le nom de feuille !!!
                End If
            Next


            If File.Exists(Urlfile) Then

                'Lecture dans le fichier EXCEL ------------------------------------

                Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Urlfile & ";Extended Properties=""Excel 8.0;IMEX=1;HDR=No;"";"
                Dim connection As New System.Data.OleDb.OleDbConnection(connectionString)

                Dim Tableau As String = ""

                'Recherche de l onglet
                'If IsNumeric(NomFeuille) Then
                '    connection.Open()
                '    Dim FeuillesExcel As System.Data.DataTable
                '    FeuillesExcel = connection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
                '    NomFeuille = FeuillesExcel.Rows(0).Item("TABLE_NAME").ToString()
                '    connection.Close()
                'End If

                'Requête sur l'onglet
                Dim cmdText As String = "SELECT * FROM [" & NomFeuille & Zone & "]"
                Dim command As New System.Data.OleDb.OleDbCommand(cmdText, connection)

                'Lecture des Valeurs
                command.Connection.Open()
                Dim reader As System.Data.OleDb.OleDbDataReader = command.ExecuteReader()

                If reader.HasRows Then
                    While reader.Read() ' Lecture ligne
                        If reader(0).ToString() Like FindTXT Then 'Analyse 1er colonne
                            ListData.Add(reader(ColData - 1).ToString()) 'ajoute la cellule selon n° Col
                        End If
                    End While
                End If

                command.Connection.Close()

            Else 'Pas de fichier existant

            End If

            Return ListData

        End Function

        Private Function VarAcad(ByVal Collprop As Collection, ByVal VarData As List(Of String), ByVal VarName As String, ByVal CurDWG As AcadDocument, ByVal LoopVarName As String) As List(Of String)

            Dim CoordX As String = ""
            Dim CoordY As String = ""
            Dim ValideIf As Boolean = True
            Dim DataList As New List(Of String)
            For i = 0 To VarData.Count - 1
                DataList.Add(i) 'add ID de tout VarData
            Next

            'Chargement des variables  :       [[Propriété]] Acad 
            For Each Prop() As String In Collprop

                Prop(1) = Replace(Prop(1), "\P", vbCrLf) 'Remplace le retour à la ligne dans toute le valeur des variables

                If Prop(0).ToLower = "if" Then 'If =/var/|=|var1 (si est égal ou |<| ou |>| ou |<>| : exécute les autres propriétés)
                    Dim VarCond() As String = Split(Prop(1), "|")
                    If VarCond.Count = 3 Then
                        ValideIf = False : DataList.Clear()
                        Dim Cond0 As String = ""
                        Dim NbreVarData As Double = VarData.Count - 1
                        If NbreVarData = -1 Then NbreVarData = 0
                        For i = 0 To NbreVarData
                            If VarCond(0) = "" Then
                                If VarData.Count > 0 Then Cond0 = VarData(i) 'si var vide complète avec une existant
                            Else
                                Cond0 = VarCond(0)
                            End If
                            If VarCond(1) = "=" Or VarCond(1) = "<>" Then
                                If Cond0 Like VarCond(2) Then
                                    If VarCond(1) = "=" Then DataList.Add(i) : ValideIf = True
                                Else
                                    If VarCond(1) = "<>" Then DataList.Add(i) : ValideIf = True
                                End If
                            ElseIf VarCond(1) = "<" Then
                                If Cond0 < VarCond(2) Then DataList.Add(i) : ValideIf = True
                            ElseIf VarCond(1) = ">" Then
                                If Cond0 > VarCond(2) Then DataList.Add(i) : ValideIf = True
                            End If
                        Next
                    End If

                    'Propriété ajoutant une valeurs

                ElseIf Prop(0).ToLower = "val" And ValideIf Then '* Val = ""

                    If Mid(Prop(1), 1, 3) = ">>>" Then 'Sélectionne la coordonnées d'un point à l'écran


                        'Suppression messages
                        Connect.Message("Sélection d'objets", "Traitement en cours ... ", False, 0, 0, "hide")

                        Dim PicX As String = ""
                        Dim PicY As String = ""
                        Dim PicZ As String = ""

                        'Saisir un points de coordonnées
                        PicX = ""
                        PicY = ""
                        PicZ = ""

                        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

                        Dim pPtRes As PromptPointResult
                        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

                        '' Prompt for the start point
                        pPtOpts.Message = vbLf & "Saisir un point:"
                        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                        Dim ptPic As Point3d = pPtRes.Value

                        '' Exit if the user presses ESC or cancels the command
                        If pPtRes.Status = PromptStatus.Cancel Then
                            PicX = "0"
                            PicY = "0"
                            PicZ = "0"
                        Else
                            PicX = ptPic(0)
                            PicY = ptPic(1)
                            PicZ = ptPic(2)
                        End If

                        'Stockage de la variable
                        VarData.Add(PicX & "," & PicY & "," & PicZ) 'ajoute à la collection !!!
                        DataList.Add(VarData.Count - 1) 'add ID

                    ElseIf Mid(Prop(1), 1, 2) = ">>" Then 'Sélection d'objet manuel
                        Prop(1) = Mid(Prop(1), 3, Len(Prop(1)) - 2)
                        ' Start for selection of objects

                        'Suppression messages
                        Connect.Message("Sélection d'objets", "Traitement en cours ... ", False, 0, 0, "hide")

                        '' Get the current document and database
                        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                        Dim acCurDb As Database = acDoc.Database
                        '' Start a transaction
                        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                            '' Request for objects to be selected in the drawing area
                            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection()
                            '' If the prompt status is OK, objects were selected
                            If acSSPrompt.Status = PromptStatus.OK Then
                                Dim acSSet As SelectionSet = acSSPrompt.Value
                                '' Step through the objects in the selection set
                                For Each acSSObj As SelectedObject In acSSet
                                    '' Check to make sure a valid SelectedObject object was returned
                                    If Not IsDBNull(acSSObj) Then
                                        '' Open the selected object for write
                                        Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, DatabaseServices.OpenMode.ForRead)
                                        If Not IsDBNull(acEnt) Then

                                            If TypeName(acEnt) = "BlockReference" Then
                                                Dim BL As BlockReference
                                                BL = acEnt
                                                Dim attRef As AttributeReference = Nothing
                                                For Each Attid As ObjectId In BL.AttributeCollection
                                                    attRef = DirectCast(acTrans.GetObject(Attid, DatabaseServices.OpenMode.ForRead), AttributeReference)
                                                    If attRef.Tag.ToUpper = VarName.ToUpper Then
                                                        'Dim tempObj As AcadObject
                                                        'Dim objectHandle As String
                                                        'objectHandle = splineObj.Handle
                                                        If CoordX <> "" Then CoordX += "|" : CoordY += "|"
                                                        CoordX += "ID:"
                                                        CoordY += BL.Handle.ToString
                                                        VarData.Add(attRef.TextString) 'ajoute à la collection !!!
                                                        DataList.Add(VarData.Count - 1) 'add ID
                                                        Exit For
                                                    End If
                                                Next
                                            ElseIf TypeName(acEnt) = "MText" Then
                                                Dim Mtxt As MText
                                                Mtxt = acEnt
                                                If CoordX <> "" Then CoordX += "|" : CoordY += "|"
                                                CoordX += Mtxt.GeometricExtents.MinPoint(0).ToString
                                                CoordY += Mtxt.GeometricExtents.MaxPoint(1).ToString
                                                VarData.Add(Mtxt.Text) 'ajoute à la collection !!!
                                                DataList.Add(VarData.Count - 1) 'add ID
                                            ElseIf TypeName(acEnt) = "DBText" Then
                                                Dim DBtxt As DBText
                                                DBtxt = acEnt
                                                If CoordX <> "" Then CoordX += "|" : CoordY += "|"
                                                CoordX += DBtxt.Position.Coordinate(0).ToString
                                                CoordY += DBtxt.Position.Coordinate(1).ToString
                                                VarData.Add(DBtxt.TextString) 'ajoute à la collection !!!
                                                DataList.Add(VarData.Count - 1) 'add ID
                                            End If
                                        End If
                                    End If
                                Next
                            End If
                            '' Save the new object to the database
                            acTrans.Commit()
                            '' Dispose of the transaction
                        End Using
                        ' End selection of objects


                    ElseIf Mid(Prop(1), 1, 1) = ">" Then 'Affiche la sélection manuel d'une valeur

                        Prop(1) = Mid(Prop(1), 2, Len(Prop(1)) - 1)
                        Dim fSelect As New frmState
                        fSelect.Text = Ass.xTitle
                        fSelect.lbl_infos.Text = "Choissisez une donnée : " & VarName
                        fSelect.ProgBar.Visible = False
                        fSelect.BtnValid.Visible = True
                        fSelect.BoxList.Visible = True
                        'fSelect.btnSelect.Visible = True
                        Dim MsgList() As String = Prop(1).Split(vbCrLf)
                        For Each Txt In MsgList
                            If Txt <> "" Then fSelect.BoxList.Items.Add(Replace(Txt, Chr(10), ""))
                        Next
                        Dim SelectedTxt As String = ""
                        If fSelect.BoxList.Items.Count <> 0 Then fSelect.BoxList.Text = fSelect.BoxList.Items.Item(0)
                        fSelect.ShowDialog()
                        SelectedTxt = (fSelect.BoxList.Text) 'ajoute à la collection !!!
                        fSelect.Dispose()
                        VarData.Add(SelectedTxt) 'ajoute à la collection !!!
                        DataList.Add(VarData.Count - 1) 'add ID

                    Else 'Valeur inscrite dans le script

                        If LoopVarName <> "" Then 'Si boucle vide les données
                            DataList.Clear()
                            VarData.Clear()
                            VarData.Add(Prop(1)) 'ajoute à la collection !!!
                            DataList.Add(VarData.Count - 1) 'add ID

                        Else 'S'il y a pas de boucle ajoute (sans vider les données)
                            VarData.Add(Prop(1)) 'ajoute à la collection !!!
                            DataList.Add(VarData.Count - 1) 'add ID
                        End If

                    End If

                ElseIf Prop(0).ToLower = "today" And ValideIf Then '* ToDay = dd.mm.yyyy
                    Try
                        Dim FormatDate As String = Prop(1)
                        VarData.Add(Format(Today, FormatDate)) 'ajoute à la collection !!!
                        DataList.Add(VarData.Count - 1) 'add ID de tout VarData
                    Catch
                    End Try
                ElseIf Prop(0).ToLower = "split" And ValideIf Then '* Split = car||nbre||/var/
                    Try
                        Dim VarSplit() As String = Split(Prop(1), "||")
                        If VarSplit.Count = 3 Then
                            'VarSplit(0) = Replace(VarSplit(0), "\P", vbCrLf)
                            Dim DataSplit() As String = Split(VarSplit(2), VarSplit(0))
                            If IsNumeric(VarSplit(1)) And VarSplit(1) <> "" Then ' C:\NY\08\000-100\"NY08014"  < nbre = 4
                                Dim DataSplitID As Double = CDbl(VarSplit(1)) - 1
                                If DataSplitID <= DataSplit.Count - 1 And DataSplitID >= 0 Then
                                    VarData.Add(DataSplit(DataSplitID)) 'ajoute à la collection !!!
                                    DataList.Add(VarData.Count - 1) 'add ID
                                End If
                            Else 'Si pas de numéro de bloc crée une collection
                                For i = 0 To DataSplit.Count - 1
                                    VarData.Add(DataSplit(i))
                                    DataList.Add(VarData.Count - 1) 'add ID
                                Next
                            End If
                        End If
                    Catch
                        '  MsgBox("Erreur Split " & "ex.Message" & " : " & Prop(1))
                    End Try

                ElseIf Prop(0).ToLower = "find" And ValideIf Then '* Find = ObjFilter||Cond: block Name / *txt*||ObjID / Txt||Space
                    Dim Findparam() As String = Split(Prop(1), "||") '0          1                       2         3
                    If Findparam.Count >= 4 Then
                        'Choix du Space
                        Dim Space As AcadBlock = SelectSpace(CurDWG, Findparam(3)) 'Par défault : Model
                        Dim refID As Double = 0
                        'Boucle dans les OBJETS block : ModelSpace ou autres
                        For Each Obj In Space
                            If TypeName(Obj) Like Findparam(0) Then
                                If Findparam(2).ToLower = "lastobjid" Then
                                    If IsNumeric(Obj.ObjectID) Then
                                        If refID < Obj.ObjectID Then
                                            refID = Obj.ObjectID : VarData.Clear() : VarData.Add(CStr(Obj.ObjectID)) '1 seul élément !!!
                                            DataList.Add(VarData.Count - 1) 'add ID
                                        End If
                                    End If
                                ElseIf Findparam(2).ToLower = "txt" Then
                                End If
                            End If
                        Next
                    End If
                ElseIf Prop(0).ToLower = "file" And ValideIf Then '* Folder = nbre depuis la racine
                    If Prop(1).ToLower = "file" Then  'File
                        VarData.Add(CurDWG.Name) 'ajoute à la collection !!!
                        DataList.Add(VarData.Count - 1) 'add ID
                    ElseIf Prop(1).ToLower = "folder" Then 'Folder
                        VarData.Add(CurDWG.FullName) 'ajoute à la collection !!!
                        DataList.Add(VarData.Count - 1) 'add ID
                    ElseIf Prop(1).ToLower = "library" Then 'Library block folder
                        VarData.Add(Ass.Library) 'ajoute à la collection !!!
                        DataList.Add(VarData.Count - 1) 'add ID
                    End If


                    'Propriété modifiant une valeur (pas d'ajout)

                ElseIf Prop(0).ToLower = "column" And ValideIf Then 'Column = car||2,3,1 (réorganise une collection)

                    Dim RepVar() As String = Split(Prop(1), "||")
                    Dim OrderCol() As String = Split(RepVar(1), ",")
                    Try
                        For i = 0 To DataList.Count - 1

                            ' Division de la ligne
                            Dim VarDataTri() As String = Split(VarData(DataList(i)), RepVar(0))
                            Dim VarDaraTriLine As String = ""
                            For u = 0 To OrderCol.Count - 1
                                If VarDataTri.Count = OrderCol.Length Then ' CDbl(OrderCol(u)) >= 1 And CDbl(OrderCol(u)) <= OrderCol.Count Then
                                    VarDaraTriLine += VarDataTri(CDbl(OrderCol(u)) - 1)
                                    If u <> OrderCol.Count - 1 Then VarDaraTriLine += RepVar(0)
                                Else
                                    '    MsgBox("Problème dans la commande #Var Column :" & OrderCol(u), vbInformation)
                                End If
                            Next
                            VarData(DataList(i)) = VarDaraTriLine
                        Next

                    Catch exep2 As COMException
                        MsgBox("Commande REVO VarAcad (Column) COM" & exep2.Message)
                    Catch exep As System.Exception
                        MsgBox("Commande REVO VarAcad (Column) SYS" & exep.Message)
                    Finally
                        ' MsgBox("Commande REVO VarAcad (Column) FINALLY")
                    End Try
                ElseIf Prop(0).ToLower = "case" Then '* Case = Lower / Upper / Capitalize
                    If Prop(1).ToLower = "lower" Then
                        For i = 0 To DataList.Count - 1
                            If VarData(DataList(i)) <> Nothing Then VarData(DataList(i)) = VarData(DataList(i)).ToLower
                        Next
                    ElseIf Prop(1).ToLower = "upper" Then
                        For i = 0 To DataList.Count - 1
                            If VarData(DataList(i)) <> Nothing Then VarData(DataList(i)) = VarData(DataList(i)).ToUpper
                        Next
                    ElseIf Prop(1).ToLower = "capitalize" Then
                        For i = 0 To DataList.Count - 1
                            If VarData(DataList(i)) <> Nothing Then VarData(DataList(i)) = StrConv(VarData(DataList(i)), vbProperCase)
                        Next
                    End If
                ElseIf Prop(0).ToLower = "replace" Then '* Replace = "xyz"
                    Dim RepVar() As String
                    If InStr(Prop(1), "||") = 0 Then Prop(1) = Prop(1) & "||"
                    RepVar = Split(Prop(1), "||")
                    For i = 0 To DataList.Count - 1
                        VarData(DataList(i)) = Replace(VarData(DataList(i)), RepVar(0), RepVar(1))
                    Next
                ElseIf Prop(0).ToLower = "mid" Then '* Mid = Start||Lenght
                    Dim MidVar() As String = Split(Prop(1), "||")
                    Dim Mid0, Mid1 As Double
                    For i = 0 To DataList.Count - 1
                        If MidVar(0) = "" Then
                            Mid0 = 1 : Mid1 = Len(VarData(DataList(i))) - MidVar(1)
                        Else
                            Mid0 = CDbl(MidVar(0)) : Mid1 = CDbl(MidVar(1))
                        End If
                        VarData(DataList(i)) = (Mid(VarData(DataList(i)), Mid0, Mid1))
                    Next
                ElseIf Prop(0).ToLower = "math" Then '* Replace = "xyz"
                    Dim MathVar() As String
                    If InStr(Prop(1), "||") = 0 Then Prop(1) = Prop(1) & "||"
                    MathVar = Split(Prop(1), "||")
                    For i = 0 To DataList.Count - 1
                        If IsNumeric(MathVar(1)) And IsNumeric(VarData(DataList(i))) Then
                            If MathVar(0) = "+" Then VarData(DataList(i)) = CDbl(VarData(DataList(i))) + CDbl(MathVar(1))
                            If MathVar(0) = "-" Then VarData(DataList(i)) = CDbl(VarData(DataList(i))) - CDbl(MathVar(1))
                        End If
                    Next
                ElseIf Prop(0).ToLower = "format" Then '* Mid = Start||Lenght
                    Dim MidVar() As String = Split(Prop(1), "||")
                    Dim Mid0, Mid1 As Double
                    Dim StrS As String = "" : Dim StrM As String = "" : Dim StrE As String = ""
                    For i = 0 To DataList.Count - 1
                        If MidVar(0) = "" Then
                            If IsNumeric(MidVar(1)) = False Then MidVar(1) = 0
                            Mid0 = MidVar(1) : Mid1 = Len(VarData(DataList(i))) - MidVar(1)
                            StrS = Mid(VarData(DataList(i)), 1, Mid1)
                            StrM = Mid(VarData(DataList(i)), Mid1 + 1, Len(VarData(DataList(i))) - Mid0 + Mid1 + 1)
                            StrE = "" 'Mid(VarData(DataList(i)), Mid1 + Mid0 + 1, Len(VarData(DataList(i))))
                        Else
                            Mid0 = CDbl(MidVar(0))
                            If Mid0 < 1 Then Mid0 = 1
                            Mid1 = CDbl(MidVar(1))
                            StrS = Mid(VarData(DataList(i)), 1, Mid0 - 1)
                            StrM = Mid(VarData(DataList(i)), Mid0, Mid1)
                            StrE = Mid(VarData(DataList(i)), Mid1 + 1, Len(VarData(DataList(i))) - Mid0 + Mid1 + 1)
                        End If

                        If InStr(MidVar(2), "#") <> 0 And IsNumeric(Replace(StrM, "'", "")) Then
                            MidVar(2) = Replace(MidVar(2), "'", "`")
                            VarData(DataList(i)) = StrS & Replace(Format(CDbl(Replace(StrM, "'", "")), Replace(MidVar(2), "'", "`")), "`", "'") & StrE
                        Else
                            VarData(DataList(i)) = StrS & Format((StrM), MidVar(2)) & StrE
                        End If
                    Next

                End If
            Next


            VarData.Add(CoordX)
            VarData.Add(CoordY)

            Return VarData

        End Function


        Public Function VarHttp(ByVal Collprop As Collection, ByVal VarData As List(Of String), ByVal CurDWG As AcadDocument, ByVal LoopVarName As String) As List(Of String)

            WebForm.WebStat = "-"
            'WebbrowserState = "-"

            '     MsgBox(WebForm.WebBrowser2.Version.Major.ToString)

            Try
                For Each Prop() As String In Collprop

                    If Prop(0).ToLower = "wurl" Then 'Wurl = "http://..."    'Test Refrech  <<<

                        Try
                            'WebForm.Show()
                            'WebForm.WebBrowser2.Navigate(Prop(1)) 'WebForm.
                            'WebForm.Text = Ass.xTitle & " Navigator"
                            'WebForm.TimerWeb.Interval = 2000    'Timer1_Tick sera déclenché toutes les secondes.
                            'WebForm.TimerWeb.Start()            'On démarre le Timer
                            'WebForm.TimerWeb.Stop()

                            WebForm.LoadPage(Prop(1))

                            'Reactivation des lignes si-dessous : 06.03.2013
                            Dim t As DateTime = DateTime.Now
                            Do While DateTime.Now < t.AddSeconds(40)
                                If WebForm.WebStat = "LOADED" Then
                                    ' MsgBox("c 'est ok")
                                    Exit Do
                                End If
                                System.Windows.Forms.Application.DoEvents()
                            Loop



                            'WebForm.TimerWeb.Interval = 1000 '1sec
                            'WebForm.TimerWeb.Start()

                            'Dim Start, Finish, Duree 'i, reponse, gagne, il_reste
                            'Duree = 40 'sec durée de pause (40sec)
                            ''Dim TimerWeb As New System.Timers.Timer '= New System.Timer
                            'Start = Timer ' top
                            'Finish = Start + Duree
                            'Do While Timer < Finish
                            '    If WebForm.WebStat = "LOADED" Then
                            '        ' MsgBox("c 'est ok")
                            '        Exit Do
                            '    End If
                            '    System.Windows.Forms.Application.DoEvents()
                            'Loop

                        Catch
                            MsgBox("Problème avec le Timer (Dans la commande Wurl)")
                            '  Control.CheckForIllegalCrossThreadCalls = True
                        End Try


                    ElseIf Prop(0).ToLower = "wfind" Then 'Wfind = a/input||Name/Value/href/class/id||*xyz*
                        Dim VarProp() As String = Split(Prop(1), "||")
                        If VarProp.Count = 3 Then
                            Dim theWElementCollection4 As System.Windows.Forms.HtmlElementCollection = WebForm.WebBrowser2.Document.GetElementsByTagName(VarProp(0))
                            For Each curElement As System.Windows.Forms.HtmlElement In theWElementCollection4
                                Dim controlValue As String = curElement.GetAttribute(VarProp(1)).ToString
                                If controlValue.ToLower Like VarProp(2).ToLower Then
                                    VarData.Add(controlValue.ToString) 'ajoute à la collection !!!
                                    Exit For
                                End If
                            Next
                        End If
                    ElseIf Prop(0).ToLower = "wsetname" Then  'WsetName = *Name*||nom
                        Dim VarProp() As String = Split(Prop(1), "||")
                        If VarProp.Count = 2 Then
                            Dim theWElementCollection4 As System.Windows.Forms.HtmlElementCollection = WebForm.WebBrowser2.Document.GetElementsByTagName("input")
                            For Each curElement As System.Windows.Forms.HtmlElement In theWElementCollection4
                                Dim controlName As String = curElement.GetAttribute("name").ToString
                                If controlName.ToLower Like VarProp(0).ToLower Then curElement.SetAttribute("value", VarProp(1))
                            Next
                        End If
                    ElseIf Prop(0).ToLower = "wsetvalue" Then  'WsetValue = *Pass*||xyz
                        Dim VarProp() As String = Split(Prop(1), "||")
                        If VarProp.Count = 2 Then
                            Dim theWElementCollection4 As System.Windows.Forms.HtmlElementCollection = WebForm.WebBrowser2.Document.GetElementsByTagName("input")
                            For Each curElement As System.Windows.Forms.HtmlElement In theWElementCollection4
                                Dim controlValue As String = curElement.GetAttribute("value").ToString
                                If controlValue.ToLower Like VarProp(0).ToLower Then curElement.SetAttribute("value", VarProp(1))
                            Next
                        End If
                    ElseIf Prop(0).ToLower = "wclickname" Then  'WclickName = *OK*               'Test Refrech  <<<
                        Dim theWElementCollection4 As System.Windows.Forms.HtmlElementCollection = WebForm.WebBrowser2.Document.GetElementsByTagName("input")
                        For Each curElement As System.Windows.Forms.HtmlElement In theWElementCollection4
                            Dim controlName As String = curElement.GetAttribute("name").ToString
                            If controlName.ToLower Like Prop(1).ToLower Then
                                curElement.InvokeMember("click")

                                WebForm.LoadPage("")

                                'Dim Start, Finish, Duree 'i, reponse, gagne, il_reste
                                'Duree = 40 'sec durée de pause
                                'Start = Timer ' top
                                'Finish = Start + Duree
                                'Do While Timer < Finish
                                '    If WebForm.WebStat = "LOADED" Then
                                '        Exit Do
                                '    End If
                                '    System.Windows.Forms.Application.DoEvents()
                                'Loop

                                Exit For
                            End If
                        Next
                    ElseIf Prop(0).ToLower = "wtable" Then  'Wtable = 1er table||dernière table||id

                        Dim VarProp() As String = Split(Prop(1), "||")
                        If VarProp.Count = 3 Then

                            Try
                                If LoopVarName <> "" Then 'Si Loop => vide la Wtable
                                    CollTableHTML.Clear() '!!! Vide TableHTML !!!
                                End If
                                Dim Tableautest As String = ""
                                Dim num As Double = 0
                                Dim theWElementCollection4 As System.Windows.Forms.HtmlElementCollection = WebForm.WebBrowser2.Document.GetElementsByTagName("table")

                                If VarProp(0).ToUpper = "DIVID" Then ' Wtable =  DIVID||nomDivID||id

                                    Dim ElementCollDiv As System.Windows.Forms.HtmlElementCollection = WebForm.WebBrowser2.Document.GetElementsByTagName("div")
                                    For Each curElement As System.Windows.Forms.HtmlElement In ElementCollDiv
                                        If curElement.Id IsNot Nothing Then
                                            If curElement.Id.ToLower = VarProp(1).ToLower Then
                                                theWElementCollection4 = curElement.GetElementsByTagName("table")
                                                VarProp(0) = 1
                                                VarProp(1) = 1000
                                            End If
                                        End If
                                    Next
                                End If

                                For Each curElement As System.Windows.Forms.HtmlElement In theWElementCollection4
                                    num += 1
                                    Dim numTR As Double = 0
                                    If num >= VarProp(0) And num <= VarProp(1) Then
                                        Dim ElementTR As System.Windows.Forms.HtmlElementCollection = curElement.GetElementsByTagName("tr")
                                        For Each curTR As System.Windows.Forms.HtmlElement In ElementTR
                                            numTR += 1
                                            Dim numTD As Double = 0
                                            Dim Col1 As String = ""
                                            Dim Col2 As String = ""
                                            Dim Col3 As String = ""
                                            Dim Col4 As String = ""
                                            Dim ElementTD As System.Windows.Forms.HtmlElementCollection = curTR.GetElementsByTagName("td")
                                            For Each curTD As System.Windows.Forms.HtmlElement In ElementTD
                                                numTD += 1

                                                'Suppresion des tableaux interne à cellule
                                                Dim NbreTab As Double = (curTD.GetElementsByTagName("table").Count)
                                                If NbreTab <> 0 Then
                                                    Dim InnerTDtext As String = ""
                                                    For Each InnerTab As System.Windows.Forms.HtmlElement In curTD.GetElementsByTagName("table")
                                                        For Each InnerTD As System.Windows.Forms.HtmlElement In InnerTab.GetElementsByTagName("td")
                                                            If InnerTD.GetElementsByTagName("table").Count = 0 Then
                                                                If Trim(InnerTDtext) <> "" Then InnerTDtext += "\P"
                                                                InnerTDtext += InnerTD.InnerText
                                                            End If
                                                        Next
                                                        Exit For
                                                    Next
                                                    curTD.InnerHtml = InnerTDtext
                                                End If

                                                Dim CurTDtext As String = curTD.OuterText

                                                'Ajout des données sur 4 colonnes (max)
                                                If numTD = 1 Then Col1 = CurTDtext
                                                If numTD = 2 Then Col2 = CurTDtext
                                                If numTD = 3 Then Col3 = CurTDtext
                                                If numTD = 4 Then Col4 = CurTDtext
                                            Next
                                            CollTableHTML.Add(New TableHTML(VarProp(2), num, numTR, Col1, Col2, Col3, Col4))
                                            '  Tableautest += num & ";" & numTR & ";" & numTD & ";" & Replace(Col1, vbCrLf, "") & ";" & Replace(Col2, vbCrLf, "") & ";" & Replace(Col3, vbCrLf, "") & ";" & Replace(Col4, vbCrLf, "") & vbCrLf ' TEST
                                        Next
                                    End If

                                Next

                                '    MsgBox(Tableautest)

                            Catch ex As Exception
                                MsgBox(ex.Message)
                            End Try
                        End If

                    ElseIf Prop(0).ToLower = "wfindt" Then  'WfindT = id||num table||*xyz*
                        Dim VarProp() As String = Split(Prop(1), "||")
                        If VarProp.Count = 3 Then
                            'Boucle dans les tableau html
                            Dim AddData As Boolean = False
                            For u = 1 To CollTableHTML.Count
                                Dim Col1 As String = ""
                                Dim Val As TableHTML = CollTableHTML(u)
                                If Val.Col1 = Nothing Then : Col1 = "" : Else : Col1 = Val.Col1 : End If
                                If Val.ID.ToLower = VarProp(0).ToLower And Val.NumTable Like VarProp(1).ToLower Then
                                    If InStr(VarProp(2), ">*") <> 0 Or InStr(VarProp(2), "*<") <> 0 Then
                                        If InStr(VarProp(2), "*<") <> 0 Then
                                            AddData = True
                                            If Col1.ToLower Like Replace(VarProp(2).ToLower, "*<", "") Then AddData = False
                                        End If
                                        If AddData Then
                                            Dim DataLigne As String = ""
                                            DataLigne = Col1
                                            If Val.Col2 <> "" Then DataLigne += vbTab & Val.Col2 : If Val.Col3 <> "" Then DataLigne += vbTab & Val.Col3 : If Val.Col4 <> "" Then DataLigne += vbTab & Val.Col4
                                            VarData.Add(DataLigne) ' Ajoute VALEUR
                                        End If
                                        If InStr(VarProp(2), ">*") <> 0 Then
                                            If Col1.ToLower Like Replace(VarProp(2).ToLower, ">*", "") Then AddData = True
                                        End If
                                    Else
                                        If Col1.ToLower Like VarProp(2).ToLower Then
                                            Dim DataLigne As String = ""
                                            DataLigne = Col1
                                            If Val.Col2 <> "" Then DataLigne += vbTab & Val.Col2 : If Val.Col3 <> "" Then DataLigne += vbTab & Val.Col3 : If Val.Col4 <> "" Then DataLigne += vbTab & Val.Col4
                                            VarData.Add(DataLigne)
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            Catch
                Connect.Message("Connexion internet", "Aucune connexion internet disponible.", False, 60, 100, "critical")
                Connect.RevoLog(Connect.DateLog & "Cmd VarHttp" & vbTab & True & vbTab & "Error http or web connexion")
            End Try

            Return VarData
        End Function

        ''' <summary>
        ''' Revo Command: Execute Autocad line command
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks></remarks>
        Public Function cmdCmd(ByVal Cmd As String)  ' Commande autocad à executer (exemple: "regen" )
            'Cmd;Commande ACAD
            ' 0       1
            '    Chaque bloc est séparé par ";"


            'EXEPCTION DE COMMANDE INTERNE
            If Mid(Cmd.ToUpper, 1, 15) = "#CMD;REVOPROPVD" Then
                PropRF("VD")
            Else

                Dim Cmdinfo() As String = SplitCmd(Cmd, 1)
                ' If Mid(Cmd.ToLower, 1, 4) = "#cmd" And Len(Cmd) > 5 Then
                'Cmd = Mid(Cmd, 6, Len(Cmd) - 5)
                Cmdinfo(1) = ReplaceVar(Cmdinfo(1))
                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Using docLock As DocumentLock = doc.LockDocument
                    'Cmdinfo(1) = Cmdinfo(1).Replace(" ", "§")
                    'Dim Cmds() As String
                    'Cmds = Split(Cmdinfo(1), "|")
                    'For Each cmdtxt As String In Cmds
                    'doc.SendStringToExecute(Replace(Cmdinfo(1), "|", vbCr) & vbCr, True, False, False)
                    'Next
                    'If Cmds.Length > 0 Then 'si la ligne n'est pas vide
                    'Dim oLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()

                    'Dim Cmds() As Object
                    ' Cmds = Split(Cmdinfo(1), "|")
                    'CommandLine.CmdsC(Cmds)
                    CommandLine.CmdsC(Cmdinfo(1))


                    'oLock.Dispose()
                    'End If
                    'End If
                End Using

            End If

            Return Connect.DateLog & "Cmd Cmd" & vbTab & True & vbTab & "Nom de la commande: " & Cmd & vbTab
        End Function

        ''' <summary>
        ''' Revo Command: Save or change name and de rights for draws
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <param name="FileData">Draw file name</param>
        ''' <remarks></remarks>
        Public Function cmdFile(ByVal Cmd As String, acDocMgr As DocumentCollection, NotCloseFile As Boolean) 'Traitement du fichier (renommer, copier, ...)

            ' #File;Nouveau ou /var/+"/FileName/;[[PropA]]Val;[[PropB]]Val;
            'Traitement du fichier (renommer, copier, ...). Executé en fin de script
            'Variables(requises)
            '0) Nom du fichier : Nouveau nom ou rien

            '[[Paramètre]]
            '1) Version = DWG2007 / DWG2010 / DXF2013
            '2) ReadOnly = 0/1

            Dim DocName As String = acDocMgr.MdiActiveDocument.Name.ToString
            Dim Cmdinfos() As String
            Dim FichierBak As String = ""
            Dim RetErr As String = ""

            Cmdinfos = SplitCmd(Cmd, 1)
            Dim FileName As String = ReplaceVar(Cmdinfos(1)) ' 0) Nom du fichier : Nouveau nom ou rien
            Dim FileVersion As String = ""
            Dim FileReadOnly As String = ""

            If Cmdinfos.Count > 2 Then
                For i = 2 To Cmdinfos.Count - 1 'Boucle prop
                    If InStr(Cmdinfos(i), "[[") <> 0 And InStr(Cmdinfos(i), "]]") <> 0 Then
                        Cmdinfos(i) = Replace(Replace(Cmdinfos(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                        Dim TabProp() As String = SplitCmd(Cmdinfos(i), 1)
                        TabProp(0) = TabProp(0).ToLower
                        If "version" = TabProp(0) Then '1) Version = DWG2007 / DWG2010 / DXF2013
                            FileVersion = TabProp(1).ToUpper
                        ElseIf "readonly" = TabProp(0) Then '2) ReadOnly = 0/1
                            FileReadOnly = Trim(TabProp(1))
                        End If
                    End If
                Next
            End If


            If Trim(FileName) <> "" Then DocName = FileName 'Sauver sous un autre nom (copy file)

            'Selection du format d'enregistrement
            Dim DWGVers As DwgVersion = DwgVersion.Current
            If FileVersion <> "" Then
                Select Case FileVersion
                    Case "DWG2007"
                        DWGVers = DwgVersion.AC1021
                    Case "DWG2010"
                        DWGVers = DwgVersion.AC1024
                    Case "DWG2013"
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                        'Error file format (DWG2013): Is not supported
                        Connect.RevoLog(Connect.DateLog & "Cmd File" & vbTab & False & vbTab & "Error file format: DWG2013" & "Is not supported" & " (" & DocName & ")")
#Else 'Version AutoCAD 2013 et +
                        DWGVers = DwgVersion.AC1027
#End If
                End Select
            End If

            Dim SavedFile As Boolean = CDbl(Application.GetSystemVariable("DWGTITLED"))
            'Uniquement pour format DWG (pas de dxf) + ferme le fichier -----------------
            If Right(DocName.ToUpper, 4) = ".DWG" And NotCloseFile = False And SavedFile = True Then

                'Saved + bak => Si déjà sauvé
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                Dim acadDoc As Object = acDocMgr.MdiActiveDocument.AcadDocument
                acadDoc.Save()
#Else 'Versio AutoCAD 2013 et +
                Dim acadDoc As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
                acadDoc.Save()
#End If

                Dim DataCopy As String = Mid(DocName, 1, DocName.Length - 4) & "_RevoBAK.dwg"
                'Changer de format  (AC1018 AutoCAD 2004/2005/2006 - AC1015 AutoCAD 2000/2000i/2002)
                acDocMgr.MdiActiveDocument.Database.SaveAs(DataCopy, DWGVers) 'Save As RevoBaK

                Try ' Close the file

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                    acDocMgr.MdiActiveDocument.CloseAndDiscard()
#Else 'Versio AutoCAD 2013 et +
                    Autodesk.AutoCAD.ApplicationServices.DocumentExtension.CloseAndDiscard(acDocMgr.MdiActiveDocument)
#End If

                    'Rename files
                    '  Revo.RevoFiler.RenFile(DocName, Mid(DocName, 1, DocName.Length - 4) & ".bak2")
                    Revo.RevoFiler.RenFile(DataCopy, DocName)

                    Dim Fi As IO.FileInfo = New IO.FileInfo(DocName)
                    If Fi.IsReadOnly = True Then
                        ' Fichier en lectur seul
                        If FileReadOnly = "0" Then Fi.IsReadOnly = False '-> write
                    Else ' n'est pas en lecture seul
                        If FileReadOnly = "1" Then Fi.IsReadOnly = True ' -> read
                    End If

                Catch ex1 As System.NullReferenceException
                    RetErr = Connect.DateLog & "Cmd File" & vbTab & False & vbTab & "Sauvegarde de dessin" & vbTab & DocName
                Catch ex As COMException
                    RetErr = Connect.DateLog & "Cmd File" & vbTab & False & vbTab & "Commande en cours d'éxecution: " & ex.ErrorCode & vbTab & DocName
                End Try

            ElseIf Right(DocName.ToUpper, 4) = ".DWG" And NotCloseFile = True Then

                'Saved + bak => Si pas encore sauvé
                If CDbl(Application.GetSystemVariable("DWGTITLED")) = 0 And FileName <> "" Then
                    Dim DataCopy As String = DocName
                    'Changer de format  (AC1018 AutoCAD 2004/2005/2006 - AC1015 AutoCAD 2000/2000i/2002)
                    acDocMgr.MdiActiveDocument.Database.SaveAs(FileName, DWGVers) 'Save As RevoBaK
                End If

            Else 'Ne traite pas le fichier car pas DWG
                Connect.RevoLog(Connect.DateLog & "Cmd File" & vbTab & False & vbTab & "Error file format: " & "Only DWG" & " (" & DocName & ")")
            End If


            If RetErr = "" Then RetErr = Connect.DateLog & "Cmd File" & vbTab & True & vbTab & "Lecture seul: " & FileReadOnly & vbTab & DocName
            Return RetErr
        End Function

        ''' <summary>
        ''' Revo Command: zoom in draws
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks>
        ''' #Zoom;Extents/Scale=1
        ''' </remarks>
        Public Function cmdZoom(ByVal Cmd As String)  ' Zoom in draw with options

            Dim Cmdinfo() As String
            Dim curDwg As Document = Application.DocumentManager.MdiActiveDocument

            Cmdinfo = SplitCmd(Cmd, 1)


            'Execute command
            If Cmdinfo(1).ToUpper = "EXTENTS" Then
                Zooming.ZoomExtents()

            ElseIf Cmdinfo(1) <> 0 And IsNumeric(Cmdinfo(1)) Then
                '' Get the current document
                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                '' Get the current view
                Using acView As ViewTableRecord = acDoc.Editor.GetCurrentView()
                    '' Get the center of the current view
                    Dim pCenter As Point3d = New Point3d(acView.CenterPoint.X, acView.CenterPoint.Y, 0)
                    '' Set the scale factor to use
                    Dim dScale As Double = 1
                    '' Scale the view using the center of the current view
                    Zooming.Zoom(New Point3d(), New Point3d(), pCenter, 1 / dScale)
                End Using

            Else 'ignore the command
                Return Connect.DateLog & "Cmd Zoom" & vbTab & False & vbTab & "Pas de paramètre" & vbTab
            End If

            Return Connect.DateLog & "Cmd Zoom" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab

        End Function

        ''' <summary>
        ''' Revo Command: Purge in draws
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks>
        ''' Cmd;'#Purge;all/block
        ''' </remarks>
        Public Function cmdPurge(ByVal Cmd As String)  ' Purge the draw with options

            Dim Cmdinfo() As String
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim curDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim curDWG As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            '' Get the current document and database
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Cmdinfo = SplitCmd(Cmd, 1)

            If Cmdinfo(1).ToUpper = "ALL" Then 'purge all
                'Execute command
                curDwg.PurgeAll()
                curDwg.PurgeAll()
                Application.DocumentManager.MdiActiveDocument.Editor.Regen()

            ElseIf Cmdinfo(1).ToUpper = "BLOCK" Then 'purge block
                Dim acBL As AcadBlock

                For Each acBL In curDwg.Blocks
                    'MsgBox(acBL.Name)

                    If acBL.IsLayout = False Then
                        Try 'MsgBox(acBL.Name)
                            acBL.Delete()
                        Catch ex As COMException
                            ' MsgBox("Un bloc n'a pu être supprimé")
                        End Try

                    End If
                Next

            ElseIf Left(Cmdinfo(1).ToUpper, 4) = "BL||" Then 'purge block avec *nom
                Dim acBL As AcadBlock
                For Each acBL In curDwg.Blocks
                    If acBL.IsLayout = False Then
                        Try
                            If acBL.Name.ToUpper Like Replace(Cmdinfo(1).ToUpper, "BL||", "") Then
                                acBL.Delete()
                            End If

                        Catch ex As COMException
                            ' Un bloc n'a pu être supprimé")
                        End Try
                    End If
                Next

            ElseIf Left(Cmdinfo(1).ToUpper, 4) = "LA||" Then 'purge layer avec *nom
                Dim strLayer As String = Replace(Cmdinfo(1).ToUpper, "LA||", "")
                DeleteLayer(strLayer)

            Else 'ignore the command
                Return Connect.DateLog & "Cmd Purge" & vbTab & False & vbTab & "Pas de paramètre" & vbTab
            End If

            Return Connect.DateLog & "Cmd Purge" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab

        End Function

        ''' <summary>
        ''' Revo Command: PaperSpace 
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks></remarks>

        Public Function cmdPS(ByVal Cmd As String)  ' Remarks
            '#PS;NomPRE;[[PropA]]Val;[[PropB]]Val;
            'Importation de présentation avec une mise en forme
            '
            'Balise
            '   0   NomPrésentation = "Like" ( ">" : import)
            '
            '[[Propriété]]
            '   1   Name = "New name"
            '   2   URL = "*.dwt"
            '   3   WinCenter = "x,y"/all
            '   4   WinScale = "0.5"
            '   5   WinRotation = dbl (défault degré ou g = grade)
            '   6   PaperScale = dbl (rapport d'impression : XXX / 1 unité dessin)
            '   7   CTB = "acad.ctb" (nom du CTB avec extention ou Aucune = "None")

            Dim PSprop(0 To 9) As String
            Dim Cmdinfo() As String
            Dim TabProp() As String
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim curDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim curDWG As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim NewCollLayouts As New Collection
            Cmdinfo = SplitCmd(Cmd, 1) '1 = min de paramètre obligatoire

            If Cmdinfo(1) <> "" Then 'si la ligne n'est pas vide
                'Chargement des propriétés
                PSprop(0) = Cmdinfo(1) ' Orignial Name / current
                If PSprop(0).ToLower = "current" Then PSprop(0) = LayoutManager.Current.CurrentLayout

                For i = 2 To Cmdinfo.Count - 1 'Boucle prop
                    If InStr(Cmdinfo(i), "[[") <> 0 And InStr(Cmdinfo(i), "]]") <> 0 Then
                        Cmdinfo(i) = Replace(Replace(Cmdinfo(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                        TabProp = SplitCmd(Cmdinfo(i), 1)
                        TabProp(0) = TabProp(0).ToLower
                       
                        If "name" = TabProp(0) Then ' 1  Name = ""  Nouveau nom
                            PSprop(1) = TabProp(1)
                        ElseIf "url" = TabProp(0) Then ' 2  URL = ""  URL
                            PSprop(2) = TabProp(1)
                        ElseIf "wincenter" = TabProp(0) Then ' 3   WinCenter = "x,y"/all
                            PSprop(3) = TabProp(1)
                        ElseIf "winscale" = TabProp(0) Then '  4   WinScale = "0.5"
                            PSprop(4) = TabProp(1)
                        ElseIf "winrotation" = TabProp(0) Then '  5   WinRotation = dbl (défault degré ou g = grade)
                            PSprop(5) = TabProp(1)
                        ElseIf "paperscale" = TabProp(0) Then '  6   PaperScale = dbl (rapport d'impression : XXX / 1 unité dessin)
                            PSprop(6) = TabProp(1)
                        ElseIf "ctb" = TabProp(0) Then '  7   CTB = "acad.ctb" (nom du CTB avec extention ou Aucune = "None")
                            PSprop(7) = TabProp(1)
                        End If
                    End If
                Next

                'Chargement des layouts  'analyse des présentations actuelles
                Dim CollLayouts As New Collection
                Try
                    For Each acLay As AcadLayout In curDwg.Layouts
                        If acLay.Name <> "Model" Then CollLayouts.Add(acLay)
                    Next
                Catch ex As COMException
                    Connect.RevoLog(Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Error read blocks: " & curDwg.Name)
                End Try

                Dim TestTemplateExist As Boolean = True ' Crée variable de test pour l'existance du fichier template

                'Importation de la présetantation : ">"
                If Mid(PSprop(0), 1, 1) = ">" And InStr(PSprop(0), "*") = 0 Then
                    Dim NouvNom As String = Replace(PSprop(0), ">", "", 1, 1)
                    If NouvNom <> "" And File.Exists(PSprop(2)) Then

                        Try ' Renommer tous les présentations
                            For Each Lay As AcadLayout In CollLayouts
                                Lay.Name = "[[" & Lay.ObjectID & "]]" & Lay.Name
                            Next
                        Catch
                            Connect.RevoLog(Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Error tempory rename layouts: " & curDwg.Name)
                        End Try

                        'importation de présentation ou plusieurs (like = *)
                        importLayouts(PSprop(2), NouvNom)

                        Try 'recharge toutes les présentations
                            CollLayouts.Clear()
                            For Each Lay As AcadLayout In curDwg.Layouts
                                If Lay.Name <> "Model" Then CollLayouts.Add(Lay)
                            Next
                        Catch ex As COMException
                            Connect.RevoLog(Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Error read2 blocks: " & curDwg.Name)
                        End Try

                        'Recherche les nouvelles et renomment
                        NewCollLayouts.Clear()
                        For Each Lay As AcadLayout In CollLayouts
                            Try
                                If InStr(Lay.Name, "[[" & Lay.ObjectID & "]]") = 0 Then
                                    NewCollLayouts.Add(Lay)
                                    Lay.Name = LayoutNametest(CollLayouts, Lay.Name, "[[" & Lay.ObjectID & "]]")
                                End If
                            Catch
                                Connect.RevoLog(Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Error rename layout1: " & Lay.Name)
                            End Try
                        Next

                        'Renommer les anciennes présentation
                        For Each Lay As AcadLayout In CollLayouts
                            Try
                                If InStr(Lay.Name, "[[" & Lay.ObjectID & "]]") <> 0 Then
                                    Lay.Name = Replace(Lay.Name, "[[" & Lay.ObjectID & "]]", "")
                                End If
                            Catch
                                Connect.RevoLog(Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Error rename layout2: " & Lay.Name)
                            End Try
                        Next

                        'Activation de la présentation importée
                        Dim newlm As LayoutManager = LayoutManager.Current
                        For Each Lay As AcadLayout In NewCollLayouts
                            If newlm.CurrentLayout <> Lay.Name Then newlm.CurrentLayout = Lay.Name
                        Next

                    Else
                        MsgBox("Le fichier modèle (.dwt) est introuvable." & vbCrLf & "Attention, la suite du script ne fonctionnera pas correctement", MsgBoxStyle.Critical, "Fichier introuvable")
                        TestTemplateExist = False
                        Connect.RevoLog(Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Template untraceable : " & vbTab & curDwg.Name)
                        'TestTemplateExist = True
                    End If
                End If



                If Mid(PSprop(0), 1, 1) = "<" Then 'Delete Layouts
                    PSprop(0) = PSprop(0).Replace("<", "")
                    
                    Using tr As Transaction = ed.Document.TransactionManager.StartTransaction()
                        For Each Lay As AcadLayout In CollLayouts
                            Try
                                If Lay.Name Like PSprop(0) Then
                                    LayoutManager.Current.DeleteLayout(Lay.Name)
                                End If
                            Catch
                                Connect.RevoLog(Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Error delete : " & Lay.Name)
                            End Try
                        Next
                        tr.Commit()
                    End Using

                End If


                If TestTemplateExist = True Then

                        'Renommer et Modifier la présentation 
                        Dim lm As LayoutManager = LayoutManager.Current
                        For Each lay As AcadLayout In CollLayouts
                            Try
                                If lay.Name Like Replace(PSprop(0), ">", "", 1, 1) Then
                                    If Mid(PSprop(0), 1, 1) <> ">" Then
                                        If lm.CurrentLayout <> lay.Name Then lm.CurrentLayout = lay.Name
                                    End If

                                    If PSprop(1) <> Nothing Then '   1   Name = "New name"
                                        If LayoutNameExist(CollLayouts, PSprop(1)) Then
                                            Connect.RevoLog(Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Error rename layout3: " & lay.Layout.Name)
                                        Else
                                            lay.Layout.Name = PSprop(1)
                                        End If
                                    End If

                                    If PSprop(5) <> Nothing Then '   5   WinRotation = dbl (défault degré ou g = grade)
                                        Try
                                            Dim WinRot As Double = 0
                                            If InStr(PSprop(5).ToLower, "g") <> 0 Then ' = grade
                                                WinRot = CDbl(Replace(PSprop(5), "g", ""))
                                                ed.SwitchToModelSpace()
                                                cmdCmd("#Cmd;_UCS|Z|" & WinRot & "|_plan|_c")
                                                ed.SwitchToPaperSpace()
                                            Else ' défault degré
                                                WinRot = CDbl(PSprop(5))
                                                ed.SwitchToModelSpace()
                                                cmdCmd("#Cmd;_UCS|Z|" & WinRot & "|_plan|_c")
                                                ed.SwitchToPaperSpace()
                                            End If

                                        Catch
                                        End Try
                                    End If

                                    If PSprop(3) <> Nothing Then '   3   WinCenter = "x,y"/all
                                        Try
                                            If PSprop(3).ToLower = "all" Then
                                                ed.SwitchToModelSpace()
                                                Zooming.ZoomExtents()
                                                ed.SwitchToPaperSpace()
                                            Else
                                                Dim XYsplit() As String = Split(PSprop(3), ",")
                                                If XYsplit.Count = 2 Then
                                                    ed.SwitchToModelSpace()
                                                    Dim Pts As New Point3d(XYsplit(0), XYsplit(1), 0)
                                                    Zooming.ZoomWindow(Pts, Pts)
                                                    ed.SwitchToPaperSpace()
                                                End If
                                            End If

                                        Catch
                                        End Try
                                    End If

                                    If PSprop(4) <> Nothing Then '   4   WinScale = "0.5" = 1:2000
                                        Try
                                            If IsNumeric(PSprop(4)) Then
                                                ed.SwitchToModelSpace()
                                                Dim oLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                                                Dim vpid As ObjectId = ed.CurrentViewportObjectId
                                                Using tr As Transaction = ed.Document.TransactionManager.StartTransaction()
                                                    Dim vport As Viewport = DirectCast(tr.GetObject(vpid, DatabaseServices.OpenMode.ForWrite), Viewport)
                                                    'ed.WriteMessage(vport.CustomScale.ToString())
                                                    vport.CustomScale = CDbl(PSprop(4))
                                                    tr.Commit()
                                                End Using
                                                oLock.Dispose()
                                                ed.SwitchToPaperSpace()
                                            End If
                                        Catch
                                        End Try
                                    End If


                                    'Modifcation du réglage de l'impression ------------------------------

                                    If PSprop(6) <> "" Then ' 6   PaperScale = dbl (rapport d'impression : XXX / 1 unité dessin)
                                        If IsNumeric(PSprop(6)) Then
                                            PlotCurrentLayout("", "", PSprop(6), "1", "", True, False)
                                        End If
                                    End If

                                    If PSprop(7) <> "" Then '   7   CTB = "acad.ctb" (nom du CTB avec extention ou Aucune = "None")
                                        PlotCurrentLayout("", "", "", "", PSprop(7), True, False)
                                    End If


                                End If
                            Catch ex As COMException
                                ' MsgBox("Un bloc n'a pu être supprimé")
                            End Try
                        Next


                        curDWG.Regen(AcRegenType.acAllViewports)
                    Else
                        Return Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Erreur: " & vbTab & curDWG.Name
                    End If
                Else 'ignore the command
                    Return Connect.DateLog & "Cmd PS" & vbTab & False & vbTab & "Erreur: " & vbTab & curDwg.Name
            End If


            Return Connect.DateLog & "Cmd PS" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab & curDwg.Name

            '   ed.SwitchToModelSpace()
            'AcadDocument.Plot.PlotToFile(DestinationPath)
            'http://forums.autodesk.com/t5/NET/AcadDocument-Plot-PlotToFile-DestinationPath/m-p/2637611/highlight/true#M17984

        End Function

        Function LayoutNametest(ByVal CollLayout As Collection, ByVal LayoutName As String, Optional ByVal IDobj As String = "", Optional ByVal Incr As Double = 0)

            Dim Num As Double = Incr + 1
            Dim NewName As String = Replace(LayoutName, IDobj, "")

            For Each Lay In CollLayout
                If "[[" & Lay.ObjectID.ToString & "]]" <> IDobj Then
                    Dim TempName As String = Replace(Lay.Name, "[[" & Lay.ObjectID.ToString & "]]", "")
                    If TempName.ToLower = LayoutName.ToLower Then

                        NewName = TempName & " (" & Num.ToString & ")"
                        If LayoutNameExist(CollLayout, NewName) Then
                            For Each LayX As AcadLayout In CollLayout
                                If "[[" & LayX.ObjectID.ToString & "]]" <> IDobj Then
                                    Num += 1
                                    NewName = TempName & " (" & Num & ")"
                                    If LayoutNameExist(CollLayout, NewName) Then
                                    Else
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                    End If
                End If
            Next

            Return NewName
        End Function
        Function LayoutNameExist(ByVal CollLayout As Collection, ByVal LayoutName As String)
            Dim Resultat As Boolean = False

            For Each Lay As AcadLayout In CollLayout
                Dim ActName As String = Replace(Lay.Name, "[[" & Lay.ObjectID & "]]", "")
                If ActName.ToLower = LayoutName.ToLower Then
                    Resultat = True
                End If
            Next

            Return Resultat
        End Function
        Function importLayouts(ByVal fileName As String, ByVal LayoutName As String)
            Dim oLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
            Dim dbSource As New Database(False, False)
            Dim db As Database = HostApplicationServices.WorkingDatabase

            Try
                Using trans As Transaction = db.TransactionManager.StartTransaction()
                    dbSource.ReadDwgFile(fileName, System.IO.FileShare.Read, True, Nothing)
                    Dim idDbdSource As ObjectId = dbSource.LayoutDictionaryId
                    Dim dbdLayout As DBDictionary = DirectCast(trans.GetObject(idDbdSource, DatabaseServices.OpenMode.ForRead, False, False), DBDictionary)
                    Dim idLayout As ObjectId

                    Dim idc As New ObjectIdCollection()
                    For Each deLayout As DictionaryEntry In dbdLayout

                        If deLayout.Key.ToString().ToUpper() <> "MODEL" Then
                            If deLayout.Key.ToString.ToLower Like LayoutName.ToLower Then

                                idLayout = DirectCast(deLayout.Value, ObjectId)
                                If Not idLayout.IsErased Then
                                    idc.Add(idLayout)
                                End If

                            End If

                        End If
                    Next
                    Dim im As New IdMapping()
                    db.WblockCloneObjects(idc, db.LayoutDictionaryId, im, DuplicateRecordCloning.Ignore, False)
                    trans.Commit()
                End Using

            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                MsgBox(ex.ToString)
            Catch ex As System.Exception

                MsgBox(ex.ToString)
            Finally

                dbSource.Dispose()
                oLock.Dispose()
            End Try

            Return True
        End Function
        Public Shared Function importLayout2(ByVal filename As String, ByVal layoutname As String) As ObjectId
            Dim idLayout As ObjectId = ObjectId.Null

            Dim oLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
            Dim dbSource As New Database(False, False)
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                dbSource.ReadDwgFile(filename, System.IO.FileShare.Read, True, Nothing)
                Dim idDbdSource As ObjectId = dbSource.LayoutDictionaryId
                Dim dbdLayout As DBDictionary = DirectCast(trans.GetObject(idDbdSource, DatabaseServices.OpenMode.ForRead, False, False), DBDictionary)
                For Each deLayout As DictionaryEntry In dbdLayout
                    Dim sLayout As String = deLayout.Key.ToString()
                    If sLayout.ToUpper() Like layoutname.ToUpper() Then
                        idLayout = DirectCast(deLayout.Value, ObjectId)
                        Exit For
                    End If
                Next
                If idLayout <> ObjectId.Null Then
                    Dim idc As New ObjectIdCollection()
                    idc.Add(idLayout)
                    Dim im As New IdMapping()
                    db.WblockCloneObjects(idc, db.LayoutDictionaryId, im, DuplicateRecordCloning.MangleName, False)
                    Dim ip As IdPair = im.Lookup(idLayout)
                    idLayout = ip.Value

                    Dim lm As LayoutManager = LayoutManager.Current
                    lm.CurrentLayout = layoutname
                    ed.Regen()
                End If
                trans.Commit()
            Catch ex As Autodesk.AutoCAD.Runtime.Exception
                trans.Dispose()
                MsgBox(ex.ToString())
            Catch ex As System.Exception
                trans.Dispose()
                MsgBox(ex.ToString())
            Finally
                dbSource.Dispose()
                oLock.Dispose()
            End Try
            Return idLayout
        End Function

        ''' <summary>
        ''' Revo Command: Layer modifier
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks></remarks>
        Public Function cmdLA(ByVal Cmd As String)  ' Remarks
            ' #LA;NomOriginal;[[PropA]]Val;[[PropB]]Val;
            ' 0    >Nom : ">" force la création ou  "<" : supp du calque
            ' 1            Name = ""  Nouveau nom
            ' 2            LayerOn = True
            ' 3            Freeze = True
            ' 4            Lock = True
            ' 5            TrueColor = dbl(1-255)/R,V,B
            ' 6            Linetype = ""
            ' 7            LineWeight = dbl
            ' 8            Plottable = True
            ' 9            Description = ""
            '10            Filter = ">Name" ( ">" : new et "<" : supp)

            Dim LAprop(0 To 12) As String
            Dim Cmdinfo() As String
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDoc As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Dim CurLA As AcadLayer
            Dim LAtest As String = ""
            Dim TabProp() As String

            Dim LAColor As New AcadAcCmColor
            Cmdinfo = SplitCmd(Cmd, 1)
            LAColor.ColorIndex = 7 'Color.FromColorIndex(ColorMethod.ByAci, 7)

            If Cmdinfo(1) <> "" Then 'si le nom du calque n'est pas vide
                LAprop(0) = Cmdinfo(1)

                'Création du calques
                If Mid(LAprop(0), 1, 1) = ">" And InStr(LAprop(0), "*") = 0 Then
                    LAprop(0) = Replace(LAprop(0), ">", "", 1, 1)
                    Try
                        CurLA = acDoc.Layers.Add(LAprop(0))
                    Catch
                        'Write log error
                        Connect.RevoLog(Connect.DateLog & "Cmd LA" & vbTab & False & vbTab & "Erreur create layer: " & LAprop(0))
                    End Try
                Else
                    LAprop(0) = Replace(LAprop(0), ">", "", 1, 1)
                End If

                'Suppresion du calques
                If Mid(LAprop(0), 1, 1) = "<" Then
                    LAprop(0) = Replace(LAprop(0), "<", "", 1, 1)
                    DeleteLayer(LAprop(0)) ' Avec * = like

                Else 'si pas de suppression continue le traitement normal

                    For i = 2 To Cmdinfo.Count - 1 'Boucle prop
                        If InStr(Cmdinfo(i), "[[") <> 0 And InStr(Cmdinfo(i), "]]") <> 0 Then
                            Cmdinfo(i) = Replace(Replace(Cmdinfo(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                            TabProp = SplitCmd(Cmdinfo(i), 1)
                            TabProp(0) = TabProp(0).ToLower

                            If "name" = TabProp(0) Then ' 1  Name = ""  Nouveau nom
                                LAprop(1) = TabProp(1)
                            ElseIf "layeron" = TabProp(0) Then ' 2  LayerOn = True
                                If IsNumeric(TabProp(1)) Then LAprop(2) = TabProp(1)
                            ElseIf "freeze" = TabProp(0) Then ' 3   Freeze = True
                                If IsNumeric(TabProp(1)) Then LAprop(3) = TabProp(1)
                            ElseIf "lock" = TabProp(0) Then ' 4   Lock = True
                                If IsNumeric(TabProp(1)) Then LAprop(4) = TabProp(1)
                            ElseIf "truecolor" = TabProp(0) Then ' 5   TrueColor = dbl(1-255)/R,V,B
                                LAprop(5) = TabProp(1)
                            ElseIf "linetype" = TabProp(0) Then ' 6   Linetype = ""
                                LAprop(6) = TabProp(1)
                            ElseIf "lineweight" = TabProp(0) Then ' 7   LineWeight = dbl
                                If IsNumeric(TabProp(1)) Then LAprop(7) = TabProp(1)
                            ElseIf "plottable" = TabProp(0) Then ' 8   Plottable = True
                                If IsNumeric(TabProp(1)) Then LAprop(8) = TabProp(1)
                            ElseIf "description" = TabProp(0) Then ' 9   Description = ""
                                LAprop(9) = TabProp(1)
                            ElseIf "filter" = TabProp(0) Then '10 Filter = ">Name" ( ">" : new et "<" : supp)
                                LAprop(10) = TabProp(1)
                            End If
                        End If
                    Next

                    'New or delete layer Filter
                    If LAprop(10) IsNot Nothing Then

                        'Suppresion
                        Dim AcadLayer As New modAcadLayer
                        If Left(LAprop(10), 1) = "<" Then
                            LAprop(10) = Replace(LAprop(10), "<", "", 1, 1)
                            AcadLayer.DeleteLayerFilter(LAprop(10))
                        ElseIf Left(LAprop(10), 1) = ">" Then 'Ajout de filtre (">")
                            LAprop(10) = Replace(LAprop(10), ">", "", 1, 1)
                            AcadLayer.CreateLayerFilters(LAprop(10), LAprop(0))
                        Else   'rien : pour filtre de groupe ?
                            AcadLayer.CreateLayerGroup(LAprop(10), LAprop(0))
                        End If

                    Else 'No Filter created

                        'Boucle dans les calques
                        Dim Calques As AcadLayers = acDoc.Layers
                        Dim LA As AcadLayer

                        For Each LA In Calques
                            Try
                                If LA.Name.ToUpper Like LAprop(0).ToUpper = True Then 'Test du nom de calques
                                    'Si trouve le calque > traite au Param.

                                    If LAprop(1) <> Nothing And LA.Name <> "0" Then ' 1  Name = ""  Nouveau nom
                                        Try
                                            If InStr(LAprop(1), "*") = 0 Then
                                                LA.Name = LAprop(1)
                                            Else
                                                If Left(LAprop(1), 1) = "*" And Right(LAprop(1), 1) = "*" Then
                                                    LA.Name = Replace(LA.Name, Replace(LAprop(0), "*", ""), Replace(LAprop(1), "*", ""))   ' *xyz* -> *abc*
                                                ElseIf Left(LAprop(1), 1) <> "*" And Right(LAprop(1), 1) = "*" Then
                                                    LA.Name = Replace(LAprop(1), "*", "") & LA.Name                       ' abc***
                                                ElseIf Left(LAprop(1), 1) = "*" And Right(LAprop(1), 1) <> "*" Then
                                                    LA.Name = LA.Name & Replace(LAprop(1), "*", "")                       ' ***abc
                                                Else
                                                    LA.Name = Replace(LAprop(1), "*", "") 'Nouveau nom
                                                End If
                                            End If
                                        Catch
                                            Connect.RevoLog(Connect.DateLog & "Cmd LA" & vbTab & False & vbTab & "Erreur rename layer: " & LAprop(1))
                                        End Try
                                    End If
                                    If LAprop(2) <> Nothing Then LA.LayerOn = CBool(LAprop(2)) ' 2  LayerOn = True
                                    If LAprop(3) <> Nothing Then
                                        Try
                                            LA.Freeze = CBool(LAprop(3)) ' 3   Freeze = True
                                        Catch '("Current layer " & LA.Name)
                                        End Try
                                    End If

                                    If LAprop(4) <> Nothing Then LA.Lock = CBool(LAprop(4)) ' 4   Lock = True
                                    If LAprop(5) <> Nothing Then
                                        If LAprop(5) <> "" Then ' 5   TrueColor = dbl
                                            If IsNumeric(LAprop(5)) And LAprop(5) > 0 And LAprop(5) < 256 Then
                                                LAColor.ColorIndex = CDbl(Trim(LAprop(5)))
                                                LA.TrueColor = LAColor ' 5   TrueColor = dbl
                                            Else
                                                Dim RBVcolor() As String
                                                RBVcolor = Split(LAprop(5), ",")
                                                If RBVcolor.Length = 3 Then
                                                    If IsNumeric(RBVcolor(0)) And IsNumeric(RBVcolor(1)) And IsNumeric(RBVcolor(2)) Then
                                                        Try
                                                            LAColor.SetRGB(RBVcolor(0), RBVcolor(1), RBVcolor(2))
                                                            LA.TrueColor = LAColor
                                                        Catch
                                                            Connect.RevoLog(Connect.DateLog & "Command ColorRVB" & vbTab & False & vbTab & "Erreur color : " & LAprop(5))
                                                        End Try
                                                    End If
                                                End If
                                            End If
                                        End If
                                    End If
                                    If LAprop(6) <> Nothing Then
                                        Try
                                            LA.Linetype = LAprop(6) ' 6   Linetype = ""
                                        Catch
                                            'Write log error
                                            Connect.RevoLog(Connect.DateLog & "Cmd LA" & vbTab & False & vbTab & "Error Line type: " & vbTab & LAprop(6))
                                        End Try
                                    End If

                                    If LAprop(7) <> Nothing Then LA.Lineweight = ConvertLineWeight(LAprop(7)) 'Lineweight = dbl (ByLayer/ByBlock/ByDefault)
                                    If LAprop(8) <> Nothing Then LA.Plottable = LAprop(8) ' 8   Plottable = True
                                    If LAprop(9) <> Nothing Then LA.Description = LAprop(9) ' 9   Description = ""

                                Else
                                    'Ne traite pas le calque
                                End If

                            Catch ex As Exception
                                MsgBox("Erreur de calque" & vbCrLf & ex.Message)
                            End Try
                        Next
                    End If

                End If
            Else 'ignore the command
                Return Connect.DateLog & "Cmd LA" & vbTab & False & vbTab & "Pas de paramètre" & vbTab
            End If

            Return Connect.DateLog & "Cmd LA" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab & acDoc.Name

        End Function

        ''' <summary>
        ''' Revo Command: Block modifier
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks>#BL;NomOriginal;[[PropA]]Val;</remarks>
        Public Function cmdBL(ByVal Cmd As String)  ' Remarks
            ' 0    >Nom : ">" force la création ( ">" : new)
            ' 1            Name = "" "abc"   
            ' 2            XInsertionPoint = dbl
            ' 3            YInsertionPoint = dbl
            ' 4            ZInsertionPoint = dbl
            ' 5            Rotation = dbl
            ' 6            XScaleFactor = dbl
            ' 7            YScaleFactor = dbl
            ' 8            ZScaleFactor = dbl
            ' 9            TrueColor = dbl/ByLayer/ByBlock
            '10            Layer = ""
            '11            Lineweight = dbl
            '12            space = ""
            '13            XYZInsertionPoint = dbl,dbl,dbl
            '14            Annotative = 1 / 0
            '15            PaperOrientation = 1 / 0
            '16            Uniformly = 1 / 0 (1 = uniforme scale)
            '17            Xplode = 1 / 0 (1 = authorize decompos)

            'XX            Attribut / propriété personnalisé

            Dim BLprop(0 To 20) As String  ' Propriété du block
            Dim Collatt As New Collection  ' Attribut ou propriété personalisé
            Dim Cmdinfo() As String
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDoc As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Dim TabProp() As String
            Dim BLColor As Color
            Dim BLidSelect As Boolean = False
            BLColor = Color.FromColorIndex(ColorMethod.ByAci, 7)

            Dim Ptinsert(2) As Double
            Ptinsert(0) = 0 : Ptinsert(1) = 0 : Ptinsert(2) = 0
            Cmdinfo = SplitCmd(Cmd, 1)

            If Cmdinfo(1) <> "" Then 'si le nom du block n'est pas vide

                'Remplacement des Var dans le nom du block
                Cmdinfo(1) = ReplaceVar(Cmdinfo(1))
                BLprop(0) = Cmdinfo(1)

                For i = 2 To Cmdinfo.Count - 1 'Boucle prop
                    If InStr(Cmdinfo(i), "[[") <> 0 And InStr(Cmdinfo(i), "]]") <> 0 Then
                        Cmdinfo(i) = Replace(Replace(Cmdinfo(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                        TabProp = SplitCmd(Cmdinfo(i), 1)
                        TabProp(0) = TabProp(0).ToLower
                        TabProp(1) = Replace(TabProp(1), "¦", ";")

                        If "name" = TabProp(0) Then ' 1  Name = ""  Nouveau nom
                            BLprop(1) = TabProp(1)
                        ElseIf "xinsertionpoint" = TabProp(0) Then ' 2   xinsertionpoint = dbl
                            If IsNumeric(TabProp(1)) Then BLprop(2) = TabProp(1)
                        ElseIf "yinsertionpoint" = TabProp(0) Then ' 3  yinsertionpoint = dbl
                            If IsNumeric(TabProp(1)) Then BLprop(3) = TabProp(1)
                        ElseIf "zinsertionpoint" = TabProp(0) Then ' 4  zinsertionpoint = dbl
                            If IsNumeric(TabProp(1)) Then BLprop(4) = TabProp(1)
                        ElseIf "rotation" = TabProp(0) Then ' 5  rotation = dbl
                            If IsNumeric(TabProp(1)) Then BLprop(5) = TabProp(1)
                        ElseIf "xscalefactor" = TabProp(0) Then ' 6  xscalefactor = dbl
                            If IsNumeric(TabProp(1)) Then BLprop(6) = TabProp(1)
                        ElseIf "yscalefactor" = TabProp(0) Then ' 7  yscalefactor = dbl
                            If IsNumeric(TabProp(1)) Then BLprop(7) = TabProp(1)
                        ElseIf "zscalefactor" = TabProp(0) Then ' 8  zscalefactor = dbl
                            If IsNumeric(TabProp(1)) Then BLprop(8) = TabProp(1)
                        ElseIf "truecolor" = TabProp(0) Then ' 9  TrueColor = dbl/ByLayer/ByBlock
                            BLprop(9) = TabProp(1)
                        ElseIf "layer" = TabProp(0) Then ' 10  Layer = ""
                            If TabProp(1) <> "" Then BLprop(10) = TabProp(1)
                        ElseIf "lineweight" = TabProp(0) Then ' 11  Lineweight = dbl
                            If IsNumeric(TabProp(1)) Then BLprop(11) = TabProp(1)
                        ElseIf "space" = TabProp(0) Then ' 12  space = ""
                            BLprop(12) = TabProp(1)
                        ElseIf "xyzinsertionpoint" = TabProp(0) Then '13  XYZInsertionPoint = dbl,dbl,dbl
                            BLprop(13) = ReplaceVar(TabProp(1))
                        ElseIf "annotative" = TabProp(0) Then '14  Annotative = 1 / 0
                            BLprop(14) = TabProp(1)
                        ElseIf "paperorientation" = TabProp(0) Then '15  PaperOrientation = 1 / 0
                            BLprop(15) = TabProp(1)
                        ElseIf "uniformly" = TabProp(0) Then '16  uniformly = 1 / 0 (1 = uniforme scale)
                            BLprop(16) = TabProp(1)
                        ElseIf "xplode" = TabProp(0) Then '17  xplode = 1 / 0 (1 = authorize decompos)
                            BLprop(17) = TabProp(1)
                        ElseIf "!" = Mid(TabProp(0), 1, 1) Then 'XX Attributs + propiété personnalisé
                            Collatt.Add(New CollBLattr(BLprop(0), BLprop(12), Mid(TabProp(0).ToUpper, 2), Replace(TabProp(1), "¦", ";")))
                        End If
                    End If
                Next



                ' ******************************************************************************************************************************


                If Mid(BLprop(0), 1, 1) = ">" And InStr(BLprop(0), "*") = 0 Then
                    BLprop(0) = Replace(BLprop(0), ">", "", 1, 1)

                    'Insertion de tout les blocks d'un dossier -----------------------
                    Dim BLfolder As String = BLprop(0)
                    Dim ListBL As New List(Of String)
                    If Directory.Exists(BLfolder) = True Then '   Listing DWG files
                        For Each fichier As String In Directory.GetFiles(BLfolder)
                            If Right(fichier.ToUpper, 4) = ".DWG" And File.Exists(fichier) = True Then
                                Try
                                    '   Insert
                                    Dim blockRefObj As AcadBlockReference
                                    blockRefObj = acDoc.ModelSpace.InsertBlock(Ptinsert, fichier, 1, 1, 1, 0)
                                    blockRefObj.Delete()
                                Catch
                                    'Write log error
                                    Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Erreur create block: " & BLprop(0))
                                End Try
                            End If
                        Next


                    Else ' Importation de la définition d'un bloc (utilisé dans WriteBlock) #BL;>BlocName;

                        Dim Fichier As String = ""
                        If File.Exists(BLfolder) Then
                            Fichier = BLfolder
                        Else
                            Fichier = Path.Combine(Ass.Library, BLfolder & ".dwg")
                            If File.Exists(Fichier) = False Then Fichier = Ass.Template 'Si pas de DWG trouvé charge le Template REVO !!!!!
                        End If

                        If File.Exists(Fichier) Then
                            Try
                                '   Insert

                                'Insertion d'un nouveau bloc
                                If BLprop(13) <> Nothing Then ' 13  xyzinsertionpoint = dbl,dbl,dbl
                                    Dim ListXYZ() As String = Split(BLprop(13), ",")
                                    If ListXYZ.Length = 3 Then
                                        Ptinsert(0) = ListXYZ(0) ' X   
                                        Ptinsert(1) = ListXYZ(1) ' Y
                                        Ptinsert(2) = ListXYZ(2) ' Z
                                        Dim blockRefObj As AcadBlockReference
                                        blockRefObj = acDoc.ModelSpace.InsertBlock(Ptinsert, Fichier, 1, 1, 1, CDbl(BLprop(5)))
                                    Else '
                                        Dim blockRefObj As AcadBlockReference
                                        blockRefObj = acDoc.ModelSpace.InsertBlock(Ptinsert, Fichier, 1, 1, 1, CDbl(BLprop(5)))
                                        'blockRefObj.Delete()
                                    End If
                                Else ' Remplace la définition et supprime le bloc
                                    Dim blockRefObj As AcadBlockReference
                                    blockRefObj = acDoc.ModelSpace.InsertBlock(Ptinsert, Fichier, 1, 1, 1, CDbl(BLprop(5)))
                                    blockRefObj.Delete()
                                End If

                            Catch
                                'Write log error
                                Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Erreur create block: " & BLprop(0))
                            End Try
                        End If
                    End If


                ElseIf Mid(BLprop(0), 1, 1) = "<" Then 'Purge du block

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                    Dim curDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                    Dim curDWG As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

                    Dim acBL As AcadBlock
                    BLprop(0) = Replace(BLprop(0), "<", "", 1, 1)
                    For Each acBL In curDwg.Blocks
                        If acBL.IsLayout = False Then
                            Try
                                If acBL.Name.ToUpper Like BLprop(0).ToUpper Then
                                    acBL.Delete()
                                End If

                            Catch ex As COMException
                                ' Un bloc n'a pu être supprimé")
                            End Try
                        End If
                    Next

                    Return Connect.DateLog & "Cmd BL" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab & acDoc.Name

                    Exit Function

                ElseIf Left(Trim(BLprop(0)), 3).ToUpper = "ID:" Then
                    'Si nom du bloc est égale à l'ID d'un bloc
                    '>>>>>> '#BL;ID:(2127672488);[[!xyz]]Val;
                    BLidSelect = True
                Else
                    BLprop(0) = Replace(BLprop(0), ">", "", 1, 1)
                End If




                'Boucle dans les calques > test lock
                Dim Calques As AcadLayers = acDoc.Layers
                Dim CollLayersLock As String = ""

                For Each LA As AcadLayer In Calques
                    If LA.Lock = True Then
                        CollLayersLock += "<" & LA.Name.ToUpper & ">"
                    End If
                Next

                'Nouvelle collection des blocks
                Dim CollBlockTmp As New Collection



                'Filtre les Spaces  -----------------------  Space = *PaperSpaceXYZ*/current/Paper/Model (défault)


                If BLprop(12) = Nothing Then BLprop(12) = "model"
                ' Dim Space As AcadBlock = Nothing
                Dim SpaceCur As AcadLayout = acDoc.ActiveLayout 'mémo du current Space
                Dim Layouts As New Collection

                For Each acLay As AcadLayout In acDoc.Layouts
                    '  If acLay.Name Like OBJprop(13) Then CurSpaceName += "<<" & acLay.Block.Name & ">>"
                    If BLprop(12).ToLower = "model" And acLay.Name = "Model" Then
                        ' Space = acDoc.ModelSpace
                        Layouts.Add(acLay)
                        Exit For
                    ElseIf BLprop(12).ToLower = "paper" And acLay.Name <> "Model" Then
                        'acDoc.ActiveLayout = acLay
                        ' Space = acDoc.PaperSpace
                        Layouts.Add(acLay)
                    ElseIf BLprop(12).ToLower = "current" And acLay.Name = acDoc.ActiveLayout.Name Then
                        Layouts.Add(acLay)
                        Exit For
                        'If acLay.Name = "Model" Then
                        ' Space = acDoc.ModelSpace
                        'Else
                        ' Space = acDoc.PaperSpace
                        'End If
                    ElseIf acLay.Name.ToLower Like BLprop(12).ToLower Then
                        Layouts.Add(acLay)
                        'acDoc.ActiveLayout = acLay
                        'If acLay.Name = "Model" Then
                        ' Space = acDoc.ModelSpace
                        'Else
                        ' Space = acDoc.PaperSpace
                        'End If
                        ' ElseIf Space Is Nothing Then
                        ' Space = acDoc.ModelSpace
                    Else
                        '  Space = Nothing 'pas de traitement
                    End If
                Next


                For Each acLay As AcadLayout In Layouts 'acDoc.Layouts

                    Try
                        If acDoc.ActiveLayout.Name <> acLay.Name Then acDoc.ActiveLayout = acLay
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                    If 1 = 1 Then ' Space IsNot Nothing Then


                        '' Edition des propriétés des block dans l'espace objet (par défault) ************************************************************


                        '' Dim acDwg As Document = Application.DocumentManager.MdiActiveDocument
                        '' Dim acCurDb As Database = acDwg.Database
                        ''Dim Values() As TypedValue = New TypedValue() {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.BlockName, BLprop(0).ToUpper)}

                        ''Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction() '' Start a transaction
                        ''    '' Request for objects to be selected in the drawing area
                        ''    Dim acSSPrompt As PromptSelectionResult = SelectAllItems(Values)
                        ''    '' If the prompt status is OK, objects were selected
                        ''    If acSSPrompt.Status = PromptStatus.OK Then
                        ''        Dim acSSet As SelectionSet = acSSPrompt.Value
                        ''        '' Step through the objects in the selection set
                        ''        For Each acSSObj As SelectedObject In acSSet
                        ''            '' Check to make sure a valid SelectedObject object was returned
                        ''            If Not IsDBNull(acSSObj) Then

                        ''                Try
                        ''                    '' Open the selected object for write
                        ''                    Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        ''                    Dim acEnt2 As Entity = acTrans.GetObject(acSSObj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                        ''                    If Not IsDBNull(acEnt) Then

                        ''                        '  If InStr(CurSpaceName.ToLower, "<<" & acEnt.BlockName.ToLower & ">>") <> 0 Then 'Contrôle si le Space est concerné

                        ''                            Dim nBL As BlockReference
                        ''                            nBL = acEnt
                        ''                            ' MsgBox(nBL.BlockName.ToString)


                        ''                            '!!! CollBlockTmp.Add(nBL) 'Ajouter à la biblithèque de bloc


                        ''                            'If nBL.BlockName.ToUpper Like BLprop(0).ToUpper Then 'Test du nom de block

                        ''                            'End If

                        ''                        'End If
                        ''                    End If
                        ''                Catch ex As Exception
                        ''                    ' Calque verrouillé : eOnLockedLayer
                        ''                    ' MsgBox(ex.Message)
                        ''                End Try
                        ''            End If
                        ''        Next
                        ''        '' Save the new object to the database
                        ''        acTrans.Commit()
                        ''    End If
                        ''    '' Dispose of the transaction
                        ''End Using


                        Try


                            'Set up the typedvalues to filter out the objects we're looking for
                            'Select all items using the filter values
                            ' Dim Values() As TypedValue = New TypedValue() {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayoutName, Space.Layout.Name), New TypedValue(CInt(DxfCode.BlockName), BLprop(0))}
                            Dim Values() As TypedValue = New TypedValue() {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayoutName, acLay.Name)} ', New TypedValue(CInt(DxfCode.BlockName), (If(BLprop(0).StartsWith("*"), "`" & BLprop(0), BLprop(0))))}
                            '  Dim Values() As TypedValue = New TypedValue() {New TypedValue(-4, "<and"), New TypedValue(0, "insert"), New TypedValue(2, BLprop(0)), New TypedValue(410, Space.Layout.Name), New TypedValue(-4, "and>")}


                            Dim pres As PromptSelectionResult = SelectAllItems(Values)
                            If pres.Status = PromptStatus.OK Then ' If we found any...
                                Dim db As Database = HostApplicationServices.WorkingDatabase
                                Using acTrans As Transaction = db.TransactionManager.StartTransaction
                                    'Get the blocktablerecord for modelspace
                                    ' Dim bt As BlockTable = DirectCast(acTrans.GetObject(db.BlockTableId, DatabaseServices.OpenMode.ForRead), BlockTable)
                                    '  Dim btr As BlockTableRecord = DirectCast(acTrans.GetObject(bt.Item(BlockTableRecord.ModelSpace), DatabaseServices.OpenMode.ForWrite), BlockTableRecord)

                                    For Each oid As ObjectId In pres.Value.GetObjectIds
                                        Dim nBL As BlockReference = acTrans.GetObject(oid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                        Dim effectiveName As String = ""
                                        If nBL.IsDynamicBlock Then
                                            Dim dbtrID As ObjectId = nBL.DynamicBlockTableRecord
                                            Dim btr As BlockTableRecord = DirectCast(acTrans.GetObject(dbtrID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockTableRecord)
                                            effectiveName = btr.Name
                                        Else
                                            effectiveName = nBL.Name
                                        End If

                                        If effectiveName.ToLower Like BLprop(0).ToLower Then ' !!!!!!!!!!!!!!!!!!!! FILTRE PAR NOM ? A CONTROLER
                                            Try
                                                CollBlockTmp.Add(nBL) 'Ajouter à la biblithèque de bloc

                                                If InStr(CollLayersLock, "<" & nBL.Layer.ToUpper & ">") = 0 Then ' ?????

                                                    'Si trouve le block > traite au Param.

                                                    ' 2  xinsertionpoint = dbl  ' 3   yinsertionpoint = dbl   ' 4   zinsertionpoint = dbl
                                                    If BLprop(2) <> Nothing And BLprop(3) <> Nothing And BLprop(4) <> Nothing Then ' Point 3D
                                                        nBL.Position = New Point3d(CDbl(BLprop(2)), CDbl(BLprop(3)), CDbl(BLprop(4)))
                                                    ElseIf BLprop(2) <> Nothing And BLprop(3) <> Nothing Then ' Point 2D
                                                        nBL.Position = New Point3d(CDbl(BLprop(2)), CDbl(BLprop(3)), nBL.Position.Z)
                                                    End If

                                                    If BLprop(5) <> Nothing Then nBL.Rotation = BLprop(5) ' 5   rotation = dbl

                                                    ' 6   xscalefactor = dbl   ' 7   yscalefactor = dbl   ' 8   zscalefactor = dbl
                                                    If BLprop(6) <> Nothing And BLprop(7) <> Nothing And BLprop(8) <> Nothing Then 'Scale 3D
                                                        nBL.ScaleFactors = New Scale3d(CDbl(BLprop(6)), CDbl(BLprop(7)), CDbl(BLprop(8)))
                                                    ElseIf BLprop(6) <> Nothing And BLprop(7) <> Nothing Then 'Scale 2D
                                                        nBL.ScaleFactors = New Scale3d(CDbl(BLprop(6)), CDbl(BLprop(7)), nBL.ScaleFactors.Z)
                                                    ElseIf BLprop(6) <> Nothing Then 'Scale 1D
                                                        nBL.ScaleFactors = New Scale3d(CDbl(BLprop(6)))
                                                    End If

                                                    If BLprop(9) <> Nothing Then 'AcadAcCmColor
                                                        BLColor = AcadTrueColor(BLprop(9))
                                                        Dim ColorBL As New AcadAcCmColor
                                                        ColorBL.ColorIndex = BLColor.ColorIndex '
                                                        nBL.ColorIndex = ColorBL.ColorIndex '2 TrueColor = dbl(1-255)/R,V,B/ByLayer/ByBlock
                                                    End If
                                                    If BLprop(10) <> Nothing Then
                                                        nBL.Layer = BLprop(10) ' 10   Layer = db
                                                    End If

                                                    If BLprop(11) <> Nothing Then nBL.LineWeight = ConvertLineWeight(BLprop(11)) ' 11   Lineweight = dbl

                                                    If BLprop(13) <> Nothing Then ' 13  xyzinsertionpoint = dbl,dbl,dbl
                                                        Dim ListXYZ() As String = Split(BLprop(13), ",")
                                                        If ListXYZ.Length = 3 Then
                                                            '    X   ' Y   ' Z
                                                            nBL.Position = New Point3d(CDbl(ListXYZ(0)), CDbl(ListXYZ(1)), CDbl(ListXYZ(2)))
                                                        Else '    X   ' Y   ' Z
                                                            nBL.Position = New Point3d(0, 0, 0)
                                                        End If
                                                    End If

                                                Else 'Error si les objets sont sur un calques  Lock
                                                    '   Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Layer Locked: " & nBL.Name & " > " & nBL.Layer)
                                                End If

                                            Catch
                                                Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Block modif: " & nBL.Name & " > " & BLprop(10))
                                            End Try
                                        End If
                                    Next

                                    'Save the changes
                                    acTrans.Commit()

                                End Using
                            End If


                        Catch 'ex As Exception
                            'MsgBox(".Message")
                        End Try


                        ''        'Boucle dans les OBJETS block : dans le Space
                        ''        For Each Objet In Space

                        ''            'Filtre les objets par Space
                        ''            ' If InStr(CurSpaceName.ToLower, "<<" & acEnt.BlockName.ToLower & ">>") <> 0 Then 'Contrôle si le Space est concerné
                        ''            '  MsgBox(acDoc.ActiveLayout.Name)
                        ''            'End If

                        ''            If TypeName(Objet) Like "IAcadBlockReference*" Then
                        ''                Dim nBL As AcadBlockReference
                        ''                nBL = Objet

                        ''                If nBL.EffectiveName.ToUpper Like BLprop(0).ToUpper Then 'Test du nom de block

                        ''                    CollBlockTmp.Add(nBL) 'Ajouter à la biblithèque de bloc !!!!!!!!!!!!!!!!!!!!! FILTRE PAR NOM ? A CONTROLER

                        ''                    Try
                        ''                        If InStr(CollLayersLock, "<" & nBL.Layer.ToUpper & ">") <> 0 Then ' ?????

                        ''                            'Si trouve le block > traite au Param.
                        ''                            If BLprop(2) <> Nothing Then nBL.InsertionPoint(0) = BLprop(2) ' 2  xinsertionpoint = dbl
                        ''                            If BLprop(3) <> Nothing Then nBL.InsertionPoint(1) = BLprop(3) ' 3   yinsertionpoint = dbl
                        ''                            If BLprop(4) <> Nothing Then nBL.InsertionPoint(2) = BLprop(4) ' 4   zinsertionpoint = dbl
                        ''                            If BLprop(5) <> Nothing Then nBL.Rotation = BLprop(5) ' 5   rotation = dbl
                        ''                            If BLprop(6) <> Nothing Then nBL.XScaleFactor = BLprop(6) ' 6   xscalefactor = dbl
                        ''                            If BLprop(7) <> Nothing Then nBL.YScaleFactor = BLprop(7) ' 7   yscalefactor = dbl
                        ''                            If BLprop(8) <> Nothing Then nBL.ZScaleFactor = BLprop(8) ' 8   zscalefactor = dbl
                        ''                            If BLprop(9) <> Nothing Then 'AcadAcCmColor
                        ''                                BLColor = AcadTrueColor(BLprop(9))
                        ''                                Dim ColorBL As New AcadAcCmColor
                        ''                                ColorBL.ColorIndex = BLColor.ColorIndex '
                        ''                                nBL.TrueColor = ColorBL '2 TrueColor = dbl(1-255)/R,V,B/ByLayer/ByBlock
                        ''                            End If
                        ''                            If BLprop(10) <> Nothing Then
                        ''                                nBL.Layer = BLprop(10) ' 10   Layer = db
                        ''                            End If

                        ''                            If BLprop(11) <> Nothing Then nBL.Lineweight = ConvertLineWeight(BLprop(11)) ' 11   Lineweight = dbl
                        ''                            If BLprop(13) <> Nothing Then ' 13  xyzinsertionpoint = dbl,dbl,dbl
                        ''                                Dim ListXYZ() As String = Split(BLprop(13), ",")
                        ''                                If ListXYZ.Length = 3 Then
                        ''                                    nBL.InsertionPoint(0) = ListXYZ(0) ' X   
                        ''                                    nBL.InsertionPoint(1) = ListXYZ(1) ' Y
                        ''                                    nBL.InsertionPoint(2) = ListXYZ(2) ' Z
                        ''                                Else
                        ''                                    nBL.InsertionPoint(0) = 0 ' X   
                        ''                                    nBL.InsertionPoint(1) = 0 ' Y
                        ''                                    nBL.InsertionPoint(2) = 0 ' Z
                        ''                                End If
                        ''                            End If

                        ''                        Else 'Error si les objets sont sur un calques  Lock
                        ''                            '   Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Layer Locked: " & nBL.Name & " > " & nBL.Layer)
                        ''                        End If

                        ''                    Catch
                        ''                        Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Block modif: " & nBL.EffectiveName & " > " & BLprop(10))
                        ''                    End Try
                        ''                Else
                        ''                    'Ne traite pas le calque
                        ''                End If
                        ''            Else
                        ''                Connect.RevoLog(TypeName(Objet))
                        ''            End If
                        ''        Next

                    Else

                    End If
                Next 'Fin boucle des Space




                If BLidSelect = False Then 'si pas d'ID cherche dans la définition des block

                    Dim ListBL As New List(Of String) 'Ajout de tout les noms des block à traiter > Annotative + PaperOrientation

                    'Boucle dans la DEFINITION du block : bibliotheque
                    For Each BL As AcadBlock In acDoc.Blocks
                        Try
                            If BL.Name.ToUpper Like BLprop(0).ToUpper = True Then 'Test du nom de block 
                                'Si trouve le block > traite au Param.
                                ListBL.Add(BL.Name)

                                If BLprop(1) <> Nothing Then ' 1  Name = ""  Nouveau nom --------------------------------
                                    Dim BLname As String = ""
                                    If InStr(BLprop(1), "*") = 0 Then
                                        BLname = BLprop(1)
                                    Else
                                        If Left(BLprop(1), 1) = "*" And Right(BLprop(1), 1) = "*" Then
                                            BLname = Replace(BL.Name, Replace(BLprop(0), "*", ""), Replace(BLprop(1), "*", ""))   ' *xyz* -> *abc*
                                        ElseIf Left(BLprop(1), 1) <> "*" And Right(BLprop(1), 1) = "*" Then
                                            BLname = Replace(BLprop(1), "*", "") & BL.Name                       ' abc***
                                        ElseIf Left(BLprop(1), 1) = "*" And Right(BLprop(1), 1) <> "*" Then
                                            BLname = BL.Name & Replace(BLprop(1), "*", "")                       ' ***abc
                                        Else
                                            BLname = Replace(BLprop(1), "*", "") 'Nouveau nom
                                        End If
                                    End If


                                    Try
                                        If IsBlockIN(BLname) = False Then
                                            BL.Name = BLname
                                        Else  'Si nouveau nom existant dans la bibliothèque remplace !
                                            Dim acCurDb As Database = Application.DocumentManager.MdiActiveDocument.Database
                                            For Each oldBLK As BlockReference In CollBlockTmp
                                                ReplaceBlock(oldBLK, BLname)
                                            Next
                                        End If

                                        Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Replace block: " & BL.Name & " > " & BLname)

                                    Catch ex1 As COMException
                                        MsgBox("Remplace Block Name : " & ex1.Message)
                                    Catch ex As Exception
                                        MsgBox("Remplace Block Name : " & ex.Message)
                                    Catch
                                        MsgBox("Remplace Block Name.")
                                    End Try
                                End If


                                If Collatt.Count > 0 Then ' XX  !att = ""  Attribut et propriété perso --------------------------------

                                    Dim ReplaceAttColl As New Collection '

                                    'Boucle sur tous les block
                                    For Each BLKnew As BlockReference In CollBlockTmp

                                        Dim objectHandle As String = BLKnew.Handle.ToString  'Mid(BLprop(0), 4, Len(BLprop(0)) - 3)
                                        Dim BLK As AcadBlockReference = acDoc.HandleToObject(objectHandle)

                                        If BLK.HasAttributes Or BLK.IsDynamicBlock Then 'Si attribut ou dynamique
                                            Dim BLAtts As Object
                                            BLAtts = BLK.GetAttributes
                                            Dim ActiveReplaceAtt As Boolean = False
                                            Dim AttrName, AttrValue As New List(Of String)

                                            For Each att As CollBLattr In Collatt 'J'AI INVERSE BOUCLE !!!!!!!!!!!!!!!!!!!  +++ rapide de traiter just le nécessaire !!!

                                                For i = 0 To UBound(BLAtts)

                                                    If BLAtts(i).TagString.ToString.ToUpper Like att.Att.ToUpper Then
                                                        Try

                                                            'Test si écriture de Valeur OU transfert de Valeur
                                                            If Left(att.Data, 3) = "((!" And Right(att.Data, 2) = "))" Then

                                                                Dim AttName As String = Mid(att.Data, 4, Len(att.Data) - 5)
                                                                Dim AttParam() As String = Split(AttName, "||")
                                                                'Formatage Spécifique

                                                                If Left(AttName.ToLower, 6) = "case||" And AttParam.Length > 1 Then
                                                                    '((!Case||Lower / Upper / Capitalize))
                                                                    If AttParam(1).ToLower = "lower" Then
                                                                        BLAtts(i).TextString = BLAtts(i).TextString.ToString.ToLower
                                                                    ElseIf AttParam(1).ToLower = "upper" Then
                                                                        BLAtts(i).TextString = BLAtts(i).TextString.ToString.ToUpper
                                                                    ElseIf AttParam(1).ToLower = "capitalize" Then
                                                                        BLAtts(i).TextString = StrConv(BLAtts(i).TextString.ToString, vbProperCase)
                                                                    End If

                                                                ElseIf Left(AttName.ToLower, 9) = "replace||" And AttParam.Length > 2 Then
                                                                    '((!Replace||"abc"||"xyz"))
                                                                    BLAtts(i).TextString = Replace(BLAtts(i).TextString, AttParam(1), AttParam(2))

                                                                ElseIf Left(AttName.ToLower, 5) = "mid||" And AttParam.Length > 2 Then
                                                                    '((!Mid||Start||Lenght)) (si Start vide supprime Lenght depuis la fin)
                                                                    If Trim(AttParam(1)) = "" Then
                                                                        BLAtts(i).TextString = Right(BLAtts(i).TextString, AttParam(2))
                                                                    Else
                                                                        BLAtts(i).TextString = Mid(BLAtts(i).TextString, AttParam(1), AttParam(2))
                                                                    End If

                                                                ElseIf Left(AttName.ToLower, 8) = "format||" And AttParam.Length > 3 Then
                                                                    '((!Format||Start||Lenght||###`##0 + d'info)) =  (si Start vide format Lenght depuis la fin)
                                                                    '((!Format|| || ||VD????00????))

                                                                    If Trim(AttParam(1)) = "" And Trim(AttParam(2)) = "" Then
                                                                        Dim ValPts As String = BLAtts(i).TextString
                                                                        Dim AttValueNew As String = ""
                                                                        Dim IncDecoupe As Double = 1
                                                                        ValPts = Replace(ValPts, " ", "0")
                                                                        For x = 1 To Len(AttParam(3))
                                                                            If Mid(AttParam(3), x, 1) = "?" Then 'ajoute le caractère ?
                                                                                If IncDecoupe <= Len(ValPts) Then
                                                                                    AttValueNew += Mid(ValPts, IncDecoupe, 1)
                                                                                    IncDecoupe += 1
                                                                                End If
                                                                            Else
                                                                                AttValueNew += Mid(AttParam(3), x, 1)
                                                                            End If
                                                                        Next
                                                                        BLAtts(i).TextString = AttValueNew
                                                                    End If

                                                                    ''* Mid = Start||Lenght
                                                                    'Dim DataVar As String = BLAtts(i).TextString
                                                                    'Dim Mid0, Mid1 As Double
                                                                    'Dim StrS As String = "" : Dim StrM As String = "" : Dim StrE As String = ""

                                                                    'If AttParam(1) = "" Then
                                                                    '    If IsNumeric(AttParam(2)) = False Then AttParam(2) = 0
                                                                    '    Mid0 = AttParam(2) : Mid1 = Len(VarData(DataList(i))) - AttParam(2)
                                                                    '    StrS = Mid(VarData(DataList(i)), 1, Mid1)
                                                                    '    StrM = Mid(VarData(DataList(i)), Mid1 + 1, Len(VarData(DataList(i))) - Mid0 + Mid1 + 1)
                                                                    '    StrE = "" 'Mid(VarData(DataList(i)), Mid1 + Mid0 + 1, Len(VarData(DataList(i))))
                                                                    'Else
                                                                    '    Mid0 = CDbl(AttParam(1))
                                                                    '    If Mid0 < 1 Then Mid0 = 1
                                                                    '    Mid1 = CDbl(AttParam(2))
                                                                    '    StrS = Mid(VarData(DataList(i)), 1, Mid0 - 1)
                                                                    '    StrM = Mid(VarData(DataList(i)), Mid0, Mid1)
                                                                    '    StrE = Mid(VarData(DataList(i)), Mid1 + 1, Len(VarData(DataList(i))) - Mid0 + Mid1 + 1)
                                                                    'End If

                                                                    'If InStr(MidVar(2), "#") <> 0 And IsNumeric(Replace(StrM, "'", "")) Then
                                                                    '    MidVar(2) = Replace(MidVar(2), "'", "`")
                                                                    '    VarData(DataList(i)) = StrS & Replace(Format(CDbl(Replace(StrM, "'", "")), Replace(MidVar(2), "'", "`")), "`", "'") & StrE
                                                                    'Else
                                                                    '    VarData(DataList(i)) = StrS & Format((StrM), MidVar(2)) & StrE
                                                                    'End If

                                                                Else 'Sans code spécifique
                                                                    ActiveReplaceAtt = True
                                                                    AttrName.Add(AttName)
                                                                    AttrValue.Add(BLAtts(i).TextString)
                                                                End If

                                                            Else 'Valeur d'attribut normal
                                                                Dim Attdata As String = att.Data
                                                                'Boucle dans les variables existantes pour le remplacement
                                                                For Each coll As ScriptVar In CollVar
                                                                    Dim DataFormatLigne As String = ""
                                                                    For Each DataX In coll.Data
                                                                        If DataFormatLigne <> "" Then DataFormatLigne += vbCrLf
                                                                        DataFormatLigne += CStr(DataX)
                                                                    Next
                                                                    Attdata = Replace(Attdata, coll.Name, DataFormatLigne)
                                                                Next
                                                                BLAtts(i).TextString = Attdata  '(att.Name & att.Space & att.Att & att.Data)
                                                                AttrName.Add(att.Att)
                                                                AttrValue.Add(Attdata)
                                                            End If
                                                        Catch ex As COMException
                                                            'Erreur calques verrouillés
                                                            Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Err attribut, layer locked: " & BLK.EffectiveName & " > " & BLAtts(i).TagString.ToString)
                                                        End Try

                                                    End If
                                                Next
                                            Next

                                            'Boucle pour les propriétés dynamique
                                            '   a faire
                                            ' ----------------------------------


                                            'Stockage provisoire des données des Attribut --------'Création de l'objet PTS
                                            If ActiveReplaceAtt Then
                                                'Ajout des attributs si inexistant -------------------------
                                                For Each att As CollBLattr In Collatt
                                                    Dim Attinexistant As Boolean = True
                                                    For idattr = 0 To AttrName.Count - 1
                                                        If att.Att.ToUpper = AttrName(idattr).ToUpper Then Attinexistant = False
                                                    Next
                                                    If Attinexistant Then 'Pas encore ajoutée 1x
                                                        AttrName.Add(att.Att)
                                                        AttrValue.Add(att.Data)
                                                    End If
                                                Next

                                                ReplaceAttColl.Add(New RevoPTS(BLK.EffectiveName, "", BLK.Handle, "", "", "", BLK.InsertionPoint(0), BLK.InsertionPoint(1), BLK.InsertionPoint(2), 1, 0, AttrName, AttrValue))
                                            End If


                                        End If
                                    Next


                                    'Transfert vers les nouveaux nom d'attribut ***************************
                                    If ReplaceAttColl.Count > 0 Then
                                        'Syncronisation des attributs du bloc
                                        cmdCmd("#cmd;_ATTSYNC|N|" & BLprop(0))

                                        'Mise à jour des valeurs > Boucle sur tous les points
                                        For Each Pt As RevoPTS In ReplaceAttColl

                                            Dim UpBlk As AcadBlockReference
                                            Dim objectHandle As String = Pt.ID  'Mid(BLprop(0), 4, Len(BLprop(0)) - 3)
                                            UpBlk = acDoc.HandleToObject(objectHandle)

                                            If UpBlk.HasAttributes Or UpBlk.IsDynamicBlock Then 'Si attribut ou dynamique
                                                Dim BLAtts As Object
                                                BLAtts = UpBlk.GetAttributes
                                                For i = 0 To UBound(BLAtts)
                                                    For u = 0 To Pt.AttrName.Count - 1
                                                        If BLAtts(i).TagString.ToString.ToUpper Like Pt.AttrName(u).ToUpper Then
                                                            Try
                                                                BLAtts(i).TextString = Pt.AttrValue(u)  '(att.Name & att.Space & att.Att & att.Data)
                                                            Catch ex As COMException
                                                                'Erreur calques verrouillés
                                                                Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Err attribut, layer locked: " & UpBlk.EffectiveName & " > " & BLAtts(i).TagString.ToString)
                                                            End Try
                                                        End If
                                                    Next
                                                Next
                                            End If
                                        Next

                                    End If
                                End If

                            Else
                                'Ne traite pas le block
                            End If

                        Catch ex1 As Exception
                            MsgBox(ex1.Message)
                        Catch ex2 As COMException
                            MsgBox(ex2.Message)
                        Catch
                            MsgBox("Erreur avec la commande : " & Cmd)
                        End Try
                    Next


                    'Traitement séparée de l'Annotative et PaperOrientation en fonction de la présention

                    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                    Dim db As Database = doc.Database
                    Dim ed As Editor = doc.Editor

                    Using trx As Transaction = db.TransactionManager.StartTransaction()
                        Dim bt As BlockTable = trx.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)

                        For Each BLname As String In ListBL
                            If bt.Has(BLname) Then
                                Dim btr As BlockTableRecord = trx.GetObject(bt(BLname), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                btr.UpgradeOpen()

                                '14  Annotative = 1 / 0
                                If BLprop(14) IsNot Nothing Then
                                    If BLprop(14).Trim = "0" Then btr.Annotative = AnnotativeStates.False
                                    If BLprop(14).Trim = "1" Then btr.Annotative = AnnotativeStates.True
                                End If

                                '15  PaperOrientation = 1 / 0
                                If BLprop(15) IsNot Nothing Then
                                    If BLprop(15).Trim = "0" Then btr.SetPaperOrientation(False) 'PaperOrientationStates.False
                                    If BLprop(15).Trim = "1" Then btr.SetPaperOrientation(True)
                                End If

                                '16  Uniformly = 1 / 0 (1 = uniforme scale)
                                If BLprop(16) IsNot Nothing Then
                                    If BLprop(16).Trim = "0" Then btr.BlockScaling = BlockScaling.Any
                                    If BLprop(16).Trim = "1" Then btr.BlockScaling = BlockScaling.Uniform
                                End If

                                '17  Xplode = 1 / 0 (1 = authorize decompos)
                                If BLprop(17) IsNot Nothing Then
                                    If BLprop(17).Trim = "0" Then btr.Explodable = False
                                    If BLprop(17).Trim = "1" Then btr.Explodable = True
                                End If


                                'Try
                                '    btr.Name = BLname & "_test"
                                'Catch ex As System.Exception
                                '    If ex.Message = "eDuplicateRecordName" Then
                                '        MsgBox("Bloc avec le même nom déjà existant")
                                '    End If
                                'End Try
                            End If
                        Next

                        trx.Commit()
                    End Using


                Else 'si ID cherche le block

                    Try


                        If Collatt.Count > 0 Then ' XX  !att = ""  Attribut et propriété perso --------------------------------

                            'Boucle sur tous les block
                            ' For Each BLK As AcadBlockReference In CollBlockTmp

                            'If "BLK.ObjectID.ToString" = Trim(Replace(Replace(BLprop(0), "ID(", ""), ")", "")) Then


                            Dim BLK As AcadBlockReference
                            'MsgBox(Mid(BLprop(0), 4, Len(BLprop(0)) - 3))
                            Dim objectHandle As String = Mid(BLprop(0), 4, Len(BLprop(0)) - 3)
                            BLK = acDoc.HandleToObject(objectHandle)

                            If BLK.HasAttributes Or BLK.IsDynamicBlock Then 'Si attribut ou dynamique
                                Dim BLAtts As Object
                                BLAtts = BLK.GetAttributes

                                For i = 0 To UBound(BLAtts)
                                    For Each att As CollBLattr In Collatt
                                        If BLAtts(i).TagString.ToString.ToUpper Like att.Att.ToUpper Then
                                            Try
                                                Dim Attdata As String = att.Data

                                                'Boucle dans les variables existantes pour le remplacement
                                                For Each coll As ScriptVar In CollVar
                                                    Dim DataFormatLigne As String = ""
                                                    For Each DataX In coll.Data
                                                        If DataFormatLigne <> "" Then DataFormatLigne += vbCrLf
                                                        DataFormatLigne += CStr(DataX)
                                                    Next
                                                    Attdata = Replace(Attdata, coll.Name, DataFormatLigne)
                                                Next

                                                BLAtts(i).TextString = Attdata  '(att.Name & att.Space & att.Att & att.Data)
                                            Catch ex As COMException
                                                'Erreur calques verrouillés
                                                Connect.RevoLog(Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Err attribut, layer locked: " & BLK.EffectiveName & " > " & BLAtts(i).TagString.ToString)
                                            End Try
                                        End If
                                    Next
                                Next

                                'Boucle pour les propriétés dynamique
                                '   a faire
                                ' ----------------------------------
                            End If

                            'End If
                            ' Next
                        End If

                    Catch ex As Exception
                        MsgBox(ex.Message)

                    End Try
                End If


                acDoc.ActiveLayout = SpaceCur 'restor du current Space


                Collatt.Clear() 'Vide la table des attribus
                CollBlockTmp.Clear() 'vide la table des blocks

            Else 'ignore the command
                Return Connect.DateLog & "Cmd BL" & vbTab & False & vbTab & "Pas de paramètre" & vbTab
            End If

            Return Connect.DateLog & "Cmd BL" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab & acDoc.Name

        End Function

        Public Function ReplaceVar(ByVal Var As String) 'Remplacement d'une variable (simple ligne)
            Dim NewVar As String = Var

            'Boucle dans les variables existantes pour le remplacement
            For Each coll As ScriptVar In CollVar
                If coll.Data.Count > 0 Then
                    If InStr(NewVar, coll.Name) <> 0 And coll.Data(0) IsNot Nothing Then
                        NewVar = Replace(NewVar, coll.Name, coll.Data(0)) 'Remplace Var (1er ligne)
                    End If
                End If
            Next

            Return NewVar
        End Function

        Public Function SelectSpace(ByVal acDoc As AcadDocument, ByVal SpaceName As String) As AcadBlock

            'Choix du Space
            Dim Space As AcadBlock = acDoc.ModelSpace
            If SpaceName <> Nothing Then
                If SpaceName.ToLower = "model" Then ' Model ou NomPresentation
                    Space = acDoc.ModelSpace
                ElseIf SpaceName.ToLower = "current" Then
                    If acDoc.ActiveLayout.Name = "Model" Then
                        Space = acDoc.ModelSpace
                    Else
                        Space = acDoc.PaperSpace
                    End If
                ElseIf SpaceName.ToLower <> "" Then
                    For Each lay As AcadLayout In acDoc.Layouts
                        If lay.Name.ToLower = SpaceName.ToLower Then
                            acDoc.ActiveLayout = lay
                            If acDoc.ActiveLayout.Name = "Model" Then
                                Space = acDoc.ModelSpace
                            Else
                                Space = acDoc.PaperSpace
                            End If
                        End If
                    Next
                End If
            End If


            Return Space
        End Function


        Public Shared Function GetLayoutName(ByVal entity As Entity) As String
            Dim blockId As ObjectId = entity.BlockId
            Dim btr As BlockTableRecord = TryCast(entity.BlockId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockTableRecord)
            If Not btr.LayoutId.IsNull Then
                Dim layout As Layout = TryCast(btr.LayoutId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Layout)
                Return layout.LayoutName
            End If
            Return String.Empty
        End Function

        ''' <summary>
        ''' Revo Command: New Hatch
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks></remarks>
        Public Function cmdHA(ByVal Cmd As String)  ' Remarks
            '#HA;filtre Calque;[[PropA]]Val;[[PropB]]Val;
            'Création des hachures dans des surfaces
            '0  Filtre Calque = "*" (recherche des surfaces)

            '[[Paramètre]]
            '1  Space = *PaperSpaceXYZ*/Current/Model/paper/ * (Model: défault)
            '2  TrueColor = dbl(1-255)/R,V,B/ByLayer/ByBlock
            '3  Layer = "" (calque de destination)
            '4  Filtre Objet = * /Polyline3d/Polyline2d/Polyline/Line
            '5  Rotation = dbl
            '6  StyleName = ""
            '7  ScaleFactor = dbl
            '8  LayerIsland = "" (calque pour la détection d'objets)
            '9  Annotative = 1 / 0

            Dim HAprop(0 To 9) As String
            Dim Cmdinfo() As String
            Dim OBJColor As Color
            OBJColor = Color.FromColorIndex(ColorMethod.ByAci, 7)

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim acDwg As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If


            Dim Ptinsert(2) As Double
            Ptinsert(0) = 0 : Ptinsert(1) = 0 : Ptinsert(2) = 0
            Cmdinfo = SplitCmd(Cmd, 1)

            Dim CollPoly As New Collection
            ' Dim CollPolyOut As ObjectIdCollection = New ObjectIdCollection()


            HAprop(0) = Cmdinfo(1)  '0  Filtre Calque = "*" (recherche des surfaces)

            For i = 1 To Cmdinfo.Count - 1 'Boucle prop
                If InStr(Cmdinfo(i), "[[") <> 0 And InStr(Cmdinfo(i), "]]") <> 0 Then
                    Cmdinfo(i) = Replace(Replace(Cmdinfo(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                    Dim TabProp() As String = SplitCmd(Cmdinfo(i), 1)
                    TabProp(0) = TabProp(0).ToLower

                    If "space" = TabProp(0) Then ' '1  Space = *PaperSpaceXYZ*/Current/Model/paper/ * (Model: défault)
                        ' HAprop(1) = TabProp(1)
                    ElseIf "trueColor" = TabProp(0) Then '2  TrueColor = dbl(1-255)/R,V,B/ByLayer/ByBlock
                        HAprop(2) = TabProp(1)
                    ElseIf "layer" = TabProp(0) Then   '3  Layer = "" (calque de déstination)
                        HAprop(3) = TabProp(1)
                    ElseIf "objfilter" = TabProp(0) Then '4 ObjFilter Filtre Objet = * /Polyline3d/Polyline2d/Polyline/Line
                        HAprop(4) = TabProp(1)
                    ElseIf "rotation" = TabProp(0) Then   '5  Rotation = dbl
                        If IsNumeric(TabProp(1)) Then HAprop(5) = TabProp(1)
                    ElseIf "stylename" = TabProp(0) Then       '6  StyleName = ""
                        HAprop(6) = TabProp(1)
                    ElseIf "scalefactor" = TabProp(0) Then '7  ScaleFactor = dbl
                        If IsNumeric(TabProp(1)) Then HAprop(7) = TabProp(1)
                    ElseIf "layerisland" = TabProp(0) Then '8  LayerIsland = "" (calque pour la détection d'objets)
                        HAprop(8) = TabProp(1)
                    ElseIf "annotative" = TabProp(0) Then '9  Annotative = 1 / 0
                        HAprop(9) = TabProp(1)
                    End If
                End If
            Next



            'Default Value
            If HAprop(6) = Nothing Then HAprop(6) = "SOLID"
            If HAprop(7) = Nothing Then HAprop(7) = 0
            If HAprop(9) = Nothing Then HAprop(9) = 0

            ' Start a transaction
            Dim LA As AcadLayer
            For Each LA In acDwg.Layers
                LA.Lock = False
            Next

            'Active/désactive Annotation
            Dim AcAnnotative As Boolean
            If HAprop(9) = 1 Then
                AcAnnotative = True
            Else
                AcAnnotative = False
            End If

            'Activer la suppression des hachures existantes
            If Mid(HAprop(3), 1, 1) = ">" Then
                HAprop(3) = Right(HAprop(3), HAprop(3).Length - 1)
                ' Suppression des hachures existantes
                DeleteHatchInLayer(HAprop(3))
            End If


            If HAprop(0) <> "" Then

                Zooming.ZoomExtents()
                Dim Values() As TypedValue = New TypedValue() {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, HAprop(0))}

                'Select all items using the filter values
                Dim pres As PromptSelectionResult = SelectAllItems(Values)
                If pres.Status = PromptStatus.OK Then ' If we found any...
                    Dim db As Database = HostApplicationServices.WorkingDatabase
                    'Declare a objectidcollection (required for the MoveToTop method of the draw order table)
                    Dim entCol As New ObjectIdCollection
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        'Get the blocktablerecord for modelspace
                        Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, DatabaseServices.OpenMode.ForRead), BlockTable)
                        Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), DatabaseServices.OpenMode.ForRead), BlockTableRecord)
                        'Loop through the found objects adding them to the objectidcollection
                        For Each oid As ObjectId In pres.Value.GetObjectIds
                            Dim Poly As Polyline = DirectCast(trans.GetObject(oid, DatabaseServices.OpenMode.ForRead), Polyline)
                            If Poly.Closed = True Then CollPoly.Add(New AcObjID(oid, Poly.Area))
                        Next
                        'trans.Commit() 'Save the changes
                    End Using
                End If

                'Traitement des polylignes
                If CollPoly.Count > 0 And HAprop(0) <> "" Then
                    Try
                        AddHatch(CollPoly, HAprop(8), HAprop(6), CDbl(HAprop(7)), HAprop(0), HAprop(3), AcAnnotative)

                    Catch ex As Exception
                        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                        Try
                            AddHatch(CollPoly, HAprop(8), "SOLID", CDbl(HAprop(7)), HAprop(0), HAprop(3), AcAnnotative)
                            ed.WriteMessage(vbLf & "Problème avec les hachures : " & HAprop(3) & " (" & HAprop(6) & ")")
                        Catch
                            ed.WriteMessage(vbLf & "Problème avec les hachures : " & HAprop(3) & " (" & HAprop(6) & ")")
                        End Try
                    End Try

                    'Mise en arrière des hachures
                    If HAprop(3) IsNot Nothing Then PlaceObjectsOnBottom("HATCH", HAprop(3))
                End If

                Zooming.ZoomPrevious()

            End If


            Return Connect.DateLog & "Cmd HA" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & "(nbre:" & CollPoly.Count & ")" & vbTab & acDwg.Name

        End Function

        ''' <summary>
        ''' Revo Command: Objet modifier
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks></remarks>
        Public Function cmdOBJ(ByVal Cmd As String)  ' Remarks
            '#OBJ;filtre. Calque;filtre. obj;[[PropA]]Val;[[PropB]]Val;
            ' Modifie la propriété d'objet d'un calque
            ' 0  Filtre Calque
            ' 1  Filtre type obj
            'Général
            ' 2  TrueColor = dbl(1-255)/R,V,B/ByLayer/ByBlock
            ' 3  Layer = ""
            ' 4  LineType = ""
            ' 5  LinetypeScale = dbl
            ' 6  Lineweight = dbl

            'Block + AcadText + AcadMText + AcadAttribute
            ' 7  Rotation = dbl

            'AcadText + AcadMText + AcadAttribute  
            ' 8  StyleName = ""
            ' 9  Alignment = TopCenter
            '10  Height = dbl

            'Block
            '11  ScaleFactor = dbl
            '12  Text = "*"||"abc" (inteli, * = tout)
            '13  Space = *PaperSpaceXYZ*/current/Model/paper/ * (all: défault)
            '14  Attribute = "*" (filtre)
            '15  Annotative = 1 / 0

            '1X  PropPerso ???


            Dim OBJprop(0 To 16) As String
            Dim Cmdinfo() As String
            Dim TabProp() As String
            Dim OBJColor As Color
            OBJColor = Color.FromColorIndex(ColorMethod.ByAci, 7)

            Dim Ptinsert(2) As Double
            Ptinsert(0) = 0 : Ptinsert(1) = 0 : Ptinsert(2) = 0
            Cmdinfo = SplitCmd(Cmd, 2)

            If Cmdinfo(0) <> "" Then 'si le nom du calque n'est pas vide

                OBJprop(0) = Cmdinfo(1)
                OBJprop(1) = Cmdinfo(2)

                For i = 1 To Cmdinfo.Count - 1 'Boucle prop
                    If InStr(Cmdinfo(i), "[[") <> 0 And InStr(Cmdinfo(i), "]]") <> 0 Then
                        Cmdinfo(i) = Replace(Replace(Cmdinfo(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                        TabProp = SplitCmd(Cmdinfo(i), 1)
                        TabProp(0) = TabProp(0).ToLower

                        If "truecolor" = TabProp(0) Then ' 2  TrueColor = dbl(1-255)/R,V,B/ByLayer/ByBlock
                            OBJprop(2) = TabProp(1)
                        ElseIf "layer" = TabProp(0) Then ' 3  Layer = ""
                            OBJprop(3) = TabProp(1)
                        ElseIf "linetype" = TabProp(0) Then ' 4  LineType = ""
                            OBJprop(4) = TabProp(1)
                        ElseIf "linetypescale" = TabProp(0) Then  ' 5  LinetypeScale = dbl
                            If IsNumeric(TabProp(1)) Then OBJprop(5) = TabProp(1)
                        ElseIf "lineweight" = TabProp(0) Then ' 6  Lineweight = dbl
                            If IsNumeric(TabProp(1)) Then OBJprop(6) = TabProp(1)
                        ElseIf "rotation" = TabProp(0) Then ' 7  Rotation = dbl
                            If IsNumeric(TabProp(1)) Then OBJprop(7) = TabProp(1)
                        ElseIf "stylename" = TabProp(0) Then ' 8  StyleName = ""
                            OBJprop(8) = TabProp(1)
                        ElseIf "alignment" = TabProp(0) Then ' 9  Alignment = TopCenter
                            OBJprop(9) = TabProp(1)
                        ElseIf "height" = TabProp(0) Then '10  Height = dbl
                            If IsNumeric(TabProp(1)) Then OBJprop(10) = TabProp(1)
                        ElseIf "scalefactor" = TabProp(0) Then '11  ScaleFactor = dbl
                            If IsNumeric(TabProp(1)) Then OBJprop(11) = TabProp(1)
                        ElseIf "text" = TabProp(0) Then '12  Text = "*"||"abc"
                            OBJprop(12) = TabProp(1)
                        ElseIf "space" = TabProp(0) Then ' 13  space = ""
                            OBJprop(13) = TabProp(1)
                        ElseIf "attribute" = TabProp(0) Then ' 14  space = ""
                            OBJprop(14) = TabProp(1)
                        ElseIf "annotative" = TabProp(0) Then ' 15 Annotative = 1 / 0
                            OBJprop(15) = TabProp(1)

                            '15  PropPerso ???

                        End If
                    End If
                Next

                ' Start a transaction
                Dim LA As AcadLayer
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                Dim acDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                Dim acDwg As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
                For Each LA In acDwg.Layers
                    LA.Lock = False
                Next

                'Filtre les Spaces  -----------------------
                Dim CurSpaceName As String = ""
                If OBJprop(13) = Nothing Then OBJprop(13) = "*" 'Par défault : tout = *

                'Choix du Space
                If OBJprop(13).ToLower = "current" Then
                    Dim CurSpace As AcadBlock = SelectSpace(acDwg, "current")
                    CurSpaceName = "<<" & CurSpace.Name & ">>"
                ElseIf OBJprop(13).ToLower = "model" Then
                    CurSpaceName = "<<*Model_Space>>"
                ElseIf OBJprop(13).ToLower = "paper" Then
                    For Each acLay As AcadLayout In acDwg.Layouts
                        If acLay.Name <> "Model" Then CurSpaceName += "<<" & acLay.Block.Name & ">>"
                    Next
                Else
                    For Each acLay As AcadLayout In acDwg.Layouts
                        If acLay.Name Like OBJprop(13) Then CurSpaceName += "<<" & acLay.Block.Name & ">>"
                    Next
                End If

                'Remplacement des variables
                'Boucle dans les variables existantes pour le remplacement
                Dim ParamVar() As String = Nothing
                Try
                    ParamVar = Split(OBJprop(12), "||")
                    For Each coll As ScriptVar In CollVar
                        ParamVar(1) = Replace(ParamVar(1), "¦", ";") 'remplacement des ;
                        If InStr(ParamVar(1), coll.Name) <> 0 Then 'Recherche s'il faut remplacer une var
                            Dim DataMultiline As String = ""
                            For Each DataX As String In coll.Data ' Transforme la list en une variable multiline
                                If DataMultiline <> "" Then DataMultiline += vbCrLf
                                DataMultiline += DataX
                            Next
                            ParamVar(1) = Replace(ParamVar(1), coll.Name, DataMultiline) ' remplacement des /var/ dans toute les lignes
                        End If
                    Next
                Catch
                End Try

                If OBJprop(12) <> Nothing And ParamVar.Count = 2 Then OBJprop(12) = ParamVar(0) & "||" & ParamVar(1)

                'Select dans les OBJETS de tout le dessins (filtre en fonction de Space OBJprop(13))
                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                Dim acCurDb As Database = acDoc.Database

                If CollLoad = False Then
                    CollLoad = True
                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction() '' Start a transaction


                        '' Request for objects to be selected in the drawing area
                        Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.SelectAll()
                        '' If the prompt status is OK, objects were selected
                        If acSSPrompt.Status = PromptStatus.OK Then
                            Dim acSSet As SelectionSet = acSSPrompt.Value
                            '' Step through the objects in the selection set
                            For Each acSSObj As SelectedObject In acSSet
                                '' Check to make sure a valid SelectedObject object was returned
                                If Not IsDBNull(acSSObj) Then

                                    Try
                                        '' Open the selected object for write
                                        Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                        Dim acEnt2 As Entity = acTrans.GetObject(acSSObj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                        If Not IsDBNull(acEnt) Then

                                            If InStr(CurSpaceName.ToLower, "<<" & acEnt.BlockName.ToLower & ">>") <> 0 Then 'Contrôle si le Space est concerné

                                                If TypeName(acEnt) = "Polyline" Then
                                                    CollPolyline.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Line" Or TypeName(acEnt) = "IAcadLine2" Then
                                                    CollLine.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "BlockReference" Then
                                                    CollBlockReference.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "DBText" Or TypeName(acEnt) = "IAcadText2" Then
                                                    CollDBText.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "MText" Or TypeName(acEnt) = "IAcadMText2" Then
                                                    CollMText.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Polyline2d" Or TypeName(acEnt) = "IAcadLWPolyline2" Then
                                                    CollPolyline2d.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Polyline3d" Then
                                                    CollPolyline3d.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "DBPoint" Then
                                                    CollDBPoint.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Hatch" Then
                                                    CollHatch.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Arc" Then
                                                    CollArc.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Circle" Or TypeName(acEnt) = "IAcadCircle2" Then
                                                    CollCircle.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "AlignedDimension" Then
                                                    CollAlignedDimension.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "RotatedDimension" Then
                                                    CollRadialDimension.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "RadialDimension" Then
                                                    CollRadialDimension.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "LineAngularDimension2" Then
                                                    CollLineAngularDimension2.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "ArcDimension" Then
                                                    CollArcDimension.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "DiametricDimension" Then
                                                    CollDiametricDimension.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Spline" Then
                                                    CollSpline.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Xline" Then
                                                    CollXline.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Ray" Then
                                                    CollRay.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Helix" Then
                                                    CollHelix.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Ellipse" Then
                                                    CollEllipse.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "AttributeDefinition" Or TypeName(acEnt) = "IAcadAttribute2" Or TypeName(acEnt) = "IAcadAttribute2" Then
                                                    CollAttribute.Add(acEnt)
                                                ElseIf TypeName(acEnt) = "Viewport" Or TypeName(acEnt) = "IAcadPViewport2" Then
                                                    'ne faire rien
                                                Else
                                                    CollOthers.Add(acEnt)
                                                    Connect.RevoLog(TypeName(acEnt))
                                                End If
                                            End If
                                        End If
                                    Catch ex As Exception
                                        ' Calque verrouillé : eOnLockedLayer
                                        ' MsgBox(ex.Message)
                                    End Try
                                End If
                            Next
                            '' Save the new object to the database
                            acTrans.Commit()
                        End If
                        '' Dispose of the transaction
                    End Using
                End If

                'Traitement des propriétés des objets

                If "DBTEXT" Like OBJprop(1).ToUpper Then ' TEXTE
                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction() '' Start a transaction
                        For Each ObjDBText As DBText In CollDBText 'Traitement des objets: DBText
                            If ObjDBText.Layer.ToUpper Like OBJprop(0).ToUpper = True Then 'Filtre calque

                                'Problème:  eNotFromThisDocument
                                Dim acEnt As Entity = acTrans.GetObject(ObjDBText.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                If Not IsDBNull(acEnt) Then
                                    If OBJprop(2) <> Nothing Then
                                        OBJColor = AcadTrueColor(OBJprop(2))
                                        ObjDBText.ColorIndex = OBJColor.ColorIndex '2 TrueColor = dbl(1-255)/R,V,B/ByLayer/ByBlock
                                    End If

                                    Try
                                        If OBJprop(3) <> Nothing Then ObjDBText.Layer = OBJprop(3) ' 3  Layer = ""
                                        If OBJprop(4) <> Nothing Then ObjDBText.Linetype = OBJprop(4) ' 4  LineType = ""
                                        If OBJprop(5) <> Nothing Then ObjDBText.LinetypeScale = OBJprop(5) ' 5  LinetypeScale = dbl
                                        If OBJprop(6) <> Nothing Then ObjDBText.LineWeight = ConvertLineWeight(OBJprop(6)) ' 6  Lineweight = dbl
                                        If OBJprop(7) <> Nothing Then ObjDBText.Rotation = OBJprop(7) ' 7  Rotation = dbl
                                        If OBJprop(8) <> Nothing Then ' 8  StyleName = ""
                                            Dim textStyleTable As TextStyleTable = TryCast(acTrans.GetObject(acCurDb.TextStyleTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), TextStyleTable)
                                            Dim textStyleId As ObjectId = ObjectId.Null
                                            If textStyleTable.Has(OBJprop(8)) Then
                                                textStyleId = textStyleTable(OBJprop(8))
                                                ObjDBText.TextStyleId = textStyleId ' acCurDb.Textstyle 'OBJprop(8)
                                            End If
                                        End If
                                        If OBJprop(15) <> "" Then
                                            If OBJprop(15) = 1 Then ObjDBText.Annotative = AnnotativeStates.True
                                            If OBJprop(15) = 0 Then ObjDBText.Annotative = AnnotativeStates.False
                                        End If


                                    Catch
                                End Try

                                    If OBJprop(9) <> Nothing Then ' 9  Alignment = TopCenter
                                        Dim InsertPt As Point3d
                                        If ObjDBText.AlignmentPoint = Nothing Then
                                            InsertPt = ObjDBText.Position
                                        Else
                                            InsertPt = ObjDBText.AlignmentPoint
                                        End If

                                        If OBJprop(9).ToUpper = "TOPCENTER" Then
                                            ObjDBText.HorizontalMode = 1 : ObjDBText.VerticalMode = 3
                                        ElseIf OBJprop(9).ToUpper = "TOPLEFT" Then
                                            ObjDBText.HorizontalMode = 0 : ObjDBText.VerticalMode = 3
                                        ElseIf OBJprop(9).ToUpper = "TOPRIGHT" Then
                                            ObjDBText.HorizontalMode = 2 : ObjDBText.VerticalMode = 3
                                        ElseIf OBJprop(9).ToUpper = "BOTTOMCENTER" Then
                                            ObjDBText.HorizontalMode = 1 : ObjDBText.VerticalMode = 1
                                        ElseIf OBJprop(9).ToUpper = "BOTTOMLEFT" Then
                                            ObjDBText.HorizontalMode = 0 : ObjDBText.VerticalMode = 1
                                        ElseIf OBJprop(9).ToUpper = "BOTTOMRIGHT" Then
                                            ObjDBText.HorizontalMode = 2 : ObjDBText.VerticalMode = 1
                                        ElseIf OBJprop(9).ToUpper = "MIDDLECENTER" Then
                                            ObjDBText.HorizontalMode = 1 : ObjDBText.VerticalMode = 2
                                        ElseIf OBJprop(9).ToUpper = "MIDDLELEFT" Then
                                            ObjDBText.HorizontalMode = 0 : ObjDBText.VerticalMode = 2
                                        ElseIf OBJprop(9).ToUpper = "MIDDLERIGHT" Then
                                            ObjDBText.HorizontalMode = 2 : ObjDBText.VerticalMode = 2
                                        End If
                                        ObjDBText.AlignmentPoint = InsertPt
                                    End If
                                    If OBJprop(10) <> Nothing Then ObjDBText.Height = OBJprop(10) '10  Height = dbl
                                End If
                            End If
                        Next
                        '' Save the new object to the database
                        acTrans.Commit()
                        '' Dispose of the transaction
                    End Using
                End If

                Dim testpr
                testpr = PropOBJ("MText", CollMText, acCurDb, OBJprop)
                testpr = PropOBJ("BlockReference", CollBlockReference, acCurDb, OBJprop)
                testpr = PropOBJ("AlignedDimension", CollAlignedDimension, acCurDb, OBJprop)
                testpr = PropOBJ("Arc", CollArc, acCurDb, OBJprop)
                testpr = PropOBJ("ArcDimension", CollArcDimension, acCurDb, OBJprop)
                testpr = PropOBJ("BlockReference", CollBlockReference, acCurDb, OBJprop)
                testpr = PropOBJ("Circle", CollCircle, acCurDb, OBJprop)
                testpr = PropOBJ("DBPoint", CollDBPoint, acCurDb, OBJprop)
                testpr = PropOBJ("DBText", CollDBText, acCurDb, OBJprop)
                testpr = PropOBJ("DiametricDimension", CollDiametricDimension, acCurDb, OBJprop)
                testpr = PropOBJ("Ellipse", CollEllipse, acCurDb, OBJprop)
                testpr = PropOBJ("Hatch", CollHatch, acCurDb, OBJprop)
                testpr = PropOBJ("Helix", CollHelix, acCurDb, OBJprop)
                testpr = PropOBJ("LineAngularDimension2", CollLineAngularDimension2, acCurDb, OBJprop)
                testpr = PropOBJ("Line", CollLine, acCurDb, OBJprop)
                testpr = PropOBJ("MText", CollMText, acCurDb, OBJprop)
                testpr = PropOBJ("Polyline3d", CollPolyline3d, acCurDb, OBJprop)
                testpr = PropOBJ("Polyline2d", CollPolyline2d, acCurDb, OBJprop)
                testpr = PropOBJ("Polyline", CollPolyline, acCurDb, OBJprop)
                testpr = PropOBJ("RadialDimension", CollRadialDimension, acCurDb, OBJprop)
                testpr = PropOBJ("Ray", CollRay, acCurDb, OBJprop)
                testpr = PropOBJ("RotatedDimension", CollRotatedDimension, acCurDb, OBJprop)
                testpr = PropOBJ("Spline", CollSpline, acCurDb, OBJprop)
                testpr = PropOBJ("Xline", CollXline, acCurDb, OBJprop)
                testpr = PropOBJ("Attribute", CollAttribute, acCurDb, OBJprop)

                testpr = PropOBJ("OTHERS", CollOthers, acCurDb, OBJprop)

            Else 'ignore the command
                Return Connect.DateLog & "Cmd OBJ" & vbTab & False & vbTab & "Pas de paramètre" & vbTab
            End If


#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDWG2 As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim acDWG2 As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Return Connect.DateLog & "Cmd OBJ" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab & acDWG2.Name

        End Function

        ''' <summary>
        ''' Revo Command: Traitement des propriétés
        ''' </summary>
        ''' <param name="Types">Type des objets</param>
        ''' <param name="OBJcoll">Collection des objets</param>
        ''' <remarks></remarks>
        Public Function PropOBJ(ByVal Types As String, ByVal OBJcoll As Collection, ByVal acCurDb As Database, ByVal OBJprop() As String)  ' Remarks

            If Types.ToUpper Like OBJprop(1).ToUpper Then

                Dim OBJColor As Color
                OBJColor = Color.FromColorIndex(ColorMethod.ByAci, 7)

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction() '' Start a transaction
                    For Each Obj In OBJcoll 'Traitement des objets: 

                        Try

                            If Obj.Layer.ToUpper Like OBJprop(0).ToUpper = True Then 'Filtre calque
                                Dim acEnt As Entity = acTrans.GetObject(Obj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                                If Not IsDBNull(acEnt) Then
                                    If OBJprop(2) <> Nothing Then
                                        OBJColor = AcadTrueColor(OBJprop(2))
                                        Obj.ColorIndex = OBJColor.ColorIndex '2  TrueColor = dbl(1-255)/R,V,B/ByLayer/ByBlock
                                    End If

                                    If OBJprop(3) <> Nothing Then Obj.Layer = OBJprop(3) ' 3  Layer = ""
                                    If OBJprop(4) <> Nothing Then Obj.Linetype = OBJprop(4) ' 4  LineType = ""
                                    If OBJprop(5) <> Nothing Then Obj.LinetypeScale = OBJprop(5) ' 5  LinetypeScale = dbl
                                    If OBJprop(6) <> Nothing Then Obj.LineWeight = ConvertLineWeight(OBJprop(6)) ' 6  Lineweight = dbl

                                    If Types = "MText" Then
                                        If OBJprop(7) <> Nothing Then Obj.Rotation = OBJprop(7) ' 7  Rotation = dbl
                                        'If OBJprop(8) <> Nothing Then ObjMText.TextStyleId = OBJprop(8) ' 8  StyleName = ""
                                        'If OBJprop(9) <> Nothing Then ObjMText.FlowDirection = OBJprop(9) ' 9  Alignment = TopCenter
                                        If OBJprop(10) <> Nothing Then Obj.Height = OBJprop(10) '10  Height = dbl
                                        'If OBJprop(11) <> Nothing Then ObjBL. = OBJprop(11) '11  ScaleFactor = dbl

                                        If OBJprop(12) <> Nothing Then '12  Text = "*"||"abc"
                                            Dim Param() As String = Split(OBJprop(12), "||")
                                            If Param.Count = 2 Then
                                                If Obj.Contents.ToUpper Like Param(0).ToUpper = True Then 'Test du texte de Mtext
                                                    If InStr(Param(1), "*") = 0 Then
                                                        Obj.Contents = Param(1)
                                                    Else
                                                        If Left(Param(1), 1) = "*" And Right(Param(1), 1) = "*" Then
                                                            Obj.Contents = Replace(Obj.Contents, Replace(Param(0), "*", ""), Replace(Param(1), "*", ""))   ' *xyz* -> *abc*
                                                        ElseIf Left(Param(1), 1) <> "*" And Right(Param(1), 1) = "*" Then
                                                            Obj.Contents = Replace(Param(1), "*", "") & Obj.Contents                       ' abc***
                                                        ElseIf Left(Param(1), 1) = "*" And Right(Param(1), 1) <> "*" Then
                                                            Obj.Contents = Obj.Contents & Replace(Param(1), "*", "")                       ' ***abc
                                                        Else
                                                            Obj.Contents = Replace(Param(1), "*", "") 'Nouveau nom
                                                        End If
                                                    End If
                                                End If
                                            End If
                                        End If

                                        If OBJprop(15) <> "" Then
                                            If OBJprop(15) = 1 Then Obj.Annotative = AnnotativeStates.True
                                            If OBJprop(15) = 0 Then Obj.Annotative = AnnotativeStates.False
                                        End If

                                        'If OBJprop(13) <> Nothing Then ObjBL.Item = OBJprop(13) '13  PropPerso ???
                                        'Index = "*"||"abc"

                                    End If
                                    If Types = "Hatch" Then
                                        If OBJprop(15) <> "" Then
                                            If OBJprop(15) = 1 Then Obj.Annotative = AnnotativeStates.True
                                            If OBJprop(15) = 0 Then Obj.Annotative = AnnotativeStates.False
                                        End If
                                    End If
                                    If Types = "BlockReference" Then
                                        If OBJprop(7) <> Nothing Then Obj.Rotation = OBJprop(7) ' 7  Rotation = dbl

                                        If OBJprop(12) <> Nothing Then '12  Text = "*"||"abc"
                                            Dim Param() As String = Split(OBJprop(12), "||")
                                            If Param.Count = 2 Then

                                                'Boucle sur tous les attributs
                                                Dim blkRef As BlockReference = DirectCast(acTrans.GetObject(Obj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockReference)
                                                Dim attCol As AttributeCollection = blkRef.AttributeCollection 'Get the block attributue collection
                                                'Dim attCol As AttributeCollection = br.AttributeCollection
                                                Dim attRef As AttributeReference = Nothing
                                                Dim CheckAtt As Boolean = False
                                                'Loop through the attribute collection
                                                For Each attId As ObjectId In attCol
                                                    'Get this attribute reference
                                                    attRef = DirectCast(acTrans.GetObject(attId, DatabaseServices.OpenMode.ForWrite), AttributeReference)

                                                    If OBJprop(14) = Nothing Then
                                                        CheckAtt = True
                                                    ElseIf attRef.Tag.ToLower Like OBJprop(14).ToLower Then
                                                        CheckAtt = True
                                                    Else
                                                        CheckAtt = False
                                                    End If

                                                    If CheckAtt = True Then 'Check if attribut filter is actived
                                                        'Store its name
                                                        'Dim attName As String = attRef.Tag
                                                        Try
                                                            If attRef.TextString.ToString.ToLower Like Param(0).ToLower Then
                                                                If InStr(Param(1), "*") = 0 Then
                                                                    attRef.TextString = Param(1)
                                                                Else
                                                                    If Left(Param(1), 1) = "*" And Right(Param(1), 1) = "*" Then
                                                                        attRef.TextString = Replace(attRef.TextString, Replace(Param(0), "*", ""), Replace(Param(1), "*", ""))   ' *xyz* -> *abc*
                                                                    ElseIf Left(Param(1), 1) <> "*" And Right(Param(1), 1) = "*" Then
                                                                        attRef.TextString = Replace(Param(1), "*", "") & attRef.TextString     ' abc***
                                                                    ElseIf Left(Param(1), 1) = "*" And Right(Param(1), 1) <> "*" Then
                                                                        attRef.TextString = attRef.TextString & Replace(Param(1), "*", "")     ' ***abc
                                                                    Else
                                                                        attRef.TextString = Replace(Param(1), "*", "") 'Nouveau nom               '
                                                                    End If
                                                                End If
                                                            End If

                                                            'Dim IDAttrname As Integer = Array.IndexOf(Pts.AttrName.ToArray, attName)
                                                            'If IDAttrname >= 0 Then attRef.TextString = Pts.AttrValue(IDAttrname) 'dr(attName)

                                                        Catch exA As Exception

                                                        Catch ex As COMException
                                                            'Erreur calques verrouillés
                                                            Connect.RevoLog(Connect.DateLog & "Cmd OBJ" & vbTab & False & vbTab & "Err attribut, layer locked: " & blkRef.Name & " > " & attRef.Tag.ToString)
                                                        End Try
                                                    End If
                                                Next


                                                'Boucle pour les propriétés dynamique
                                                '   a faire
                                                ' ----------------------------------
                                            End If

                                        End If

                                    End If

                                End If
                        End If

                        Catch
                        End Try
                    Next
                    '' Save the new object to the database
                    acTrans.Commit()
                    '' Dispose of the transaction
                End Using

                Return True
            Else
                Return False
            End If

        End Function

        ''' <summary>
        ''' Revo Command: Style modifier
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks></remarks>
        Public Function cmdSTY(ByVal Cmd As String)  ' Remarks
            'STY; Modif. le style de texte
            '#STY;filtre Style;filtre Type;[[PropA]]Val;[[PropB]]Val;
            '
            '0 : Filtre Nom Style: "Like"
            '1 : Filtre Type: "Like", Text, Cotation
            '
            'Text
            '2 : Fonts = ""
            '3 : Height = dbl

            Dim STYprop(0 To 3) As String
            Dim TabProp() As String
            Dim Cmdinfo() As String = SplitCmd(Cmd, 1)

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDoc As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Dim STY As AcadTextStyle
            Dim Styles As AcadTextStyles = acDoc.TextStyles


            If Cmdinfo(1) <> "" Then 'si la ligne n'est pas vide

                STYprop(0) = Cmdinfo(1) ' Name style
                STYprop(1) = Cmdinfo(2) ' Type style

                For i = 2 To Cmdinfo.Count - 1 'Boucle prop
                    If InStr(Cmdinfo(i), "[[") <> 0 And InStr(Cmdinfo(i), "]]") <> 0 Then
                        Cmdinfo(i) = Replace(Replace(Cmdinfo(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                        TabProp = SplitCmd(Cmdinfo(i), 1)
                        TabProp(0) = TabProp(0).ToLower

                        If "fonts" = TabProp(0) Then ' 2 : Fonts = ""
                            STYprop(2) = TabProp(1)
                        ElseIf "height" = TabProp(0) Then ' 3 : Height = dbl
                            If IsNumeric(TabProp(1)) Then STYprop(3) = TabProp(1)
                        End If
                    End If
                Next

                If "TEXT" Like STYprop(1).ToUpper Then 'Si type de style Texte
                    For Each STY In Styles 'Boucle sur les style de textes
                        If STY.Name.ToUpper Like STYprop(0).ToUpper = True Then 'Test du nom de style

                            'Si trouve le nom du style > traite au Param.
                            Try
                                If File.Exists(STYprop(2)) Then
                                    STY.fontFile = STYprop(2) ' 2 : Fonts = ""
                                Else
                                    Connect.RevoLog(Connect.DateLog & "Cmd STY" & vbTab & False & vbTab & "Erreur: file is missing" & vbTab & STYprop(2))
                                End If

                            Catch ex As Exception
                                Connect.RevoLog(Connect.DateLog & "Cmd STY" & vbTab & False & vbTab & "Erreur: " & vbTab & STYprop(2))
                            End Try
                        End If
                    Next
                End If

            Else 'ignore the command
                Return Connect.DateLog & "Cmd STY" & vbTab & False & vbTab & "Erreur: " & vbTab & STYprop(2)
            End If

            Return Connect.DateLog & "Cmd STY" & vbTab & True & vbTab & "Type: " & Cmdinfo(1)

        End Function


        Public Function AcadTrueColor(ByVal nColor As String)
            Dim ObjColor As Color
            ObjColor = Color.FromColorIndex(ColorMethod.None, 7)
            'Dim ObjColors As Autodesk.AutoCAD.Colors.Color
            'ObjColors = ACAD_COLOR '  AcColor.acWhite '  AcColor.acWhite

            If nColor <> "" Then
                If nColor.ToUpper = "BYLAYER" Then
                    ObjColor = Color.FromColorIndex(ColorMethod.ByLayer, 256)
                ElseIf nColor.ToUpper = "BYBLOCK" Then
                    ObjColor = Color.FromColorIndex(ColorMethod.ByBlock, 0)
                Else
                    If IsNumeric(nColor) And nColor > 0 And nColor < 256 Then
                        ObjColor = Color.FromColorIndex(ColorMethod.ByAci, CDbl(Trim(nColor)))
                    Else
                        Dim RBVcolor() As String
                        RBVcolor = Split(nColor, ",")
                        If RBVcolor.Length = 3 Then
                            If IsNumeric(RBVcolor(0)) And IsNumeric(RBVcolor(1)) And IsNumeric(RBVcolor(2)) Then
                                Try
                                    ObjColor = Color.FromRgb(RBVcolor(0), RBVcolor(1), RBVcolor(2))
                                Catch
                                    Connect.RevoLog(Connect.DateLog & "Command ColorRVB" & vbTab & False & vbTab & "Erreur color : " & nColor)
                                End Try
                            End If
                        End If
                    End If
                End If
            End If


            Dim ObjColorAcad As New ACAD_COLOR
            ObjColor = Color.FromColorIndex(ColorMethod.None, 7)
            'Dim ObjColors As Autodesk.AutoCAD.Colors.Color
            'ObjColors = ACAD_COLOR '  AcColor.acWhite '  AcColor.acWhite

            If nColor <> "" Then
                If nColor.ToUpper = "BYLAYER" Then
                    ObjColor = Color.FromColorIndex(ColorMethod.ByLayer, 256)
                ElseIf nColor.ToUpper = "BYBLOCK" Then
                    ObjColor = Color.FromColorIndex(ColorMethod.ByBlock, 0)
                Else
                    If IsNumeric(nColor) And nColor > 0 And nColor < 256 Then
                        ObjColor = Color.FromColorIndex(ColorMethod.ByAci, CDbl(Trim(nColor)))
                    Else
                        Dim RBVcolor() As String
                        RBVcolor = Split(nColor, ",")
                        If RBVcolor.Length = 3 Then
                            If IsNumeric(RBVcolor(0)) And IsNumeric(RBVcolor(1)) And IsNumeric(RBVcolor(2)) Then
                                Try
                                    ObjColor = Color.FromRgb(RBVcolor(0), RBVcolor(1), RBVcolor(2))
                                Catch
                                    Connect.RevoLog(Connect.DateLog & "Command ColorRVB" & vbTab & False & vbTab & "Erreur color : " & nColor)
                                End Try
                            End If
                        End If
                    End If
                End If
            End If

            Return ObjColor
        End Function


        ''' <summary>
        ''' Revo Command: Export
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks>#Export;Format;NomFichier;[[PropB]]Val;</remarks>

        Public Function cmdExport(ByVal Cmd As String)  ' Remarks
            'Exportation de données (si NomFichier vide demande dans le nom de fichier)

            Dim Cmdinfo() As String = SplitCmd(Cmd, 1)

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim curDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim curDwg As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Dim XMLformat As String = Ass.XMLformatBase
            Dim NbreFormat As Double = 0
            Dim ConfirmOK As Double = -1
            Dim Filters As String = ""
            Dim ExportName As String = ""
            Dim SaveFileDialogExport As New System.Windows.Forms.SaveFileDialog
            Dim OKExport As Boolean = False
            Dim NumFormat As Double = -1
            Dim CollBlock As New Collection
            Dim CheckDoublon As Boolean = False
            Dim Check3D As Boolean = True
            Dim ThemeValue As String = "MUT" ' REVOTHEME
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim FiltreActived As Boolean = False

            '1) Choix du format et fichier -----------------------------------------

            If Cmdinfo(1) = "" Then 'Sélection manuel du format d'exportation

                'Importation des formats points : format.xml **********
                'If File.Exists(XMLformat) Then 'Test de l'existance du fichier XML Revo
                '    Dim x As New RevoXML(XMLformat) 'lecture des extentions
                '    NbreFormat = x.countElements("/data-format/format")
                '    Dim VarExt As String = ""
                '    For i = 1 To NbreFormat
                '        VarExt = x.getElementValue("/data-format/format/extension", i)
                '        Filters += "|" & x.getElementValue("/data-format/format/name", i) & " (*." & VarExt & ")" & "|*." & VarExt
                '    Next
                '    If Left(Filters, 1) = "|" Then Filters = Mid(Filters, 2, Filters.Length - 1)
                'End If

                'SaveFileDialogExport.Filter = Filters
                'SaveFileDialogExport.Title = "Exporter des données"
                'SaveFileDialogExport.FilterIndex = Ass.FormatExport
                'SaveFileDialogExport.FileName = ""
                'ConfirmOK = SaveFileDialogExport.ShowDialog()

                Connect.Message("Exportation de points", "Ecriture des points en cours ... ", False, 0, 0, "hide")

                Dim ImportExportFrm As New frmImportExport
                ImportExportFrm.ActiveImport = False
                ImportExportFrm.ShowDialog()

                FiltreActived = True

                If ParamImportExportFrm.Count = 7 Then

                    '1 (TBNameFile.Text) 'Sélection du fichier + Chargement des formats
                    ExportName = ParamImportExportFrm(1)

                    '2) NumFormat
                    NumFormat = ParamImportExportFrm(2)

                    '3 (SelectedOjs) 'Sélection des objets manuelle (défault tout le fichiers)
                    CollBlock = ParamImportExportFrm(3)

                    '4 (cboTheme.Text) 'Points du thèmes
                    Select Case ParamImportExportFrm(4)
                        Case 1, 2, 3, 4, 5, 6 ' MO
                            ThemeValue = "MO"
                        Case 7, 8, 9, 10, 11, 12 'MUT
                            ThemeValue = "MUT"
                        Case Else '0, 13, 14, 15, 16
                            ThemeValue = "MUT"
                    End Select

                    '5 (TBCommunes.Text) 'Points des communes

                    '6 (cboxDoublons.CheckState) 'Détection des doublons
                    CheckDoublon = ParamImportExportFrm(6)

                    '7 Points 3D : OK
                    Check3D = ParamImportExportFrm(7)

                    If NumFormat < 0 Then NumFormat = 0
                    Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "FormatExport".ToLower, NumFormat.ToString)

                    If ExportName <> "" And CollBlock.Count > 0 Then
                        OKExport = True
                    End If

                End If


            Else 'Cherche les formats disponible dans format.xml

                '...

                ' Return Connect.DateLog & "Cmd Exemple" & vbTab & False & vbTab & "Erreur: " & vbTab & "NomFichier"
            End If


            '2) Choix des objets ------------------------------------------------

            'Choix avec filtre 
            If FiltreActived = False Then

                'if "select" = "all" then ....  or ""

                ' Start for selection of objects

                'Suppression messages
                Connect.Message("Exportation de points", "Ecriture des points en cours ... ", False, 0, 0, "hide")


                '' Start a transaction
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Dim acOPrompt As PromptSelectionOptions = New PromptSelectionOptions()
                    acOPrompt.MessageForAdding = "Sélectionner les points à exporter"
                    '' Request for objects to be selected in the drawing area
                    Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(acOPrompt)
                    '' If the prompt status is OK, objects were selected
                    If acSSPrompt.Status = PromptStatus.OK Then
                        Dim acSSet As SelectionSet = acSSPrompt.Value
                        '' Step through the objects in the selection set
                        For Each acSSObj As SelectedObject In acSSet
                            '' Check to make sure a valid SelectedObject object was returned
                            If Not IsDBNull(acSSObj) Then
                                '' Open the selected object for write
                                Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                                If Not IsDBNull(acEnt) Then
                                    If TypeName(acEnt) = "BlockReference" Then
                                        CollBlock.Add(acEnt)
                                    End If
                                End If
                            End If
                        Next
                    End If
                    '' Save the new object to the database
                    acTrans.Commit()
                    '' Dispose of the transaction
                End Using
                ' End selection of objects
            End If


            '3) Exportation OK

            'Check si AUCUN altitude
            '' Start a transaction
            Dim CheckAllZero As Boolean = True
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Try   'Traitement des block
                    For Each BLent As BlockReference In CollBlock 'Traitement des objets: BlockReference
                        If BLent.Position.Z <> 0 Then
                            CheckAllZero = False
                        End If
                    Next
                Catch ex As Exception
                End Try
            End Using
            If CollBlock.Count > 0 Then
                If Check3D And CheckAllZero Then
                    If MsgBox("Tous les objets ont une valeur Z = 0." & vbCrLf & "Souhaitez-vous activer l'exportation 2D" _
                              & vbCrLf & "(Oui : aucune Altitude reportée)", vbInformation + vbYesNo, "Erreur de détection altimétrique") = MsgBoxResult.Yes Then
                        Check3D = False
                    End If
                End If
            End If

            If OKExport = True Then ExportPTS(ExportName, NumFormat, CollBlock, acCurDb, ThemeValue, CheckDoublon, Check3D)


            Return Connect.DateLog & "Cmd Export" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab & ExportName


        End Function

        ''' <summary>
        ''' Revo Command: Export data
        ''' </summary>
        ''' <param name="ExportFileName">Exort name file</param>
        ''' <param name="NumFormat">Num for XML formats</param>
        ''' <param name="CollBlock">Object selected</param>
        ''' <remarks>Export de PTS</remarks>
        Public Function ExportPTS(ByVal ExportFileName As String, ByVal NumFormat As Double, ByVal CollBlock As Collection, ByVal acCurDb As Database, Theme As String, CheckDoublon As Boolean, Check3D As Boolean)
            'Exportation du format sélectionné

            If File.Exists(Ass.XMLformatBase) Then 'Si format XML exist
                Dim docXML As New System.Xml.XmlDocument
                docXML.Load(Ass.XMLformatBase)
                Dim NodeSymbols As System.Xml.XmlNodeList
                Dim Xroot As String = "/data-format/format[" & NumFormat & "]"
                Dim NodeTypes As System.Xml.XmlNodeList
                Dim NodeDatas As System.Xml.XmlNodeList
                Dim NodeTopics As System.Xml.XmlNodeList

                'Load Block Name (LOAD format.xml
                Dim AutorizName As String = ""
                NodeSymbols = docXML.SelectNodes(Xroot & "/topic/data/type/symbol/block")
                For Each NodeSymbol As System.Xml.XmlNode In NodeSymbols
                    AutorizName += "[" & NodeSymbol.InnerText & "]"  '[PTS_BF_BORNE] ' Avec REVOTHEME_....  !!!!
                Next

                ' Lecture du caractère de split
                Dim SplitCar As String = ""
                Try 'Charge le car Split (, ou . etc)
                    SplitCar = docXML.SelectSingleNode(Xroot & "/split").InnerText()
                Catch
                    MsgBox("Erreur de définition du caractère de Split")
                End Try


                ' Lecture du tri (sort)
                Dim SortCharStart As Double = 0
                Dim SortCharLong As Double = 1
                Try 'Charge le car Split (, ou . etc)
                    SortCharStart = docXML.SelectSingleNode(Xroot & "/sort/start").InnerText()
                    SortCharLong = docXML.SelectSingleNode(Xroot & "/sort/len").InnerText()
                Catch
                    '  MsgBox("Erreur de définition du caractère de tri")
                End Try


                Dim LinePts(0 To CollBlock.Count - 1) As String
                Dim NumPts As Double = 0

                '' Start a transaction
                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                    Try

                        'Traitement des block
                        For Each BLent As BlockReference In CollBlock 'Traitement des objets: BlockReference

                            Dim Btr As BlockTableRecord = CType(acTrans.GetObject(BLent.BlockTableRecord.ConvertToRedirectedId, DatabaseServices.OpenMode.ForRead, False), BlockTableRecord)
                            Dim dynObjectID As ObjectId = BLent.DynamicBlockTableRecord
                            Dim dynBtr As BlockTableRecord = CType(acTrans.GetObject(dynObjectID, DatabaseServices.OpenMode.ForRead, False), BlockTableRecord)

                            If InStr(Replace(AutorizName.ToUpper, "REVOTHEME_", ""), "[" & Replace(dynBtr.Name.ToUpper, "MUT_", "") & "]") <> 0 Then 'Traitement si nom de block OK

                                Dim BlockName, VarColor, VarLayer, VarLink, PositionX, PositionY, PositionZ, VarScale, Rotation, VarSymbol As String
                                Dim TypeTag As String = ""


                                'Général
                                BlockName = dynBtr.Name
                                VarColor = BLent.ColorIndex.ToString
                                VarLayer = BLent.Layer
                                VarLink = "" 'BLent.Hyperlinks.Count
                                PositionX = BLent.Position.Coordinate(0).ToString
                                PositionY = BLent.Position.Coordinate(1).ToString
                                PositionZ = BLent.Position.Coordinate(2).ToString


                                VarScale = BLent.ScaleFactors(0)
                                Rotation = BLent.Rotation
                                VarSymbol = dynBtr.Name

                                'Lecture XML --------------------------------------------

                                'Boucle dans les TOPIC
                                Dim NumTopic As Double = 0
                                NodeTopics = docXML.SelectNodes(Xroot & "/topic")
                                Dim NoTopic As Boolean = True

                                For Each NodeTopic As System.Xml.XmlNode In NodeTopics
                                    NumTopic += 1 '(1er bloc = 1)

                                    Dim ValideTopic As Boolean = False
                                    Dim TopicValue As String = docXML.SelectNodes(Xroot & "/topic[" & NumTopic & "]").ItemOf(0).Attributes.ItemOf(0).InnerText

                                    'Test les /topic
                                    If TopicValue = "" Then 'Sans code Topic
                                        ValideTopic = True
                                    Else ' Avec code Topic
                                        Dim DataLayerNode As System.Xml.XmlNodeList

                                        DataLayerNode = docXML.SelectNodes(Xroot & "/topic[" & NumTopic & "]/layer")
                                        For Each DataLayer As System.Xml.XmlNode In DataLayerNode
                                            If BLent.Layer.ToLower Like DataLayer.InnerText.ToLower Then
                                                ValideTopic = True
                                                NoTopic = False
                                                Exit For
                                            End If
                                        Next

                                        If NoTopic = True And NumTopic = NodeTopics.Count Then
                                            NumTopic = 1 ' Pas trouvé de Topic = 1er par default
                                            ValideTopic = True
                                        End If

                                    End If

                                    If ValideTopic = True Then

                                        'Boucle dans les DATA
                                        Dim NumDatas As Double = 0
                                        NodeDatas = docXML.SelectNodes(Xroot & "/topic[" & NumTopic & "]/data")
                                        For Each NodeData As System.Xml.XmlNode In NodeDatas
                                            NumDatas += 1 '(1er bloc = 1)

                                            Dim ValideData As Boolean = False
                                            Dim DataValue As String = docXML.SelectNodes(Xroot & "/topic[" & NumTopic & "]/data[" & NumDatas & "]").ItemOf(0).Attributes.ItemOf(0).InnerText

                                            'Test les /topic/data
                                            If DataValue = "" Then 'Sans code Data
                                                ValideData = True
                                            Else ' Avec code Data
                                                Dim DataLayerNode As System.Xml.XmlNodeList
                                                DataLayerNode = docXML.SelectNodes(Xroot & "/topic[" & NumTopic & "]/data[" & NumDatas & "]/layer")
                                                For Each DataLayer As System.Xml.XmlNode In DataLayerNode
                                                    If BLent.Layer.ToLower Like DataLayer.InnerText.ToLower Then
                                                        ValideData = True
                                                        Exit For
                                                    End If
                                                Next
                                            End If

                                            If ValideData = True Then

                                                'Boucle dans les Type
                                                NodeTypes = docXML.SelectNodes(Xroot & "/topic[" & NumTopic & "]/data[" & NumDatas & "]/type")
                                                For Each NodeType As System.Xml.XmlNode In NodeTypes
                                                    TypeTag = NodeType.Attributes.ItemOf(0).InnerText()

                                                    'Type paramètre
                                                    Dim AttValue, PropertyPts, xStart, xLen, xRound, xAlign, IDtype, PtsVisible, PtsFormat, PtsScale As String
                                                    AttValue = "" : PropertyPts = "" : xStart = "" : xLen = "" : xRound = "" : xAlign = "" : IDtype = "" : PtsVisible = "" : PtsFormat = "" : PtsScale = ""
                                                    Dim PtsReplaceVal As New List(Of String)
                                                    Dim PtsReplace As New List(Of String)
                                                    Dim ListLayerVal As New List(Of String)
                                                    Dim ListLayer As New List(Of String)

                                                    'Boucle dans les paramètres d'un type
                                                    Dim SymbolOK As Boolean = False

                                                    For Each Prop As System.Xml.XmlNode In NodeType
                                                        Dim Var As String = Prop.InnerText
                                                        Dim PropName As String = Prop.Name.ToLower
                                                        'Chargement des paramètres
                                                        If PropName = "start" Then
                                                            xStart = Trim(Var)
                                                        ElseIf PropName = "len" Then
                                                            xLen = Trim(Var)
                                                        ElseIf PropName = "round" Then
                                                            xRound = Trim(Var)
                                                        ElseIf PropName = "align" Then
                                                            xAlign = Trim(Var.ToLower)
                                                        ElseIf PropName = "property" Then
                                                            PropertyPts = Trim(Var.ToLower)
                                                        ElseIf PropName = "id" Then
                                                            IDtype = Trim(Var.ToLower)
                                                        ElseIf PropName = "visible" Then
                                                            PtsVisible = Trim(Var.ToLower)
                                                        ElseIf PropName = "replace" Then
                                                            PtsReplaceVal.Add(Prop.Attributes.ItemOf(0).InnerText)
                                                            PtsReplace.Add(Var)
                                                        ElseIf PropName = "layer" Then
                                                            ListLayerVal.Add(Prop.Attributes.ItemOf(0).InnerText) ' MO_BF_PTS
                                                            ListLayer.Add(Var) ' 06
                                                        ElseIf PropName = "format" Then
                                                            PtsFormat = Var ' VD????00????
                                                        ElseIf PropName = "scale" Then
                                                            PtsScale = Var ' * 1000
                                                        ElseIf PropName = "symbol" And SymbolOK = False Then
                                                            If Prop.Attributes.ItemOf(0).InnerText.ToLower = "default" Then
                                                                For Each SymbInfo As System.Xml.XmlNode In Prop
                                                                    If SymbInfo.Name.ToLower = "block" Then BlockName = Replace(Replace(SymbInfo.InnerText, "REVOTHEME_", Theme & "_"), "MO_", "") ' Ajout THEME 06.10.2015 THA
                                                                    If SymbInfo.Name.ToLower = "layer" Then VarLayer = Replace(SymbInfo.InnerText, "REVOTHEME_", Theme & "_") ' Ajout THEME 06.10.2015 THA
                                                                    If SymbInfo.Name.ToLower = "color" Then VarColor = SymbInfo.InnerText
                                                                    If SymbInfo.Name.ToLower = "rotation" Then Rotation = SymbInfo.InnerText
                                                                Next
                                                            ElseIf Prop.Attributes.ItemOf(0).InnerText.ToLower <> "default" Then
                                                                For Each SymbInfo As System.Xml.XmlNode In Prop
                                                                    Dim BlockNameXMKL As String = Replace(SymbInfo.InnerText.ToUpper, "REVOTHEME_", "*") ' Ajout THEME 06.10.2015 THA
                                                                    If SymbInfo.Name.ToLower = "block" And dynBtr.Name.ToUpper Like BlockNameXMKL Then 'LIKE Ajout THEME 06.10.2015 THA
                                                                        VarSymbol = Prop.Attributes.ItemOf(0).InnerText.ToLower
                                                                        SymbolOK = True
                                                                        Exit For
                                                                    End If
                                                                Next
                                                            End If
                                                        End If
                                                    Next ' fin de la boucle des propriétés dans un type


                                                    If dynBtr.IsDynamicBlock Then
                                                        For Each BLdyn As DynamicBlockReferenceProperty In BLent.DynamicBlockReferencePropertyCollection
                                                            If TypeTag.ToUpper = BLdyn.PropertyName.ToUpper Then
                                                                AttValue = BLdyn.Value
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If

                                                    If dynBtr.HasAttributeDefinitions Then
                                                        Dim ent As Entity
                                                        For Each BLattrID As ObjectId In BLent.AttributeCollection
                                                            ent = acTrans.GetObject(BLattrID, DatabaseServices.OpenMode.ForRead, True, True)
                                                            If TypeOf ent Is AttributeReference Then
                                                                Dim attref As AttributeReference = ent
                                                                If TypeTag.ToUpper = attref.Tag.ToUpper Then
                                                                    AttValue = attref.TextString
                                                                    Exit For
                                                                End If
                                                            End If
                                                        Next
                                                    End If


                                                    'ID type
                                                    'If IDtype = "numeric" Then
                                                    '   NumID = Trim(ValPts)
                                                    'End If

                                                    'Test des Property  (écrase l'ancienne valeur)
                                                    If PropertyPts = "positionx" Then
                                                        AttValue = PositionX
                                                    ElseIf PropertyPts = "positiony" Then
                                                        AttValue = PositionY
                                                    ElseIf PropertyPts = "positionz" Then

                                                        If Check3D = True Then
                                                            Dim ValAlt As Double = PositionZ 'Par défault
                                                            If InStr(AttValue, "?") <> 0 Then AttValue = 0 'Supp ?
                                                            If IsNumeric(AttValue) = False Then AttValue = 0 'si pas numérique

                                                            'Contrôle si Altitude OK
                                                            If CDbl(PositionZ) = 0 And CDbl(AttValue) <> 0 Then 'Z = 0 et Alt <> 0  => Alt
                                                                ValAlt = AttValue
                                                            ElseIf CDbl(AttValue) = 0 And CDbl(PositionZ) <> 0 Then 'Alt = 0 et Z <> 0 => Z
                                                                ValAlt = PositionZ
                                                            ElseIf CDbl(PositionZ) <> 0 And CDbl(AttValue) <> 0 And CDbl(PositionZ) <> CDbl(AttValue) Then 'Z <> 0 et Alt <> 0 et Alt <> Z = Questions
                                                                If MsgBox("Souhaitez-vous garder l'altitude du bloc ? ( Z = " & PositionZ & ")" & vbCrLf & vbCrLf & "non : garder l'altitude de l'attribut (" & AttValue & ")", vbInformation + vbYesNo, "Conflit de la définition de l'altitude") = MsgBoxResult.Yes Then
                                                                    ValAlt = PositionZ
                                                                Else
                                                                    ValAlt = AttValue
                                                                End If
                                                            End If
                                                            AttValue = ValAlt
                                                        Else
                                                            AttValue = ""
                                                        End If

                                                    ElseIf PropertyPts = "rotation" Then
                                                        AttValue = Rotation
                                                    ElseIf PropertyPts = "color" Then
                                                        AttValue = VarColor
                                                    ElseIf PropertyPts = "layer" Then
                                                        AttValue = VarLayer
                                                    ElseIf PropertyPts = "link" Then
                                                        AttValue = VarLink
                                                    ElseIf PropertyPts = "scale" Then
                                                        AttValue = VarScale
                                                    ElseIf PropertyPts = "symbol" Then
                                                        AttValue = VarSymbol
                                                    End If



                                                    'Format : Scale  export multiply (only num) => *1000
                                                    If PtsScale <> "" And IsNumeric(PtsScale) And IsNumeric(AttValue) Then
                                                        AttValue = CDbl(AttValue) * CDbl(PtsScale)
                                                    End If

                                                    'Format : Round
                                                    If xRound <> "" And IsNumeric(xRound) And IsNumeric(AttValue) Then
                                                        AttValue = System.Math.Round(CDbl(AttValue), CInt(xRound))
                                                        AttValue = Format(CDbl(AttValue), Replace("#0.R", "R", Replace(Space(xRound), " ", "0")))
                                                    End If

                                                    If PtsFormat <> "" Then 'VD????00????  >>  'VD 0090 00 0600
                                                        Dim Replace0Space As Boolean = False
                                                        Dim ActiveBloc As Double = 0

                                                        'Formatage export (full format)
                                                        If Left(PtsFormat, 1) = System.Convert.ToChar(34) And Right(PtsFormat, 1) = System.Convert.ToChar(34) Then
                                                            Dim FormX As String = Mid(PtsFormat, 2, Len(PtsFormat) - 2)

                                                            'MsgBox(NumPts.ToString)
                                                            If Len((NumPts + 1).ToString) < 5 Then FormX = Replace(FormX, "####", Replace(Space(4 - Len((NumPts + 1).ToString)), " ", "0") & (NumPts + 1).ToString)
                                                            If Len((NumPts + 1).ToString) < 4 Then FormX = Replace(FormX, "###", Replace(Space(3 - Len((NumPts + 1).ToString)), " ", "0") & (NumPts + 1).ToString)
                                                            If Len((NumPts + 1).ToString) < 3 Then FormX = Replace(FormX, "##", Replace(Space(2 - Len((NumPts + 1).ToString)), " ", "0") & (NumPts + 1).ToString)

                                                            ' MsgBox(Replace(Space(4 - Len(LinePts.Length + 1)), " ", "0") & LinePts.Length + 1)
                                                            ' MsgBox(LinePts.Length)

                                                            'Correction du signe - (>0)
                                                            If IsNumeric(AttValue) Then If CDbl(AttValue) < 0 Then FormX = Replace(FormX, "+", "-") : AttValue = CDbl(AttValue * (-1))

                                                            'Encolonnement avec Alignement
                                                            If xAlign.ToLower = "right" Then
                                                                AttValue = Mid(FormX, 1, Len(FormX) - Len(AttValue)) & AttValue
                                                            Else 'Align = left
                                                                AttValue = AttValue & Mid(FormX, Len(AttValue) + 1, Len(FormX) - Len(AttValue))
                                                            End If

                                                        Else 'Formatage perso import/export
                                                            Dim AttValueNew As String = ""
                                                            For i = 1 To Len(PtsFormat)
                                                                If Mid(PtsFormat, i, 1) = "?" Then   'ajoute le caractère ?
                                                                    If ActiveBloc = 0 Then ActiveBloc = 1
                                                                    If ActiveBloc = 1 And Mid(AttValue, i, 1) = "0" Then
                                                                        Replace0Space = True
                                                                    ElseIf ActiveBloc = 1 And Mid(AttValue, i, 1) <> "0" Then
                                                                        ActiveBloc = 2
                                                                        Replace0Space = False
                                                                    Else
                                                                        Replace0Space = False
                                                                    End If

                                                                    If Replace0Space = True Then 'remplace 0 par Espace
                                                                        AttValueNew += Replace(Mid(AttValue, i, 1), "0", " ")
                                                                    Else ' Sinon ajout la valeur
                                                                        AttValueNew += Mid(AttValue, i, 1)
                                                                    End If
                                                                Else
                                                                    ActiveBloc = 0
                                                                    Replace0Space = False
                                                                End If
                                                            Next
                                                            AttValue = AttValueNew
                                                        End If

                                                    End If

                                                    'Replace Data: Var => Value
                                                    For i = 0 To PtsReplaceVal.Count - 1
                                                        If PtsReplaceVal(i).ToLower = "zero" Then
                                                            'If IsNumeric(AttValue) Then
                                                            If SplitCar = "" Then
                                                                Try
                                                                    If AttValue = "" Then AttValue = 0
                                                                    If CDbl(AttValue) = 0 Then AttValue = PtsReplace(i)
                                                                Catch
                                                                End Try
                                                            Else 'Si séparateur = découpe à l'intérieur du bloc
                                                                If CDbl(Replace(Replace(Mid(AttValue, xStart, xLen), "+", ""), "-", "")) = 0 Then
                                                                    AttValue = PtsReplace(i)
                                                                End If
                                                            End If
                                                            'End If

                                                        ElseIf PtsReplaceVal(i) <> "" Then
                                                            AttValue = Replace(AttValue, PtsReplaceVal(i), PtsReplace(i))
                                                        Else 'Si Value "" et pas de texte "zero"
                                                            If PtsReplace(i) <> "zero" Then AttValue = AttValue & PtsReplace(i)
                                                        End If
                                                    Next

                                                    'Select Layer
                                                    For i = 0 To ListLayerVal.Count - 1
                                                        If BLent.Layer.ToUpper = ListLayerVal(i).ToUpper Then
                                                            AttValue = ListLayer(i)
                                                            Exit For
                                                        End If
                                                    Next



                                                    'Format : Len
                                                    If SplitCar = "" Then 'Si pas de Split
                                                        If xLen <> "" And IsNumeric(xLen) Then
                                                            'Align + space
                                                            If xAlign.ToLower = "right" Then
                                                                If Len(AttValue) <= xLen Then
                                                                    AttValue = Space(CDbl(xLen) - Len(AttValue)) & AttValue
                                                                Else
                                                                    AttValue = Right(AttValue, CDbl(xLen))
                                                                End If
                                                            Else 'Align = left
                                                                If Len(AttValue) <= xLen Then
                                                                    AttValue = AttValue & Space(CDbl(xLen) - Len(AttValue))
                                                                Else
                                                                    AttValue = Left(AttValue, CDbl(xLen))
                                                                End If
                                                            End If
                                                        End If
                                                    Else 'Si Split
                                                        'Pas de start
                                                        'Pas de len
                                                        'Pas d' Align (sauf avec Format)
                                                    End If

                                                    'Add value to the line + Split
                                                    If LinePts(NumPts) = "" Then 'Si vide sans Split
                                                        LinePts(NumPts) += AttValue
                                                    Else
                                                        LinePts(NumPts) += SplitCar & AttValue 'Avec Split
                                                    End If

                                                    'Statue bar
                                                    If Right(NumPts.ToString, 2) = "50" Or Right(NumPts.ToString, 2) = "00" Then Connect.Message("Exportation de points", "Ecriture des points en cours ... " & vbCrLf & "Nombre de point(s) : " & NumPts & "/" & CollBlock.Count, False, 70, 100)

                                                Next 'Fin de la boucle des types

                                            End If 'Fin test si traitement Data

                                        Next 'Fin de la boucle des Data

                                    End If 'Fin test si traitement Topic

                                Next 'Fin de la boucle des Topic

                            End If

                            If LinePts(NumPts) <> "" Then NumPts += 1 ' incrémentation

                        Next 'Fin du la boucle des blocks
                        '' Dispose of the transaction

                    Catch ex As COMException
                        MsgBox("Erreur d'exportation" & vbCrLf & ex.Message)
                    End Try
                End Using

                'redimensionnent les points
                ReDim Preserve LinePts(0 To NumPts - 1)

                'Tri Points : Sort
                If SortCharStart <> 0 And LinePts.Count <> 0 Then
                    Try
                        Dim LinePtsSort = New List(Of KeyValuePair(Of String, Integer))()
                        For i = 0 To LinePts.Count - 1
                            Dim IDsort As String = ""
                            If SplitCar <> "" Then 'Par séparateur
                                Dim NumLigne() As String = Split(LinePts(i), SplitCar)
                                IDsort = NumLigne(SortCharStart - 1)
                            Else 'Par colonne
                                If LinePts(i) IsNot Nothing Then IDsort = Mid(LinePts(i), SortCharStart, SortCharLong)
                            End If
                            If IDsort = "" Or IsNumeric(IDsort) = False Then IDsort = 100000 + i
                            If LinePts(i) Is Nothing Then LinePts(i) = ""
                            LinePtsSort.Add(New KeyValuePair(Of String, Integer)(LinePts(i), CInt(IDsort)))
                        Next
                        LinePtsSort.Sort(AddressOf Compare1)
                        For i = 0 To LinePts.Count - 1
                            LinePts(i) = LinePtsSort.Item(i).Key
                        Next
                    Catch
                    End Try
                End If

                'Write data file

                If Revo.RevoFiler.EcritureFichier(ExportFileName, LinePts, False) Then
                End If


                Return True
            Else
                Return False
            End If


        End Function

        ''' <summary>
        ''' Revo Command: Import data
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks>#Import;Format;NomFichier;@XYZ;Echelle /Echelle/;Rotation;</remarks>

        Public Function cmdImport(ByVal Cmd As String, ByVal ImportFiles As List(Of String))
            '#Import;Format;[[PropA]]Val;[[PropB]]Val;

            'Importation de données (si NomFichier vide import uniquement dans le fichier courant)

            'Variables(requises)

            '0 : Format()
            '[[Paramètre]]

            '1 : Name = "NomFichier"
            '2 : Space = *PaperSpaceXYZ*/Current/Paper/Model (défault), uniq. DWG & DXF
            '3 : XYZInsertionPoint = dbl,dbl,dbl (@)
            '4 : ScaleFactor = /Echelle/
            '5 : Rotation = dbl

            Dim Importprop(0 To 5) As String

            Dim TestReturn As Boolean = True
            Dim Cmdinfos() As String = SplitCmd(Cmd, 1)
            Dim Test As String = Connect.DateLog & "Cmd Import" & vbTab & True & vbTab & "Imported file: "
            Dim SelFiles As String = ""
            Dim ActivateITF As Boolean = False
            Dim TestXML As Boolean = False
            Dim XMLformat As String = Ass.XMLformatBase
            Dim docXML As New System.Xml.XmlDocument
            Dim FormatInfo As New List(Of String)
            Dim NbreFormat As Double = 0
            Dim NumFormat As Double = 0 ' = Valeur par défault pour le format de fichier
            Dim FormatExt() As String

            Importprop(0) = Cmdinfos(1) ' 0 : Format()

            For i = 1 To Cmdinfos.Count - 1 'Boucle prop
                If InStr(Cmdinfos(i), "[[") <> 0 And InStr(Cmdinfos(i), "]]") <> 0 Then
                    Cmdinfos(i) = Replace(Replace(Cmdinfos(i), "[[", "", 1, 1), "]]", ";", 1, 1)
                    Dim TabProp() As String = SplitCmd(Cmdinfos(i), 1)
                    TabProp(0) = TabProp(0).ToLower

                    If "name" = TabProp(0) Then '     1 : Name = "NomFichier"
                        Importprop(1) = TabProp(1)
                    ElseIf "space" = TabProp(0) Then '2 : Space = *PaperSpaceXYZ*/Current/Paper/Model (défault), uniq. DWG & DXF
                        Importprop(2) = TabProp(1)
                    ElseIf "xyzinsertionpoint" = TabProp(0) Then '3 : XYZInsertionPoint = dbl,dbl,dbl (@)
                        Importprop(3) = TabProp(1)
                    ElseIf "scalefactor" = TabProp(0) Then '4 : ScaleFactor = /Echelle/
                        If IsNumeric(TabProp(1)) Then Importprop(4) = TabProp(1)
                    ElseIf "rotation" = TabProp(0) Then '5 : Rotation = dbl
                        If IsNumeric(TabProp(1)) Then Importprop(5) = TabProp(1)
                    End If
                End If
            Next

            If Importprop(2) Is Nothing Then Importprop(2) = "Model"

            If File.Exists(XMLformat) Then
                docXML.Load(XMLformat)
                FormatInfo = conn.ReadXMLformat(docXML)
                NbreFormat = FormatInfo(0)
            Else
                FormatInfo.Add(0)
                FormatInfo.Add("")
                FormatInfo.Add("")
                FormatInfo.Add("")
            End If

            'Si un fichier définit dans la commande (supprime le listing de fichier)
            If Importprop(1) IsNot Nothing And File.Exists(Importprop(1)) Then
                ImportFiles.Clear()
                ImportFiles.Add(Importprop(1)) 'CORRECTION en (1) au lieu de (2)
            End If

            'Si pas de fichiers sélectionné
            If ImportFiles.Count < 1 Then

                Dim ConfirmOK As Double = 0
                Dim DialogBoxOpen As New System.Windows.Forms.OpenFileDialog
                Dim FormatSelect As Double = 1

                Try
                    'Importation des formats points : format.xml **********

                    Dim Filters As String = ""

                    'Suppression messages
                    Connect.Message("Importation de données", "... ", False, 0, 0, "hide")
                    ' Connect.Message(Ass.xTitle, "En cours d'importation ...", False, 100, 100, "close")

                    Revo.connect.ActDWG = True
                    Revo.connect.ActDXF = True
                    Revo.connect.ActPTS = True
                    Revo.connect.ActITF = True


                    'Test si fichier compatible Interlis-MOVD
                    Revo.connect.ActITF = InterlisProjectTest()

                    'Chargement des filtres
                    If Revo.connect.ActDWG = True Then Filters += "*.dwg"
                    '    If Revo.connect.ActDXF = True Then Filters += ";*.dxf"
                    If Revo.connect.ActITF = True Then Filters += ";*.itf"
                    If FormatInfo(2) <> "" And Revo.connect.ActPTS = True Then Filters += ";" & FormatInfo(2)

                    If Filters <> "" Then Filters = "Formats compatibles|" & Filters
                    If Revo.connect.ActDWG = True Then Filters += "|Dessin (*.dwg)|*.dwg" : FormatSelect += 1
                    '   If Revo.connect.ActDXF = True Then Filters += "|DXF (*.dxf)|*.dxf" : FormatSelect += 1
                    If Revo.connect.ActITF = True Then Filters += "|Interlis (*.itf)|*.itf" : FormatSelect += 1


                    If FormatInfo(1) <> "" And Revo.connect.ActPTS = True Then Filters += "|" & FormatInfo(1)
                    If Left(Filters, 1) = "|" Then Filters = Mid(Filters, 2, Filters.Length - 1)

                    Try
                        DialogBoxOpen.Filter = Filters
                        DialogBoxOpen.Title = "Importer des données"
                        DialogBoxOpen.Multiselect = True
                        DialogBoxOpen.FilterIndex = Ass.FormatImport
                        DialogBoxOpen.FileName = ""
                        ' DialogBoxOpen.InitialDirectory = "c:\"
                        ConfirmOK = DialogBoxOpen.ShowDialog()
                    Catch 'ex As Exception
                        '  MsgBox("Erreur : ") '& ex.Message)
                    End Try

                Catch 'ex As Exception
                    '  MsgBox(ex.Message)
                End Try

                Try

                    If ConfirmOK = 1 Then
                        ImportFiles.Clear()
                        NumFormat = DialogBoxOpen.FilterIndex - FormatSelect
                        If NumFormat < 0 Then NumFormat = 0
                        Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "FormatImport".ToLower, FormatSelect + NumFormat)

                        Dim strFiles() As String = DialogBoxOpen.FileNames
                        For i = 0 To strFiles.GetUpperBound(0)
                            ImportFiles.Add(strFiles(i) & "|" & NumFormat)
                        Next
                    End If

                Catch 'ex As Exception

                End Try
            End If



            'Importation des formats de points : format.xml **********
            'Test de l'existance du fichier XML Revo
            If File.Exists(XMLformat) Then
                'lecture des extentions
                If Importprop(0) IsNot Nothing Then
                    Dim FormatCodes() As String = Split(FormatInfo(3), "|")
                    For i = 0 To NbreFormat - 1 'Lecture des code de formats <format code="MD01_PTP"
                        If FormatCodes(i).ToLower = Importprop(0).ToLower Then
                            NumFormat = i + 1
                            Exit For
                        End If
                    Next
                End If
                FormatExt = Split(FormatInfo(2), ";")
                TestXML = True
            Else
                FormatExt = Split("", ";")
            End If


            ' Boucle dans tout les fichiers à importer
            Dim OKimport As String = ""
            For Each Datafile In ImportFiles
                NumFormat = 0 ' = Valeur par défault pour le format de fichier

                'Chargement du format de fichier (selon sélection)
                If InStr(Datafile, "|") <> 0 Then
                    Dim SplitFormat() As String = Split(Datafile, "|")
                    Datafile = SplitFormat(0)
                    NumFormat = SplitFormat(1)
                End If

                If IO.File.Exists(Datafile) = True Then

                    If Right(Datafile.ToUpper, 4) = ".ITF" Then ' Import ITF Interlis ********************************

                        If InterlisProjectTest() = True Then
                            Connect.Message("Importation", "... ", False, 0, 0, "hide")
                            MyCommands.Revo_ImportITF(Datafile)

                        Else 'Format du fichier non compatible
                            Connect.Message("Importation", "... ", False, 0, 0, "hide")
                            MsgBox("Pour charger un fichier Interlis, il est nécessaire d'avoir un dessin basé sur un Gabarit compatible" & vbCrLf _
                                                                 & "(créer un nouveau projet ou lancer une migration)", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Impossible d'importer le fichier")
                            Test = "STOP"
                        End If


                    ElseIf Right(Datafile.ToUpper, 4) = ".DWG" Then '  DWG import     *********************************

                        'Importer l'espace Objet et papier
                        InsertDWG(Datafile, Importprop(2), Importprop(3))

                    ElseIf Right(Datafile.ToUpper, 4) = ".DXF" Then '  DXF import     *********************************
                        ' à développer


                        'Importation des fichiers points : format.xml
                    ElseIf TestXML Then 'test de l'initialisation du format.xml

                        'Sélection des paramètres
                        Dim ImportExportFrm As New frmImportExport
                        ImportExportFrm.ActiveImport = True
                        ImportExportFrm.ShowDialog()


                        Dim ThemeValue As String = "MUT" ' REVOTHEME
                        Dim ParamCheckDoublon As Boolean = False 'Détection des doublons
                        Dim ParamCheck3D As Boolean = True  'Points 3D : OK


                        If ParamImportExportFrm.Count = 7 Then

                            If ParamImportExportFrm(4) = 1 Then 'Points du thèmes
                                ThemeValue = "MO"
                            ElseIf ParamImportExportFrm(4) = 2 Then
                                ThemeValue = "MUT"
                            Else ' 0 Automatique
                            End If

                            ParamCheckDoublon = ParamImportExportFrm(6) 'Détection des doublons
                            ParamCheck3D = ParamImportExportFrm(7)  'Points 3D : OK

                        End If

                        If NumFormat = 0 Then 'Recherche du 1er format via l'extention
                            For i = 0 To NbreFormat - 1
                                If Right(FormatExt(i).ToLower, 4) = Right(Datafile.ToLower, 4) Then
                                    NumFormat = i + 1
                                    Exit For
                                End If
                            Next
                            ImportFormats(Datafile, docXML, NumFormat, ThemeValue, ParamCheckDoublon, ParamCheck3D)
                        Else 'importation avec le format définit
                            ImportFormats(Datafile, docXML, NumFormat, ThemeValue, ParamCheckDoublon, ParamCheck3D)
                        End If

                        OKimport += ", " & Datafile
                        'Test = Connect.DateLog & "Cmd Import" & vbTab & True & vbTab & "Imported file: " & vbTab & OKimport
                    End If

                    'End    Importation des fichiers points : format.xml **********


                Else 'ignore the command
                    'If Test <> "" Then Test = Test & vbCrLf
                    'Test = Connect.DateLog & "Cmd Import" & vbTab & False & vbTab & "File non-existent: " & vbTab & Datafile
                End If
            Next

            Return Test
        End Function

        Public Function ImportFormats(ByVal Datafile As String, ByVal docXML As System.Xml.XmlDocument, ByVal NumFormat As Double, Theme As String, CheckDoublon As Boolean, Check3D As Boolean)

            ' Collection : BlockName - Action - N°ID - Color - Layer - Link - X - Y - Z - Scale , Rotation , Attibuts list (Name,Value)
            Dim PtsColl As New Collection
            Dim TestImport As Boolean
            Dim Xroot As String = "/data-format/format[" & NumFormat & "]"
            Dim Ligne As String = ""

            If File.Exists(Datafile) And NumFormat <> 0 Then
                If docXML.SelectSingleNode(Xroot & "/structure").InnerText.ToLower = "line" Then
                    Try
                        Dim sr As StreamReader = New StreamReader(Datafile, System.Text.Encoding.Default)
                        '--- Traitement du fichier ligne par ligne

                        While Not sr.EndOfStream()
                            Ligne = sr.ReadLine()
                            'Découpage/Formatage des lignes
                            If Ligne <> "" Then PtsColl.Add(ImportLine(Ligne, docXML, NumFormat, Theme, CheckDoublon, Check3D))
                            If Right(PtsColl.Count.ToString, 2) = "50" Or Right(PtsColl.Count.ToString, 2) = "00" Then Connect.Message("Importation de points", "Lecture des points en cours ... " & vbCrLf & "Nombre de point(s) : " & PtsColl.Count, False, 60, 100)
                        End While
                        '--- Referme StreamReader
                        sr.Close()

                        'Insertion des blocks + Dessin avec les Actions
                        TestImport = WriteBlock(PtsColl)


                    Catch ex As Exception
                        'Throw ex
                        Connect.RevoLog(Connect.DateLog & "Cmd Import" & vbTab & False & vbTab & "Reading error: " & ex.Message & vbTab & Datafile)
                    Catch ex As COMException
                        ' MsgBox("Erreur COM")
                    End Try

                ElseIf docXML.SelectSingleNode(Xroot & "/structure").InnerText.ToLower = "xlsmatrix" Then
                    Try
                        PtsColl = ImportXLSmatrix(Datafile, docXML, NumFormat)
                        TestImport = WriteBlock(PtsColl) 'Insertion des blocks + Dessin avec les Actions

                    Catch ex As Exception
                        'Throw ex
                        Connect.RevoLog(Connect.DateLog & "Cmd Import" & vbTab & False & vbTab & "Reading error: " & ex.Message & vbTab & Datafile)
                    Catch ex As COMException
                        ' MsgBox("Erreur COM")
                    End Try
                End If
            Else
                Connect.RevoLog(Connect.DateLog & "Cmd Import" & vbTab & False & vbTab & "Reading error: " & "file not found" & vbTab & Datafile)
            End If

            Return True
        End Function
        Public Function ImportLine(ByVal Line As String, ByVal x As System.Xml.XmlDocument, ByVal NumFormat As Double, Theme As String, CheckDoublon As Boolean, Check3D As Boolean) As RevoPTS

            'Propriety >> Type objet : BlockName - Action - N°ID - Color - Layer - Link - X - Y - Z - Scale , Rotation , Attibuts list (Name,Value)
            Dim BlockName, VarAction, NumID, VarColor, VarLayer, VarLink, PositionX, PositionY, PositionZ, VarScale, Rotation As String
            Dim AttrName, AttrValue As New List(Of String)
            BlockName = "" : VarAction = "" : NumID = "" : VarColor = "" : VarLayer = "" : VarLink = ""
            PositionX = "" : PositionY = "" : PositionZ = "" : VarScale = "" : Rotation = ""

            Dim Xroot As String = "/data-format/format[" & NumFormat & "]" ' 1er bloc = 1

            Dim NodeTypes As System.Xml.XmlNodeList
            Dim TypeName As String = ""
            Dim NodeDatas As System.Xml.XmlNodeList
            Dim NodeTopics As System.Xml.XmlNodeList
            Dim SplitCar As String = ""
            Try 'Charge le car Split (, ou . etc)
                SplitCar = x.SelectSingleNode(Xroot & "/split").InnerText()
            Catch
                MsgBox("Erreur de définition du caractère de Split")
            End Try


            'Boucle sur le bloc si Split activé
            Dim LineRef As String = Line
            Dim SplitLine() As String = Split(LineRef, SplitCar)
            For Nbloc As Double = 0 To SplitLine.Length - 1
                Line = SplitLine(Nbloc)

                'Boucle dans les TOPIC
                Dim NumTopic As Double = 0
                NodeTopics = x.SelectNodes(Xroot & "/topic")
                For Each NodeTopic As System.Xml.XmlNode In NodeTopics
                    NumTopic += 1 '(1er bloc = 1)

                    Dim ValideTopic As Boolean = False
                    Dim TopicValue As String = x.SelectNodes(Xroot & "/topic[" & NumTopic & "]").ItemOf(0).Attributes.ItemOf(0).InnerText

                    'Test les /topic
                    If TopicValue = "" Then 'Sans code Topic
                        ValideTopic = True
                    Else ' Avec code Topic
                        Dim TopicStart As String = ""
                        Dim TopicLen As String = ""
                        Try
                            TopicStart = NodeTopic.SelectSingleNode(Xroot & "/topic[" & NumTopic & "]/start").InnerText
                            TopicLen = NodeTopic.SelectSingleNode(Xroot & "/topic[" & NumTopic & "]/len").InnerText
                        Catch
                            ValideTopic = False
                        End Try

                        If TopicStart <> "" And IsNumeric(TopicStart) And TopicLen <> "" And IsNumeric(TopicLen) Then
                            If Len(TopicValue) >= CDbl(TopicLen) Then
                                If Mid(Line.ToLower, CDbl(TopicStart), CDbl(TopicLen)) = TopicValue.ToLower Then ValideTopic = True ' Si code est égale à Topic <Value>
                            End If
                        End If
                    End If

                    If ValideTopic = True Then

                        'Boucle dans les DATA
                        Dim NumDatas As Double = 0
                        NodeDatas = x.SelectNodes(Xroot & "/topic[" & NumTopic & "]/data")
                        For Each NodeData As System.Xml.XmlNode In NodeDatas
                            NumDatas += 1 '(1er bloc = 1)
                            Dim ValideData As Boolean = False
                            Dim DataValue As String = x.SelectNodes(Xroot & "/topic[" & NumTopic & "]/data[" & NumDatas & "]").ItemOf(0).Attributes.ItemOf(0).InnerText

                            'Test les /topic/data
                            If DataValue = "" Then 'Sans code Data
                                ValideData = True
                            Else ' Avec code Data
                                Dim DataStart As String = ""
                                Dim DataLen As String = ""
                                Try
                                    DataStart = NodeData.SelectSingleNode(Xroot & "/topic[" & NumTopic & "]/data[" & NumDatas & "]/start").InnerText
                                    DataLen = NodeData.SelectSingleNode(Xroot & "/topic[" & NumTopic & "]/data[" & NumDatas & "]/len").InnerText
                                Catch
                                    ValideData = False
                                End Try

                                If DataStart <> "" And IsNumeric(DataStart) And DataLen <> "" And IsNumeric(DataLen) Then
                                    If Len(DataValue) >= CDbl(DataLen) Then
                                        If Mid(Line.ToLower, CDbl(DataStart), CDbl(DataLen)) = DataValue.ToLower Then ValideData = True ' Si code est égale à Data <Value>
                                    End If
                                End If
                            End If

                            If ValideData = True Then

                                'Boucle dans les Type
                                Dim nIDtype As Double = 0
                                NodeTypes = x.SelectNodes(Xroot & "/topic[" & NumTopic & "]/data[" & NumDatas & "]" & "/type")
                                For Each NodeType As System.Xml.XmlNode In NodeTypes
                                    nIDtype += 1
                                    TypeName = NodeType.Attributes.ItemOf(0).InnerText()

                                    'Type paramètre
                                    Dim PropertyPts, xStart, xLen, xRound, xAlign, IDtype, PtsVisible, PtsFormat, PtsScale As String
                                    PropertyPts = "" : xStart = "" : xLen = "" : xRound = "" : xAlign = "" : IDtype = "" : PtsVisible = "" : PtsFormat = "" : PtsScale = ""

                                    Dim PtsReplaceVal As New List(Of String)
                                    Dim PtsReplace As New List(Of String)
                                    Dim ListLayerVal As New List(Of String)
                                    Dim ListLayer As New List(Of String)
                                    Dim NameSymbol As String = ""

                                    'Boucle dans les paramètres d'un type
                                    For Each Prop As System.Xml.XmlNode In NodeType
                                        Dim Var As String = Prop.InnerText
                                        Dim PropName As String = Prop.Name.ToLower
                                        'Chargement des paramètres
                                        If PropName = "start" Then
                                            xStart = Trim(Var)
                                        ElseIf PropName = "len" Then
                                            xLen = Trim(Var)
                                        ElseIf PropName = "round" Then
                                            xRound = Trim(Var)
                                        ElseIf PropName = "align" Then
                                            xAlign = Trim(Var.ToLower)
                                        ElseIf PropName = "property" Then
                                            PropertyPts = Trim(Var.ToLower)
                                        ElseIf PropName = "id" Then
                                            IDtype = Trim(Var.ToLower)
                                        ElseIf PropName = "visible" Then
                                            PtsVisible = Trim(Var.ToLower)
                                        ElseIf PropName = "replace" Then
                                            PtsReplaceVal.Add(Prop.Attributes.ItemOf(0).InnerText)
                                            PtsReplace.Add(Var)
                                        ElseIf PropName = "layer" Then
                                            ListLayerVal.Add(Prop.Attributes.ItemOf(0).InnerText) ' MO_BF_PTS
                                            ListLayer.Add(Var) ' 06
                                        ElseIf PropName = "format" Then
                                            PtsFormat = Var ' VD????00????
                                        ElseIf PropName = "scale" Then
                                            PtsScale = Var ' / 1000
                                        ElseIf PropName = "symbol" Then
                                            If Prop.Attributes.ItemOf(0).InnerText.ToLower = "default" Then
                                                NameSymbol = TypeName
                                                For Each SymbInfo As System.Xml.XmlNode In Prop
                                                    If SymbInfo.Name.ToLower = "block" Then BlockName = Replace(Replace(SymbInfo.InnerText, "REVOTHEME_", Theme & "_"), "MO_", "") ' AJOUT du THEME 6.10.2015 THA
                                                    If SymbInfo.Name.ToLower = "layer" Then VarLayer = Replace(SymbInfo.InnerText, "REVOTHEME_", Theme & "_") ' AJOUT du THEME 6.10.2015 THA
                                                    If SymbInfo.Name.ToLower = "color" Then VarColor = SymbInfo.InnerText
                                                    If SymbInfo.Name.ToLower = "rotation" Then Rotation = SymbInfo.InnerText
                                                Next
                                            End If
                                        End If
                                    Next ' fin de la boucle des propriétés dans un type


                                    'Découpage et formatage des données
                                    Dim ActiveFormat As Boolean = False
                                    Dim ValPts As String = ""
                                    If SplitCar = "" Then 'Si pas de Split
                                        'Découpage / split de la valeur du points
                                        If xStart <> "" And xLen <> "" Then
                                            If IsNumeric(xStart) And IsNumeric(xLen) Then
                                                ValPts = Mid(Line, CDbl(xStart), CDbl(xLen))
                                                ActiveFormat = True
                                            End If
                                        End If
                                    Else 'Si Split

                                        If xStart <> "" And xLen <> "" Then ' Si Split + Start +Len
                                            If IsNumeric(xStart) And IsNumeric(xLen) Then
                                                ValPts = Mid(Line, CDbl(xStart), CDbl(xLen))
                                                ActiveFormat = True
                                            End If
                                        Else 'Si Split SANS Start + Len 
                                            If Nbloc = nIDtype - 1 Then
                                                ValPts = SplitLine(Nbloc) 'Line
                                                ActiveFormat = True
                                            End If
                                        End If

                                    End If


                                    If ActiveFormat Then

                                        'Chargement du symbol
                                        If NameSymbol <> "" Then
                                            'docXML.SelectNodes(Xroot & "/topic/data/type/symbol/block")

                                            If Trim(ValPts) = "" Then ValPts = "default" 'Sans définition du bloc => envoi par défault

                                            Dim RootSymbol As String = Xroot & "/topic[" & NumTopic & "]/data[" & NumDatas & "]" & "/type[@name='" & NameSymbol & "']/symbol[@name='" & Trim(ValPts) & "']/"
                                            Dim sybBlockName, sybVarLayer, sybVarColor, sybRotation As String
                                            Try
                                                sybBlockName = x.SelectSingleNode(RootSymbol & "block").InnerText
                                                If sybBlockName <> "" Then BlockName = Replace(Replace(sybBlockName, "REVOTHEME_", Theme & "_"), "MO_", "") ' AJOUT du THEME 6.10.2015 THA
                                                sybVarLayer = x.SelectSingleNode(RootSymbol & "layer").InnerText
                                                If sybVarLayer <> "" Then VarLayer = Replace(sybVarLayer, "REVOTHEME_", Theme & "_") ' AJOUT du THEME 6.10.2015 THA
                                                sybVarColor = x.SelectSingleNode(RootSymbol & "color").InnerText
                                                If sybVarColor <> "" Then VarColor = sybVarColor
                                                sybRotation = x.SelectSingleNode(RootSymbol & "rotation").InnerText
                                                If sybRotation <> "" Then Rotation = sybRotation
                                            Catch
                                                'Erreur lecture XML" probablement impossibilté de trouvé une clé
                                            End Try
                                        End If

                                        If PtsFormat <> "" Then '0133   9  >>  VD????00????

                                            'Formatage Export (full format)
                                            If Left(PtsFormat, 1) = System.Convert.ToChar(34) And Right(PtsFormat, 1) = System.Convert.ToChar(34) Then
                                                'non utilisé à l'importation

                                            Else 'Formatage perso import/export
                                                Dim AttValueNew As String = ""
                                                Dim IncDecoupe As Double = 1
                                                ValPts = Replace(ValPts, " ", "0")
                                                For i = 1 To Len(PtsFormat)
                                                    If Mid(PtsFormat, i, 1) = "?" Then   'ajoute le caractère ?
                                                        If IncDecoupe <= Len(ValPts) Then
                                                            AttValueNew += Mid(ValPts, IncDecoupe, 1)
                                                            IncDecoupe += 1
                                                        End If
                                                    Else
                                                        AttValueNew += Mid(PtsFormat, i, 1)
                                                    End If
                                                Next
                                                ValPts = AttValueNew
                                            End If

                                        End If

                                        'Replace Data: Var => Value
                                        For i = 0 To PtsReplaceVal.Count - 1
                                            If PtsReplace(i).ToLower = "zero" Then
                                                Dim ZeroCount As Double = 0
                                                For u = 1 To Len(ValPts)
                                                    If Mid(ValPts, u, 1) = "0" Then   'Supprime les 0
                                                        ZeroCount += 1
                                                    Else
                                                        Exit For
                                                    End If
                                                Next
                                                ValPts = Mid(ValPts, ZeroCount + 1, Len(ValPts) - ZeroCount)
                                            Else
                                                ValPts = Replace(ValPts, PtsReplace(i), PtsReplaceVal(i))
                                            End If
                                        Next

                                        'Scale : import division(only num) => /1000
                                        If PtsScale <> "" And IsNumeric(PtsScale) And IsNumeric(ValPts) Then
                                            ValPts = CDbl(ValPts) / CDbl(PtsScale)
                                        End If

                                        'Select Layer
                                        For i = 0 To ListLayerVal.Count - 1
                                            If ValPts = ListLayer(i).ToUpper Then
                                                VarLayer = ListLayerVal(i)
                                                Exit For
                                            End If
                                        Next


                                        'Chargement des attributs
                                        AttrName.Add(TypeName)
                                        AttrValue.Add(Trim(ValPts)) ' ATTENTION Trim toute les valeurs

                                        'ID type
                                        If IDtype = "numeric" Then
                                            NumID = Trim(ValPts)
                                        End If

                                        'Test des Property
                                        If PropertyPts = "positionx" Then
                                            PositionX = Trim(ValPts)
                                        ElseIf PropertyPts = "positiony" Then
                                            PositionY = Trim(ValPts)
                                        ElseIf PropertyPts = "positionz" Then
                                            If Check3D Then
                                                PositionZ = (Trim(ValPts))
                                            Else
                                            End If

                                        ElseIf PropertyPts = "rotation" Then
                                            Rotation = Trim(ValPts)
                                        ElseIf PropertyPts = "color" Then
                                            VarColor = Trim(ValPts)
                                        ElseIf PropertyPts = "layer" Then
                                            VarLayer = ValPts
                                        ElseIf PropertyPts = "link" Then
                                            VarLink = ValPts
                                        ElseIf PropertyPts = "scale" Then
                                            VarScale = Trim(ValPts)
                                        End If

                                    End If

                                Next 'Fin de la boucle des types

                            End If 'Fin test si traitement Data

                        Next 'Fin de la boucle des Data

                    End If 'Fin test si traitement Topic

                Next 'Fin de la boucle des Topic

            Next 'Fin de la boucle des bloc d'une ligne


            'Création de l'objet PTS
            Dim Pts As New RevoPTS(BlockName, VarAction, NumID, VarColor, VarLayer, VarLink, PositionX, PositionY, PositionZ, VarScale, Rotation, AttrName, AttrValue)



            Return Pts
        End Function

        Public Function ImportXLSmatrix(ByVal Datafile As String, ByVal x As System.Xml.XmlDocument, ByVal NumFormat As Double) As Collection

            'Propriety >> Type objet : BlockName - Action - N°ID - Color - Layer - Link - X - Y - Z - Scale , Rotation , Attibuts list (Name,Value)
            Dim BlockName, VarAction, NumID, VarColor, VarLayer, VarLink, PositionX, PositionY, PositionZ, VarScale, Rotation As String
            Dim AttrName, AttrValue As New List(Of String)
            BlockName = "" : VarAction = "" : NumID = "" : VarColor = "" : VarLayer = "" : VarLink = ""
            PositionX = "" : PositionY = "" : PositionZ = "" : VarScale = "" : Rotation = ""
            Dim NumRow As Integer = 1
            Dim NumCol As Integer = 1
            Dim NameAtt As String = ""

            Dim PtsColl As New Collection

            Dim Xroot As String = "/data-format/format[" & NumFormat & "]"
            Dim NodeTypes As System.Xml.XmlNodeList
            Dim TypeName As String = ""

            If x.SelectNodes(Xroot & "/topic").ItemOf(0).Attributes.ItemOf(0).InnerText = "" Then 'Sans code Topic

                If x.SelectNodes(Xroot & "/topic/data").ItemOf(0).Attributes.ItemOf(0).InnerText = "" Then 'Sans code Data


                    'Boucle dans les Type
                    NodeTypes = x.SelectNodes(Xroot & "/topic/data/type")
                    For Each NodeType As System.Xml.XmlNode In NodeTypes
                        TypeName = NodeType.Attributes.ItemOf(0).InnerText()

                        'Type paramètre
                        Dim PropertyPts, xStart, xLen, xRound, xAlign, IDtype, PtsVisible As String
                        PropertyPts = "" : xStart = "" : xLen = "" : xRound = "" : xAlign = "" : IDtype = "" : PtsVisible = ""
                        Dim ActiveRow As Boolean = False
                        Dim ActiveCol As Boolean = False

                        'Boucle dans les paramètres d'un type
                        For Each Prop As System.Xml.XmlNode In NodeType
                            Dim Var As String = Prop.InnerText
                            Dim PropName As String = Prop.Name.ToLower
                            'Chargement des paramètres
                            If PropName = "start" Then
                                xStart = Trim(Var)
                            ElseIf PropName = "len" Then
                                xLen = Trim(Var)
                            ElseIf PropName = "round" Then
                                xRound = Trim(Var)
                            ElseIf PropName = "align" Then
                                xAlign = Trim(Var.ToLower)
                            ElseIf PropName = "property" Then
                                PropertyPts = Trim(Var.ToLower)
                            ElseIf PropName = "id" Then
                                IDtype = Trim(Var.ToLower)
                                If IDtype <> "" Then NameAtt = TypeName
                            ElseIf PropName = "visible" Then
                                PtsVisible = Trim(Var.ToLower)
                            ElseIf PropName = "symbol" Then
                                If Prop.Attributes.ItemOf(0).InnerText.ToLower = "default" Then
                                    For Each SymbInfo As System.Xml.XmlNode In Prop
                                        If SymbInfo.Name.ToLower = "block" Then BlockName = SymbInfo.InnerText
                                        If SymbInfo.Name.ToLower = "layer" Then VarLayer = SymbInfo.InnerText
                                        If SymbInfo.Name.ToLower = "color" Then VarColor = SymbInfo.InnerText
                                        If SymbInfo.Name.ToLower = "rotation" Then Rotation = SymbInfo.InnerText
                                    Next
                                End If

                                'UNIQUEMENT EXCEL --------
                            ElseIf PropName = "row" Then
                                If IsNumeric(Var) Then NumRow = CDbl(Var) : ActiveRow = True
                            ElseIf PropName = "col" Then
                                If IsNumeric(Var) Then NumCol = CDbl(Var) : ActiveCol = True
                                '------------------------------

                            End If
                        Next ' fin de la boucle des propriétés dans un type


                        'ID type
                        ' If IDtype = "numeric" Then
                        'NumID = Trim(ValPts)
                        'End If

                        'Test des Property
                        If PropertyPts = "positionx" Then
                            If ActiveRow = True Then
                                PositionX = "ROW"
                            ElseIf ActiveCol = True Then
                                PositionX = "COL"
                            End If
                        ElseIf PropertyPts = "positiony" Then
                            If ActiveRow = True Then
                                PositionY = "ROW"
                            ElseIf ActiveCol = True Then
                                PositionY = "COL"
                            End If
                        ElseIf PropertyPts = "positionz" Then
                            '  PositionZ = Trim(ValPts)
                        ElseIf PropertyPts = "rotation" Then
                            'Rotation = Trim(ValPts)
                        ElseIf PropertyPts = "color" Then
                            'VarColor = Trim(ValPts)
                        ElseIf PropertyPts = "layer" Then
                            'VarLayer = ValPts
                        ElseIf PropertyPts = "link" Then
                            'VarLink = ValPts
                        ElseIf PropertyPts = "scale" Then
                            'VarScale = Trim(ValPts)
                        End If

                    Next 'Fin de la boucle des types

                Else ' Avec code Data

                End If
            Else 'Avec un code Topic

            End If


            'Lecture dans le fichier EXCEL ------------------------------------

            Dim filePath As String = Datafile
            Dim connectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & filePath & ";Extended Properties=""Excel 8.0;HDR=No;"";"
            Dim connection As New System.Data.OleDb.OleDbConnection(connectionString)

            Dim Tableau As String = ""
            Dim NomFeuille As String = (x.SelectSingleNode(Xroot & "/sheet").InnerText).ToString ' x.getElementValue(Xroot & "/sheet")

            'Recherche de l onglet
            If IsNumeric(NomFeuille) Then
                connection.Open()
                Dim FeuillesExcel As System.Data.DataTable
                FeuillesExcel = connection.GetOleDbSchemaTable(System.Data.OleDb.OleDbSchemaGuid.Tables, New Object() {Nothing, Nothing, Nothing, "TABLE"})
                NomFeuille = FeuillesExcel.Rows(0).Item("TABLE_NAME").ToString()
                connection.Close()
            End If

            'Requête sur l'onglet
            Dim cmdText As String = "SELECT * FROM [" & NomFeuille & "]"
            Dim command As New System.Data.OleDb.OleDbCommand(cmdText, connection)


            'Lecture des Valeurs

            command.Connection.Open()
            Dim reader As System.Data.OleDb.OleDbDataReader = command.ExecuteReader()
            Dim CurRow As Double = 0
            Dim RowValue As String = ""
            Dim ColValue As New List(Of String)
            Dim CellValue As String = ""
            If reader.HasRows Then
                While reader.Read() ' CurRow = ROW
                    If CurRow >= NumRow - 1 Then 'Lecture depuis la 1er ligne définie

                        AttrName.Clear() : AttrValue.Clear() 'Vide tableau

                        For i As Integer = 0 To reader.FieldCount - 1  ' i = COL :Lecture depuis la 1er colonne définie

                            If CurRow = NumRow - 1 Then
                                'Load head row
                                ColValue.Add(reader(i).ToString())
                            Else

                                If i > NumCol - 1 And CurRow > NumRow - 1 Then 'Supprimer les entêtes
                                    Dim CoordX As String = ""
                                    Dim CoordY As String = ""

                                    If PositionX = "ROW" Or PositionY = "COL" Then
                                        'Row = X  /   Col = Y 
                                        CoordX = ColValue(i)
                                        CoordY = reader(NumCol - 1).ToString()
                                    Else 'Col = X  /   Row = Y 
                                        CoordX = reader(NumCol - 1).ToString()
                                        CoordY = ColValue(i)
                                    End If

                                    'Création de l'objet PTS
                                    If CoordX IsNot Nothing And CoordY IsNot Nothing Then
                                        NumID = Trim(reader(i).ToString()) ' ATTENTION Trim toute les valeurs
                                        AttrName.Add(NameAtt) 'Chargement des attributs
                                        AttrValue.Add(NumID)

                                        ' If CDbl(NumID) >= 1 Then 'Condition sur la valeurs
                                        Dim Pts As New RevoPTS(BlockName, "", NumID, "", "", "", CoordX, CoordY, 0, 1, 0, AttrName, AttrValue)
                                        PtsColl.Add(Pts)
                                        'End If

                                    End If

                                End If

                            End If

                        Next
                    End If


                    '                    Console.WriteLine("{0}" & vbTab & "{1}", reader(0).ToString(), reader(1).ToString())
                    Connect.Message("Importation de points", "Lecture des points en cours ... " & vbCrLf & "Nombre de point(s) : " & CurRow, False, 60, 100)
                    CurRow += 1
                End While
            End If

            '  MsgBox(Tableau)
            command.Connection.Close()


            Return PtsColl
        End Function
        Public Function WriteBlock(ByVal PtsColl As Collection)

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Dim Nbre As Double = 0
            doc.LockDocument()

            ' Start a transaction
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                ' Test if block exists in the block table
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockTable)
                Dim BlocName As String = ""

                For Each Pts As RevoPTS In PtsColl
                    BlocName = Pts.BlockName
                    Nbre += 1
                    'Test validité du points d'origine ' X + Y
                    If IsNumeric(Pts.PositionX) And IsNumeric(Pts.PositionY) Then

                        If Not bt.Has(BlocName) Then
                            cmdBL("#BL;>" & BlocName & ";") ' Importation de la définition d'un bloc
                        End If

                        If Not bt.Has(BlocName) Then
                            ed.WriteMessage(vbLf & "Block not found : " & BlocName)
                        Else

                            Dim id As ObjectId = bt(BlocName) 'Block Name = pr.stringresult

                            'insertion point
                            Dim Orig(0 To 2) As Double
                            Orig(0) = CDbl(Pts.PositionX)
                            Orig(1) = CDbl(Pts.PositionY)
                            If IsNumeric(Pts.PositionZ) Then
                                Orig(2) = CDbl(Pts.PositionZ)
                            Else
                                Orig(2) = 0 'Altitude 0
                            End If
                            Dim PtOrig As Point3d = New Point3d(Orig(0), Orig(1), Orig(2))

                            ' Get Model space
                            Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), BlockTableRecord)

                            ' Add the block reference to Model space
                            Dim oID As ObjectId = InsertBlockRef(PtOrig, btr, BlocName, Path.Combine(Ass.Library, BlocName & ".dwg"), db).ObjectId

                            'If the block was inserted successfully then call the function to add its attributes
                            If Not oID.IsNull Then
                                'Boucle dans les attributs
                                Dim blkRef As BlockReference = DirectCast(tr.GetObject(oID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockReference)

                                'Rotation
                                If Pts.Rotation <> "" And IsNumeric(Pts.Rotation) Then blkRef.Rotation = ConvGisTopoTrigo(CDbl(Pts.Rotation) * (Math.PI / 200)) 'br.Rotation
                                'Scale
                                If Pts.Scale <> "" And IsNumeric(Pts.Scale) Then blkRef.ScaleFactors = New Scale3d(CDbl(Pts.Scale)) 'br.Rotation
                                'Color
                                'If Pts.Color <> "" And IsNumeric(Pts.Color) Then blkRef.ColorIndex = CDbl(Pts.Color) 'test color
                                'Layer
                                If Pts.Layer <> "" Then
                                    Dim lt As LayerTable = DirectCast(tr.GetObject(db.LayerTableId, DatabaseServices.OpenMode.ForWrite), LayerTable)
                                    If lt.Has(Pts.Layer) Then 'Test si Layer existe
                                        blkRef.Layer = Pts.Layer
                                    Else
                                        Dim acLyrTblRec As LayerTableRecord = New LayerTableRecord()
                                        'acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 1) ' Assign the layer the ACI color 1 and a name
                                        acLyrTblRec.Name = Pts.Layer
                                        lt.UpgradeOpen() '' Upgrade the Layer table for write
                                        lt.Add(acLyrTblRec)  '' Append the new layer to the Layer table and the transaction
                                        tr.AddNewlyCreatedDBObject(acLyrTblRec, True)
                                        blkRef.Layer = Pts.Layer
                                    End If
                                End If

                                'Link
                                'If Pts.Link <> "" Then blkRef.Hyperlink = Pts.Link '


                                'Get the block attributue collection
                                Dim attCol As AttributeCollection = blkRef.AttributeCollection
                                'Dim attCol As AttributeCollection = br.AttributeCollection
                                Dim attRef As AttributeReference = Nothing
                                'Loop through the attribute collection
                                For Each attId As ObjectId In attCol
                                    'Get this attribute reference
                                    attRef = DirectCast(tr.GetObject(attId, DatabaseServices.OpenMode.ForWrite), AttributeReference)
                                    'Store its name
                                    Dim attName As String = attRef.Tag
                                    Try
                                        Dim IDAttrname As Integer = Array.IndexOf(Pts.AttrName.ToArray, attName)
                                        If IDAttrname >= 0 Then attRef.TextString = Pts.AttrValue(IDAttrname) 'dr(attName)
                                    Catch ex As Exception
                                    End Try
                                Next
                            End If
                        End If
                    End If

                    If Right(Nbre.ToString, 2) = "50" Or Right(Nbre.ToString, 2) = "00" Or Right(Nbre.ToString, 2) = "25" Or Nbre = PtsColl.Count Then
                        Connect.Message("Importation de points", "Ecriture des points en cours ... " & vbCrLf & "Identifiant de point(s) : " & Nbre & "/" & PtsColl.Count, False, 60, 100)
                    End If
                Next

                ' Commit the transaction
                tr.Commit()
            End Using

            doc.LockDocument.Dispose()

            Return True
        End Function




        Public Overridable Function OLD_InsertBlock(ByVal dblInsert As Point3d, ByVal btrSpace As BlockTableRecord, ByVal strSourceBlockName As String, ByVal strSourceBlockPath As String) As BlockReference

            Dim doc As Document = Application.DocumentManager.MdiActiveDocument

            Dim dlock As DocumentLock = Nothing
            Dim bt As BlockTable
            Dim btr As BlockTableRecord = Nothing
            Dim br As BlockReference
            Dim id As ObjectId
            Dim db As Autodesk.AutoCAD.DatabaseServices.Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor
                'insert block and rename it
                Try
                    Try
                        dlock = doc.LockDocument
                    Catch ex As Exception
                        Dim aex As New System.Exception("Error locking document for InsertBlock: " & strSourceBlockName & ": ", ex)
                        Throw aex
                    End Try
                    bt = trans.GetObject(db.BlockTableId, DatabaseServices.OpenMode.ForWrite)
                    If bt.Has(strSourceBlockName) Then
                        'block found, get instance for copying
                        btr = trans.GetObject(bt.Item(strSourceBlockName), DatabaseServices.OpenMode.ForRead)
                    Else
                        'block not found, insert into drawing
                        Using sourcedb As Database = New Database(False, False)
                            Try
                                sourcedb.ReadDwgFile(strSourceBlockPath, IO.FileShare.Read, True, "")
                                id = db.Insert(strSourceBlockPath, sourcedb, True)
                                btr = trans.GetObject(id, DatabaseServices.OpenMode.ForWrite)
                                btr.Name = strSourceBlockName
                            Catch ex As System.Exception
                                Dim aex As New System.Exception("Block file not found " & strSourceBlockPath & ": ", ex)
                                Throw aex
                                Exit Function
                            End Try
                            sourcedb.Dispose()
                        End Using
                        'CloneBlock(strSourceBlockPath, "*ModelSpace") ''clone block will not work when inserting the entire 'current drawing'
                    End If
                    'If bt.Has(strSourceBlockName) Then MsgBox("Got it: " & strSourceBlockName)
                    btrSpace = trans.GetObject(btrSpace.ObjectId, DatabaseServices.OpenMode.ForWrite)
                    'Set the Attribute Value
                    Dim attColl As AttributeCollection
                    Dim ent As Entity
                    Dim btrenum As BlockTableRecordEnumerator
                    br = New BlockReference(dblInsert, btr.ObjectId)
                    btrSpace.AppendEntity(br)
                    trans.AddNewlyCreatedDBObject(br, True)
                    attColl = br.AttributeCollection
                    btrenum = btr.GetEnumerator

                    While btrenum.MoveNext
                        ent = btrenum.Current.GetObject(DatabaseServices.OpenMode.ForWrite)
                        If TypeOf ent Is AttributeDefinition Then
                            Dim attdef As AttributeDefinition = ent
                            Dim attref As New AttributeReference
                            attref.SetAttributeFromBlock(attdef, br.BlockTransform)
                            attref.TextString = attref.Tag
                            attColl.AppendAttribute(attref)
                            trans.AddNewlyCreatedDBObject(attref, True)
                        End If
                    End While
                    trans.Commit()
                Catch ex As System.Exception
                    Dim aex2 As New System.Exception("Error in inserting new block: " & strSourceBlockName & ": ", ex)
                    Throw aex2
                Finally
                    If Not trans Is Nothing Then trans.Dispose()
                    If Not dlock Is Nothing Then dlock.Dispose()
                End Try
            End Using
            Return br
        End Function




        ''' <summary>
        ''' Revo Command: Exemple
        ''' </summary>
        ''' <param name="Cmd">Commande line</param>
        ''' <remarks></remarks>

        Public Function Exemple(ByVal Cmd As String)  ' Remarks
            'Cmd;
            '
            ' 


            Dim Cmdinfo() As String
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim curDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim curDwg As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Cmdinfo = SplitCmd(Cmd, 1)

            If Cmdinfo(1) <> "" Then 'si la ligne n'est pas vide
                'Execute command

            Else 'ignore the command
                Return Connect.DateLog & "Cmd Exemple" & vbTab & False & vbTab & "Erreur: " & vbTab & "NomFichier"
            End If

            Return Connect.DateLog & "Cmd Exemple" & vbTab & True & vbTab & "Type: " & Cmdinfo(1) & vbTab & "NomFichier"


        End Function

        Private Function BLname() As ObjectId
            Throw New NotImplementedException
        End Function


    End Class

    Public Class CollBLattr
        Private vName As String
        Private vSpace As String
        Private vAtt As String
        Private vData As String

        Public Sub New(ByVal Name As String, ByVal Space As String, ByVal Att As String, ByVal Data As String)
            vName = Name
            vSpace = Space
            vAtt = Att
            vData = Data
        End Sub

        Public Property Name() As String
            Get
                Return vName
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Space() As String
            Get
                Return vSpace
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Att() As String
            Get
                Return vAtt
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Data() As String
            Get
                Return vData
            End Get
            Set(ByVal value As String)
            End Set
        End Property
    End Class

    Public Class ScriptVar
        Private vName As String
        Private vData As List(Of String)
        Private vType As String
        Private vCollProp As Collection
        Private vCoordX As List(Of String)
        Private vCoordY As List(Of String)
        'Private vInit As Boolean

        Public Sub New(ByVal Name As String, ByVal Data As List(Of String), ByVal Type As String, ByVal CollProp As Collection, ByVal X As List(Of String), ByVal Y As List(Of String))
            vName = Name
            vData = Data
            vType = Type
            vCollProp = CollProp
            vCoordX = X
            vCoordY = Y
            '  vInit = init
        End Sub

        Public Property Name() As String
            Get
                Return vName
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Data() As List(Of String)
            Get
                Return vData
            End Get
            Set(ByVal value As List(Of String))
                vData = value
            End Set
        End Property
        Public Property Type() As String
            Get
                Return vType
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property CollProp() As Collection
            Get
                Return vCollProp
            End Get
            Set(ByVal value As Collection)
            End Set
        End Property
        Public Property CoordX() As List(Of String)
            Get
                Return vCoordX
            End Get
            Set(ByVal value As List(Of String))
            End Set
        End Property
        Public Property CoordY() As List(Of String)
            Get
                Return vCoordY
            End Get
            Set(ByVal value As List(Of String))
            End Set
        End Property
    End Class
    Public Class TableHTML
        Private xID As String
        Private xNumTable As Double
        Private xNumLigne As Double
        Private xCol1 As String
        Private xCol2 As String
        Private xCol3 As String
        Private xCol4 As String

        Public Sub New(ByVal ID As String, ByVal NumTable As Double, ByVal NumLigne As Double, ByVal Col1 As String, ByVal Col2 As String, ByVal Col3 As String, ByVal Col4 As String)
            xID = ID
            xNumTable = NumTable
            xNumLigne = NumLigne
            xCol1 = Col1
            xCol2 = Col2
            xCol3 = Col3
            xCol4 = Col4
        End Sub
        Public Property ID() As String
            Get
                Return xID
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property NumTable() As Double
            Get
                Return xNumTable
            End Get
            Set(ByVal value As Double)
            End Set
        End Property
        Public Property NumLigne() As Double
            Get
                Return xNumLigne
            End Get
            Set(ByVal value As Double)
            End Set
        End Property
        Public Property Col1() As String
            Get
                Return xCol1
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Col2() As String
            Get
                Return xCol2
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Col3() As String
            Get
                Return xCol3
            End Get
            Set(ByVal value As String)
            End Set
        End Property

        Public Property Col4() As String
            Get
                Return xCol4
            End Get
            Set(ByVal value As String)
            End Set
        End Property
    End Class


    ' Type objet : BlockName - Action - N°ID - Color - Layer - Link - X - Y - Z - Scale , Rotation , Attibuts list (Name,Value)
    Public Class RevoPTS
        Private xBlockName As String
        Private xAction As String
        Private xID As String
        Private xColor As String
        Private xLayer As String
        Private xLink As String
        Private xPositionX As String
        Private xPositionY As String
        Private xPositionZ As String
        Private xScale As String
        Private xRotation As String
        Private xAttrName As List(Of String)
        Private xAttrValue As List(Of String)


        Public Sub New(ByVal BlockName As String, ByVal Action As String, ByVal ID As String, ByVal Color As String, ByVal Layer As String, ByVal Link As String, ByVal PositionX As String, ByVal PositionY As String, ByVal PositionZ As String, ByVal Scale As String, ByVal Rotation As String, ByVal AttrName As List(Of String), ByVal AttrValue As List(Of String))
            xBlockName = BlockName
            xAction = Action
            xID = ID
            xColor = Color
            xLayer = Layer
            xLink = Link
            xPositionX = PositionX
            xPositionY = PositionY
            xPositionZ = PositionZ
            xScale = Scale
            xRotation = Rotation
            xAttrName = AttrName
            xAttrValue = AttrValue
        End Sub
        Public Property BlockName() As String
            Get
                Return xBlockName
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Action() As String
            Get
                Return xAction
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property ID() As String
            Get
                Return xID
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Color() As String
            Get
                Return xColor
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Layer() As String
            Get
                Return xLayer
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Link() As String
            Get
                Return xLink
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property PositionX() As String
            Get
                Return xPositionX
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property PositionY() As String
            Get
                Return xPositionY
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property PositionZ() As String
            Get
                Return xPositionZ
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Scale() As String
            Get
                Return xScale
            End Get
            Set(ByVal value As String)
            End Set
        End Property
        Public Property Rotation() As String
            Get
                Return xRotation
            End Get
            Set(ByVal value As String)
            End Set
        End Property

        Public Property AttrName() As List(Of String)
            Get
                Return xAttrName
            End Get
            Set(ByVal value As List(Of String))
            End Set
        End Property

        Public Property AttrValue() As List(Of String)
            Get
                Return xAttrValue
            End Get
            Set(ByVal value As List(Of String))
            End Set
        End Property
    End Class


    Public Class AcObjID
        Private xObjID As ObjectId
        Private xVal As String

        Public Sub New(ByVal objID As ObjectId, ByVal Val As String)
            xObjID = objID
            xVal = Val
        End Sub

        Public Property objID As ObjectId
            Get
                Return xObjID
            End Get
            Set(ByVal value As ObjectId)
            End Set
        End Property
        Public Property Value() As String
            Get
                Return xVal
            End Get
            Set(ByVal value As String)
            End Set
        End Property
    End Class

End Namespace
