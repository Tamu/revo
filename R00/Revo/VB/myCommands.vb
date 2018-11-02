Option Strict Off

Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Runtime.InteropServices
Imports System.Net
Imports Autodesk.AutoCAD.Windows

' This line is not mandatory, but improves loading performances
<Assembly: CommandClass(GetType(Revo.MyCommands))> 

Namespace Revo


    ' This class is instantiated by AutoCAD for each document when
    ' a command is called by the user the first time in the context
    ' of a given document. In other words, non static data in this class
    ' is implicitly per-document!
    Public Class MyCommands

        Public Shared m_ps1 As Autodesk.AutoCAD.Windows.PaletteSet = Nothing
        Public Shared m_ps2 As Autodesk.AutoCAD.Windows.PaletteSet = Nothing
        Public Shared m_ps3 As Autodesk.AutoCAD.Windows.PaletteSet = Nothing
        Public Shared PaletteTxT As New FrmPaletteMO(m_ps1)
        Public Shared PaletteGeoTools As New FrmGeoTools(m_ps3)

        '' Global variable for AddCOMEvent and RemoveCOMEvent commands
        ' Dim acAppCom As Autodesk.AutoCAD.Interop.AcadApplication 'AcadApplication
        Public Shared StatFrmRevoFinder As Boolean = True
        Public Shared ReloadMenu As Boolean = False
        ' Public Const REVOprofil As String = "REVO"
        Public Shared ActiveEventsCmd As Boolean = True

        Public Const NomCmd As String = "revo" 'Préfixe pour le lancement des commande (peux-etre diff de xproduit)
        Private TestVariableCmd As String = ""
        Private Small As Boolean = False


        <CommandMethod(NomCmd, CommandFlags.Modal + CommandFlags.Session)> _
        Public Sub RevoFinder() ' This method can have any name'  


            Dim Ass As New Revo.RevoInfo
            Dim ProjITF As Boolean = False
            StatFrmRevoFinder = False

            Dim Connect As New Revo.connect

            UpdateSharedSpace() 'Mise à jour de l'espace partagé

                Dim Fenetre As New frmRevoFinder


            Fenetre.Preflist = Connect.PrefUser()
                If Small Then 'Taille de la fenêtre mini
                    Fenetre.TableLayoutFileData.ColumnStyles.Item(0).Width = 0
                    Fenetre.Width = Fenetre.MinFormWidth
                    Fenetre.btnHide.Text = "<"
                Else 'Taille perso
                    If InStr(Fenetre.Preflist(5), "/") <> 0 Then
                        Dim Taille() As String
                        Taille = Split(Fenetre.Preflist(5), "/")
                        If Taille(0) <> "" And IsNumeric(Taille(0)) = True Then Fenetre.Height = Taille(0)
                        If Taille(1) <> "" And IsNumeric(Taille(1)) = True Then
                            Fenetre.Width = Taille(1)
                            If Taille(1) <= Fenetre.MinFormWidth Then Fenetre.btnHide.Text = "<" 'ResizeForm
                        End If

                    End If
                End If
                Fenetre.ShowDialog() ' Application.ShowModalDialog(  <--- code débile
                System.Windows.Forms.Application.DoEvents()
                Small = False


                If StatFrmRevoFinder = True Then 'Si fermeture correct


                    Try ' Suppression du XMLflow + création
                        System.IO.File.Delete(Ass.XMLflow)
                        If System.IO.File.Exists(Ass.XMLflow) = False Then
                            Dim Txts() As String = Split(Replace("<?xml version='1.0' encoding='ISO-8859-1'?>|<" & Ass.xProduct.ToLower & ">|  <dwg>|  </dwg>|  <flow>|  </flow>|  <import>|  </import>|  <fileaction>|  </fileaction>|</" & Ass.xProduct.ToLower & ">", "'", Chr(34)), "|")
                            Revo.RevoFiler.EcritureFichier(Ass.XMLflow, Txts, False)
                        End If
                    Catch
                    End Try

                    'Load listing of DWG or DXF
                    For i = 0 To Fenetre.GridData.RowCount - 1
                        conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower & "/dwg", "file", IO.Path.Combine(Fenetre.GridData.Rows(i).Cells(2).Value, Fenetre.GridData.Rows(i).Cells(1).Value), i + 1)
                        'DWGFiles.Add(IO.Path.Combine(Fenetre.GridData.Rows(i).Cells(2).Value, Fenetre.GridData.Rows(i).Cells(1).Value))
                    Next

                    'Load listing of imported file (itf, ptp, pts, ...)
                    For i = 0 To Fenetre.GridImport.RowCount - 1
                        conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower & "/import", "file", IO.Path.Combine(Fenetre.GridImport.Rows(i).Cells(2).Value, Fenetre.GridImport.Rows(i).Cells(1).Value) _
                                                         & "|" & Fenetre.GridImport.Rows(i).Cells(3).Value, i + 1)
                        'ImportedFiles.Add(IO.Path.Combine(Fenetre.GridImport.Rows(i).Cells(2).Value, Fenetre.GridImport.Rows(i).Cells(1).Value) _
                        '& "|" & Fenetre.GridImport.Rows(i).Cells(3).Value)
                        If Right(Fenetre.GridImport.Rows(i).Cells(1).Value, 4).ToUpper = ".ITF" Then ProjITF = True ' Start module interlis  (migrate_interlis.dvb)
                    Next

                    'Start of selected scripts
                    For i = 0 To Fenetre.GridCmd.RowCount - 1
                        If Fenetre.GridCmd.Rows(i).Selected = True Then
                            conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower & "/flow", "file", IO.Path.Combine(Fenetre.GridCmd.Rows(i).Cells(4).Value, Fenetre.GridCmd.Rows(i).Cells(3).Value), i + 1)
                            'ScriptFiles.Add(IO.Path.Combine(Fenetre.GridCmd.Rows(i).Cells(4).Value, Fenetre.GridCmd.Rows(i).Cells(3).Value))
                        End If
                    Next

                    If ProjITF = True Then conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower, "projinterlis", "True")
                    conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower, "startfilename", Application.DocumentManager.MdiActiveDocument.Name.ToString)

                    Fenetre.Close()
                    Fenetre.Dispose()

                    'Chargement des scripts pour chaque script
                    LoopDWG()

                Else

                    Fenetre.Close()
                    Fenetre.Dispose()
                End If



                If Revo.MyCommands.ReloadMenu = True Then
                Dim Ribbon As New Revo.AdskApplication()
                Ribbon.removeRibbon()
                Ribbon.createRibbon()
                Ribbon.createQuickStart()
            End If


        End Sub

        <CommandMethod("-" & NomCmd, CommandFlags.Modal + CommandFlags.Session)> _
        Public Sub RevoCmd() 'Commande Revo en ligne de commande

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim ProjITF As Boolean = False ' Projet interlis

            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim opts As New PromptStringOptions("Fichier script")
            Dim Pres As PromptResult '= ed.GetString(opts)
            Dim CmdsList As String = ""
            Dim SelFileString As String = ""
            Dim CmdFiles() As String
            opts.DefaultValue = "#"

            'The string result will contain the SelFiles string we passed when we called SendStringToExecute from RevoFinder
            opts.AllowSpaces = True
            Do
                Pres = ed.GetString(opts)     '--------->never stop there if showpalette is executed.
                If Left(Pres.StringResult, 1) = "#" And Pres.StringResult <> "#" And Pres.StringResult <> "" Then
                    CmdsList += (Pres.StringResult) & vbCrLf 'Stock command REVO
                Else
                    SelFileString = Pres.StringResult 'Stock fichier
                End If
                If Pres.StringResult = "" Or Pres.StringResult = "#" Then Exit Do
            Loop While Left(Pres.StringResult, 1) = "#"  ' pr.Status = PromptStatus.OK


            'Choix entre fichier script OU commande ********

            If CmdsList.Count = 0 Then 'Fichier
                'Replace the question marks in the string with spaces
                SelFileString = SelFileString.Replace("?", " ")
                CmdFiles = Split(SelFileString, "|")

            Else 'Commande > création des commandes
                Dim Cmds() As String = Split(CmdsList, vbCrLf)
                Dim CSVname As String = IO.Path.Combine(System.IO.Path.GetTempPath(), "REVO-live-cmds.csv")
                Revo.RevoFiler.EcritureFichier(CSVname, Cmds, False)
                CmdFiles = Split(CSVname, "|")

                'Test projet interlis validé par REVO
                If InStr(CmdsList.ToUpper, "#IMPORT;MD01MOVD;") <> 0 Then ProjITF = True
            End If


            Dim Connect As New Revo.connect
            Dim Ass As New Revo.RevoInfo
            StatFrmRevoFinder = False


            UpdateSharedSpace() 'Mise à jour de l'espace partagé

                'test des options de la licence. ' DESACTIVATION PROVISOIRE BUG : KB2962872
                '''Dim LicReg As String = Ass.VLIC
                '''If LicReg <> "" And Trim(LicReg).Length >= 5 Then
                StatFrmRevoFinder = True
                '''End If


                If StatFrmRevoFinder = True Then 'Si ok
                    Try ' Suppression du XMLflow + création
                        System.IO.File.Delete(Ass.XMLflow)
                        If System.IO.File.Exists(Ass.XMLflow) = False Then
                            Dim Txts() As String = Split(Replace("<?xml version='1.0' encoding='ISO-8859-1'?>|<" & Ass.xProduct.ToLower & ">|  <dwg>|  </dwg>|  <flow>|  </flow>|  <import>|  </import>|  <fileaction>|  </fileaction>|</" & Ass.xProduct.ToLower & ">", "'", Chr(34)), "|")
                            Revo.RevoFiler.EcritureFichier(Ass.XMLflow, Txts, False)
                        End If
                    Catch
                    End Try


                    'Load listing of DWG or DXF
                    ' --- For i = 0 To Fenetre.GridData.RowCount - 1
                    Dim curDwg As Document = Application.DocumentManager.MdiActiveDocument
                    conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower & "/dwg", "file", curDwg.Name, 1)
                    'DWGFiles.Add(IO.Path.Combine(Fenetre.GridData.Rows(i).Cells(2).Value, Fenetre.GridData.Rows(i).Cells(1).Value))
                    ' --- Next


                    'Start of selected scripts
                    For i = 0 To CmdFiles.Length - 1
                        conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower & "/flow", "file", CmdFiles(i), i + 1)
                    Next

                    If ProjITF = True Then conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower, "projinterlis", "True")
                    conn.XMLwriteX(Ass.XMLflow, "/" & Ass.xProduct.ToLower, "startfilename", Application.DocumentManager.MdiActiveDocument.Name.ToString)

                    'Chargement des scripts pour chaque script
                    LoopDWG()

                End If




        End Sub


        <CommandMethod(NomCmd & "LoopDWG", CommandFlags.Modal + CommandFlags.Session)> _
        Sub LoopDWG() ' Boucle sur chaque fichier

            Dim LoadOK As Boolean = True
            Dim Ass As New Revo.RevoInfo
            Dim docXML As New System.Xml.XmlDocument
            If IO.File.Exists(Ass.XMLflow) Then
                docXML.Load(Ass.XMLflow)
            Else
                Exit Sub
            End If


            Try
                Dim acDocMgrs As DocumentCollection = Application.DocumentManager
                Dim RVscript As New Revo.RevoScript
                Dim docXML2 As New System.Xml.XmlDocument
                docXML2.Load(Ass.XMLflow)
                Dim StartFileName As String = docXML2.SelectSingleNode(Ass.xProduct.ToLower & "/startfilename[1]").InnerText
                Dim FileStarting As Boolean = False
                If StartFileName.ToUpper = acDocMgrs.MdiActiveDocument.Name.ToUpper Then FileStarting = True 'Différent du fichier ouvert

                If docXML.SelectNodes(Ass.xProduct.ToLower & "/fileaction/action").Count <> 0 Then
                    For x = 1 To docXML.SelectNodes(Ass.xProduct.ToLower & "/fileaction/action").Count
                        Dim ActionStr As String = docXML2.SelectSingleNode(Ass.xProduct.ToLower & " /fileaction/action[" & x & "]").InnerText
                        conn.RevoLog(RVscript.cmdFile(ActionStr, acDocMgrs, FileStarting))
                    Next
                End If
            Catch
                MsgBox("Erreur de la commande : #File")
            End Try

            'Recherche si un fichier dans DWGfiles
            If docXML.SelectNodes(Ass.xProduct.ToLower & "/dwg/file").Count > 0 Then


                Dim ActiveDWG As Boolean = Revo.RevoFiler.ActiveDraw(docXML.SelectSingleNode(Ass.xProduct.ToLower & " /dwg/file[1]").InnerText)
                If ActiveDWG Then 'Active the DRAW

                    Dim Projili As Boolean = False
                    Try
                        If docXML.SelectSingleNode(Ass.xProduct.ToLower & " /projinterlis").InnerText = "True" Then Projili = True
                    Catch
                    End Try

                    'Test if need Interlis
                    If Projili = True Then 'Active import interlis
                        If InterlisProjectTest() Then
                            'c'est OK
                            LoadOK = True
                        Else 'Le dessin n'est pas compatible
                            If MsgBox("Pour profiter de toute les fonctions, le dessin actuel, doit être migré" & vbCrLf & _
                                   "Voulez-vous annuler l'opération en cours et executer la migration du dessin ?", MsgBoxStyle.YesNo + MsgBoxStyle.Information, _
                                   "Dessin non valide") = MsgBoxResult.Yes Then
                                LoadOK = False
                                Try
                                    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                                    Using docLock As DocumentLock = acDoc.LockDocument
                                        acDoc.SendStringToExecute(String.Format(NomCmd & "MigrationProject") & vbCr, True, False, False)
                                    End Using
                                Catch
                                    MsgBox("ERREUR EXE SCRIPT")
                                End Try
                            Else
                                LoadOK = True 'continue sans test
                            End If
                        End If
                    End If

                    If LoadOK Then
                        Try
                            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                            Using docLock As DocumentLock = acDoc.LockDocument
                                acDoc.SendStringToExecute(String.Format(NomCmd & "Start") & vbCr, True, False, False)
                            End Using
                        Catch
                            MsgBox("ERREUR EXE SCRIPT")
                        End Try


                    Else
                        If docXML.SelectNodes(Ass.xProduct.ToLower & "/dwg/file").Count > 0 Then conn.XMLdelete("/" & Ass.xProduct.ToLower & "/dwg", "file", 1, Ass.XMLflow) ' DWGFiles.RemoveAt(0) 'SI ERREUR
                    End If

                End If

            Else
                'Ne lance plus d'action
            End If

        End Sub

        <CommandMethod(NomCmd & "MigrationProject", CommandFlags.Modal + CommandFlags.Session)> _
        Public Sub RevoMigrationProject() ' Migration du projet

            Dim Ass As New Revo.RevoInfo
            Dim MigrYes As Boolean = True
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            'Test si modèles pas compatible
            If InterlisProjectTest() Then
                If MsgBox("Le dessin est déjà dans les normes du modèles" & vbCrLf & _
                         "Voulez-vous annuler la migration", vbYesNo + MsgBoxStyle.Information, "Migration validée") = MsgBoxResult.Yes Then MigrYes = False
            End If



            If MigrYes Then 'Migration du dessin

                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                Dim strDWGName As String = acDoc.Name
                Dim NameBack As String = IO.Path.GetFileName(acDoc.Name)
                If Right(NameBack.ToUpper, 4) = ".DWG" Then NameBack = Left(NameBack, Len(NameBack) - 4) & "_backup"
                If Right(strDWGName.ToUpper, 4) = ".DWG" Then strDWGName = Left(strDWGName, Len(strDWGName) - 4) & ".dwg"
                Dim strDWGNameBak As String = Replace(strDWGName, ".dwg", "_backup.dwg")


                'Test si fichier sauver
                If CDbl(Application.GetSystemVariable("DWGTITLED")) = 0 Or Right(strDWGName, 4) <> ".dwg" Then
                    MsgBox("Merci d'enregistrer le fichier au format DWG avant de faire la migration", vbOKOnly + MsgBoxStyle.Information, "Fichier non-enregistré")

                Else
                    Try

                        'Sauvegarde du fichiers actuel en backup
                        Revo.RevoFiler.SaveAsDraw(strDWGNameBak) ' Revo.RevoFiler.CloseAndSaveDraw(strDWGNameBak, NomCmd & "MigrationProject")

                        Dim docs As DocumentCollection = Application.DocumentManager
                        For Each doc As Document In docs
                            If strDWGNameBak.ToUpper = doc.Name.ToUpper Then
                                ' First cancel any running command

                                ' Activate the document, so we can check DBMOD
                                If docs.MdiActiveDocument <> doc Then
                                    docs.MdiActiveDocument = doc
                                End If

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                                doc.CloseAndDiscard()
#Else 'Versio AutoCAD 2013 et +
                                Autodesk.AutoCAD.ApplicationServices.DocumentExtension.CloseAndDiscard(docs.MdiActiveDocument)
#End If
                            End If
                        Next


                        'Nouveau fichier selon DWT + SaveAS
                        Revo.RevoFiler.NewDraw(Ass.Template)
                        Revo.RevoFiler.SaveAsDraw(strDWGName)

                        '    Renommer les présentations du model -> TRASHMIGRATION-A3_Paysage
                        Dim strDelete As String = ""

                        '' Get the current document and database, and start a transaction
                        Dim acDoc3 As Document = Application.DocumentManager.MdiActiveDocument
                        Dim acCurDb As Database = acDoc3.Database
                        Using docLock As DocumentLock = acDoc3.LockDocument
                            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                                '' Reference the Layout Manager
                                Dim myBT As BlockTable
                                Dim myBTR As BlockTableRecord
                                Dim myBTE As SymbolTableEnumerator
                                myBT = acCurDb.BlockTableId.GetObject(OpenMode.ForRead)
                                myBTE = myBT.GetEnumerator

                                While myBTE.MoveNext
                                    myBTR = myBTE.Current.GetObject(OpenMode.ForRead)
                                    If myBTR.IsLayout Then
                                        If Not myBTR.Name = "*Model_Space" Then
                                            Dim LayoutMgr As LayoutManager = LayoutManager.Current
                                            Dim layoutObject As Layout = acTrans.GetObject(myBTR.LayoutId, OpenMode.ForWrite)
                                            ' Get the layout's name
                                            Dim layoutName As String = layoutObject.LayoutName
                                            Dim newLayoutName As String
                                            newLayoutName = "TRASHMIGRATION-" & layoutName
                                            LayoutMgr.RenameLayout(layoutName, newLayoutName)
                                            strDelete += "_-Layout _D " & newLayoutName & vbCr
                                        End If
                                    End If
                                End While
                                acTrans.Commit()   '' Save the new objects to the database
                                acDoc3.Editor.Regen()
                            End Using
                        End Using


                        'Importer l'espace Objet et papier
                        Try
                            Dim acDocV As Document = Application.DocumentManager.MdiActiveDocument
                            Using docLock As DocumentLock = acDocV.LockDocument
                                InsertDWG(strDWGNameBak, "model", "")

                                acDocV.Editor.WriteMessage(vbLf & "Importation des présentations")

                                '  RVscript.cmdCmd("#Cmd;filedia|0|_-Layout|_T|" & DataFile & "|*|filedia|1")
                                acDocV.SendStringToExecute("filedia 0 _-Layout _T " & strDWGNameBak & vbCr & "*" & vbCr & "filedia 1" & vbCr & strDelete, True, False, False)
                                acDocV.Editor.Regen()

                                'Supprimer block
                                Dim acCurDbV As Database = acDocV.Database
                                Using acTrans As Transaction = acCurDbV.TransactionManager.StartTransaction()
                                    '' Returns the layer table for the current database
                                    Dim acBlTbl As BlockTable
                                    acBlTbl = acTrans.GetObject(acCurDbV.BlockTableId, OpenMode.ForRead)

                                    '' Check to see if MyLayer exists in the Layer table
                                    If acBlTbl.Has(NameBack) = True Then
                                        Dim acBlTblRec As BlockTableRecord
                                        acBlTblRec = acTrans.GetObject(acBlTbl(NameBack), OpenMode.ForWrite)

                                        Try
                                            acBlTblRec.Erase()
                                            acDocV.Editor.WriteMessage(vbLf & NameBack & " est supprimé")

                                            '' Commit the changes
                                            acTrans.Commit()
                                        Catch
                                            acDocV.Editor.WriteMessage(vbLf & NameBack & " ne peux être supprimé")
                                        End Try
                                    Else
                                        acDocV.Editor.WriteMessage(vbLf & NameBack & " n'exsite pas")
                                    End If
                                End Using


                            End Using
                        Catch
                            MsgBox("Erreur : Importation présentation")
                        End Try

                    Catch ex As System.Exception
                        MsgBox("Migration annulée : " & vbCrLf & ex.Message, MsgBoxStyle.Critical, "Erreur d'enregistrement")
                    End Try
                End If
            End If



        End Sub


        <CommandMethod(NomCmd & "Start")> _
        Sub StartRevo()  'Boot script Revo in only ONE draw at the same time

            Dim Connect As New Revo.connect

            If Application.DocumentManager.Count <> 0 Then

                Dim booSuccess As Boolean = True
                Dim Ass As New Revo.RevoInfo
                Dim FlowFiles As New List(Of String)
                Dim ImportedFiles As New List(Of String)
                Dim docXML As New System.Xml.XmlDocument
                If IO.File.Exists(Ass.XMLflow) Then
                    docXML.Load(Ass.XMLflow)
                Else
                    Exit Sub
                End If

                Try
                    Dim NodeFormats As System.Xml.XmlNodeList
                    NodeFormats = docXML.SelectNodes(Ass.xProduct.ToLower & "/flow/file")

                    For Each NodeFormat As System.Xml.XmlNode In NodeFormats
                        FlowFiles.Add(NodeFormat.InnerText)
                    Next

                    NodeFormats = docXML.SelectNodes(Ass.xProduct.ToLower & "/import/file")
                    For Each NodeFormat As System.Xml.XmlNode In NodeFormats
                        ImportedFiles.Add(NodeFormat.InnerText)
                    Next
                Catch
                End Try


                Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                Dim DocName As String = doc.Name
                Using docLock As DocumentLock = doc.LockDocument
                    Dim CmdRevo As New Revo.RevoScript
                    booSuccess = booSuccess And CmdRevo.StartScript(FlowFiles, ImportedFiles)
                End Using


                If booSuccess Then 'OK


                    If docXML.SelectNodes(Ass.xProduct.ToLower & "/dwg/file").Count > 0 Then

                        conn.XMLdelete("/" & Ass.xProduct.ToLower & "/dwg", "file", 1, Ass.XMLflow)
                    End If

                    '''Dim oLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                    '''CommandLine.Command(NomCmd & "LoopDWG")
                    '''oLock.Dispose()
                    Try
                        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                        acDoc.SendStringToExecute(String.Format(NomCmd & "LoopDWG") & vbCr, True, False, False)
                    Catch
                    End Try

                Else

                    'ERREUR
                    If docXML.SelectNodes(Ass.xProduct.ToLower & "/dwg/file").Count > 0 Then

                        conn.XMLdelete("/" & Ass.xProduct.ToLower & "/dwg", "file", 1, Ass.XMLflow) ' DWGFiles.RemoveAt(0)
                        'End If
                    End If

                End If

                'Ouverture de l'historique
                If connect.ActLog = True Then RevoFiler.OpenExe(Ass.LogPath)

            End If


        End Sub



        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "Doc", CommandFlags.Modal + CommandFlags.Session)>
        Public Sub RevoDoc()
            Dim Ass As New Revo.RevoInfo
            System.Diagnostics.Process.Start(Ass.urlDEV)
        End Sub

        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "Contact", CommandFlags.Modal + CommandFlags.Session)>
        Public Sub RevoForm()
            Dim Ass As New Revo.RevoInfo
            System.Diagnostics.Process.Start(Ass.urlLicRight)
        End Sub



        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "Horiz")>
        Public Sub RevoHoriz()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim pPtRes As PromptPointResult
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            '' Prompt for the start point
            pPtOpts.Message = vbLf & "Saisir le 1er point"
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            Dim ptPic1 As Point3d = pPtRes.Value
            ' Dim Pts_1 As Point2d = New Point2d(ptPic1.X, ptPic1.Y)

            pPtOpts.Message = vbLf & "Saisir le 2ème point"
            pPtOpts.BasePoint = ptPic1
            pPtOpts.UseBasePoint = True
            pPtOpts.UseDashedLine = True
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            Dim ptPic2 As Point3d = pPtRes.Value
            ' Dim Pts_2 As Point2d = New Point2d(ptPic2.X, ptPic2.Y)

            'Dim Dist2d As String = Math.Round(Dist(Pts_1, Pts_2), acDoc.Database.Luprec)

            Dim ed As Editor = acDoc.Editor
            'ed.WriteMessage(vbLf & "Distance 2D = " & Dist2d)

            UcsByLine(ptPic1, ptPic2)


            'acVportTblRec.UcsFollowMode = True

        End Sub

        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "Dist2D", CommandFlags.Modal + CommandFlags.Session)>
        Public Sub RevoDist2D()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim pPtRes As PromptPointResult
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            '' Prompt for the start point
            pPtOpts.Message = vbLf & "Saisir le 1er point"
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            Dim ptPic1 As Point3d = pPtRes.Value
            Dim Pts_1 As Point2d = New Point2d(ptPic1.X, ptPic1.Y)

            pPtOpts.Message = vbLf & "Saisir le 2ème point"
            pPtOpts.BasePoint = ptPic1
            pPtOpts.UseBasePoint = True
            pPtOpts.UseDashedLine = True
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            Dim ptPic2 As Point3d = pPtRes.Value
            Dim Pts_2 As Point2d = New Point2d(ptPic2.X, ptPic2.Y)

            Dim Dist2d As String = Math.Round(Dist(Pts_1, Pts_2), acDoc.Database.Luprec)

            Dim ed As Editor = acDoc.Editor
            ed.WriteMessage(vbLf & "Distance 2D = " & Dist2d)

        End Sub

        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "Arc2Lines")>
        Public Sub RevoArc2Lines()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


            ''Selection polyligns
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Dim collection_retour As New Collection

            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor

            Try

                Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                    Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                    Dim acBlkTblRec As BlockTableRecord
                    acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)


                    Dim acOPrompt As PromptSelectionOptions = New PromptSelectionOptions()

                    acOPrompt.MessageForAdding = "Sélectionner les poylignes à transformer)"
                    acOPrompt.SingleOnly = False

                    '' Request for objects to be selected in the drawing area
                    Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(acOPrompt)

                    '' If the prompt status is OK, objects were selected
                    If acSSPrompt.Status = PromptStatus.OK Then
                        Dim acSSet As SelectionSet = acSSPrompt.Value
                        '' Step through the objects in the selection set

                        Dim NbreArc As Double = 0

                        For Each acSSObj As SelectedObject In acSSet
                            '' Check to make sure a valid SelectedObject object was returned
                            If Not IsDBNull(acSSObj) Then


                                '' Open the selected object for write
                                Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                If Not IsDBNull(acEnt) Then
                                    If TypeName(acEnt) Like "Polyline*" Then

                                        Dim Poly As Polyline = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                                        Dim NewPoly As Polyline = New Polyline() '' Create a polyline with two segments
                                        Dim NbreIndex As Double = 0

                                        For nCnt As Double = 0 To Poly.NumberOfVertices - 1


                                            If Poly.GetSegmentType(nCnt) = 1 Then ' Courbe

                                                Try

                                                    'Coordonnée du centre
                                                    Dim PtCentre As Point2d = New Point2d(Poly.GetArcSegment2dAt(nCnt).Center.X, Poly.GetArcSegment2dAt(nCnt).Center.Y)
                                                    'Coordonnée du pt départ
                                                    Dim PtStart As Point2d = New Point2d(Poly.GetArcSegment2dAt(nCnt).StartPoint.X, Poly.GetArcSegment2dAt(nCnt).StartPoint.Y)
                                                    'Coordonnée du pt de fin
                                                    Dim PtEnd As Point2d = New Point2d(Poly.GetArcSegment2dAt(nCnt).EndPoint.X, Poly.GetArcSegment2dAt(nCnt).EndPoint.Y)

                                                    'Rayon
                                                    Dim Rayon As Double = Poly.GetArcSegment2dAt(nCnt).Radius

                                                    'Droite centre - départ
                                                    Dim LineStart As Line2d = New Line2d(PtCentre, PtStart)
                                                    Dim LineEnd As Line2d = New Line2d(PtCentre, PtEnd)

                                                    'Gis. pt de fin - départ = angle / 3

                                                    Dim AngStart As Double = Poly.GetArcSegment2dAt(nCnt).StartAngle
                                                    Dim AngEnd As Double = Poly.GetArcSegment2dAt(nCnt).EndAngle
                                                    'AngStard + ((AngStard - AngEnd) / 2)



                                                    Dim DivArc As Double = 5
                                                    Dim arc As CircularArc2d = Poly.GetArcSegment2dAt(nCnt)
                                                    Dim LongArc As Double = arc.GetLength(arc.GetParameterOf(arc.StartPoint), arc.GetParameterOf(arc.EndPoint))

                                                    'Tolérence longeur de l'arc : max 5 x 3m = 15m d'arc
                                                    If LongArc > DivArc * 3 Then
                                                        DivArc = Math.Round(LongArc / 3, 0)
                                                    End If

                                                    'Tolérance longeur du rayon ' min 1m
                                                    If Rayon < 1 Then
                                                        DivArc = 6
                                                    End If

                                                    Dim DiffAngle As Double = (AngEnd - AngStart) / DivArc
                                                    Dim AngleBase As Double = DiffAngle


                                                    '  If NbreIndex = 0 Then 'Début
                                                    NewPoly.AddVertexAt(NbreIndex, PtStart, 0, 0, 0)
                                                    NbreIndex += 1
                                                    ' End If

                                                    Dim SenseMontre As Boolean = True
                                                    If LineEnd.Direction.Angle - LineStart.Direction.Angle < 0 Then
                                                        SenseMontre = False
                                                    End If


                                                    For o = 2 To DivArc

                                                        'Test sense
                                                        Dim Pt1 As Point2d = PolarPoints2D(PtCentre, LineStart.Direction.Angle + AngleBase, Rayon)
                                                        Dim Pt2 As Point2d = PolarPoints2D(PtCentre, LineStart.Direction.Angle - AngleBase, Rayon)
                                                        Dim Dist1Test As Double = Dist(PtEnd, Pt1)
                                                        Dim Dist2Test As Double = Dist(PtEnd, Pt2)
                                                        If Dist1Test > Dist2Test Then SenseMontre = False
                                                        If Dist1Test < Dist2Test Then SenseMontre = True

                                                        Dim AngleCorr As Double = 0
                                                        If SenseMontre Then
                                                            AngleCorr = LineStart.Direction.Angle + AngleBase
                                                        Else
                                                            AngleCorr = LineStart.Direction.Angle - AngleBase
                                                        End If

                                                        Dim PtsNew As Point2d = PolarPoints2D(PtCentre, AngleCorr, Rayon)
                                                        NewPoly.AddVertexAt(NbreIndex, PtsNew, 0, 0, 0)
                                                        NbreIndex += 1
                                                        AngleBase = AngleBase + DiffAngle

                                                    Next


                                                    ' MsgBox(PtsNew.X & "-" & PtsNew.Y)

                                                    'Coordonnées du nouv. points

                                                    'x = x0 + r*cos(t)
                                                    'y = y0 + r * sin(t)


                                                    'Distance du points début <-> nouv

                                                    '  MsgBox(Poly.GetArcSegment2dAt(nCnt).Radius.ToString)
                                                    '   MsgBox(Poly.GetArcSegment2dAt(nCnt).Center.Y.ToString)

                                                    'Dim PtsCurve As PointOnCurve3d()
                                                    'PtsCurve = Poly.GetArcSegmentAt(2).GetNewSamplePoints(nCnt, nCnt + 1, 10)
                                                    'For num As Double = 0 To PtsCurve.Count - 1
                                                    '    '   MsgBox(PtsCurve(num).Point.Y & " - " & PtsCurve(num).Point.X)
                                                    'Next
                                                Catch
                                                    MsgBox("Erreur de courbe")
                                                End Try

                                                NbreArc += 1

                                                '  Poly.Color = Autodesk.AutoCAD.Colors.Color.FromColor(Drawing.Color.Violet)
                                                ' Exit For

                                            Else ' Ligne

                                                Dim pt As Point2d = New Point2d(Poly.GetPoint2dAt(nCnt).X, Poly.GetPoint2dAt(nCnt).Y)
                                                NewPoly.AddVertexAt(NbreIndex, pt, 0, 0, 0)
                                                NbreIndex += 1


                                            End If
                                        Next

                                        NewPoly.ColorIndex = 1
                                        NewPoly.Closed = Poly.Closed

                                        '' Add the new object to the block table record and the transaction
                                        acBlkTblRec.AppendEntity(NewPoly)
                                        acTrans.AddNewlyCreatedDBObject(NewPoly, True)



                                    End If
                                End If

                            End If
                        Next


                        '  console("Nombre d'arc : " & NbreArc)

                    Else


                    End If


                    acTrans.Commit()
                    acTrans.Dispose()

                    ''End of the selection of the polylign 
                End Using


            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try
        End Sub

        <CommandMethod(NomCmd & "PlanTopo", CommandFlags.Modal + CommandFlags.Session)>
        Public Sub RevoPlanTopo()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim FrmPT As New frmPlanTopo
            FrmPT.Show()

        End Sub

        <CommandMethod(NomCmd & "Quadrillage")>
        Public Sub RevoQuadrillage()

            ActiveEventsCmd = False 'Désactivation de mise à jour auto

            'Test Si PaperSpace
            If LayoutManager.Current.CurrentLayout = "Model" Then
                MsgBox("Cette fonction est disponible uniquement dans l'espace papier" & vbCrLf & vbCrLf &
                       "(lancer la commande dans la présentation souhaitée)", vbOKOnly + vbInformation, "REVO")

            Else 'OK Espace papier


                'Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                'Dim acCurDb As Database = acDoc.Database


                Dim FenScale As Double = 1
                Dim FenCenter As Point3d
                Dim FenCenterObj As Point3d
                Dim FenHeight As Double = 297
                Dim FenWidth As Double = 420
                Dim FenValide As Boolean = False
                Dim FactPS As Double = 1 '1
                Dim Update As Boolean = False
                Dim objBlockRef As Autodesk.AutoCAD.Interop.Common.AcadBlockReference
                Dim ViewPortObjectID As String = 0
                Dim QuadriColl As New Collection
                Dim ViewRotation As Double = 0

                If IsNumeric(GetScalePS) Then FactPS = GetScalePS()


                'Insertion du bloc
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
                Try

                    'Contrôle si Quadri est existant dans la fênetre

                    QuadriColl = CheckExistQuadri()

                    If QuadriColl.Count = 0 Then ' Création normal
                        Update = False
                    ElseIf QuadriColl.Count = 1 Then ' Mise à jour

                        Update = True
                    Else
                        MsgBox("Plusieurs réseaux de coordonnées ont été trouvés, conserver uniquement un réseau par présentation", vbExclamation + vbOKOnly, "Interruption du traitement")
                        Exit Sub
                    End If



                    ''Selection de la fenêtre automatique ---------ATTENTION 1 seul fenêtre sinon sélection ------------







                    'If Update = True Then 'Mise à jour : Recherche de la fenêtre

                    '    ' Dim objBlockRefY As Autodesk.AutoCAD.Interop.Common.AcadBlockReference
                    '    ' Dim objectHandle As String = QuadriColl.Item(1).Handle.ToString  'Mid(BLprop(0), 4, Len(BLprop(0)) - 3)
                    '    ' Dim BLKY As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = acDocX.HandleToObject(objectHandle)
                    '    ' objBlockRefY = BLKY

                    '    'Dim BlocID As Long
                    '    ' Dim varAttributes As Object
                    '    ' varAttributes = objBlockRefY.GetAttributes

                    '    'For j = LBound(varAttributes) To UBound(varAttributes)
                    '    '    If "ViewPortObjectID".ToUpper = varAttributes(j).TagString.ToString.ToUpper Then
                    '    '        ViewPortObjectID = varAttributes(j) ' Replace(Replace(varAttributes(j).TextString, "(", ""), ")", "")
                    '    '        BlocID = ViewPortObjectID '(CLng(ViewPortObjectID))
                    '    '        Exit For
                    '    '    End If
                    '    'Next

                    '    Dim ListV As List(Of String) = SelectViewPort("")

                    '    If ListV(0) = 1 Then

                    '        FenValide = True
                    '        FenScale = ListV(1)
                    '        Dim FenCenterL() As String = Split(ListV(2), ",")
                    '        FenCenter = New Point3d(FenCenterL(0), FenCenterL(1), FenCenterL(2))
                    '        Dim FenCenterObjL() As String = Split(ListV(3), ",")
                    '        FenCenterObj = New Point3d(FenCenterObjL(0), FenCenterObjL(1), FenCenterObjL(2))
                    '        FenHeight = ListV(4)
                    '        FenWidth = ListV(5)
                    '        ViewPortObjectID = ListV(6)

                    '    End If

                    'Else 'Sélection fenêtre

                    Dim ListV As List(Of String) = SelectViewPort()
                    If ListV(0) = 1 Then

                        FenValide = True
                        FenScale = ListV(1)
                        Dim FenCenterL() As String = Split(ListV(2), ",")
                        FenCenter = New Point3d(FenCenterL(0), FenCenterL(1), FenCenterL(2))
                        Dim FenCenterObjL() As String = Split(ListV(3), ",")
                        FenCenterObj = New Point3d(FenCenterObjL(0), FenCenterObjL(1), FenCenterObjL(2))
                        FenHeight = ListV(4)
                        FenWidth = ListV(5)
                        ViewPortObjectID = ""

                        'Rotation de la fenêtre depuis le centre
                        ViewRotation = 0 'CDbl(ListV(6))  ' << parametre pour la rotation !!!!!!!

                    End If

                    'End If

                Catch ex As System.Exception
                    MsgBox(ex.Message)
                End Try


                ' Validation de l'échelle -----------------------------------------------------------------
                Dim ScaleTest As Double = Math.Round(FenScale * FactPS, 7)
                If Math.Round(ScaleTest, 0) - ScaleTest = 0 Or ScaleTest = 0.5 Or ScaleTest = 0.4 _
                     Or ScaleTest = 0.2 Or ScaleTest = 0.1 Or ScaleTest = 0.04 Then
                Else 'SI échelle bizarre qui après question
                    '   If MsgBox("L'échelle ne semble pas être réglée, voulez-vous continuer ?" & vbCrLf & vbCrLf & _
                    '            "Attention : si vous annuler le quadriage ne sera pas à jour", vbOKCancel + vbQuestion, "Erreur d'échelle") = vbCancel Then Exit Sub
                End If



                ''Traitement du bloc :  de la fenêtre -----------------------------------------------------------------

                If FenValide Then


                    'Calcul des coordonnées  -----------------------------------------------------------------
                    Dim AttList As New List(Of String)
                    Dim ValueList As New List(Of String)
                    Dim DistQuadri As Double = 100 / FenScale / FactPS
                    Dim Ybase As Double = FenCenterObj.X - ((FenWidth / 2) / FenScale)
                    Dim Xbase As Double = FenCenterObj.Y - ((FenHeight / 2) / FenScale)
                    Dim YbaseRound As Double = 0
                    Dim XbaseRound As Double = 0


                    ' ARRONDI(C16 /5;1)*5
                    YbaseRound = Math.Round((Ybase / DistQuadri), 0) * DistQuadri
                    XbaseRound = Math.Round((Xbase / DistQuadri), 0) * DistQuadri

                    'Check si Y Round est trop petit (sort de la fênetre)
                    If YbaseRound < Ybase Then
                        YbaseRound += DistQuadri
                    End If

                    'Check si X Round est trop petit (sort de la fênetre)
                    If XbaseRound < Xbase Then
                        XbaseRound += DistQuadri
                    End If

                    Dim QuadriPosY As Double = ((YbaseRound - Ybase) * FenScale)  '(X math)
                    Dim QuadriPosX As Double = ((XbaseRound - Xbase) * FenScale)  '(Y math)

                    Dim NbreY As Double = Math.Round((FenWidth - QuadriPosY + (30 / FactPS)) / (100 / FactPS), 0)
                    Dim NbreX As Double = Math.Round((FenHeight - QuadriPosX + (30 / FactPS)) / (100 / FactPS), 0)



                    'Ajout des nom attribut
                    For i = 0 To 14

                        AttList.Add("COORDYT" & Format(i, "00"))
                        AttList.Add("COORDYB" & Format(i, "00"))
                        AttList.Add("COORDXL" & Format(i, "00"))
                        AttList.Add("COORDXR" & Format(i, "00"))

                    Next

                    For i = 0 To 14

                        ' Load Coord Y
                        If i < NbreY Then
                            ValueList.Add(Replace(Format(YbaseRound + (DistQuadri * i), "###`##0"), "`", "'"))
                            ValueList.Add(Replace(Format(YbaseRound + (DistQuadri * i), "###`##0"), "`", "'"))
                        Else
                            ValueList.Add("")
                            ValueList.Add("")
                        End If

                        ' Load Coord X
                        If i < NbreX Then
                            ValueList.Add(Replace(Format(XbaseRound + (DistQuadri * i), "###`##0"), "`", "'"))
                            ValueList.Add(Replace(Format(XbaseRound + (DistQuadri * i), "###`##0"), "`", "'"))
                        Else
                            ValueList.Add("")
                            ValueList.Add("")
                        End If

                    Next




                    'Insertion du bloc -----------------------------------------------------------------
                    Try

                        Dim dblInsertPt(2) As Double 'New Autodesk.AutoCAD.Geometry.Point3d(Val(Me.txtX.Text), Val(Me.txtY.Text), Val(Me.txtZ.Text))
                        dblInsertPt(0) = FenCenter.X - (FenWidth / 2) + QuadriPosY '- 30
                        dblInsertPt(1) = FenCenter.Y - (FenHeight / 2) + QuadriPosX '- 30
                        dblInsertPt(2) = 0

                        If Update = True Then 'Update bloc
                            Dim objectHandle As String = QuadriColl.Item(1).Handle.ToString  'Mid(BLprop(0), 4, Len(BLprop(0)) - 3)
                            Dim BLK As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = acDocX.HandleToObject(objectHandle)
                            objBlockRef = BLK
                            objBlockRef.InsertionPoint = dblInsertPt
                            objBlockRef.Rotation = ViewRotation
                        Else '                Create new bloc
                            objBlockRef = acDocX.PaperSpace.InsertBlock(dblInsertPt, "QUADRI_DYN", 1 / FactPS, 1 / FactPS, 1 / FactPS, ViewRotation)

                            Dim LayerCadre As String = "HAB_CADRE"
                            Dim RVscipt As New Revo.RevoScript
                            RVscipt.cmdLA("#LA;>" & LayerCadre)
                            objBlockRef.Layer = LayerCadre

                        End If



                        'Ecriture des attributs ---------------------
                        Dim varAttributes As Object
                        Dim strNomAttr As String = ""
                        varAttributes = objBlockRef.GetAttributes

                        For j = LBound(varAttributes) To UBound(varAttributes)
                            For i = 0 To AttList.Count - 1
                                If AttList(i).ToUpper = varAttributes(j).TagString.ToString.ToUpper Then
                                    varAttributes(j).TextString = ValueList(i)
                                ElseIf "ViewPortObjectID".ToUpper = varAttributes(j).TagString.ToString.ToUpper Then
                                    varAttributes(j).TextString = ViewPortObjectID
                                End If
                            Next
                        Next


                        ' Ecriture des propriété Dynamique ---------------------
                        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                        Dim db As Database = doc.Database
                        Dim hn As Handle = New Handle(Convert.ToInt64(objBlockRef.Handle.ToString, 16))
                        Dim id As ObjectId = db.GetObjectId(False, hn, 0)

                        Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
                            Dim BL As BlockReference = trans.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

                            Dim FenetreLargeur As Double = CDbl(FenWidth - QuadriPosY)
                            Dim FenetreHauteur As Double = CDbl(FenHeight - QuadriPosX)
                            Dim ZeroVal As Double = CDbl(0.00001)

                            If objBlockRef.IsDynamicBlock Then
                                For Each BLdyn As DynamicBlockReferenceProperty In BL.DynamicBlockReferencePropertyCollection

                                    '  Dim FenetreLargeurDecal As Double = CDbl(BLdyn.Value + QuadriPosY)



                                    If "DL" = Mid(BLdyn.PropertyName, 1, 2) Or "DR" = Mid(BLdyn.PropertyName, 1, 2) Or
                                       "DT" = Mid(BLdyn.PropertyName, 1, 2) Or "DD" = Mid(BLdyn.PropertyName, 1, 2) Then
                                        BLdyn.Value = RevoQuadrillageCheck(BLdyn.PropertyName, FenetreHauteur, FenetreLargeur, QuadriPosY, QuadriPosX)
                                    End If


                                    'If "QuadriVert".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    'BLdyn.Value = CDbl(FenHeight - QuadriPosX + (100 / FactPS))

                                    'ElseIf "QuadriHoriz".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    'BLdyn.Value = CDbl(FenWidth - QuadriPosY + (100 / FactPS))

                                    'If "QuadriPos X".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '    BLdyn.Value = CDbl(QuadriPosY) 'X Math

                                    'ElseIf "QuadriPos Y".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '    BLdyn.Value = CDbl(QuadriPosX) ' Y Math

                                    ' ElseIf "QuadriHeight".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '    BLdyn.Value = CDbl(FenHeight - QuadriPosX)

                                    'ElseIf "QuadriWidth".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '    BLdyn.Value = CDbl(FenWidth - QuadriPosY)

                                    'ElseIf "QuadriHeightMin".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '    BLdyn.Value = CDbl((30 / FactPS) + QuadriPosX)

                                    'ElseIf "QuadriWidthMin".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '    BLdyn.Value = CDbl((30 / FactPS) + QuadriPosY)

                                    'ElseIf "CoordY".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '    BLdyn.Value = CDbl(QuadriPosY)

                                    'ElseIf "CoordX".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '    BLdyn.Value = CDbl(QuadriPosX)

                                    'ElseIf "CoordYTop".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '   BLdyn.Value = CDbl(FenHeight - (40 / FactPS))

                                    'ElseIf "CoordXRight".ToUpper = BLdyn.PropertyName.ToUpper Then
                                    '  BLdyn.Value = CDbl(FenWidth - (88 / FactPS))

                                    'End If


                                    If "ViewGrip".ToUpper = BLdyn.PropertyName.ToUpper Then

                                        'Activation du nombre de param (Grid01 à Grid14)
                                        If NbreX > NbreY Then 'Choix du plus grand !

                                            BLdyn.Value = "Grid" & Format(NbreX, "00")

                                        Else ' Sinon valeur Y

                                            BLdyn.Value = "Grid" & Format(NbreY, "00")

                                        End If
                                    End If


                                Next

                            End If

                            trans.Commit()
                            trans.Dispose()

                        End Using



                    Catch ex As System.Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End If

            ActiveEventsCmd = True 'Activation de mise à jour auto


        End Sub
        Function RevoQuadrillageCheck(Label As String, FenetreHauteur As Double, FenetreLargeur As Double, QuadriPosY As Double, QuadriPosX As Double) As Double
            Dim BLValue As Double = 0.000001
            Dim MaxLimit As Double = 100

            Try
                MaxLimit = CDbl(Mid(Label, 3, 2)) * 100
            Catch

            End Try


            'Horizontal
            If Mid(Label, 1, 2) = "DL" Then '400+
                If FenetreHauteur > MaxLimit Then BLValue = FenetreLargeur
            ElseIf Mid(Label, 1, 2) = "DR" Then
                If FenetreHauteur > MaxLimit Then BLValue = FenetreLargeur + QuadriPosY

                'Vertical
            ElseIf Mid(Label, 1, 2) = "DT" Then '400+
                If FenetreLargeur > MaxLimit Then BLValue = FenetreHauteur
            ElseIf Mid(Label, 1, 2) = "DD" Then
                If FenetreLargeur > MaxLimit Then BLValue = FenetreHauteur + QuadriPosX
            End If

            Return BLValue

        End Function


        <CommandMethod(NomCmd & "RotationVD", CommandFlags.Modal + CommandFlags.Session)>
        Public Sub RevoRotationVD()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim fSelect As New frmState
            Dim Ass As New Revo.RevoInfo
            fSelect.Text = Ass.xTitle
            fSelect.lbl_infos.Text = "Angle de rotation [g] : (selon angle de base du dessin = gisement)"
            fSelect.ProgBar.Visible = False
            fSelect.BtnValid.Visible = True
            fSelect.BoxList.Visible = True
            fSelect.BoxList.Items.Add(GetCurrentRotation())

            If fSelect.BoxList.Items.Count <> 0 Then fSelect.BoxList.Text = fSelect.BoxList.Items.Item(0)
            fSelect.ShowDialog()
            Dim Angle As String = Trim(fSelect.BoxList.Text)
            fSelect.Hide()
            fSelect.Dispose()


            If Angle <> "" And IsNumeric(Angle) Then

                Dim Conn As New Revo.connect
                Conn.Message("Rotation MO", "Rotation des objets en cours", False, 2, 100)
                System.Windows.Forms.Application.DoEvents()


                'Sélection des blocs

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

                Dim objSelSet As Autodesk.AutoCAD.Interop.AcadSelectionSet
                Dim objBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference

                'Restrictions pour la sélection
                Dim adDXFCode(1) As Short
                Dim adDXFGroup(1) As Object


                'Chargement des calques -------------------------
                Dim RVinfo As New Revo.RevoInfo
                Dim MutationXML As String = RVinfo.ConfigMOVD '(Config-MOVD.xml) anciennement : rotation.dat

                Dim ListCalque As New List(Of String)
                Try
                    If IO.File.Exists(MutationXML) Then  ' <validate-mutation>  <exeption layer="MUT_ADHOC">DELETE
                        Dim configXML As New System.Xml.XmlDocument
                        configXML.Load(MutationXML)
                        Dim Xroot As String = "revo/rotation"
                        Dim NodeLayer As System.Xml.XmlNodeList
                        NodeLayer = configXML.SelectNodes(Xroot & "/layer")
                        For Each NodeExeption As System.Xml.XmlNode In NodeLayer
                            ListCalque.Add(NodeExeption.InnerText.ToUpper)
                        Next
                    End If
                Catch
                End Try


                'Boucle sur les calques  -------------------------
                Dim icount As Double = 0
                For Each strLA As String In ListCalque


                    objSelSet = acDoc.SelectionSets.Add("REVO" & acDoc.SelectionSets.Count) 'Nom de la sélection
                    adDXFCode(0) = 8
                    adDXFGroup(0) = strLA 'Calque
                    adDXFCode(1) = 0
                    adDXFGroup(1) = "INSERT" 'Type d'objet

                    'Effectue la sélection
                    objSelSet.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , adDXFCode, adDXFGroup)


                    'Boucle pour tous les éléments sélectionnés (blocs)

                    If objSelSet.Count <> 0 Then

                        'Déclarations
                        For Each objBlock In objSelSet

                            Dim varAttributes As Object
                            Dim dblRot0, dblDiffRot As Double

                            Select Case objBlock.Layer

                                'Attention: on ne modifie pas les numéros de DP (sens d'une route ou d'un ruisseau)
                                Case "MO_BF_IMMEUBLE"
                                    'Type de BF: pas de rotation si DP communal ou cantonal
                                    varAttributes = objBlock.GetAttributes
                                    For j = LBound(varAttributes) To UBound(varAttributes)
                                        If UCase(varAttributes(j).TagString) = "GENRE" Then
                                            Select Case varAttributes(j).TextString
                                                Case "bien fonds.DP cantonal", "bien fonds.DP communal"
                                                    'Rotation de 0 ou 180
                                                    GoTo Rotation180
                                                Case Else
                                                    'On modifie la rotation des objets
                                                    objBlock.Rotation = ConvGisTopoTrigo((Angle) * (Math.PI / 200))
                                            End Select
                                            Exit For
                                        End If
                                    Next


                                    'Numéro de bâtiment (ou ODS) et de police: rotation par 90 degrés
                                Case "MO_CS_BAT_NUM", "MO_ODS_BATS_NUM", "MO_ODS_COUV_NUM", "MO_AD_ID_BATIMENT"

                                    'Rotation actuelle (initiale) en g
                                    dblRot0 = ConvAngTrigoTopoG(objBlock.Rotation * 200 / Math.PI)

                                    'Différence nouvelle rotation - rotation actuelle
                                    If dblRot0 >= 350 Or dblRot0 < 50 Then
                                        dblDiffRot = Angle '- 0
                                    ElseIf dblRot0 >= 50 And dblRot0 < 150 Then
                                        dblDiffRot = Angle - 100
                                    ElseIf dblRot0 >= 150 And dblRot0 < 250 Then
                                        dblDiffRot = Angle - 200
                                    ElseIf dblRot0 >= 250 And dblRot0 < 350 Then
                                        dblDiffRot = Angle - 300
                                    End If
                                    dblDiffRot = Modulo400(dblDiffRot)

                                    'Rotation de 0-90-180-270 degrés
                                    If dblDiffRot >= 350 Or dblDiffRot < 50 Then
                                        objBlock.Rotation = ConvGisTopoTrigo(dblRot0 * (Math.PI / 200))
                                    ElseIf dblDiffRot >= 50 And dblDiffRot < 150 Then
                                        objBlock.Rotation = ConvGisTopoTrigo((dblRot0 + 100) * (Math.PI / 200))
                                    ElseIf dblDiffRot >= 150 And dblDiffRot < 250 Then
                                        objBlock.Rotation = ConvGisTopoTrigo((dblRot0 + 200) * (Math.PI / 200))
                                    ElseIf dblDiffRot >= 250 And dblDiffRot < 350 Then
                                        objBlock.Rotation = ConvGisTopoTrigo((dblRot0 + 300) * (Math.PI / 200))
                                    End If


                                    'Noms de CS ou OD ou BF orientés (DP): rotation 0 ou 180 degrés
                                Case "MO_CS_NOM", "MO_CS_PROJ_NOM", "MO_ODL_NOM"
Rotation180:

                                    'Rotation actuelle (initiale) en g
                                    dblRot0 = ConvAngTrigoTopoG(objBlock.Rotation * 200 / Math.PI)

                                    'Différence nouvelle rotation - rotation actuelle
                                    If dblRot0 >= 350 Or dblRot0 < 50 Then
                                        dblDiffRot = Angle '- 0
                                    ElseIf dblRot0 >= 50 And dblRot0 < 150 Then
                                        dblDiffRot = Angle - 100
                                    ElseIf dblRot0 >= 150 And dblRot0 < 250 Then
                                        dblDiffRot = Angle - 200
                                    ElseIf dblRot0 >= 250 And dblRot0 < 350 Then
                                        dblDiffRot = Angle - 300
                                    End If
                                    dblDiffRot = Modulo400(dblDiffRot)

                                    'Rotation de 0-180 degrés
                                    If dblDiffRot >= 350 Or dblDiffRot < 150 Then
                                        objBlock.Rotation = ConvGisTopoTrigo(dblRot0 * (Math.PI / 200))
                                    ElseIf dblDiffRot >= 150 And dblDiffRot < 350 Then
                                        objBlock.Rotation = ConvGisTopoTrigo((dblRot0 + 200) * (Math.PI / 200))
                                    End If


                                Case Else
                                    'On modifie la rotation des objets
                                    objBlock.Rotation = ConvGisTopoTrigo((Angle) * (Math.PI / 200))

                            End Select


                        Next objBlock

                    End If

                    'Effacement de la sélection
                    objSelSet.Delete()
                    icount += 1
                    Conn.Message("Rotation MO", "Rotation des objets en cours ...", False, icount, ListCalque.Count)
                Next

                Conn.Message("Rotation MO", "Rotation des objets terminés", False, 100, 100, "info")
                Conn.Message("Rotation MO", "Rotation des objets en cours.", True, 100, 100)

                SetFileProperty("Rotation", Format(Val(Angle), "0.0"))

            Else
                MsgBox("La valeur n'est pas numérique ou vide", vbOKOnly + vbInformation, "Erreur de saisie")

            End If



        End Sub
        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "ProjMOVD", CommandFlags.Modal + CommandFlags.Session)> _
        Public Sub RevoProjiterlisVD() ' Demande de ProjMOVD
            Try
                UpdateSharedSpace() 'Mise à jour de l'espace partagé
                Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

                Dim Ass As New Revo.RevoInfo
                Dim connect As New connect
                Dim TemplatePath As String = Ass.Template
                If System.IO.File.Exists(TemplatePath) Then
                    Dim docMan As DocumentCollection = Application.DocumentManager

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                    Dim docDWG As Document = docMan.Add(TemplatePath)
#Else 'Versio AutoCAD 2013 et +
                    Dim docDWG As Document = Autodesk.AutoCAD.ApplicationServices.DocumentCollectionExtension.Add(docMan, TemplatePath)
#End If
                Else
                    connect.Message("Problème de création du projet Interlis", "Le fichier template est manquant: " & vbCrLf & TemplatePath, False, 99, 100, "critical")
                    connect.RevoLog(connect.DateLog & "Cmd Import" & vbTab & False & vbTab & "Le fichier template est manquant: " & vbTab & TemplatePath)
                End If
            Catch

            End Try

        End Sub

        <CommandMethod(NomCmd & "PropVD")> _
        Public Sub RevoPropVD()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Connect As New Revo.connect

            UpdateSharedSpace() 'Mise à jour de l'espace partagé

                PropRF("VD")


        End Sub


        <CommandMethod(NomCmd & "MOradier")> _
        Public Sub RevoMOradier() ' Migr fonction MOradier

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Connect As New Revo.connect

            UpdateSharedSpace() 'Mise à jour de l'espace partagé

                RadierElements()


        End Sub
        <CommandMethod(NomCmd & "MOvaliderMut")> _
        Public Sub RevoMOvaliderMut() ' Migr fonction MOvaliderMut

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Connect As New Revo.connect

            UpdateSharedSpace() 'Mise à jour de l'espace partagé

            ValiderMutations() 'Module


        End Sub
        <CommandMethod(NomCmd & "MOfiabPlan")> _
        Public Sub RevoMOfiabPlan() ' Migr fonction MOfiabPlan

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Connect As New Revo.connect

            UpdateSharedSpace() 'Mise à jour de l'espace partagé

                ''Module
                Dim PointsFiables As New frmPointsFiables
                Application.ShowModalDialog(PointsFiables)
                PointsFiables.Close()
                PointsFiables.Dispose()

        End Sub

        <CommandMethod(NomCmd & "OrdreTraceVD")> _
        Public Sub RevoOrdreTraceVD() ' Migr fonction MOfiabPlan

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim Connect As New Revo.connect

            'Message
            Connect.Message("Ordre de tracé", "Réorganisation des calques pour l'impression du plan cadastral, veuillez patienter...", False, 5, 100)

            'Active l'espace objet si nécessaire ???
            PlaceObjectsOnBottom("HATCH") 'Envoie les hachures à l'arrière-plan
            ManageTerritoryBoundaries() 'Réorganises les limites politiques et leurs caches
            PlaceObjectsOnTop("INSERT") 'Envoie les blocs au premier plan

            'Regnérer le dessins
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            doc.Editor.Regen()

            Connect.Message("Ordre de tracé", " Le dessin est prêt pour l'impression !", False, 100, 100, "info")
            Connect.Message("Fin", "... ", True, 0, 0, "hide") 'Fin


        End Sub

        <CommandMethod(NomCmd & "BatXdata")> _
        Public Sub RevoBatXdata()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            UpdateXdataBatiment(Nothing)

        End Sub

        <CommandMethod(NomCmd & "CheckMO")> _
        Public Sub RevoCheckMO()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            CheckMO()

        End Sub

        <CommandMethod(NomCmd & "GeoTools")> _
        Public Sub RevoGeoTools()

            '  Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


            If m_ps3 Is Nothing Then
                'no so create it
                m_ps3 = New Autodesk.AutoCAD.Windows.PaletteSet("GeoTools", New Guid("{ECBFEC80-2FE4-2aa3-9E4B-3068E24A8BFA}"))
                m_ps3.Icon = My.Resources.plug
                m_ps3.Style = Autodesk.AutoCAD.Windows.PaletteSetStyles.NameEditable Or Autodesk.AutoCAD.Windows.PaletteSetStyles.ShowPropertiesMenu Or Autodesk.AutoCAD.Windows.PaletteSetStyles.ShowAutoHideButton Or Autodesk.AutoCAD.Windows.PaletteSetStyles.ShowCloseButton
                m_ps3.TitleBarLocation = PaletteSetTitleBarLocation.Right
                m_ps3.Name = "GeoTools"
                m_ps3.Opacity = 80
                m_ps3.Dock = DockSides.Left

                m_ps3.Add("GeoTools", PaletteGeoTools)

                PaletteGeoTools.LoadVariable()

            End If

            'turn it on
            m_ps3.Visible = True

        End Sub

        <CommandMethod(NomCmd & "EDT")> _
        Public Sub RevoEDT()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Connect As New Revo.connect

            UpdateSharedSpace() 'Mise à jour de l'espace partagé

                Dim EDT As New frmEDT
                Application.ShowModalDialog(EDT)
                EDT.Close()
                EDT.Dispose()


        End Sub


        <CommandMethod(NomCmd & "EDT2")> _
        Public Sub RevoEDT2()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Connect As New Revo.connect

            UpdateSharedSpace() 'Mise à jour de l'espace partagé
                EDTStart()



        End Sub


        <CommandMethod(NomCmd & "PER")> _
        Public Sub RevoPER()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim Connect As New Revo.connect


            UpdateSharedSpace() 'Mise à jour de l'espace partagé

                Dim PER As New frmPER
                Application.ShowModalDialog(PER)
                PER.Close()
                PER.Dispose()



        End Sub

        <CommandMethod(NomCmd & "PER2")> _
        Public Sub RevoPER2()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Connect As New Revo.connect


            UpdateSharedSpace() 'Mise à jour de l'espace partagé
                PERStart()


        End Sub

        <CommandMethod(NomCmd & "InsertionPoints")> _
        Public Sub InsertionPoints()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Connect As New Revo.connect


            UpdateSharedSpace() 'Mise à jour de l'espace partagé

                Dim FrmPts As New frmInsertionPts
                Application.ShowModalDialog(FrmPts)
                FrmPts.Close()
                FrmPts.Dispose()



        End Sub

        <CommandMethod(NomCmd & "InsertionTextes")> _
        Public Sub InsertionTextes()

            RevoPalette()


        End Sub
        Public Function RevoPalette(Optional ObjID As ObjectId = Nothing, Optional EditMod As Boolean = False)

            Dim RVinfo As New Revo.RevoInfo

            If m_ps1 Is Nothing Then
                'no so create it
                m_ps1 = New Autodesk.AutoCAD.Windows.PaletteSet("REVO", New Guid("{ECBFEC73-9FE4-4aa2-8E4B-3068E94A2BFA}"))
                m_ps1.Icon = RVinfo.Icon
                m_ps1.Style = Autodesk.AutoCAD.Windows.PaletteSetStyles.NameEditable Or Autodesk.AutoCAD.Windows.PaletteSetStyles.ShowPropertiesMenu Or Autodesk.AutoCAD.Windows.PaletteSetStyles.ShowAutoHideButton Or Autodesk.AutoCAD.Windows.PaletteSetStyles.ShowCloseButton
                m_ps1.TitleBarLocation = PaletteSetTitleBarLocation.Right
                m_ps1.Name = "REVO"
                m_ps1.Opacity = 80
                m_ps1.Dock = DockSides.Left


                m_ps1.Add("Insertion de texte ou symbole", PaletteTxT)

                PaletteTxT.txtOri.Text = Format(Val(GetCurrentRotation()), "0.0")

            Else
                PaletteTxT.cmdNext.Text = "Valider"
                PaletteTxT.txtX.Text = ""
                PaletteTxT.txtY.Text = ""
                PaletteTxT.txtBlockID.Text = ""
                PaletteTxT.txtOri.Text = Format(Val(GetCurrentRotation()), "0.0")
            End If

            'turn it on
            m_ps1.Visible = True

            Return True
        End Function


        ' Modal Command with pickfirst selection
        <CommandMethod(NomCmd & "Options", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        Public Sub RevoOptions() ' This method can have any name
            Dim Fenetre As New frmOptions
            Fenetre.ShowDialog()
        End Sub
        ' Modal Command with pickfirst selection
        <CommandMethod(NomCmd & "UpdateAutoCADprofile", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        Public Sub RevoUpdateProfile() ' This method can have any name

            '<SyncProfilAcad>1</SyncProfilAcad>
            Dim Conn As New Revo.connect
            Conn.RevoUpdateProfile()

        End Sub

        ' Modal Command with pickfirst selection
        <CommandMethod(NomCmd & "small", CommandFlags.Modal)> _
        Public Sub RevoSmall() ' This method can have any name
            Small = True
            RevoFinder()
        End Sub



        <CommandMethod(NomCmd & "Menu")> _
        Public Sub RevoMenu() ' This method can have any name
            Dim Ribbon As New Revo.AdskApplication()
            Ribbon.removeRibbon()
            Ribbon.createRibbon()
            Ribbon.createQuickStart()

        End Sub

        <CommandMethod(NomCmd & "CurrentFolder")> _
        Public Sub RevoCurFolder() ' This method can have any name

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            'RevoFiler.OpenExe(IO.Path.GetDirectoryName(acDoc.Name))
            If System.IO.File.Exists(acDoc.Name) Then
                System.Diagnostics.Process.Start(IO.Path.GetDirectoryName(acDoc.Name))
            Else
                MsgBox("Le dossier n'existe pas")
            End If

        End Sub

        <CommandMethod(NomCmd & "Update")> _
        Public Sub RevoUpdate() ' This method can have any name

            Dim Rib As New Revo.AdskApplication
            Dim Ass As New Revo.RevoInfo
            Dim Check As Double = Rib.CheckUpdate

            If Check = -1 Then ' -1 pas de besoin de mis à jour
                MsgBox("La version actuelle est à jour : " & vbCrLf & vbCrLf & Ass.xTitle & "  " & Ass.xVersion, vbInformation, "Mise à jour")
            ElseIf Check = 0 Then ' 0 pas de connexion
                If MsgBox("Vérifier votre connexion Internet, impossible de contrôler la mise à jour. " & vbCrLf & vbCrLf & _
                          "Souhaitez-vous nous contacter pour trouver une solution ?", vbInformation + vbYesNo, "Mise à jour") = MsgBoxResult.Yes Then

                End If
            Else '  < 0 num lic.
                'La mise à jour est effecutée (dans le module CheckUpdate)
            End If

        End Sub
        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "MySQL", CommandFlags.Modal + CommandFlags.Session)>
        Public Sub RevoMySQL()
            Dim FrmSQL As New MySqlConnect
            FrmSQL.Show()
        End Sub

        <CommandMethod(NomCmd & "DivSurf", CommandFlags.Modal)> _
        Public Sub RevoDivSurf() ' This method can have any name

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            ' Put your command code here
            Dim FenetreDiv As New frmDivSurf
            FenetreDiv.ShowDialog()

            If FenetreDiv.clic_calcul Then

                If FenetreDiv.chkLayerMO.Checked = True Then ActivateLayer("MUT_BF")

                If (FenetreDiv.rbnt_dir_par.Checked And FenetreDiv.meth = "par") Or (FenetreDiv.rbnt_dir_perp.Checked And FenetreDiv.meth = "par") Or (FenetreDiv.rbtn_methode_pt_fixe.Checked And FenetreDiv.meth = "pt") Then
                    ReDim FenetreDiv.T_surf(FenetreDiv.N_parcelle - 1)
                    If FenetreDiv.rbtn_calcul_surface_choisie.Checked Then
                        FenetreDiv.meth_surf = "choisie"
                        For i = 0 To FenetreDiv.N_parcelle - 2
                            FenetreDiv.T_surf(i) = FenetreDiv.DataGridView_suface.Rows(i).Cells(0).Value
                        Next
                        FenetreDiv.T_surf(FenetreDiv.N_parcelle - 1) = FenetreDiv.surf_affic
                    Else
                        FenetreDiv.meth_surf = "egale"
                    End If
                    FenetreDiv.Close()
                    FenetreDiv.Dispose()

                    'module de création du tableau de surface à créer 
                    Dim T_surf = Calcul_surf(FenetreDiv.meth_surf, FenetreDiv.Poly, CInt(FenetreDiv.N_parcelle), FenetreDiv.T_surf)
                    'module de calcul
                    Calcul(FenetreDiv.meth, FenetreDiv.orient, FenetreDiv.N_parcelle, FenetreDiv.Poly, FenetreDiv.ptPic1_2d, FenetreDiv.ptPic2_2d, T_surf)

                    ControleSurf(T_surf, FenetreDiv.meth, FenetreDiv.orient, FenetreDiv.N_parcelle, FenetreDiv.Poly, FenetreDiv.ptPic1_2d, FenetreDiv.ptPic2_2d)
                Else
                    MsgBox("Les éléments saisis ne correspondent pas aux méthodes sélectionnées. Vous avez surement changer des paramètres en cours de route")
                    FenetreDiv.Show()
                End If
            End If

        End Sub






        ' <CommandMethod("ImportITF")> _
        'Public Shared Sub Revo_ImportITF()
        Public Shared Sub Revo_ImportITF(ByVal SelFileString As String)

            Dim Ass As New Revo.RevoInfo

            ' Try 'Copie la base de donnée interlis.db3 (DB3interlisOrig -> DB3interlis)
            System.IO.File.Copy(System.IO.Path.Combine(Ass.SystemPath, Ass.DB3interlisOrig), System.IO.Path.Combine(Ass.SystemPath, Ass.DB3interlis), True)

            ' Try 'Copie la base de donnée System.db3 (DB3interlisOrig -> DB3interlis)
            System.IO.File.Copy(System.IO.Path.Combine(Ass.SystemPath, Ass.DB3systemOrig), System.IO.Path.Combine(Ass.SystemPath, Ass.DB3system), True)

            ' Active Model space
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Dim SpaceCur As Autodesk.AutoCAD.Interop.Common.AcadLayout = acDoc.ActiveLayout 'mémo du current Space
            If SpaceCur.Name <> "Model" Then
                For Each acLay As Autodesk.AutoCAD.Interop.Common.AcadLayout In acDoc.Layouts
                    If acLay.Name = "Model" Then
                        acDoc.ActiveLayout = acLay 'Space = acdoc.ModelSpace
                    End If
                Next
            End If


            'connect.Message("Revo", "Traitement en cours ...", False, 50, 100) ' include SQLite in Revo THA
            'System.Windows.Forms.Application.DoEvents()
            Dim booSuccess As Boolean = True
            'Get the current editor
            '-  Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            '-  Dim opts As New PromptStringOptions("")
            '-  Dim Pres As PromptResult = ed.GetString(opts)
            'The string result will contain the SelFiles string we passed when we called SendStringToExecute from RevoFinder
            '-  Dim SelFileString As String = Pres.StringResult
            'Replace the question marks in the string with spaces
            '-  SelFileString = SelFileString.Replace("?", " ")
            'Split the string up by the pipe character 
            Dim SelFiles() As String = SelFileString.Split("|")
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            'Loop through the selected files importing them as before
            For Each ITFFile As String In SelFiles
                'Reset the connection
                dbCon = Nothing
                'And clear the main dataset
                MainDS = Nothing
                Using docLock As DocumentLock = doc.LockDocument
                    booSuccess = booSuccess And ImportITFFile(ITFFile, True)
                End Using
            Next

            Dim connect As New connect

            If booSuccess Then

                If SpaceCur.Name <> "Model" Then
                    For Each acLay As Autodesk.AutoCAD.Interop.Common.AcadLayout In acDoc.Layouts
                        If acLay.Name = SpaceCur.Name Then
                            acDoc.ActiveLayout = acLay 'Space = acdoc.ModelSpace
                        End If
                    Next
                End If

                connect.RevoLog(connect.DateLog & "Cmd Import ITF" & vbTab & "True" & vbTab & "Type: " & "imported successfully" & vbTab & SelFileString)
                'connect.Message(Ass.xTitle, "Importation terminée avec succès !", False, 100, 100, "info") ' include SQLite in Revo THA
                'connect.Message(Ass.xTitle, "", True, 100, 100)

                'System.Windows.Forms.MessageBox.Show("End Procedure !!!", "Import ITF File", _
                'Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Information)
            Else
                connect.RevoLog(connect.DateLog & "Cmd Import ITF" & vbTab & "False" & vbTab & "Type: " & "Not all files were imported successfully" & vbTab & SelFileString)
                connect.Message(Ass.xTitle, "Des objets n'ont pas été importé correctement", False, 100, 100, "critical") ' include SQLite in Revo THA
                'System.Windows.Forms.MessageBox.Show("End Procedure !!!" & vbCrLf & "Not all files were imported successfully.", _
                '  "Import ITF File", Windows.Forms.MessageBoxButtons.OK, _
                ' Windows.Forms.MessageBoxIcon.Exclamation)
            End If

            'Catch
            '    MsgBox("Impossible de copier la base de donnée :" & Ass.DB3interlisOrig, MsgBoxStyle.Critical, "Erreur de traitement")
            'End Try

            'Ouverture de l'historique
            If connect.ActLog = True Then RevoFiler.OpenExe(Ass.LogPath)

        End Sub

        <CommandMethod(NomCmd & "UpdateSharedSpace")> _
        Private Sub UpdateSharedSpace() 'Mise à jour de l'espace partagé()
            Dim Connect As New connect

            'Concept de mise à jour de l'espace partagé
            ' actionspath : *.csv (cadastre.csv, ExportData.csv, ImportData.csv )
            ' sharedpath : (XMLformat)
            ' supportpath : 
            ' template : Revo11.dwt
            ' plotters :
            ' Library

            'Mise à jour des fichiers de l'espace partagé
            Dim Ass As New Revo.RevoInfo
            Try
                Dim UpdateFileXML As String = IO.Path.Combine(Ass.PluginPath, "update.xml") ' Dossier Local racine

                If IO.File.Exists(UpdateFileXML) Then

                    Connect.Message("Mise à jour des dossiers partagés", "Les fichiers et dossiers partagé sont mises à jour !" & vbCrLf & "Une copie de sauvegarde est effectué au même emplacement.", False, 50, 100)

                    'Intérogation version ' 0.9.2.0 = 920
                    Dim VersLocal As Double = Trim(Replace(Ass.xVersion, ".", ""))
                    Dim XMLversion As New System.Xml.XmlDocument
                    XMLversion.Load(UpdateFileXML)
                    Dim VersUp As String = XMLversion.SelectSingleNode("/update/version").InnerText.ToLower

                    If VersLocal = CDbl(VersUp) Then ' Si la version = donc mise à jour des nouveautés listé dans le XML

                        Dim ListParam As System.Xml.XmlNodeList
                        ListParam = XMLversion.SelectNodes("/update/folder")

                        For Each Param As System.Xml.XmlNode In ListParam
                            If Param.Attributes.ItemOf(0).InnerText.ToLower = "ActionsPath".ToLower Then
                                UpdateShared(Ass.ActionsPath(True), Ass.ActionsPath, True, Param, "ActionsPath".ToLower)
                            ElseIf Param.Attributes.ItemOf(0).InnerText.ToLower = "SharedPath".ToLower Then
                                UpdateShared(Ass.SharedPath(True), Ass.SharedPath, True, Param, "SharedPath".ToLower)
                            ElseIf Param.Attributes.ItemOf(0).InnerText.ToLower = "SupportPath".ToLower Then
                                UpdateShared(Ass.SupportPath(True), Ass.SupportPath, True, Param, "SupportPath".ToLower)
                            ElseIf Param.Attributes.ItemOf(0).InnerText.ToLower = "Template".ToLower Then
                                UpdateShared(Ass.Template(True), Ass.Template, False, Param, "Template".ToLower)
                            ElseIf Param.Attributes.ItemOf(0).InnerText.ToLower = "Plotters".ToLower Then
                                UpdateShared(Ass.Library(True), Ass.Library, True, Param, "Plotters".ToLower)
                            ElseIf Param.Attributes.ItemOf(0).InnerText.ToLower = "Library".ToLower Then
                                UpdateShared(Ass.Library(True), Ass.Library, True, Param, "Library".ToLower)
                            End If
                        Next
                    Else
                        '  MsgBox("pas de mise à jour : " & VersLocal & " = " & numVersUp)
                    End If
                    'Si Oui 
                    'Recherche mise à jour
                    IO.File.Copy(UpdateFileXML, IO.Path.Combine(Ass.SharedPath, "update.xml"), True)
                    IO.File.Delete(UpdateFileXML)

                    Connect.Message(Ass.xTitle, "Mise à jour teminée avec succès!" & vbCrLf & "Une copie de sauvegarde a été effectué au même emplacement.", False, 100, 100, "info")
                    Connect.Message(Ass.xTitle, "", True, 100, 100)

                End If

            Catch

                Connect.Message("Mise à jour interrompu !", "Problème de mise à jour", True, 99, 100)

            End Try

        End Sub

        Private Sub UpdateShared(ByVal Source As String, ByVal Destination As String, ByVal Folder As Boolean, NodParam As System.Xml.XmlNode, NameDest As String)

            'Si Paramètre modifié
            If Source.ToUpper <> Destination.ToUpper Then


                'Chargement du noeuds 
                Dim RVinfo As New RevoInfo
                Dim XmlVersDest As String = IO.Path.Combine(RVinfo.SharedPath, "update.xml") ' Dossier distant (Shared)
                Dim ListParam As System.Xml.XmlNodeList
                ListParam = NodParam.SelectNodes("param")
                Dim ListParamDest As System.Xml.XmlNodeList = Nothing

                If IO.File.Exists(XmlVersDest) Then
                    'Chargement de l'arbo distante
                    Dim XMLversion As New System.Xml.XmlDocument
                    XMLversion.Load(XmlVersDest)
                    Dim ListFolderDest As System.Xml.XmlNodeList
                    ListFolderDest = XMLversion.SelectNodes("/update/folder")
                    For Each FolderDest As System.Xml.XmlNode In ListFolderDest
                        If FolderDest.Attributes.ItemOf(0).InnerText.ToLower = NameDest Then
                            ListParamDest = FolderDest.SelectNodes("param")
                            Exit For
                        End If
                    Next
                End If


                For Each Param As System.Xml.XmlNode In ListParam


                    Try
                        Dim Vers As Double = CDbl(Param.Attributes.ItemOf(0).InnerText)
                        Dim SourceFile As String = Param.InnerText


                        If Folder Then ' Si un dossier
                            If IO.Directory.Exists(Destination) Then 'Si existant copy un dossier-backup

                                Dim TestVersion As Boolean = False
                                Dim VersDest As Double = 0

                                If ListParamDest IsNot Nothing Then
                                    For Each ParamDest As System.Xml.XmlNode In ListParamDest
                                        If ParamDest.InnerText.ToLower = SourceFile.ToLower Then
                                            VersDest = CDbl(ParamDest.Attributes.ItemOf(0).InnerText)
                                            Exit For
                                        End If
                                    Next
                                Else
                                    TestVersion = True
                                End If
                                If VersDest = 0 Then TestVersion = True 'Si pas de version : copie

                                'Test si remplacer Vers(à jour) > VersDest (existant)
                                If Vers > VersDest Then TestVersion = True


                                'Si version manquant ou plus ancienne écrase le fichier
                                If TestVersion Then
                                    Dim SourceFileCopy As String = IO.Path.Combine(RVinfo.PluginPath, SourceFile)
                                    Dim BackupFile As String = IO.Path.Combine(Destination, IO.Path.GetFileName(SourceFile) & "-backup")
                                    Dim DestFile As String = IO.Path.Combine(Destination, IO.Path.GetFileName(SourceFile))

                                    'Copie le fichier > backup
                                    If IO.File.Exists(DestFile) Then IO.File.Copy(DestFile, BackupFile, True)
                                    'Puis Ecrase le fichier
                                    IO.File.Copy(SourceFileCopy, DestFile, True)
                                End If

                            Else 'Si existe pas : Copie le dossier
                                Revo.RevoFiler.CopyDirectory(Source, Destination, True)
                            End If

                        Else ' Si un fichier

                            If IO.File.Exists(Destination) Then 'Si existant copy un fichier-backup

                                Dim TestVersion As Boolean = False
                                Dim VersDest As Double = 0

                                If ListParamDest IsNot Nothing Then
                                    For Each ParamDest As System.Xml.XmlNode In ListParamDest
                                        If ParamDest.InnerText.ToLower = SourceFile.ToLower Then
                                            VersDest = CDbl(ParamDest.Attributes.ItemOf(0).InnerText)
                                            Exit For
                                        End If
                                    Next
                                Else
                                    TestVersion = True
                                End If

                                If VersDest = 0 Then TestVersion = True 'Si pas de version : copie

                                'Test si remplacer Vers(à jour) > VersDest (existant)
                                If Vers > VersDest Then TestVersion = True

                                If TestVersion Then
                                    Dim SourceFileCopy As String = IO.Path.Combine(RVinfo.PluginPath, SourceFile)
                                    Dim BackupFile As String = IO.Path.Combine(IO.Path.GetDirectoryName(Destination), IO.Path.GetFileName(SourceFile) & "-backup")
                                    Dim DestFile As String = IO.Path.Combine(IO.Path.GetDirectoryName(Destination), IO.Path.GetFileName(SourceFile))

                                    'Copie le fichier > backup
                                    If IO.File.Exists(DestFile) Then IO.File.Copy(DestFile, BackupFile, True)
                                    'Puis Ecrase le fichier
                                    IO.File.Copy(SourceFileCopy, DestFile, True)
                                End If

                            Else 'Si existe pas : Copie le fichier
                                IO.Directory.CreateDirectory(IO.Path.GetDirectoryName(Destination)) 'Crée un repertoire
                                IO.File.Copy(IO.Path.Combine(RVinfo.PluginPath, SourceFile), Destination, True)
                            End If
                        End If


                    Catch ex As System.Exception
                        MsgBox(ex.Message)
                    End Try

                Next
            End If
        End Sub



        ' Modal Command with pickfirst selection
        <CommandMethod(NomCmd & "ClearPW", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        Public Sub RevoCearPW() ' This method can have any name
            Dim Ass As New Revo.RevoInfo
            Dim con As New SQLite.SQLiteConnection
            Dim DBname As String = InputBox("Nom de la base de donnée (*.db3)", "Suppresion du mot de passe d'une DB", "")
            '3212-Interlis.db3
            Dim DBPath As String = Ass.SystemPath ' include SQLite in Revo THA

            If MsgBox("Supprimer le mot de passe ? " & vbCrLf & "(Non : ajout du mot de passe)", vbYesNo + vbInformation, "Reset BD") = MsgBoxResult.Yes Then
                Try
                    con.ConnectionString = String.Format("Data Source = {0}" & DBname & "; PASSWORD={1}", DBPath, DBPwd)
                    con.Open()
                    con.ChangePassword("")
                    con.Dispose()
                Catch ex As System.Exception
                    MsgBox(ex.Message)
                End Try
            Else
                Try
                    con.ConnectionString = String.Format("Data Source = {0}" & DBname, DBPath)
                    con.Open()
                    con.ChangePassword(DBPwd)
                    con.Dispose()
                Catch ex As System.Exception
                    MsgBox(ex.Message)
                End Try
            End If



        End Sub




        ' Modal Command with pickfirst selection
        <CommandMethod(NomCmd & "List", CommandFlags.Modal + CommandFlags.UsePickSet)> _
        Public Sub RevoList() ' This method can have any name

            Dim RvInfo As New RevoInfo

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            'Dim CurLA As Autodesk.AutoCAD.Interop.Common.AcadLayer
            Dim Textes As String = ""
            Textes += "#Title;Listing " & acDoc.Name & vbCrLf & "#Description;Listing des objets du dessin" & vbCrLf & "#Groupe;Revo" & vbCrLf & "#Version;" & CDbl(Replace(RvInfo.xVersion, ".", "")) / 1000 & vbCrLf & vbCrLf & vbCrLf


            'Boucle dans les calques ---------------------------------------------------------------------------------------------------------------
            Textes += "' Listing des calques" & vbCrLf

            Dim Calques As Autodesk.AutoCAD.Interop.Common.AcadLayers = acDoc.Layers
            Dim LA As Autodesk.AutoCAD.Interop.Common.AcadLayer
            Dim i As Double = 0
            Dim Plottable As Integer = 1
            Dim Epaisseur As String = ""
            Dim Couleur As String = ""

            For Each LA In Calques

                'nom ; couleur ; type de ligne ; tracé
                If LA.Plottable = False Then
                    Plottable = 0
                Else
                    Plottable = 1
                End If
                If LA.Lineweight.ToString = "acLnWtByBlock" Then 'ByBlock
                    Epaisseur = "ByBlock"
                ElseIf LA.Lineweight.ToString = "acLnWtByLayer" Then 'ByLayer
                    Epaisseur = "ByLayer"
                ElseIf LA.Lineweight.ToString = "acLnWtByLwDefault" Then 'ByDefault
                    Epaisseur = "ByDefault"
                Else ' ACAD_LWEIGHT.acLnWt018
                    Epaisseur = CDbl(Replace(LA.Lineweight.ToString, "acLnWt", "")) / 100 ' = 0.18
                End If

                If LA.TrueColor.ColorMethod.ToString = "acColorMethodByRGB" Then 'RGB
                    Couleur = LA.TrueColor.Red & "," & LA.TrueColor.Green & "," & LA.TrueColor.Blue
                ElseIf LA.TrueColor.ColorMethod.ToString = "acColorMethodByACI" Then 'Palette AutoCAD
                    Couleur = LA.TrueColor.ColorIndex.ToString
                    Couleur = Replace(Couleur, "acRed", 1)
                    Couleur = Replace(Couleur, "acYellow", 2)
                    Couleur = Replace(Couleur, "acGreen", 3)
                    Couleur = Replace(Couleur, "acCyan", 4)
                    Couleur = Replace(Couleur, "acBlue", 5)
                    Couleur = Replace(Couleur, "acMagenta", 6)
                    Couleur = Replace(Couleur, "acWhite", 7)
                End If

                Try
                    ' #LA;NomOriginal;[[PropA]]Val;[[PropB]]Val;
                    Textes += ("#LA;>" & LA.Name & ";[[TrueColor]];" & Couleur & ";[[Linetype]];" & LA.Linetype & ";[[Plottable]];" & Plottable & ";[[Lineweight]];" & Epaisseur & ";") & vbCrLf
                Catch
                End Try

                'Test du nom de calques
                'Si trouve le calque > traite au Param.
                i += 1
            Next


            'Boucle dans les blocs ---------------------------------------------------------------------------------------------------------------
            'Boucle dans la DEFINITION du block : bibliotheque
            Textes += vbCrLf & vbCrLf & "' Listing des blocs" & vbCrLf
            Dim PosX As Double = 0

            For Each BL As Autodesk.AutoCAD.Interop.Common.AcadBlock In acDoc.Blocks

                ' #BL;NomOriginal;[[PropA]]Val;[[PropB]]Val;
                Textes += ("#BL;>" & BL.Name & ";[[XYZInsertionPoint]];" & PosX & "," & 0 & "," & 0 & ";" & vbCrLf)
                PosX += 1
            Next



            'Boucle de "Gel de la fenêtre" dans les présentations ---------------------------------------------------------------------------------------------------------------
            Textes += vbCrLf & vbCrLf & "' Listing des calques Gelé dans la fenêtre" & vbCrLf


            ' Dim Calques As Autodesk.AutoCAD.Interop.Common.AcadLayers = acDoc.Layers
            '  Dim LA As Autodesk.AutoCAD.Interop.Common.AcadLayer
            Dim u As Double = 0
            Dim GelFen As Integer = 1
            Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = doc.Editor
            Dim db As Database = doc.Database
            Dim idLay As ObjectId
            Dim idLayTblRcd As ObjectId
            Dim lt As LayerTableRecord
            Dim layOut As Layout
            Dim tm As Autodesk.AutoCAD.ApplicationServices.TransactionManager = db.TransactionManager
            Dim ta As Transaction = tm.StartTransaction()



            Dim LayoutColl As New Collection

            Dim layoutMgr As LayoutManager = LayoutManager.Current
            ed.WriteMessage([String].Format("{0}Active Layout is : {1}", Environment.NewLine, layoutMgr.CurrentLayout))
            ed.WriteMessage([String].Format("{0}Number of Layouts: {1}{0}List of all Layouts:", Environment.NewLine, layoutMgr.LayoutCount))
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim layDico As DBDictionary = TryCast(tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, False), DBDictionary)
                For Each entry As DBDictionaryEntry In layDico
                    Dim layoutId As ObjectId = entry.Value
                    Dim layoutL As Layout = TryCast(tr.GetObject(layoutId, OpenMode.ForRead), Layout)
                    ed.WriteMessage([String].Format("{0}--> {1}", Environment.NewLine, layoutL.LayoutName))

                    LayoutColl.Add(layoutL)
                Next
                ' tr.Commit()
            End Using



            Try

                Dim acLayoutMgr As LayoutManager '= LayoutManager.Current

                '  For Each CurLayout As Layout In LayoutColl

                'If CurLayout.LayoutName <> "Model" Then

                ' Textes += vbCrLf & vbCrLf & "' Espace papier  :  " & CurLayout.LayoutName & vbCrLf

                '    MsgBox(CurLayout.LayoutName)
                '  acLayoutMgr.CurrentLayout = CurLayout.LayoutName


                acLayoutMgr = LayoutManager.Current


                Dim layDict As DBDictionary = DirectCast(ta.GetObject(db.LayoutDictionaryId, OpenMode.ForRead, False), DBDictionary)

                For Each itmdict As DBDictionaryEntry In layDict



                    layOut = DirectCast(ta.GetObject(itmdict.Value, OpenMode.ForRead, False), Layout)
                    ed.WriteMessage(vbLf + "Layout: {0}" + vbLf, layOut.LayoutName)



                    If layOut.LayoutName <> "Model" Then


                        acLayoutMgr.CurrentLayout = layOut.LayoutName

                        '   Textes += vbCrLf & vbCrLf & "' Espace papier  :  " & layOut.LayoutName & vbCrLf

                        Try
                            ed.SwitchToModelSpace()
                        Catch
                            ed.WriteMessage(vbLf + "Layout hasn't a window" + vbLf)
                        End Try

                        Dim ltt As LayerTable = DirectCast(ta.GetObject(db.LayerTableId, OpenMode.ForRead, False), LayerTable)
                        For Each LA In Calques '  For Each lname As String In layers
                            idLay = ltt(LA.Name)
                            lt = ta.GetObject(idLay, OpenMode.ForRead)
                            If ltt.Has(LA.Name) Then
                                idLayTblRcd = ltt.Item(LA.Name)
                            Else
                                ed.WriteMessage("Layer: """ + LA.Name + """ not available")

                                Exit For '   Return

                            End If

                            Dim idCol As ObjectIdCollection = New ObjectIdCollection
                            idCol.Add(idLayTblRcd)

                            ' Check that we are in paper space 

                            Dim vpid As ObjectId = ed.CurrentViewportObjectId
                            If vpid.IsNull() Then
                                ed.WriteMessage("No Viewport current.")
                                Exit For ' Return
                            End If

                            'VP need to be open for write 

                            Dim oViewport As Viewport = DirectCast(tm.GetObject(vpid, OpenMode.ForWrite, False), Viewport)
                            If Not oViewport.IsLayerFrozenInViewport(idLayTblRcd) Then
                                ' oViewport.FreezeLayersInViewport(idCol.GetEnumerator())
                                GelFen = 0
                            Else
                                'oViewport.ThawLayersInViewport(idCol.GetEnumerator())
                                GelFen = 1
                            End If


                            Textes += ("#LA;>" & LA.Name & ";[[Space]];" & layOut.LayoutName & ";[[Freeze]];" & GelFen) & vbCrLf

                        Next

                        ed.SwitchToPaperSpace()

                    End If



                Next
                '  ta.Commit()




                '   End If



                ' Next




            Finally

                ta.Dispose()

            End Try







            'Boucle dans les export   ---------------------------------------------------------------------------------------------------------------

            If 1 = 1 Then


                Textes += vbCrLf & vbCrLf & "' Listing des polylignes" & vbCrLf
                Textes += "project_id###parent_task_id###task_title###task_description###shape_lv03_path###shape_type" & vbCrLf

                Dim Inc As Double = 1
                Dim connect As New Revo.connect

                Try

                    Dim Objs As New Collection
                    Objs = SelectObj(False, "Sélectionner les objets (bloc ou polyligne 2D", True, False, True)


                    Dim acDocX As Document = Application.DocumentManager.MdiActiveDocument
                    Dim acCurDb As Database = acDocX.Database

                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                        '' Open the Block table for read
                        Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                        For Each Obj As Entity In Objs

                            connect.Message("Génération des polyligne", "Traitement teminée avec succès!", False, Inc, Objs.Count, "info")

                            Dim Num As String = ""
                            Dim Layer As String = ""
                            Dim Description As String = ""
                            Dim pathXY As String = ""
                            Dim ShapeValide As Double = 1

                            Num = Obj.Layer & " " & Inc
                            Layer = Obj.Layer

                            Dim Project As String = "839"
                            If Layer = "STN_AEL_COLL" Then Layer = 40065
                            If Layer = "STN_AEL_BRAN" Then Layer = 40064
                            If Layer = "STN_AEN_CHAM" Then Layer = 40066


                            'Si Polyligne
                            If TypeName(Obj) Like "*Polyline*" Then

                                Dim Poly As Polyline = Obj

                                For i = 0 To Poly.NumberOfVertices - 1
                                    If pathXY <> "" Then pathXY += ";"
                                    pathXY += Math.Round(Poly.GetPoint2dAt(i).X, 3) & "," & Math.Round(Poly.GetPoint2dAt(i).Y, 3)
                                Next

                                Description = "Longueur : " & Math.Round(Poly.Length, 2) & " m" & "###" & "Aire : " & Math.Round(Poly.Area, 2) & " m2"
                                If Poly.NumberOfVertices < 3 Then ShapeValide = 1

                                Textes += (Project & "###" & Layer & "###" & Num & "###" & Description & "###" & pathXY & "###" & ShapeValide) & vbCrLf


                            ElseIf TypeName(Obj) = "BlockReference" Then

                                Dim BL As BlockReference = Obj
                                pathXY = Math.Round(BL.Position.X, 3) & "," & Math.Round(BL.Position.Y, 3)


                                'Lecture Xdata
                                Dim Value As New List(Of String)
                                Dim rb As ResultBuffer = Obj.XData
                                If rb Is Nothing Then
                                    ed.WriteMessage(vbLf & "Entity does not have XData attached.")
                                Else
                                    Dim n As Integer = 0
                                    For Each tv As TypedValue In rb
                                        ' ed.WriteMessage(vbLf & "TypedValue {0} - type: {1}, value: {2}", System.Math.Max(System.Threading.Interlocked.Increment(n), n - 1), tv.TypeCode, tv.Value)
                                        Value.Add(tv.Value)
                                    Next
                                    rb.Dispose()
                                End If

                                Dim XdataValue As String = ""

                                For x = 0 To Value.Count - 1
                                    If Value(x).ToString = "GEN_ASS" Then
                                        If Description <> "" Then Description += "\n"
                                        Description += "GEN_ASS : " & Value(x + 1).ToString
                                    ElseIf Value(x).ToString = "REMARQUE" Then
                                        If Description <> "" Then Description += "\n"
                                        Description += "Remarque : " & Value(x + 1).ToString
                                    End If

                                Next

                                ShapeValide = 0

                                Textes += (Project & "###" & Layer & "###" & Num & "###" & Description & "###" & pathXY & "###" & ShapeValide) & vbCrLf

                            End If

                            Inc += 1



                        Next

                        '' Save the new objects to the database
                        acTrans.Commit()
                    End Using

                Catch ex As System.Exception
                    MsgBox(ex.Message)
                End Try
                connect.Message("Génération des données", "", True, 100, 100, "info")


            End If







            Dim SaveFileDialogExport As New System.Windows.Forms.SaveFileDialog
            Dim CheminFichier As String '= "c:\export-calque.txt"
            SaveFileDialogExport.Filter = "CSV (séparateur: point-virgule) (*.csv)|*.csv"
            SaveFileDialogExport.Title = "Exporter des données"
            SaveFileDialogExport.FileName = ""
            CheminFichier = SaveFileDialogExport.ShowDialog()

            Dim Confirm
            Dim ListText() As String = Split(Textes, vbCrLf)
            Confirm = RevoFiler.EcritureFichier(SaveFileDialogExport.FileName, ListText, False)

            'FIN Traitement des couches -----------------------------------------------------------------------------------



        End Sub


        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "Analyse")>
        Public Sub RevoAnalyse()

            Dim Ass As New Revo.RevoInfo
            Dim Connect As New connect
            Dim ChoixAnalyse As String = ""

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            'Choix du type d'analyse ***********
            Dim optstring As New PromptKeywordOptions("Choisir le type d'analyse")
            optstring.Keywords.Add("Suppression points superposés")   ' Suppression points superposés (supp. point inférieur)
            optstring.Keywords.Add("Intersection de polyligne") ' Intersection de polyligne
            optstring.Keywords.Add("Objet à l'intérieur d'une polyligne") ' Objet à l'intérieur d'une polyligne 
            optstring.Keywords.Add("Polyligne fermée (surface)") ' Créer Surface avec les polyligne affiché
            optstring.Keywords.Add("Extrait altitude d'un périmètre") 'AnaylseMaxMinZ
            'optstring.AllowArbitraryInput = False
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim pr As PromptResult = acDoc.Editor.GetKeywords(optstring)
            If pr.StringResult.ToLower = "suppression" Or pr.StringResult.ToLower = "s" Then ChoixAnalyse = "SuppPointDoubleZ"
            If pr.StringResult.ToLower = "intersection" Or pr.StringResult.ToLower = "i" Then ChoixAnalyse = "IntersPoly"
            If pr.StringResult.ToLower = "objet" Or pr.StringResult.ToLower = "i" Then ChoixAnalyse = "ObjInterPoly"
            If pr.StringResult.ToLower = "polyligne" Or pr.StringResult.ToLower = "p" Then ChoixAnalyse = "CreerSurface"
            If pr.StringResult.ToLower = "extrait" Or pr.StringResult.ToLower = "e" Then ChoixAnalyse = "AnylseMaxMinZ"


            Dim Analy As New RevoAnalyse

            If ChoixAnalyse = "IntersPoly" Then ' Intersection de polyligne ************
                Analy.IntersPoly()
            ElseIf ChoixAnalyse = "ObjInterPoly" Then       ' Objet à l'intérieur d'une polyligne 
                Analy.ObjInterPoly()
            ElseIf ChoixAnalyse = "SuppPointDoubleZ" Then       ' Suppression points superposés (supp. point inférieur)
                Analy.SuppPointDoubleZ()
            ElseIf ChoixAnalyse = "CreerSurface" Then       ' Creer Surface
                Analy.CreerSurface()
            ElseIf ChoixAnalyse = "AnylseMaxMinZ" Then 'AnylseMaxMinZ
                Analy.AnylseMaxMinZ()
            End If


        End Sub


        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "CumulDist")>
        Public Sub RevoCumulDist()

            Dim Ass As New Revo.RevoInfo
            Dim Connect As New connect
            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()


            'Sélectionner le point de départ = 1er polyligne

            'Sélectionner la trajectoire
            Dim Polys As Collection = SelectPolyline(False, False, "Sélectionner les polylignes définisant le trajectoire")
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim Prec As Integer = acDoc.Database.Luprec
            Dim TrajPts As New Collection


            'Recherche la suite des sommets : Création d'un trajectoire
            Dim PolyA As Polyline = Polys(0) 'Initialise la polyligne de départ
            Dim NbreA As Double = PolyA.NumberOfVertices - 1
            Dim LoadP0 As Boolean = False

            For u = 1 To Polys.Count - 2 'Boucle sans la 1er polyligne


                For i = 1 To Polys.Count - 2 'Boucle sans la 1er polyligne
                    Dim NbreB As Double = Polys(i).NumberOfVertices - 1

                    '    Si début PolyA = début PolyB
                    If Math.Round(PolyA.GetPoint2dAt(0).X, Prec) = Math.Round(Polys(i).GetPoint2dAt(0).X, Prec) And
                       Math.Round(PolyA.GetPoint2dAt(0).Y, Prec) = Math.Round(Polys(i).GetPoint2dAt(0).Y, Prec) Then

                        'Chargement de la Poly(0)
                        If LoadP0 = False Then LoadP0 = True : For x = 0 To Polys(0).NumberOfVertices - 1 : TrajPts.Add(Polys(0).GetPoint2dAt(x)) : Next

                        For x = 0 To Polys(i).NumberOfVertices - 1
                        Next
                        Exit For

                        'Si début PolyA = fin PolyB
                    ElseIf Math.Round(PolyA.GetPoint2dAt(0).X, Prec) = Math.Round(Polys(i).GetPoint2dAt(NbreB).X, Prec) And
                           Math.Round(PolyA.GetPoint2dAt(0).Y, Prec) = Math.Round(Polys(i).GetPoint2dAt(NbreB).Y, Prec) Then

                        'Chargement de la Poly(0)
                        If LoadP0 = False Then LoadP0 = True : For x = 0 To Polys(0).NumberOfVertices - 1 : TrajPts.Add(Polys(0).GetPoint2dAt(x)) : Next


                        'Si fin PolyA = début PolyB
                    ElseIf Math.Round(PolyA.GetPoint2dAt(NbreA).X, Prec) = Math.Round(Polys(i).GetPoint2dAt(0).X, Prec) And
                           Math.Round(PolyA.GetPoint2dAt(NbreA).Y, Prec) = Math.Round(Polys(i).GetPoint2dAt(0).Y, Prec) Then

                        'Chargement de la Poly(0)
                        If LoadP0 = False Then LoadP0 = True : For x = 0 To Polys(0).NumberOfVertices - 1 : TrajPts.Add(Polys(0).GetPoint2dAt(x)) : Next

                        'Si fin PolyA = fin PolyB
                    ElseIf Math.Round(PolyA.GetPoint2dAt(NbreA).X, Prec) = Math.Round(Polys(i).GetPoint2dAt(NbreB).X, Prec) And
                           Math.Round(PolyA.GetPoint2dAt(NbreA).Y, Prec) = Math.Round(Polys(i).GetPoint2dAt(NbreB).Y, Prec) Then

                        'Chargement de la Poly(0)
                        If LoadP0 = False Then LoadP0 = True : For x = 0 To Polys(0).NumberOfVertices - 1 : TrajPts.Add(Polys(0).GetPoint2dAt(x)) : Next


                    End If
                Next

            Next


            'For i = 0 To Poly.NumberOfVertices - 1
            '    '  poly3d.AppendVertex(New PolylineVertex3d(New Point3d(Poly.GetPoint2dAt(i).X, Poly.GetPoint2dAt(i).Y, 0)))
            '    'poly3d.(New PolylineVertex3d(New Point3d(1, 1, 1)))

            '    acPts3dPoly.Add(New Point3d(Poly.GetPoint2dAt(i).X, Poly.GetPoint2dAt(i).Y, 0))
            'Next






        End Sub


        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "Find")> _
        Public Sub RevoFind() ' Recherche de surface (arrondie)

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Ass As New Revo.RevoInfo
            Dim Connect As New connect
            Dim ChoixAnalyse As String = ""

            'Choix du type d'analyse ***********
            Dim optstring As New PromptKeywordOptions("Choisir le type de recherche")
            optstring.Keywords.Add("Surfacique")
            ' optstring.Keywords.Add("Intersection de polyligne")
            ' optstring.Keywords.Add("Objet à l'intérieur d'une polyligne")
            'optstring.AllowArbitraryInput = False
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim pr As PromptResult = acDoc.Editor.GetKeywords(optstring)
            If pr.StringResult.ToLower = "surfacique" Or pr.StringResult.ToLower = "s" Then ChoixAnalyse = "Surfacique"
            ' If pr.StringResult.ToLower = "intersection" Or pr.StringResult.ToLower = "i" Then ChoixAnalyse = "IntersPoly"
            ' If pr.StringResult.ToLower = "objet" Or pr.StringResult.ToLower = "i" Then ChoixAnalyse = "ObjInterPoly"

            If ChoixAnalyse = "Surfacique" Then ' Intersection de polyligne ************

                Dim Polys As Collection = SelectPolyline(False, False, "Sélectionner les polylignes définisant des surfaces")
                Dim ValSurf As String = ""

                ''Saisir surface
                Dim pso As New PromptStringOptions("Saisir la surface [unité courante] :")

                pso.AllowSpaces = True
                'Dim pr As PromptResult
                Do
                    pr = acDoc.Editor.GetString(pso)     '--------->never stop there if showpalette is executed.
                    If pr.StringResult = "" Then Exit Do
                Loop While IsNumeric(pr.StringResult) = False  ' pr.Status = PromptStatus.OK
                If IsNumeric(pr.StringResult) Then ValSurf = pr.StringResult

                If Polys.Count <> 0 And ValSurf <> "" Then
                    Dim arrondi As Integer = Len(ValSurf) - Len(CStr(Math.Round(CDbl(ValSurf), 0))) - 1
                    If arrondi < 0 Then arrondi = 0

                    For Each Poly As Polyline In Polys
                        If Math.Round(CDbl(ValSurf), arrondi) = Math.Round(Poly.Area, arrondi) Then
                            Zooming.ZoomToObject(Poly.ObjectId, 10)
                            SetBlockColour(Poly.ObjectId, Autodesk.AutoCAD.Colors.Color.FromColor(Drawing.Color.Magenta))
                            Exit For
                        End If
                    Next
                End If


                'ElseIf ChoixAnalyse = "ObjInterPoly" Then       ' Objet à l'intérieur d'une polyligne 
                '    Analy.ObjInterPoly()
                'ElseIf ChoixAnalyse = "SuppPointDoubleZ" Then       ' Suppression points superposés 
                '    Analy.SuppPointDoubleZ()
            End If


        End Sub

        ' Application Session Command with localized name
        <CommandMethod(NomCmd & "Poly")>
        Public Sub RevoPoly()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

            Dim Polys As New Collection
            Polys = SelectPolyline(False, False, "Sélectionner les polylignes 2D")

            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                '' Open the Block table for read
                Dim acBlkTbl As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                '' Open the Block table record Model space for write
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)






                ''' Create a polyline with two segments (3 points)
                'Dim acPoly As Polyline = New Polyline()
                'acPoly.AddVertexAt(0, New Point2d(1, 1), 0, 0, 0)
                'acPoly.AddVertexAt(1, New Point2d(1, 2), 0, 0, 0)
                'acPoly.AddVertexAt(2, New Point2d(2, 2), 0, 0, 0)
                'acPoly.ColorIndex = 1

                ''' Add the new object to the block table record and the transaction
                'acBlkTblRec.AppendEntity(acPoly)
                'acTrans.AddNewlyCreatedDBObject(acPoly, True)





                '' Create a 3D polyline with two segments (3 points)
                Dim acPoly3d As Polyline3d = New Polyline3d()
                acPoly3d.ColorIndex = 5

                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acPoly3d)
                acTrans.AddNewlyCreatedDBObject(acPoly3d, True)

                '' Before adding vertexes, the polyline must be in the drawing
                Dim acPts3dPoly As Point3dCollection = New Point3dCollection()

                For Each Poly As Polyline In Polys
                    Dim poly3d As New Polyline3d
                    poly3d.SetDatabaseDefaults()
                    For i = 0 To Poly.NumberOfVertices - 1
                        '  poly3d.AppendVertex(New PolylineVertex3d(New Point3d(Poly.GetPoint2dAt(i).X, Poly.GetPoint2dAt(i).Y, 0)))
                        'poly3d.(New PolylineVertex3d(New Point3d(1, 1, 1)))
                        acPts3dPoly.Add(New Point3d(Poly.GetPoint2dAt(i).X, Poly.GetPoint2dAt(i).Y, 0))
                    Next
                Next


                For Each acPt3d As Point3d In acPts3dPoly
                    Dim acPolVer3d As PolylineVertex3d = New PolylineVertex3d(acPt3d)
                    acPoly3d.AppendVertex(acPolVer3d)
                    acTrans.AddNewlyCreatedDBObject(acPolVer3d, True)
                Next

                ''' Get the coordinates of the lightweight polyline
                'Dim acPts2d As Point2dCollection = New Point2dCollection()
                'For nCnt As Integer = 0 To acPoly.NumberOfVertices - 1
                '    acPts2d.Add(acPoly.GetPoint2dAt(nCnt))
                'Next

                '' Get the coordinates of the 3D polyline
                Dim acPts3d As Point3dCollection = New Point3dCollection()
                For Each acObjIdVert As ObjectId In acPoly3d
                    Dim acPolVer3d As PolylineVertex3d
                    acPolVer3d = acTrans.GetObject(acObjIdVert,
                                                   OpenMode.ForRead)

                    acPts3d.Add(acPolVer3d.Position)
                Next

                ''' Display the Coordinates
                'Application.ShowAlertDialog("2D polyline (red): " & vbLf & _
                '                            acPts2d(0).ToString() & vbLf & _
                '                            acPts2d(1).ToString() & vbLf & _
                '                            acPts2d(2).ToString())

                Application.ShowAlertDialog("3D polyline (blue): " & vbLf &
                                            acPts3d(0).ToString() & vbLf &
                                            acPts3d(1).ToString() & vbLf &
                                            acPts3d(2).ToString())

                '' Save the new objects to the database
                acTrans.Commit()
            End Using


        End Sub


        <CommandMethod("REVOztestvariable")> _
        Public Sub testvariable()

            Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim ed As Editor = acDoc.Editor
            Dim RVinfo As New Revo.RevoInfo



            ed.WriteMessage(vbLf & "EDTCheckSurf : " & RVinfo.EDTCheckSurf)
            ed.WriteMessage(vbLf & "EDTFolder : " & RVinfo.EDTFolder)
            ed.WriteMessage(vbLf & "EDTIgnoreRPL : " & RVinfo.EDTIgnoreRPL)
            ed.WriteMessage(vbLf & "EDTtemplate : " & RVinfo.EDTtemplate)
            ed.WriteMessage(vbLf & "FormatExport : " & RVinfo.FormatExport)
            ed.WriteMessage(vbLf & "FormatImport : " & RVinfo.FormatImport)
            ed.WriteMessage(vbLf & "Library : " & RVinfo.Library)
            ed.WriteMessage(vbLf & "LogPath : " & RVinfo.LogPath)
            ed.WriteMessage(vbLf & "PERFolder : " & RVinfo.PERFolder)
            ed.WriteMessage(vbLf & "PalletteXML : " & RVinfo.PalletteXML)
            ed.WriteMessage(vbLf & "PlotLogFilePath : " & RVinfo.PlotLogFilePath)
            ed.WriteMessage(vbLf & "Plotters : " & RVinfo.Plotters)
            ed.WriteMessage(vbLf & "PluginPath : " & RVinfo.PluginPath)
            ed.WriteMessage(vbLf & "PluginPersoXML : " & RVinfo.PluginPersoXML)
            ed.WriteMessage(vbLf & "PluginSharedXML : " & RVinfo.PluginSharedXML)
            ed.WriteMessage(vbLf & "SharedPath : " & RVinfo.SharedPath)
            ed.WriteMessage(vbLf & "SupportPath : " & RVinfo.SupportPath)
            ed.WriteMessage(vbLf & "SyncProfilAcad : " & RVinfo.SyncProfilAcad)
            ed.WriteMessage(vbLf & "SystemPath : " & RVinfo.SystemPath)
            ed.WriteMessage(vbLf & "Template : " & RVinfo.Template)
            ed.WriteMessage(vbLf & "TemplateVersion : " & RVinfo.TemplateVersion)
            ed.WriteMessage(vbLf & "ToolPalettePath : " & RVinfo.ToolPalettePath)
            ed.WriteMessage(vbLf & "XMLflow : " & RVinfo.XMLflow)
            ed.WriteMessage(vbLf & "XMLformat : " & RVinfo.XMLformatPerso)
            ed.WriteMessage(vbLf & "urlDEV : " & RVinfo.urlDEV)
            ed.WriteMessage(vbLf & "urlUpdate : " & RVinfo.urlUpdate)
            ed.WriteMessage(vbLf & "xCompany : " & RVinfo.xCompany)
            ed.WriteMessage(vbLf & "xCopyright : " & RVinfo.xCopyright)
            ed.WriteMessage(vbLf & "xDescription : " & RVinfo.xDescription)
            ed.WriteMessage(vbLf & "xProduct : " & RVinfo.xProduct)
            ed.WriteMessage(vbLf & "xTitle : " & RVinfo.xTitle)
            ed.WriteMessage(vbLf & "xVersion : " & RVinfo.xVersion)
            'ed.WriteMessage(vbLf & "VKEY : " & RVinfo.VKEY)


            ' List all the layouts in the current drawing

            ' Get the current document and database
            Dim acDocX As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDocX.Database

            ' Get the layout dictionary of the current database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim lays As DBDictionary = _
                    acTrans.GetObject(acCurDb.LayoutDictionaryId, OpenMode.ForRead)

                acDocX.Editor.WriteMessage(vbLf & "Layouts:")

                ' Step through and list each named layout and Model
                For Each item As DBDictionaryEntry In lays
                    acDocX.Editor.WriteMessage(vbLf & "  " & item.Key & " : " & item.Value.Database.Unitmode.ToString())

                Next


                ' Abort the changes to the database
                acTrans.Abort()
            End Using


        End Sub


        '<CommandMethod(NomCmd & "Plot")> _
        Public Sub RevoPlot()

            PlotCurrentLayout("Autodesk-PDF.pc3", "ISO A3 (420.00 x 297.00 mm)", "1000", "1", "Autodesk-CH.ctb", True, False)

            'PlotCurrentLayout("", "", "1000", "1", True, False)

            'Dim rvsflux As New RevoScript
            'rvsflux.cmdCmd("#Cmd;_SCALE|_all||0,0|0.001|_zoom|ET|_.MSPACE|_AttSync|_N|*|_.PSPACE")

        End Sub
        <CommandMethod(NomCmd & "Annotatif")> _
        Public Sub RevoAnnotatif()

            Dim Val As Double = MsgBox("Souhaitez-vous ajouter le mode annnotatif ?" & vbCrLf & vbCrLf & _
                                      "Oui : Ajout de l'annotativité" & vbCrLf & _
                                      "Non : Suppression de l'annotativité" & vbCrLf & _
                                      "(Annuler pour quitter la fonction)", MsgBoxStyle.Information + MsgBoxStyle.YesNoCancel, "Mode annotatif")


            Dim RVscript As New Revo.RevoScript
            If Val = 6 Then

                RVscript.cmdOBJ("#OBJ;*;Hatch;[[Annotative]]1;") 'Ajout Annotation Hachures
                MOAnnotatif(True) ' Ajoute
                ChangeScale("1000")
            ElseIf Val = 7 Then

                RVscript.cmdOBJ("#OBJ;*;Hatch;[[Annotative]]0;") 'Supp. Annotation Hachures
                RVscript.cmdCmd("#Cmd;_CANNOSCALE|1:25000|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:10000|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:5000|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:2500|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:2000|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:1000|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:500|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:250|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:200|_AIOBJECTSCALEREMOVE|_ALL||_CANNOSCALE|1:100|_AIOBJECTSCALEREMOVE|_ALL||")
                RVscript.cmdCmd("#Cmd;ANNOAUTOSCALE|4|_CANNOSCALE|1:1|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1")

                '_AIOBJECTSCALEREMOVE|_ALL||
                MOAnnotatif(False) ' Supprime
                ChangeScale("1000")
            End If

        End Sub
        <CommandMethod(NomCmd & "Ech100")> _
        Public Sub RevoEch100()
            ChangeScale("100")
        End Sub
        <CommandMethod(NomCmd & "Ech200")> _
        Public Sub RevoEch200()
            ChangeScale("200")
        End Sub
        <CommandMethod(NomCmd & "Ech250")> _
        Public Sub RevoEch250()
            ChangeScale("250")
        End Sub
        <CommandMethod(NomCmd & "Ech500")> _
        Public Sub RevoEch500()
            ChangeScale("500")
        End Sub
        <CommandMethod(NomCmd & "Ech1000")> _
        Public Sub RevoEch1000()
            ChangeScale("1000")
        End Sub
        <CommandMethod(NomCmd & "Ech2000")> _
        Public Sub RevoEch2000()
            ChangeScale("2000")
        End Sub
        <CommandMethod(NomCmd & "Ech2500")> _
        Public Sub RevoEch2500()
            ChangeScale("2500")
        End Sub
        <CommandMethod(NomCmd & "Ech4000")> _
        Public Sub RevoEch4000()
            ChangeScale("4000")
        End Sub
        <CommandMethod(NomCmd & "Ech5000")> _
        Public Sub RevoEch5000()
            ChangeScale("5000")
        End Sub
        <CommandMethod(NomCmd & "Ech10000")> _
        Public Sub RevoEch10000()
            ChangeScale("10000")
        End Sub




    End Class

End Namespace