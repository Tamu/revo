
Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text

Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Windows
Imports Autodesk.Windows
Imports Autodesk.AutoCAD.EditorInput

Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.IO
Imports System.Windows.Media.Imaging
Imports System.Reflection
Imports Autodesk.AutoCAD.ApplicationServices

<Assembly: ExtensionApplication(GetType(Revo.AdskApplication))> 
<Assembly: CommandClass(GetType(Revo.AdskApplication))> 

Namespace Revo

    Public Class AdskApplication
        Implements IExtensionApplication
        Dim Ass As New Revo.RevoInfo
        Dim ribBtnRevo As Autodesk.Windows.RibbonSplitButton = New RibbonSplitButton()
        Dim LoadedPluginRibbon As Boolean = False
        Dim OkRibbonState As Boolean = False
        Dim OkOnIdle As Boolean = False
      
        Overridable Sub Initialize() Implements IExtensionApplication.Initialize

            ' We defer the creation of our Application Menu to when
            ' the menu is next accessed
            'AddHandler ComponentManager.ApplicationMenu.Opening, AddressOf ApplicationMenu_Opening
            Try

                Application.SetSystemVariable("DBLCLKEDIT", 0)
                AddHandler Autodesk.AutoCAD.ApplicationServices.Application.BeginDoubleClick, AddressOf Application_DblClick

                'surveille commande
                Dim acAppCom As Autodesk.AutoCAD.Interop.AcadApplication 'AcadApplication
                acAppCom = Application.AcadApplication
                AddHandler acAppCom.EndCommand, AddressOf EndCmd


                If Autodesk.Windows.ComponentManager.Ribbon Is Nothing Then
                    'load the custom Ribbon on startup, but at this point 
                    'the Ribbon control is not available, so register for 
                    'an event and wait 
                    AddHandler Autodesk.Windows.ComponentManager.ItemInitialized, AddressOf ComponentManager_ItemInitialized


                    ' We defer the creation of our Quick Access Toolbar item
                    ' to when the application is next idle
                    AddHandler Application.Idle, AddressOf Application_OnIdle

                    'Check switching workspace
                    'not in published sample
                    'Autodesk.AutoCAD.ApplicationServices.Application.SystemVariableChanged += New Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventHandler(Application_SystemVariableChanged)
                    AddHandler Application.SystemVariableChanged, AddressOf Application_SystemVariableChanged

                Else

                    'the assembly was loaded using NETLOAD, so the ribbon is available and we just create the ribbon
                    'Your command for creating your ribbon 'createRibbon();'
                    'createRibbon()

                    'Autodesk.AutoCAD.ApplicationServices.Application.SystemVariableChanged += New Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventHandler(Application_SystemVariableChanged)
                    AddHandler Application.SystemVariableChanged, AddressOf Application_SystemVariableChanged

                End If

                'initialisation de Handler pour detecter la fermeture de ACAD
                AddHandler Application.BeginQuit, AddressOf CloseFunction

            Catch

            End Try


        End Sub

        Private Sub ComponentManager_ItemInitialized(ByVal sender As Object, ByVal e As RibbonItemEventArgs)

            Try
                RemoveHandler Autodesk.Windows.ComponentManager.ItemInitialized, AddressOf ComponentManager_ItemInitialized
                createRibbon()
                createQuickStart()
                ' UpdateSpaceShared()
            Catch
            End Try

        End Sub
        Private Sub EndCmd(ByVal cmds As String)


            If cmds = "MSPACE" Then 'Entré espace papier

                'Entré espace papier
                '  MsgBox("Cmd : " & cmds)

            ElseIf cmds = "PSPACE" Then 'Sortie espace papier

                If Revo.MyCommands.ActiveEventsCmd = True Then

                    'Contrôle si Quadri est existant dans la fênetre
                    Dim QuadriColl As New Collection
                    QuadriColl = CheckExistQuadri()
                    If QuadriColl.Count = 1 Then ' Mise à jour
                        Dim MyCmd As New Revo.MyCommands
                        MyCmd.RevoQuadrillage()
                        ' Revo.MyCommands.RevoQuadrillage()
                    End If

                End If

            End If


        End Sub



        Private Sub Application_DblClick(ByVal sender As Object, ByVal e As BeginDoubleClickEventArgs)

            Application.SetSystemVariable("DBLCLKEDIT", 1)

            Dim MyDwg As Document = Application.DocumentManager.MdiActiveDocument
            Dim Myed As Editor = MyDwg.Editor
            Dim ppr As PromptSelectionResult = Myed.SelectImplied
            Dim ss As SelectionSet = ppr.Value
            If ss IsNot Nothing Then
                If ss.Count = 1 Then
                    Using tr As Autodesk.AutoCAD.DatabaseServices.Transaction = MyDwg.TransactionManager.StartTransaction
                        Try
                            Dim Myentity As Autodesk.AutoCAD.DatabaseServices.Entity = TryCast(tr.GetObject(ss(0).ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), Autodesk.AutoCAD.DatabaseServices.Entity)
                            If Myentity IsNot Nothing Then
                                ' MsgBox(Myentity.Layer.ToString)

                                If TypeName(Myentity) = "BlockReference" Then
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                                    Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                                    Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
                                    Dim objBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = acDoc.HandleToObject(Myentity.Handle.ToString)

                                    ' Si Block REVO Editable
                                    If objBlock.HasAttributes Then
                                        Dim EditPt As Double = -1

                                        Dim StrBlockValide() As String = Split("IMPL_B_PTS_CONTROL%IMPL_B_PTS_PROJET%IMPL_PTS_CONTROL%IMPL_PTS_PROJET%MUT_PFP3_BORNE%MUT_PFP3_CHEVILLE%MUT_PFP3_CROIX%MUT_PFP3_PIEU%MUT_PTS_BF_BORNE%MUT_PTS_BF_CHEVILLE%MUT_PTS_BF_CROIX%MUT_PTS_BF_NON_MATE_DEFINI%MUT_PTS_BF_NON_MATE_NONDEFINI%MUT_PTS_BF_PIEU%MUT_PTS_COM_BORNE%MUT_PTS_COM_CHEVILLE%MUT_PTS_COM_CROIX%MUT_PTS_COM_NON_MATE_DEFINI%MUT_PTS_COM_NON_MATE_NONDEFINI%MUT_PTS_COM_PIEU%MUT_PTS_CS%MUT_PTS_OD%PFA1%PFA2%PFA3%PFP1_INACCESSIBLE%PFP1_STATIONNABLE%PFP2_INACCESSIBLE%PFP2_STATIONNABLE%PFP3_BORNE%PFP3_CHEVILLE%PFP3_CROIX%PFP3_PIEU%PTS_BF_BORNE%PTS_BF_CHEVILLE%PTS_BF_CROIX%PTS_BF_NON_MATE_DEFINI%PTS_BF_NON_MATE_NONDEFINI%PTS_BF_PIEU%PTS_COM_BORNE%PTS_COM_CHEVILLE%PTS_COM_CROIX%PTS_COM_NON_MATE_DEFINI%PTS_COM_NON_MATE_NONDEFINI%PTS_COM_PIEU%PTS_CS%PTS_OD%SOUT_B_EC_CE%SOUT_B_EC_CH%SOUT_B_EC_CU%SOUT_B_EC_DE%SOUT_B_EC_EX%SOUT_B_EC_GR%SOUT_B_EC_GU%SOUT_B_EC_OS%SOUT_B_EC_PR%SOUT_B_EC_RE%SOUT_B_EC_SE%SOUT_B_EG_CB%SOUT_B_EG_CD%SOUT_B_EG_CE%SOUT_B_EG_PR%SOUT_B_EL_CD%SOUT_B_EL_POT%SOUT_B_EL_PR%SOUT_B_EM_CE%SOUT_B_EM_CH%SOUT_B_EM_OS%SOUT_B_EM_PR%SOUT_B_EP_BH%SOUT_B_EP_CC%SOUT_B_EP_CO%SOUT_B_EP_CVA%SOUT_B_EP_FO%SOUT_B_EP_MA%SOUT_B_EP_PR%SOUT_B_EP_PU%SOUT_B_EP_RE%SOUT_B_EP_VA%SOUT_B_EU_CE%SOUT_B_EU_CH%SOUT_B_EU_FS%SOUT_B_EU_OS%SOUT_B_EU_PR%SOUT_B_EU_RE%SOUT_B_EU_SE%SOUT_B_GA_MA%SOUT_B_GA_PR%SOUT_B_GA_RE%SOUT_B_GA_VA%SOUT_B_PROJ_EC_CE%SOUT_B_PROJ_EC_CH%SOUT_B_PROJ_EC_CU%SOUT_B_PROJ_EC_DE%SOUT_B_PROJ_EC_EX%SOUT_B_PROJ_EC_GR%SOUT_B_PROJ_EC_GU%SOUT_B_PROJ_EC_OS%SOUT_B_PROJ_EC_PR%SOUT_B_PROJ_EC_RE%SOUT_B_PROJ_EC_SE%SOUT_B_PROJ_EP_BH%SOUT_B_PROJ_EP_CC%SOUT_B_PROJ_EP_CO%SOUT_B_PROJ_EP_CVA%SOUT_B_PROJ_EP_FO%SOUT_B_PROJ_EP_PR%SOUT_B_PROJ_EP_PU%SOUT_B_PROJ_EP_RE%SOUT_B_PROJ_EP_VA%SOUT_B_PROJ_EU_CE%SOUT_B_PROJ_EU_CH%SOUT_B_PROJ_EU_FS%SOUT_B_PROJ_EU_OS%SOUT_B_PROJ_EU_PR%SOUT_B_PROJ_EU_RE%SOUT_B_PROJ_EU_SE%SOUT_B_TT_DTT%SOUT_B_TT_POT%SOUT_B_TT_PR%SOUT_B_TT_RTT%SOUT_B_TV_ART%SOUT_B_TV_ATT%SOUT_B_TV_CT%SOUT_B_TV_PR%TOPO_B_ARBRE%TOPO_B_PTS%TOPO_PTS", "%")
                                        For i = 0 To StrBlockValide.Length - 1
                                            If StrBlockValide(i) = objBlock.EffectiveName.ToUpper Then
                                                EditPt = 1
                                                Exit For
                                            End If
                                        Next

                                        If EditPt <> 1 Then
                                            Dim StrBlockValideTxt() As String = Split("B_CS_BAT%B_CS_PROJ_BAT%B_CS_NOM%B_CS_PROJ_NOM%B_CS_SENS%B_CS_PROJ_SENS%B_ODL_BAT%B_ODL_COUV%B_ODL_NOM%B_ODL_SENS%B_ODL_BAC%B_NO_LIEU%B_NO_LIEUDIT%B_NO_LOCAL%IMMEUBLE_NUM%IMMEUBLE_NUM_PROJ%IMMEUBLE_NUM_DDP%IMMEUBLE_NUM_DDP_PROJ%NOM_COMMUNE%NOM_COMMUNE_PROJ%Nom_district%Nom_canton%Nom_pays%PLAN%NT%AD%NORD%MUT_B_CS_BAT%MUT_B_CS_NOM%MUT_B_CS_SENS%MUT_B_ODL_BAT%MUT_B_ODL_COUV%MUT_B_ODL_NOM%MUT_B_ODL_SENS%MUT_B_ODL_BAC%MUT_B_NO_LIEU%MUT_B_NO_LIEUDIT%MUT_B_NO_LOCAL%MUT_IMMEUBLE_NUM%MUT_IMMEUBLE_NUM_DDP%MUT_NOM_COMMUNE%MUT_Nom_district%MUT_Nom_canton%MUT_Nom_pays%MUT_PLAN%MUT_NT%MUT_AD%CHECK_ITF_PFP%CHECK_ITF_CS%CHECK_ITF_OD%CHECK_ITF_NOM%CHECK_ITF_BF%CHECK_ITF_DN%CHECK_ITF_LC%CHECK_ITF_RP%CHECK_ITF_AD", "%")
                                            For i = 0 To StrBlockValideTxt.Length - 1
                                                If StrBlockValideTxt(i) = objBlock.EffectiveName.ToUpper Then
                                                    EditPt = 2
                                                    Exit For
                                                End If
                                            Next
                                        End If

                                        If EditPt = 1 Then
                                            Application.SetSystemVariable("DBLCLKEDIT", 0)
                                            EditPointBlock(objBlock)

                                        ElseIf EditPt = 2 Then
                                            Application.SetSystemVariable("PICKFIRST", 0)
                                            Application.SetSystemVariable("DBLCLKEDIT", 0)
                                            EditTxtBlock(objBlock)
                                            Application.SetSystemVariable("PICKFIRST", 1)
                                        End If

                                    End If

                                Else

                                End If

                                ' ss = Nothing                      ' need these 2 lines to remove grips after double click finished
                                ' Myed.SetImpliedSelection(ss)
                            End If

                        Catch ex As System.Exception
                            MsgBox(ex.Message)
                        End Try
                    End Using
                End If
            End If


        End Sub
        'Private Sub ApplicationMenu_Opening(ByVal sender As Object, ByVal e As EventArgs)

        ' Remove the event when it is fired
        '  RemoveHandler ComponentManager.ApplicationMenu.Opening, AddressOf ApplicationMenu_Opening

        ' Add our Application Menu
        'AddApplicationMenu()
        '  createRibbon()
        'createQuickStart()
        ' MsgBox("ApplicationMenu_Opening")

        'End Sub

        Private Sub Application_OnIdle(ByVal sender As Object, ByVal e As EventArgs)

            Try
                ' Remove the event when it is fired
                RemoveHandler Application.Idle, AddressOf Application_OnIdle

                If OkOnIdle = False And OkRibbonState = False Then OkRibbonState = True
                OkOnIdle = True

                If OkRibbonState = True And LoadedPluginRibbon = False Then
                    createQuickStart() ' Add our Quick Access Toolbar item
                    createRibbon() ' Add Ribbon
                    LoadedPluginRibbon = True
                    ' UpdateSpaceShared()
                End If

            Catch

            End Try

        End Sub

        Private Sub Application_SystemVariableChanged(ByVal sender As Object, ByVal e As Autodesk.AutoCAD.ApplicationServices.SystemVariableChangedEventArgs)
            'Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            'ed.WriteMessage("\n-Sys Var Changed: " + e.Name);

            ' MsgBox("Application_SystemVariableChanged : " & e.Name)
            Try
                If e.Name = "RIBBONSTATE" Then
                    If OkOnIdle = True And OkRibbonState = False And LoadedPluginRibbon = False Then
                        createRibbon() ' Add Ribbon
                        createQuickStart() ' Add our Quick Access Toolbar item
                        LoadedPluginRibbon = True
                    End If
                    OkRibbonState = True
                End If

                If OkRibbonState = True And OkOnIdle = True Then
                    If e.Name.ToLower() = "wscurrent" Then
                        createRibbon()
                        createQuickStart()
                    End If
                End If

            Catch

            End Try
        End Sub


        Public Sub createQuickStart()

            Try
                'Quick Access Toolbar
                Dim qat As Autodesk.Windows.ToolBars.QuickAccessToolBarSource = ComponentManager.QuickAccessToolBar
                If qat IsNot Nothing Then
                    Dim ActiveAddBtn As Boolean = True
                    If qat.Items.Count <> 0 Then
                        For i = 0 To qat.Items.Count - 1
                            If qat.Items(i).Id.ToString = MY_QUICKSTART_ID Then
                                ActiveAddBtn = False
                            End If
                        Next

                        If ActiveAddBtn = True Then
                            If ribBtnRevo.Id Is Nothing Then CreateBtnRevo()
                            qat.AddStandardItem(ribBtnRevo)

                        Else
                            Dim msg As String = (ribBtnRevo.Description)
                        End If
                    Else
                        'MsgBox(qat.Items.Count)
                    End If
                    '   MsgBox(qat.Items.Count)
                    'MsgBox(qat.Items.Item(qat.Items.Count - 1).Id)
                Else
                    ' MsgBox("Lancer Quick ")
                End If
            Catch 'ex As System.Exception
                'erreur de chargement
                ' MsgBox("ERREUR : " & ex.Message)
                MsgBox("Pour charger le menu saisissez dans la barre de commande:  " & Ass.xProduct.ToUpper & "MENU  puis valider.", vbInformation, "Impossible de charger le menu " & Ass.xTitle)

            End Try
        End Sub

        Overridable Sub Terminate() Implements IExtensionApplication.Terminate

            'Lancement en cloture de AutoCAD

        End Sub
        Public Sub CloseFunction() 'Opération à la fermeture d'AutoCAD


            'Nettoyage des données temporaires
            Try
                Dim urlXMLFlow As String = Ass.XMLflow
                Dim urlDB3interlis As String = System.IO.Path.Combine(Ass.SystemPath, Ass.DB3interlis)
                Dim urlDB3System As String = System.IO.Path.Combine(Ass.SystemPath, Ass.DB3system)


                'Supprime le flux de travail revo-flow.xml (XMLflow)
                If File.Exists(urlXMLFlow) Then System.IO.File.Delete(Ass.XMLflow)

                'Déconexion du la BD interlis
                CloseDB()

                'Supprime la base de donnée interlis.db3 (DB3interlis)
                If File.Exists(urlDB3interlis) Then System.IO.File.Delete(urlDB3interlis)

                'Supprime la base de donnée System.db3 (DB3System)
                If File.Exists(urlDB3System) Then System.IO.File.Delete(urlDB3System)


            Catch ex As System.Exception
                MsgBox(ex.Message & vbCrLf & "(" & ex.InnerException.Message.ToString & ")", vbInformation, "Erreur de nettoyage des fichiers " & Ass.xTitle)
            End Try



            'Initialisation des dossier Support
            Dim oApp As Autodesk.AutoCAD.Interop.AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication
            'Ajout du support (avec les types d'hachures)
            Dim SupportPaths As String = oApp.Preferences.Files.SupportPath
            Dim RevoSupportPath As String = Ass.SupportPath
            If Right(RevoSupportPath, 1) = "\" Then RevoSupportPath = Left(RevoSupportPath, Len(RevoSupportPath) - 1)

            If InStr(SupportPaths.ToLower, RevoSupportPath.ToLower) = 0 Then
                oApp.Preferences.Files.SupportPath = SupportPaths & ";" & RevoSupportPath
            End If




            'Mise à jour du profil AutoCAD ------------------------------------------------------------------------
            Dim Conn As New Revo.connect
            If CDbl(Ass.SyncProfilAcad) = 1 Then Conn.RevoUpdateProfile()



        End Sub


        Function CheckUpdate() 'Check version online

            Dim UpdateYes As Double = 0 ' -1 pas de besoin de mis à jour, 0 pas de connexion , < 0 num lic.

            'Check si config est OK ' Mise à jour du templates
            If InStr(Ass.Template, "Revo10.dwt") <> 0 Or InStr(Ass.Template, "Revo11.dwt") <> 0 _
                Or InStr(Ass.Template, "Revo12.dwt") Then
                Dim Connect As New Revo.connect
                Dim NouvChemin As String = ""
                NouvChemin = Replace(Ass.Template, "Revo10.dwt", "Revo13.dwt")
                NouvChemin = Replace(Ass.Template, "Revo11.dwt", "Revo13.dwt")
                NouvChemin = Replace(Ass.Template, "Revo12.dwt", "Revo13.dwt")
                If File.Exists(NouvChemin) Then
                    Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "template", NouvChemin)
                End If

            End If '<template>K:\acad\Revo\Template\Revo11.dwt</template>


            'Si un Acad ouvert
            If (Revo.RevoFiler.CheckRunApp("acad")) <= 1 Then

                ' MsgBox("recherche mise à jour")

                'Intérogation version ' 0.9.2.0 = 920
                Dim VersLocal As Double = Trim(Replace(Ass.xVersion, ".", ""))
                Dim XMLversion As New RevoXML(Ass.urlUpdate)
                Dim VersUp As String = XMLversion.getElementValue("/update/version")
                Dim UrlUp As String = XMLversion.getElementValue("/update/file")

                If IsPlatformX64() Then UrlUp = XMLversion.getElementValue("/update/filex64")

                Dim numVersUp As Double = 0
                If IsNumeric(VersUp) Then numVersUp = CDbl(VersUp) : UpdateYes = numVersUp



                'Mise à jour de REVO ------------------------------------------------------------------------

                If VersLocal < numVersUp Then
                    ' MsgBox("Après avoir validé ce message, installé la mise à jour en cliquant sur le bouton" & vbCrLf & " 'Exécuter' (ou 'Enregistrer' puis lancer le fichier)", MsgBoxStyle.Information, "Mise à jour disponible")
                    If UrlUp <> "" Then

                        Dim REMOTE_URL As String = UrlUp ' adresse de la page ou du fichier à récuperer
                        Dim F As Integer = FreeFile()
                        Dim WEB_CLIENT As New System.Net.WebClient()
                        WEB_CLIENT.Credentials = Net.CredentialCache.DefaultCredentials 'Sécurité Windows !!!
                        Dim FileName1 As String = Path.Combine(System.IO.Path.GetTempPath, Ass.xProduct & "-" & numVersUp & ".exe")

                        Try
                            conn.Message("Mise à jour en cours", "La mise à jour des plugins est en cours de téléchargement", False, 20, 100)
                            WEB_CLIENT.DownloadFile(REMOTE_URL, FileName1)
                            conn.Message("Mise à jour en cours", "La mise à jour des plugins est en cours de téléchargement", False, 99, 100)
                            conn.Message("Mise à jour en cours", "", True, 99, 100)

                            System.Diagnostics.Process.Start(FileName1)
                        Catch ex As System.Exception
                            conn.RevoLog(conn.DateLog & "Update" & vbTab & "False" & vbTab & "Error update: " & ex.Message)
                        End Try
                    End If
                    UpdateYes = numVersUp ' Mise à jour effectué à la version X
                Else
                    '  MsgBox("pas de mise à jour : " & VersLocal & " = " & numVersUp)
                    If numVersUp <> 0 Then UpdateYes = -1 ' Pas besoins de mise à jour

                End If
                'Si Oui 
                'Recherche mise à jour

            End If

            Return UpdateYes

        End Function

        Public Shared Function IsPlatformX64() As Boolean
            Dim fOk As Boolean = False
            If IntPtr.Size = 8 Then
                fOk = True
            End If
            Return fOk
        End Function


        Private MY_TAB_ID As [String] = Ass.xProduct.ToUpper & "_TAB_ID"
        Private MY_QUICKSTART_ID As [String] = Ass.xProduct.ToUpper & "_QUICKSTARTMINI_ID"

        'RevoAddTab
        Public Sub createRibbon()
            Try

                Dim ribCntrl As Autodesk.Windows.RibbonControl = Autodesk.AutoCAD.Ribbon.RibbonServices.RibbonPaletteSet.RibbonControl
                Dim ExistTabPlugin As Boolean = False

                'Recherche Si Tab existant
                'find the custom tab using the Id
                Dim NbreTab As Double = ribCntrl.Tabs.Count - 1
                For i As Double = 0 To NbreTab
                    If ribCntrl.Tabs(i).Id.Equals(MY_TAB_ID) Then
                        ExistTabPlugin = True
                        Exit For
                    End If
                Next


                If ExistTabPlugin = False Then 'If don't exist add Ribbon

                    'can also be Autodesk.Windows.ComponentManager.Ribbon; 

                    '   For Each TabX As Autodesk.Windows.RibbonTab In ribCntrl.Tabs
                    '  If TabX.IsPanelEnabled Then

                    'MsgBox(TabX.Id)

                    ' End If

                    'Next

                    'add the tab
                    Dim ribTab As New RibbonTab()
                    ribTab.Title = Ass.xTitle
                    ribTab.Id = MY_TAB_ID
                    ribCntrl.Tabs.Add(ribTab)
                    'ribTab.Theme = Autodesk.Internal.Windows.TabThemes.Purple

                    'create and add both panels 
                    'addPanel1(ribTab)
                    'addPanel2(ribTab)
                    addPanelGeneral(ribTab)
                    'addPanelQuickLaunch(ribTab)
                    If Ass.MenuType.ToUpper Like "REVO*" Then addPanelGeomVD(ribTab) 'Désactiver pour les package non REVO
                    addPanelOptions(ribTab)

                    'Dim RibCollection As New RibbonItemCollection
                    'set as active tab 
                    'ribTab.IsActive = True
                    'For Each RibbAcad As RibbonItem In RibCollection
                    ' MsgBox(RibbAcad.ToString)
                    ' Next

                End If

            Catch 'ex As NullReferenceException
                '  MsgBox("Pour charger le menu saisissez dans la barre de commande:  " & Ass.xProduct.ToUpper & "MENU  puis valider.", MsgBoxStyle.Information, "Impossible de charger le menu " & Ass.xTitle)
                '& vbCrLf & "( err:" & ex.Message & ")"
            End Try

        End Sub
        Private Sub CreateBtnRevo()

            ribBtnRevo.Id = MY_QUICKSTART_ID
            ribBtnRevo.Text = "Liste des flux"
            ribBtnRevo.Tag = "Liste des flux"
            ribBtnRevo.Size = RibbonItemSize.Standard
            ribBtnRevo.ShowText = True
            ribBtnRevo.ListImageSize = RibbonImageSize.Standard


            Dim ListFichierXML As New List(Of String)
            'Chargement des fichier script CSV : revo-perso.xml à distance (lecture seul)
            If Ass.PluginSharedXML(True).ToUpper <> Ass.PluginSharedXML(False).ToUpper And File.Exists(Ass.PluginSharedXML) Then
                Dim xcld As New RevoXML(Ass.PluginSharedXML)
                For i = 1 To xcld.countElements("/" & Ass.xProduct.ToLower & "/cmdFiles/url")
                    ListFichierXML.Add(xcld.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url", i))
                Next
            End If
            'Chargement des fichier script CSV : revo-perso.xml en local
            If File.Exists(Ass.PluginPersoXML) Then
                Dim xloc As New RevoXML(Ass.PluginPersoXML)
                For i = 1 To xloc.countElements("/" & Ass.xProduct.ToLower & "/cmdFiles/url")
                    ListFichierXML.Add(xloc.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url", i))
                Next
            End If

            'create a Ribbon List Button 
            ' Dim ribListBtn As Autodesk.Windows.RibbonSplitButton = New RibbonSplitButton()
            Try
                For Each FichierXML As String In ListFichierXML
                    ' For i = 1 To x.countElements("/" & Ass.xProduct.ToLower & "/cmdFiles/url")

                    Dim FileName As String = FichierXML '  x.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url", i)
                    If FileName <> "" Then
                        Dim NameFile As String = Path.GetFileName(FileName)
                        Dim NamePath As String = Ass.ActionsPath
                        Dim Filetype As String = "*.CSV"
                        Dim NameInfo(), Nom(), Description(), Groupe() As String
                        If NamePath.ToUpper <> Path.GetDirectoryName(FileName).ToUpper & "\" Then
                            If InStr(FileName, "\") <> 0 Then
                                NamePath = Path.GetDirectoryName(FileName).ToUpper & "\"
                            End If
                        End If
                        FileName = Path.Combine(NamePath, NameFile)
                        If Microsoft.VisualBasic.Right(NameFile.ToUpper, 4) = ".CSV" Then
                            If File.Exists(FileName) Then
                                NameInfo = Revo.RevoFiler.NameDescScript(FileName)
                                If NameInfo(0) <> "ErrFile" Then

                                    Nom = Split(NameInfo(0), ";")
                                    Description = Split(NameInfo(1), ";")
                                    Groupe = Split(NameInfo(2), ";")
                                    If Nom.Length >= 2 Then Nom(0) = Nom(1)
                                    If Description.Length >= 2 Then Description(0) = Description(1)
                                    If Groupe.Length >= 2 Then Groupe(0) = Groupe(1)

                                    Dim ribButton1 As Autodesk.Windows.RibbonButton = New RibbonButton()
                                    ribButton1.Text = Nom(0) 'x.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url"
                                    ribButton1.Description = Description(0)
                                    ribButton1.ShowText = True
                                    'ribButton1.CommandParameter = String.Format("-" & Ass.xProduct & " {0}", Replace(NamePath & NameFile, " ", "?")) & vbCrLf
                                    ribButton1.CommandParameter = "-" & Ass.xProduct & " " & Replace(NamePath & NameFile, " ", "?") & vbCr
                                    ribButton1.CommandHandler = New AdskCommandHandler()
                                    ribButton1.ShowImage = True
                                    ribButton1.Size = RibbonItemSize.Standard
                                    ribButton1.Image = Images.getBitmap(My.Resources.ribbon_plug16)
                                    ribButton1.LargeImage = Images.getBitmap(My.Resources.ribbon_plug32)
                                    ribBtnRevo.Items.Add(ribButton1)

                                End If
                            End If
                        End If
                    End If
                Next


                'Ajouts de fonctions complémentaires
                Dim BtnREVO As Autodesk.Windows.RibbonButton = New RibbonButton()
                BtnREVO.Text = "Action multi-fichiers"
                BtnREVO.Description = "Traitement de flux sur plusieurs dessins en même temps."
                BtnREVO.ShowText = True
                BtnREVO.ShowImage = True
                BtnREVO.Size = RibbonItemSize.Standard
                BtnREVO.Orientation = Windows.Controls.Orientation.Vertical
                BtnREVO.Image = Images.getBitmap(My.Resources.ribbon_plug16)
                BtnREVO.LargeImage = Images.getBitmap(My.Resources.ribbon_plug32)
                BtnREVO.CommandParameter = "REVO "
                BtnREVO.CommandHandler = New AdskCommandHandler()
                ribBtnRevo.Items.Add(BtnREVO)






            Catch 'ex As System.Exception
                'MsgBox("Super")
            End Try

        End Sub


        Private Sub addPanelGeneral(ByVal ribTab As RibbonTab)
            'create the panel source 
            Dim ribSourcePanel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
            ribSourcePanel.Title = "Flux REVO"
            'now the panel 
            Dim ribPanel As New RibbonPanel()
            ribPanel.Source = ribSourcePanel
            ribTab.Panels.Add(ribPanel)


            'now add and expanded panel (with 1 button) 
            Dim ribExpPanel As Autodesk.Windows.RibbonRowPanel = New RibbonRowPanel()
            ribSourcePanel.Items.Add(ribExpPanel)

            Dim RevoMultiFiles As Autodesk.Windows.RibbonButton = New RibbonButton()

            'needs to be here 
            RevoMultiFiles.Text = "Action multi-fichiers"
            RevoMultiFiles.Description = "Traitement de flux sur plusieurs dessins en même temps."
            RevoMultiFiles.ShowText = True
            RevoMultiFiles.ShowImage = True
            RevoMultiFiles.Size = RibbonItemSize.Large
            RevoMultiFiles.Orientation = Windows.Controls.Orientation.Vertical
            RevoMultiFiles.Width = 400
            RevoMultiFiles.Image = Images.getBitmap(My.Resources.ribbon_plug16)
            RevoMultiFiles.LargeImage = Images.getBitmap(My.Resources.ribbon_plug32)
            RevoMultiFiles.CommandParameter = MyCommands.NomCmd & " "
            RevoMultiFiles.CommandHandler = New AdskCommandHandler()
            ribExpPanel.Items.Add(RevoMultiFiles)


            Dim ListFichierXML As New List(Of String)
            'Chargement des fichier script CSV : revo-perso.xml à distance (lecture seul)
            If Ass.PluginSharedXML(True).ToUpper <> Ass.PluginSharedXML(False).ToUpper And File.Exists(Ass.PluginSharedXML) Then
                Dim xcld As New RevoXML(Ass.PluginSharedXML)
                For i = 1 To xcld.countElements("/" & Ass.xProduct.ToLower & "/cmdFiles/url")
                    ListFichierXML.Add(xcld.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url", i))
                Next
            End If
            'Chargement des fichier script CSV : revo-perso.xml en local
            If File.Exists(Ass.PluginPersoXML) Then
                Dim xloc As New RevoXML(Ass.PluginPersoXML)
                For i = 1 To xloc.countElements("/" & Ass.xProduct.ToLower & "/cmdFiles/url")
                    ListFichierXML.Add(xloc.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url", i))
                Next
            End If


            'create a Ribbon List Button 
            Dim ribListBtn As Autodesk.Windows.RibbonSplitButton = New RibbonSplitButton()
            ribListBtn.ListImageSize = RibbonImageSize.Standard
            Try
                Dim Nbre As Double = 0
                For Each FichierXML As String In ListFichierXML '
                    'For i = 1 To x.countElements("/" & Ass.xProduct.ToLower & "/cmdFiles/url")

                    Dim FileName As String = FichierXML 'x.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url", i)
                    If FileName <> "" Then
                        Dim NameFile As String = Path.GetFileName(FileName)
                        Dim NamePath As String = Ass.ActionsPath
                        Dim Filetype As String = "*.CSV"
                        Dim NameInfo(), Nom(), Description(), Groupe() As String
                        If NamePath.ToUpper <> Path.GetDirectoryName(FileName).ToUpper & "\" Then
                            If InStr(FileName, "\") <> 0 Then
                                NamePath = Path.GetDirectoryName(FileName).ToUpper & "\"
                            End If
                        End If
                        FileName = Path.Combine(NamePath, NameFile)
                        If Microsoft.VisualBasic.Right(NameFile.ToUpper, 4) = ".CSV" Then
                            If File.Exists(FileName) Then
                                NameInfo = Revo.RevoFiler.NameDescScript(FileName)
                                If NameInfo(0) <> "ErrFile" Then

                                    Nom = Split(NameInfo(0), ";")
                                    Description = Split(NameInfo(1), ";")
                                    Groupe = Split(NameInfo(2), ";")
                                    If Nom.Length >= 2 Then Nom(0) = Nom(1)
                                    If Description.Length >= 2 Then Description(0) = Description(1)
                                    If Groupe.Length >= 2 Then Groupe(0) = Groupe(1)

                                    Dim ribButton1 As Autodesk.Windows.RibbonButton = New RibbonButton()
                                    ribButton1.Text = Nom(0) 'x.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url"
                                    ribButton1.Description = Description(0)
                                    ribButton1.ShowText = True
                                    'ribButton1.CommandParameter = String.Format("-" & Ass.xProduct & " {0} ", Replace(NamePath & NameFile, " ", "?")) & vbCrLf
                                    ribButton1.CommandParameter = "-" & Ass.xProduct & " " & Replace(NamePath & NameFile, " ", "?") & vbCr
                                    ribButton1.CommandHandler = New AdskCommandHandler()
                                    ribButton1.ShowImage = True
                                    ribButton1.Size = RibbonItemSize.Large
                                    ribButton1.Image = Images.getBitmap(My.Resources.ribbon_plug16)
                                    ribButton1.LargeImage = Images.getBitmap(My.Resources.ribbon_plug32)
                                    ribListBtn.Items.Add(ribButton1)

                                End If
                            End If
                        End If
                    End If

                    Nbre += 1 'ajout de la boucle

                Next

                If Nbre > 0 Then

                    'create the panel source 
                    'Dim ribSourcePanel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
                    'ribSourcePanel.Title = "Lancement rapide"

                    'now the panel 
                    'Dim ribPanel As New RibbonPanel()
                    'ribPanel.Source = ribSourcePanel
                    'ribTab.Panels.Add(ribPanel)

                    'now add and expanded panel (with 1 button) 
                    'Dim ribExpPanel As Autodesk.Windows.RibbonRowPanel = New RibbonRowPanel()
                    'ribSourcePanel.Items.Add(ribExpPanel)

                    ribListBtn.Text = "Liste des scripts"
                    ribListBtn.Tag = "Liste des scripts"
                    ribListBtn.Size = RibbonItemSize.Large
                    ribListBtn.ShowText = True
                    ribExpPanel.Items.Add(ribListBtn)


                End If
            Catch 'ex As System.Exception
            End Try

            Dim item As New RibbonButton()
            item.IsCheckable = True
            item.CommandParameter = MyCommands.NomCmd & "small "
            item.Text = "Flux " & Ass.xProduct
            item.ShowImage = True
            item.ShowText = True
            item.CommandHandler = New AdskCommandHandler()
            ribSourcePanel.DialogLauncher() = item

        End Sub

        Private Sub addPanelGeomVD(ByVal ribTab As RibbonTab)

            'create the panel source 
            Dim ribSourcePanel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
            ribSourcePanel.Title = "Geomatics VD"

            'now the panel 
            Dim ribPanel As New RibbonPanel()
            ribPanel.Source = ribSourcePanel
            ribTab.Panels.Add(ribPanel)

            'now add and expanded panel (with 1 button) 
            Dim ribExpPanel As Autodesk.Windows.RibbonRowPanel = New RibbonRowPanel()
            ribSourcePanel.Items.Add(ribExpPanel)

            Dim ribExpPanel2 As Autodesk.Windows.RibbonRowPanel = New RibbonRowPanel()
            ribSourcePanel.Items.Add(ribExpPanel2)

            Dim ribListProj As Autodesk.Windows.RibbonSplitButton = New RibbonSplitButton()
            ribListProj.ListImageSize = RibbonImageSize.Standard

            'Bouton Create Interlis project
            Dim ribExpButton2 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton2.Text = "Créer un projet" & vbCrLf & "interlis VD"
            ribExpButton2.Description = "Créer un nouveau dessin au format interlis MO VD.01"
            ribExpButton2.ShowText = True
            ribExpButton2.ShowImage = True
            ribExpButton2.Size = RibbonItemSize.Large
            ribExpButton2.Orientation = Windows.Controls.Orientation.Vertical
            ribExpButton2.Image = Images.getBitmap(My.Resources.icon_itf)
            ribExpButton2.LargeImage = Images.getBitmap(My.Resources.icon_itf_nb)
            ribExpButton2.CommandParameter = MyCommands.NomCmd & "ProjMOVD "
            ribExpButton2.CommandHandler = New AdskCommandHandler()
            ribListProj.Items.Add(ribExpButton2)

            'Bouton Migration Interlis project
            Dim ribExpButton8 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton8.Text = "Migration" & vbCrLf & "de dessin"
            ribExpButton8.Description = "Migration de dessin au format interlis MO VD.01"
            ribExpButton8.ShowText = True
            ribExpButton8.ShowImage = True
            ribExpButton8.Size = RibbonItemSize.Large
            ribExpButton8.Orientation = Windows.Controls.Orientation.Vertical
            ribExpButton8.Image = Images.getBitmap(My.Resources.icon_itf)
            ribExpButton8.LargeImage = Images.getBitmap(My.Resources.icon_itf_nb)
            ribExpButton8.CommandParameter = MyCommands.NomCmd & "MigrationProject "
            ribExpButton8.CommandHandler = New AdskCommandHandler()
            ribListProj.Items.Add(ribExpButton8)

            ribListProj.Text = "Gestion des projets"
            ribListProj.Tag = "Gestion des projets"
            ribListProj.Size = RibbonItemSize.Large
            ribListProj.ShowText = True
            ribExpPanel.Items.Add(ribListProj)



            Dim ribListImportExport As Autodesk.Windows.RibbonSplitButton = New RibbonSplitButton()
            ribListImportExport.ListImageSize = RibbonImageSize.Standard

            'Bouton Import Interlis or data
            Dim ribExpButton7 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton7.Text = "Import" & vbCrLf & "interlis + pts"
            ribExpButton7.Description = "Importation d'un fichier format interlis MO VD.01"
            ribExpButton7.ShowText = True
            ribExpButton7.ShowImage = True
            ribExpButton7.Size = RibbonItemSize.Large
            ribExpButton7.Orientation = Windows.Controls.Orientation.Vertical
            ribExpButton7.Image = Images.getBitmap(My.Resources.icon_itf_nb)
            ribExpButton7.LargeImage = Images.getBitmap(My.Resources.icon_itf_nb)
            ribExpButton7.CommandParameter = "-" & Ass.xProduct & " " & _
                                                           "#Import;MD01MOVD;[[Space]];*;" & vbCr & _
                                                           "#Cmd;ANNOAUTOSCALE|4|_CANNOSCALE|1:500|_CANNOSCALE|1:1000|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1" & vbCr & _
                                                           "#Cmd;regen" & vbCr & vbCr
            '                                               "#cmd;_-Layout|_D|A3_Paysage" & vbCr & _
            ribExpButton7.CommandHandler = New AdskCommandHandler()
            ribListImportExport.Items.Add(ribExpButton7)

            'Bouton Export  data
            Dim ribExpButton15 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton15.Text = "Export points"
            ribExpButton15.Description = "Exportation de fichier points"
            ribExpButton15.ShowText = True
            ribExpButton15.ShowImage = True
            ribExpButton15.Size = RibbonItemSize.Large
            ribExpButton15.Orientation = Windows.Controls.Orientation.Vertical
            ribExpButton15.Image = Images.getBitmap(My.Resources.icon_pts)
            ribExpButton15.LargeImage = Images.getBitmap(My.Resources.icon_pts)
            ribExpButton15.CommandParameter = "-" & Ass.xProduct & " " & _
                                                           "#Export;" & vbCr & vbCr
            ribExpButton15.CommandHandler = New AdskCommandHandler()
            ribListImportExport.Items.Add(ribExpButton15)

            ribListImportExport.Text = "Import et export"
            ribListImportExport.Tag = "Import et export"
            ribListImportExport.Size = RibbonItemSize.Large
            ribListImportExport.ShowText = True
            ribExpPanel.Items.Add(ribListImportExport)


            'RevoInsertionPoints
            Dim ribExpButton4 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton4.Text = "Insertion" & vbCrLf & "de points"
            ribExpButton4.Description = "Insertion de points cadastral et topographique"
            ribExpButton4.ShowText = True
            ribExpButton4.ShowImage = True
            ribExpButton4.Size = RibbonItemSize.Large
            ribExpButton4.Orientation = Windows.Controls.Orientation.Vertical
            ribExpButton4.Image = Images.getBitmap(My.Resources.BtnInsererPoint)
            ribExpButton4.LargeImage = Images.getBitmap(My.Resources.BtnInsererPoint)
            ribExpButton4.CommandParameter = MyCommands.NomCmd & "InsertionPoints "
            ribExpButton4.CommandHandler = New AdskCommandHandler()
            ribExpPanel.Items.Add(ribExpButton4)



            Dim ribCustomBrk As Autodesk.Windows.RibbonRowBreak = New Autodesk.Windows.RibbonRowBreak

            'Nouveau groupe 1
            ' Dim ribInsert As New Autodesk.Windows.RibbonRowPanel

            'RevoInsertionTextes
            Dim ribExpButton5 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton5.Text = "Insertion" & vbCrLf & "de textes"
            ribExpButton5.Description = "Insertion de textes et symboles"
            ribExpButton5.ShowText = True
            ribExpButton5.ShowImage = True
            ribExpButton5.Size = RibbonItemSize.Large
            ribExpButton5.Orientation = Windows.Controls.Orientation.Vertical
            ribExpButton5.Image = Images.getBitmap(My.Resources.BtnInsererSymbole)
            ribExpButton5.LargeImage = Images.getBitmap(My.Resources.BtnInsererSymbole)
            ribExpButton5.CommandParameter = MyCommands.NomCmd & "InsertionTextes "
            ribExpButton5.CommandHandler = New AdskCommandHandler()
            ribExpPanel.Items.Add(ribExpButton5)

            'Nouveau groupe 1 bis
            Dim ribInsert0 As New Autodesk.Windows.RibbonRowPanel

            'Contour Etat descriptif technique (EDT)
            Dim ribExpButton3 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton3.Text = "EDT"
            ribExpButton3.Description = "Etat descriptif technique (EDT): Création d'un listing descriptif"
            ribExpButton3.ShowText = True
            ribExpButton3.ShowImage = True
            ribExpButton3.Size = RibbonItemSize.Standard
            ribExpButton3.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton3.Image = Images.getBitmap(My.Resources.BtnEDT)
            ribExpButton3.LargeImage = Images.getBitmap(My.Resources.BtnEDT)
            ribExpButton3.CommandParameter = MyCommands.NomCmd & "EDT2 "
            ribExpButton3.CommandHandler = New AdskCommandHandler()
            ribInsert0.Items.Add(ribExpButton3)


            ribInsert0.Items.Add(ribCustomBrk)

            'Contour PER
            Dim ribExpButton1 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton1.Text = "PER"
            ribExpButton1.Description = "Contour périmétrique (PER): Création d'un listing descriptif"
            ribExpButton1.ShowText = True
            ribExpButton1.ShowImage = True
            ribExpButton1.Size = RibbonItemSize.Standard
            ribExpButton1.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton1.Image = Images.getBitmap(My.Resources.BtnPER)
            ribExpButton1.LargeImage = Images.getBitmap(My.Resources.BtnPER)
            ribExpButton1.CommandParameter = MyCommands.NomCmd & "PER2 "
            ribExpButton1.CommandHandler = New AdskCommandHandler()
            ribInsert0.Items.Add(ribExpButton1)


            ribInsert0.Items.Add(ribCustomBrk)


            ' Nom des propriétaires (VD)
            Dim ribExpButton9 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton9.Text = "Prop. (VD)"
            ribExpButton9.Description = "Inscrit le nom des propriétaires (VD) sous le numéro de parcelle"
            ribExpButton9.ShowText = True
            ribExpButton9.ShowImage = True
            ribExpButton9.Size = RibbonItemSize.Standard
            ribExpButton9.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton9.Image = Images.getBitmap(My.Resources.rf16)
            ribExpButton9.LargeImage = Images.getBitmap(My.Resources.rf16)
            ribExpButton9.CommandParameter = MyCommands.NomCmd & "PropVD "
            ribExpButton9.CommandHandler = New AdskCommandHandler()
            ribInsert0.Items.Add(ribExpButton9)

            ribExpPanel2.Items.Add(ribInsert0)



            'Nouveau groupe 2
            Dim ribInsert2 As New Autodesk.Windows.RibbonRowPanel

            'Radiation d'objets MO...
            Dim ribExpButton10 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton10.Text = "Radiation d'objets MO"
            ribExpButton10.Description = "Radiation d'objets MO..."
            ribExpButton10.ShowText = True
            ribExpButton10.ShowImage = True
            ribExpButton10.Size = RibbonItemSize.Standard
            ribExpButton10.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton10.Image = Images.getBitmap(My.Resources.MOradier)
            'ribExpButton10.LargeImage = Images.getBitmap(My.Resources.MOradier)
            ribExpButton10.CommandParameter = MyCommands.NomCmd & "MOradier "
            ribExpButton10.CommandHandler = New AdskCommandHandler()
            ribInsert2.Items.Add(ribExpButton10)
            ribInsert2.Items.Add(ribCustomBrk)

            'Validation des mutations ouvertes
            Dim ribExpButton11 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton11.Text = "Validation des mutations"
            ribExpButton11.Description = "Validation des mutations ouvertes..."
            ribExpButton11.ShowText = True
            ribExpButton11.ShowImage = True
            ribExpButton11.Size = RibbonItemSize.Standard
            ribExpButton11.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton11.Image = Images.getBitmap(My.Resources.MOvaliderMut)
            'ribExpButton11.LargeImage = Images.getBitmap(My.Resources.MOvaliderMut)
            ribExpButton11.CommandParameter = MyCommands.NomCmd & "MOvaliderMut "
            ribExpButton11.CommandHandler = New AdskCommandHandler()
            ribInsert2.Items.Add(ribExpButton11)
            ribInsert2.Items.Add(ribCustomBrk)

            'Mise en évidence des points fiables...
            Dim ribExpButton12 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton12.Text = "Mise en évidence: Pts fiables"
            ribExpButton12.Description = "Mise en évidence des points fiables..."
            ribExpButton12.ShowText = True
            ribExpButton12.ShowImage = True
            ribExpButton12.Size = RibbonItemSize.Standard
            ribExpButton12.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton12.Image = Images.getBitmap(My.Resources.MOfiabPlan)
            ' ribExpButton12.LargeImage = Images.getBitmap(My.Resources.MOfiabPlan)
            ribExpButton12.CommandParameter = MyCommands.NomCmd & "MOfiabPlan "
            ribExpButton12.CommandHandler = New AdskCommandHandler()
            ribInsert2.Items.Add(ribExpButton12)

            ribExpPanel2.Items.Add(ribInsert2)

            'Nouveau groupe 3
            Dim ribInsert3 As New Autodesk.Windows.RibbonRowPanel

            'Distance 2D
            Dim ribExpButton14 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton14.Text = "Distance 2D"
            ribExpButton14.Description = "Distance 2D (avec projection)"
            ribExpButton14.ShowText = True
            ribExpButton14.ShowImage = True
            ribExpButton14.Size = RibbonItemSize.Standard
            ribExpButton14.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton14.Image = Images.getBitmap(My.Resources.Dist2D)
            ribExpButton14.LargeImage = Images.getBitmap(My.Resources.Dist2D)
            ribExpButton14.CommandParameter = MyCommands.NomCmd & "Dist2D "
            ribExpButton14.CommandHandler = New AdskCommandHandler()
            ribInsert3.Items.Add(ribExpButton14)
            ribInsert3.Items.Add(ribCustomBrk)


            'Division de surface
            Dim ribExpButton6 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton6.Text = "Division de surfaces"
            ribExpButton6.Description = "Division de surfaces"
            ribExpButton6.ShowText = True
            ribExpButton6.ShowImage = True
            ribExpButton6.Size = RibbonItemSize.Standard
            ribExpButton6.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton6.Image = Images.getBitmap(My.Resources.DivSurf)
            ribExpButton6.LargeImage = Images.getBitmap(My.Resources.DivSurf)
            ribExpButton6.CommandParameter = MyCommands.NomCmd & "DivSurf "
            ribExpButton6.CommandHandler = New AdskCommandHandler()
            ribInsert3.Items.Add(ribExpButton6)
            ribInsert3.Items.Add(ribCustomBrk)

            'Division de surface
            Dim ribExpButton18 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton18.Text = "Plan topographique"
            ribExpButton18.Description = "Création d'un plan topographique"
            ribExpButton18.ShowText = True
            ribExpButton18.ShowImage = True
            ribExpButton18.Size = RibbonItemSize.Standard
            ribExpButton18.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton18.Image = Images.getBitmap(My.Resources.plantopo)
            ribExpButton18.LargeImage = Images.getBitmap(My.Resources.plantopo)
            ribExpButton18.CommandParameter = MyCommands.NomCmd & "PlanTopo "
            ribExpButton18.CommandHandler = New AdskCommandHandler()
            ribInsert3.Items.Add(ribExpButton18)


            ribExpPanel2.Items.Add(ribInsert3)

            'Nouveau groupe Mise en forme
            'Nouveau groupe 3
            Dim ribMiseEnForme As New Autodesk.Windows.RibbonRowPanel



            'Rotation
            Dim ribExpButton19 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton19.Text = "Rotation"
            ribExpButton19.Description = "Rotation des blocks VD"
            ribExpButton19.ShowText = True
            ribExpButton19.ShowImage = True
            ribExpButton19.Size = RibbonItemSize.Standard
            ribExpButton19.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton19.Image = Images.getBitmap(My.Resources.rotation)
            ribExpButton19.LargeImage = Images.getBitmap(My.Resources.rotation)
            ribExpButton19.CommandParameter = MyCommands.NomCmd & "RotationVD "
            ribExpButton19.CommandHandler = New AdskCommandHandler()
            ribMiseEnForme.Items.Add(ribExpButton19)
            ribMiseEnForme.Items.Add(ribCustomBrk)

            'Ordre du tracé
            Dim ribExpButton13 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton13.Text = "Ordre du tracé"
            ribExpButton13.Description = "Réorganisation des calques pour l'impression du plan cadastral"
            ribExpButton13.ShowText = True
            ribExpButton13.ShowImage = True
            ribExpButton13.Size = RibbonItemSize.Standard
            ribExpButton13.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton13.Image = Images.getBitmap(My.Resources.OrdreTrace)
            ribExpButton13.LargeImage = Images.getBitmap(My.Resources.OrdreTrace)
            ribExpButton13.CommandParameter = MyCommands.NomCmd & "OrdreTraceVD "
            ribExpButton13.CommandHandler = New AdskCommandHandler()
            ribMiseEnForme.Items.Add(ribExpButton13)
            ribMiseEnForme.Items.Add(ribCustomBrk)


            'Ajout d’un réseau de coordonnées 
            Dim ribExpButton20 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton20.Text = "Réseau de coord."
            ribExpButton20.Description = "Ajout d’un réseau de coordonnées"
            ribExpButton20.ShowText = True
            ribExpButton20.ShowImage = True
            ribExpButton20.Size = RibbonItemSize.Standard
            ribExpButton20.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton20.Image = Images.getBitmap(My.Resources.Quadrillage)
            ribExpButton20.LargeImage = Images.getBitmap(My.Resources.Quadrillage)
            ribExpButton20.CommandParameter = MyCommands.NomCmd & "Quadrillage "
            ribExpButton20.CommandHandler = New AdskCommandHandler()
            ribMiseEnForme.Items.Add(ribExpButton20)

            ribExpPanel2.Items.Add(ribMiseEnForme)


            'Nouveau groupe Echelle ------------------------

            'Nouveau groupe 3
            Dim ribEchelle As New Autodesk.Windows.RibbonRowPanel


            'Geo Tools
            Dim ribExpButton21 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton21.Text = "Geo Tools"
            ribExpButton21.Description = "Outil pour le traitent topologique"
            ribExpButton21.ShowText = True
            ribExpButton21.ShowImage = True
            ribExpButton21.IsChecked = True
            ribExpButton21.Size = RibbonItemSize.Standard
            ' ribExpButton17.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton21.Image = Images.getBitmap(My.Resources.ribbon_plug16)
            ribExpButton21.LargeImage = Images.getBitmap(My.Resources.ribbon_plug32)
            ribExpButton21.CommandParameter = MyCommands.NomCmd & "GeoTools "
            ribExpButton21.CommandHandler = New AdskCommandHandler()
            ribEchelle.Items.Add(ribExpButton21)
            ribEchelle.Items.Add(ribCustomBrk)


            'Ech annotative
            Dim ribExpButton17 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton17.Text = "Ech. annotative"
            ribExpButton17.Description = "Activer ou supprimer l'annotativité dans le dessin"
            ribExpButton17.ShowText = True
            ribExpButton17.ShowImage = True
            ribExpButton17.IsChecked = True
            ribExpButton17.Size = RibbonItemSize.Standard
            ' ribExpButton17.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton17.Image = Images.getBitmap(My.Resources.Echelle)
            ribExpButton17.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribExpButton17.CommandParameter = MyCommands.NomCmd & "Annotatif "
            ribExpButton17.CommandHandler = New AdskCommandHandler()
            ribEchelle.Items.Add(ribExpButton17)
            ribEchelle.Items.Add(ribCustomBrk)


            Dim ribExpButton16 As Autodesk.Windows.RibbonSplitButton = New RibbonSplitButton()


            ' "100"
            Dim ribBtn100 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn100.Text = "1:100"
            ribBtn100.ShowText = True
            ribBtn100.ShowImage = True
            ribBtn100.Size = RibbonItemSize.Standard
            ribBtn100.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn100.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn100.CommandHandler = New AdskCommandHandler()
            ribBtn100.CommandParameter = MyCommands.NomCmd & "Ech100 "
            ribExpButton16.Items.Add(ribBtn100)


            ' "200" 
            Dim ribBtn200 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn200.Text = "1:200"
            ribBtn200.ShowText = True
            ribBtn200.ShowImage = True
            ribBtn200.Size = RibbonItemSize.Standard
            ribBtn200.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn200.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn200.CommandHandler = New AdskCommandHandler()
            ribBtn200.CommandParameter = MyCommands.NomCmd & "Ech200 "
            ribExpButton16.Items.Add(ribBtn200)


            ' "250"
            Dim ribBtn250 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn250.Text = "1:250"
            ribBtn250.ShowText = True
            ribBtn250.ShowImage = True
            ribBtn250.Size = RibbonItemSize.Standard
            ribBtn250.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn250.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn250.CommandHandler = New AdskCommandHandler()
            ribBtn250.CommandParameter = MyCommands.NomCmd & "Ech250 "
            ribExpButton16.Items.Add(ribBtn250)


            ' "500"
            Dim ribBtn500 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn500.Text = "1:500"
            ribBtn500.ShowText = True
            ribBtn500.ShowImage = True
            ribBtn500.Size = RibbonItemSize.Standard
            ribBtn500.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn500.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn500.CommandHandler = New AdskCommandHandler()
            ribBtn500.CommandParameter = MyCommands.NomCmd & "Ech500 "
            ribExpButton16.Items.Add(ribBtn500)


            ' "1000" 
            Dim ribBtn1000 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn1000.Text = "1:1000"
            ribBtn1000.ShowText = True
            ribBtn1000.ShowImage = True
            ribBtn1000.Size = RibbonItemSize.Standard
            ribBtn1000.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn1000.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn1000.CommandHandler = New AdskCommandHandler()
            ribBtn1000.CommandParameter = MyCommands.NomCmd & "Ech1000 "
            ribExpButton16.Items.Add(ribBtn1000)


            ' "2000" 
            Dim ribBtn2000 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn2000.Text = "1:2000"
            ribBtn2000.ShowText = True
            ribBtn2000.ShowImage = True
            ribBtn2000.Size = RibbonItemSize.Standard
            ribBtn2000.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn2000.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn2000.CommandHandler = New AdskCommandHandler()
            ribBtn2000.CommandParameter = MyCommands.NomCmd & "Ech2000 "
            ribExpButton16.Items.Add(ribBtn2000)


            ' "2500"
            Dim ribBtn2500 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn2500.Text = "1:2500"
            ribBtn2500.ShowText = True
            ribBtn2500.ShowImage = True
            ribBtn2500.Size = RibbonItemSize.Standard
            ribBtn2500.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn2500.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn2500.CommandHandler = New AdskCommandHandler()
            ribBtn2500.CommandParameter = MyCommands.NomCmd & "Ech2500 "
            ribExpButton16.Items.Add(ribBtn2500)

            ' "4000" 
            Dim ribBtn4000 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn4000.Text = "1:4000"
            ribBtn4000.ShowText = True
            ribBtn4000.ShowImage = True
            ribBtn4000.Size = RibbonItemSize.Standard
            ribBtn4000.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn4000.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn4000.CommandHandler = New AdskCommandHandler()
            ribBtn4000.CommandParameter = MyCommands.NomCmd & "Ech4000 "
            ribExpButton16.Items.Add(ribBtn4000)

            ' "5000" 
            Dim ribBtn5000 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn5000.Text = "1:5000"
            ribBtn5000.ShowText = True
            ribBtn5000.ShowImage = True
            ribBtn5000.Size = RibbonItemSize.Standard
            ribBtn5000.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn5000.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn5000.CommandHandler = New AdskCommandHandler()
            ribBtn5000.CommandParameter = MyCommands.NomCmd & "Ech5000 "
            ribExpButton16.Items.Add(ribBtn5000)

            ' "10000" 
            Dim ribBtn10000 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribBtn10000.Text = "1:10000"
            ribBtn10000.ShowText = True
            ribBtn10000.ShowImage = True
            ribBtn10000.Size = RibbonItemSize.Standard
            ribBtn10000.Image = Images.getBitmap(My.Resources.Echelle)
            ribBtn10000.LargeImage = Images.getBitmap(My.Resources.Echelle)
            ribBtn10000.CommandHandler = New AdskCommandHandler()
            ribBtn10000.CommandParameter = MyCommands.NomCmd & "Ech10000 "
            ribExpButton16.Items.Add(ribBtn10000)


            ribExpButton16.Width = 100
            ribExpButton16.Image = Images.getBitmap(My.Resources.Echelle)
            ribExpButton16.LargeImage = Images.getBitmap(My.Resources.Echelle)



            ribExpButton16.Text = "Echelle"
            ribExpButton16.Tag = "Changement des échelles (sans le mode annotatif)"
            ribExpButton16.Size = RibbonItemSize.Standard
            ribExpButton16.ShowText = True
            ' ribExpButton16.ShowImage = True

            ribExpButton16.CommandHandler = New AdskCommandHandler()
            ribEchelle.Items.Add(ribExpButton16)

            ribExpButton16.Current = ribBtn1000

            ribExpPanel2.Items.Add(ribEchelle)

        End Sub


        Private Sub addPanelOptions(ByVal ribTab As RibbonTab)


            'create the panel source 
            Dim ribSourcePanel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
            ribSourcePanel.Title = "Options"
            Dim ribCustomBrk As Autodesk.Windows.RibbonRowBreak = New Autodesk.Windows.RibbonRowBreak



            'Nouveau groupe
            ' Dim ribInfo As New Autodesk.Windows.RibbonRowPanel
            Dim ribSupport As New Autodesk.Windows.RibbonRowPanel

            'now the panel
            Dim ribPanel As New RibbonPanel()
            ribPanel.Source = ribSourcePanel
            ribTab.Panels.Add(ribPanel)

            'now add and expanded panel (with 1 button) 
            Dim ribExpPanel As Autodesk.Windows.RibbonRowPanel = New RibbonRowPanel()
            ribSourcePanel.Items.Add(ribExpPanel)


            Dim ribListBtnOption As Autodesk.Windows.RibbonSplitButton = New RibbonSplitButton()
            ribListBtnOption.ListImageSize = RibbonImageSize.Standard


            'Bouton Create Update
            Dim ribExpButton1 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton1.Text = "Mise à jour"
            ribExpButton1.Description = "Contrôler si le plugin est à jour"
            ribExpButton1.ShowText = True
            ribExpButton1.ShowImage = True
            ribExpButton1.Size = RibbonItemSize.Standard
            ribExpButton1.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton1.Image = Images.getBitmap(My.Resources.ribbon_feedback16)
            ribExpButton1.LargeImage = Images.getBitmap(My.Resources.ribbon_feedback32)
            ribExpButton1.CommandParameter = MyCommands.NomCmd & "Update "
            ribExpButton1.CommandHandler = New AdskCommandHandler()
            'ribInfo.Items.Add(ribExpButton1)
            ribListBtnOption.Items.Add(ribExpButton1)


            'ribInfo.Items.Add(ribCustomBrk)


            Dim ribExpButton3 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton3.Text = "Documentations"
            ribExpButton3.Description = "Documentations et aide à l'utilisations"
            ribExpButton3.ShowText = True
            ribExpButton3.ShowImage = True
            ribExpButton3.Size = RibbonItemSize.Standard
            ribExpButton3.Orientation = Windows.Controls.Orientation.Horizontal
            ribExpButton3.Image = Images.getBitmap(My.Resources.ribbon_feedback16)
            ribExpButton3.LargeImage = Images.getBitmap(My.Resources.ribbon_feedback32)
            ribExpButton3.CommandParameter = MyCommands.NomCmd & "Doc "
            ribExpButton3.CommandHandler = New AdskCommandHandler()
            'ribInfo.Items.Add(ribExpButton3)
            ribListBtnOption.Items.Add(ribExpButton3)
            'ribInfo.Items.Add(ribCustomBrk)





            Dim item As New RibbonButton()
            item.IsCheckable = True
            item.CommandParameter = MyCommands.NomCmd & "Options "
            item.Text = "Options"
            item.ShowImage = True
            item.ShowText = True
            item.CommandHandler = New AdskCommandHandler()
            ribSourcePanel.DialogLauncher() = item


            ribListBtnOption.Text = "Support"
            ribListBtnOption.Tag = "Support et documentation"
            ribListBtnOption.Size = RibbonItemSize.Large
            ribListBtnOption.ShowText = True
            ribExpPanel.Items.Add(ribListBtnOption)

            ' ribExpPanel.Items.Add(ribInfo)



        End Sub




        Private Sub addPanel1(ByVal ribTab As RibbonTab)
            'create the panel source 
            Dim ribSourcePanel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
            ribSourcePanel.Title = "Gestion"
            'now the panel 
            Dim ribPanel As New RibbonPanel()
            ribPanel.Source = ribSourcePanel
            ribTab.Panels.Add(ribPanel)

            'create button 1 
            Dim ribButton1 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribButton1.Text = "Register me!"
            ribButton1.CommandParameter = "REGISTERME "
            ribButton1.ShowText = True
            ribButton1.CommandHandler = New AdskCommandHandler()

            'now create button 2 
            Dim ribButton2 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribButton2.Text = "Unregister me!"
            ribButton2.CommandParameter = "UNREGISTERME "
            ribButton2.ShowText = True
            ribButton2.CommandHandler = New AdskCommandHandler()

            'create a tooltip 
            Dim ribToolTip As Autodesk.Windows.RibbonToolTip = New RibbonToolTip()
            ribToolTip.Command = "REGISTERME"
            ribToolTip.Title = "Register me command"
            ribToolTip.Content = "Register this assembly on AutoCAD startup"
            ribToolTip.ExpandedContent = "Add the necessary registry key to allow this assembly to auto load on startup. Also check which event you should add handle to add custom ribbon on AutoCAD startup."
            ribButton1.ToolTip = ribToolTip

            'now add the 2 button with a RowBreak 
            ribSourcePanel.Items.Add(ribButton1)
            ribSourcePanel.Items.Add(New RibbonRowBreak())
            ribSourcePanel.Items.Add(ribButton2)

            'now add and expanded panel (with 1 button) 
            Dim ribExpPanel As Autodesk.Windows.RibbonRowPanel = New RibbonRowPanel()

            ribSourcePanel.Items.Add(ribExpPanel)
            'needs to be here 
            Dim ribExpButton1 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribExpButton1.Text = "On expanded panel"
            ribExpButton1.ShowText = True
            ' wb added 
            ribExpButton1.CommandParameter = "LINE "
            ribExpButton1.CommandHandler = New AdskCommandHandler()
            'wb end added 
            'wb commented 
            'and add it to source 
            ' ribSourcePanel.Items.Add(new RibbonPanelBreak()); 
            'ribExpPanel.Items.Add(ribExpButton1); 
            'ribSourcePanel.Items.Add(ribExpPanel);//wb commented 
            ribExpPanel.Items.Add(ribExpButton1)

        End Sub

        Private Shared Sub addPanel2(ByVal ribTab As RibbonTab)
            'create the panel source 
            Dim ribSourcePanel As Autodesk.Windows.RibbonPanelSource = New RibbonPanelSource()
            ribSourcePanel.Title = "Controls"
            'now the panel 
            Dim ribPanel As New RibbonPanel()
            ribPanel.Source = ribSourcePanel
            ribTab.Panels.Add(ribPanel)

            'create a Ribbon text 
            Dim ribTxt As Autodesk.Windows.RibbonTextBox = New RibbonTextBox()
            ribTxt.Width = 100
            ribTxt.IsEmptyTextValid = False
            ribTxt.InvokesCommand = True
            ribTxt.CommandHandler = New AdskCommandHandler()

            'create a Ribbon List Button 
            Dim ribListBtn As Autodesk.Windows.RibbonSplitButton = New RibbonSplitButton()
            Dim ribButton1 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribButton1.Text = "Call LINE command"
            ribButton1.ShowText = True
            ribButton1.CommandParameter = "LINE "
            ribButton1.CommandHandler = New AdskCommandHandler()
            Dim ribButton2 As Autodesk.Windows.RibbonButton = New RibbonButton()
            ribButton2.Text = "Call CIRCLE command"
            ribButton2.ShowText = True
            ribButton2.CommandParameter = "CIRCLE "
            ribButton2.CommandHandler = New AdskCommandHandler()
            ribListBtn.Text = "Some options"
            ribListBtn.ShowText = True
            ribListBtn.Items.Add(ribButton1)
            ribListBtn.Items.Add(ribButton2)

            ribSourcePanel.Items.Add(ribTxt)
            ribSourcePanel.Items.Add(New RibbonRowBreak())
            ribSourcePanel.Items.Add(ribListBtn)
        End Sub

        'RevoRemoveTab
        Public Sub removeRibbon()
            Dim ribCntrl As Autodesk.Windows.RibbonControl = Autodesk.AutoCAD.Ribbon.RibbonServices.RibbonPaletteSet.RibbonControl
            'find the custom tab using the Id
            Dim NbreTab As Double = ribCntrl.Tabs.Count - 1

            For i As Double = 0 To NbreTab
                If ribCntrl.Tabs(i).Id.Equals(MY_TAB_ID) Then
                    ribCntrl.Tabs.Remove(ribCntrl.Tabs(i))
                    'Exit Sub
                    Exit For
                    'NbreTab = NbreTab - 1
                End If
            Next


            'Quick Access Toolbar
            Try
                Dim qat As Autodesk.Windows.ToolBars.QuickAccessToolBarSource = ComponentManager.QuickAccessToolBar
                If qat IsNot Nothing Then
                    For i As Integer = 0 To qat.Items.Count() - 1
                        If qat.Items(i).Id = MY_QUICKSTART_ID Then
                            qat.RemoveStandardItem(qat.Items(i))
                            ribBtnRevo.Items.Clear()
                            CreateBtnRevo()
                        End If
                    Next
                Else
                    ' MsgBox("Lancer Quick ")
                End If
            Catch 'ex As System.Exception
                'erreur de chargement
                'MsgBox(ex.Message)
            End Try


        End Sub

        '<CommandMethod("registerMe")> _
        Public Sub RegisterMe()
            'AutoCAD (or vertical) and Application keys
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acadKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.RegistryProductRootKey)
#Else 'Versio AutoCAD 2013 et +
            Dim acadKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.MachineRegistryProductRootKey)
#End If
            Dim acadAppKey As Microsoft.Win32.RegistryKey = acadKey.OpenSubKey("Applications", True)

            'already registered? 
            Dim curAssemblyName As String = Ass.xProduct.ToLower
            Dim subKeys As [String]() = acadAppKey.GetSubKeyNames()
            For Each subKey As [String] In subKeys
                If subKey.Equals(curAssemblyName) Then
                    acadAppKey.Close()
                    Exit Sub
                End If
            Next

            'assembly location and description (as full name) 
            Dim curAssemblyPath As String = System.Reflection.Assembly.GetExecutingAssembly().Location
            Dim curAssemblyFullName As String = Ass.xTitle

            'create the addin key 
            Dim acadAppAddInKey As Microsoft.Win32.RegistryKey = acadAppKey.CreateSubKey(curAssemblyName)
            acadAppAddInKey.SetValue("DESCRIPTION", curAssemblyFullName, Microsoft.Win32.RegistryValueKind.[String])
            acadAppAddInKey.SetValue("LOADCTRLS", 14, Microsoft.Win32.RegistryValueKind.DWord)
            acadAppAddInKey.SetValue("LOADER", curAssemblyPath, Microsoft.Win32.RegistryValueKind.[String])
            acadAppAddInKey.SetValue("MANAGED", 1, Microsoft.Win32.RegistryValueKind.DWord)

            acadAppKey.Close()
        End Sub

        '   <CommandMethod("UnregisterMe")> _
        Public Sub UnregisterMe()
            'AutoCAD (or vertical) and Application keys 
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acadKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.RegistryProductRootKey)
#Else 'Versio AutoCAD 2013 et +
            Dim acadKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(Autodesk.AutoCAD.DatabaseServices.HostApplicationServices.Current.MachineRegistryProductRootKey)
#End If
            Dim acadAppKey As Microsoft.Win32.RegistryKey = acadKey.OpenSubKey("Applications", True)

            'get assembly name and delete 
            Dim curAssemblyName As String = Ass.xProduct.ToLower
            acadAppKey.DeleteSubKeyTree(curAssemblyName)
            acadAppKey.Close()
        End Sub

    End Class

    Public Class AdskCommandHandler

        Implements System.Windows.Input.ICommand

        Function CanExecute(ByVal parameter As Object) As Boolean Implements System.Windows.Input.ICommand.CanExecute
            Return True
        End Function

        Public Event CanExecuteChanged As EventHandler Implements System.Windows.Input.ICommand.CanExecuteChanged

        Public Sub Execute(ByVal parameter As Object) Implements System.Windows.Input.ICommand.Execute
            'is from a Ribbon Button? 
            Dim ribBtn As RibbonButton = TryCast(parameter, RibbonButton)
            If ribBtn IsNot Nothing Then

                If Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument = Nothing Then
                    MsgBox("Pour utiliser cette commande, il est nécessaire d'ouvrir un dessin", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "REVO")
                Else
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.SendStringToExecute(DirectCast(ribBtn.CommandParameter, [String]), True, False, True)
                End If


            End If

            'is from s Ribbon Textbox? 
            Dim ribTxt As RibbonTextBox = TryCast(parameter, RibbonTextBox)
            If ribTxt IsNot Nothing Then
                System.Windows.Forms.MessageBox.Show(ribTxt.TextValue)
            End If
        End Sub



    End Class

    Public Class Images
        Public Shared Function getBitmap(ByVal image As Bitmap) As BitmapImage
            Dim stream As New MemoryStream()
            image.Save(stream, ImageFormat.Png)
            Dim bmp As New BitmapImage()
            bmp.BeginInit()
            bmp.StreamSource = stream
            bmp.EndInit()
            Return bmp
        End Function

    End Class

End Namespace

