
Imports System
Imports System.Net
Imports System.IO
Imports System.Management
Imports System.Security.Cryptography
Imports System.Text
Imports Autodesk.AutoCAD.ApplicationServices

Namespace Revo

    Public Class connect
        Public FenStat As New frmState
        Dim Ass As New RevoInfo
        Public Shared ActLog As Boolean = False
        Public Shared ActDWG As Boolean = False
        Public Shared ActDXF As Boolean = False
        Public Shared ActPTS As Boolean = False
        Public Shared ActITF As Boolean = False
        Public Shared StateITF As Boolean = False
        Public Shared ITFfiles As String = ""

        ''' <summary>
        ''' Revo Message: boite de dialogue
        ''' </summary>
        '''         ''' <param name="Titel">Titel</param>
        ''' <param name="Msg">Message</param>
        ''' <param name="Close">Commande pour fermer la fenêtre</param>
        ''' <param name="CurVal">Ajout</param>
        ''' <param name="Som">Total</param>
        ''' <param name="Type">Optionel Titel</param>
        ''' <remarks>NomCMD + True,False + Commentaires + Fichier concerné </remarks>
        Public Sub Message(ByVal Titel As String, ByVal Msg As String, ByVal Close As Boolean, ByVal CurVal As Double, ByVal Som As Double, Optional ByVal Type As String = "", Optional ByVal ModalAff As Boolean = True)

            If Close = True Then 'Fermeture de la fenêtre
                Dim Start, Finish, Duree 'i, reponse, gagne, il_reste
                Duree = 3 'sec durée de pause
                Start = Timer ' top
                Finish = Start + Duree
                Do While Timer < Finish
                    'il_reste = Int(Finish - Timer)

                    System.Windows.Forms.Application.DoEvents() ' Donne le contrôle à d'autres processus.
                Loop
                FenStat.Hide()
                'FenStat.Close()
                'FenStat.Dispose()

            ElseIf Type.ToLower = "hide" Then
                FenStat.Hide()
            Else

                If ModalAff = True Then System.Windows.Forms.Application.DoEvents()

                'Chargement des messages
                FenStat.BringToFront()
                FenStat.Text = Titel
                FenStat.lbl_infos.Text = Msg

                If Type.ToLower = "critical" Then
                    FenStat.lbl_infos.ForeColor = Drawing.Color.Gold 'Yellow
                    FenStat.lbl_infos.Font = New Drawing.Font(FenStat.lbl_infos.Font.Name, 12, Drawing.FontStyle.Bold)
                ElseIf Type.ToLower = "info" Then
                    FenStat.lbl_infos.ForeColor = Drawing.Color.SkyBlue
                    FenStat.lbl_infos.Font = New Drawing.Font(FenStat.lbl_infos.Font.Name, 12, Drawing.FontStyle.Bold)
                ElseIf Type.ToLower = "select" Then
                    FenStat.lbl_infos.Text = "Choissisez une donnée:"
                    FenStat.ProgBar.Visible = False
                    FenStat.BtnValid.Visible = True
                    FenStat.BoxList.Visible = True
                    Dim MsgList() As String = Msg.Split(vbCrLf)
                    For Each Txt In MsgList
                        FenStat.BoxList.Items.Add(Txt)
                    Next
                    FenStat.BoxList.Text = FenStat.BoxList.Items.Item(0)
                ElseIf Type.ToLower = "close" Then
                    Try
                        FenStat.Close()
                    Catch 'ex As Exception
                        ' MsgBox(ex)
                    End Try

                End If


                ' Chargement de la barre de progression
                Try
                    If ModalAff = True Then
                        FenStat.Show() 'FenStat.ProgBar.Step = Math.Round(100 / (Som + 0.0001), 0)
                        ' FenStat.TopMost = True
                    Else
                        FenStat.ShowDialog() ' mettre un non modal
                    End If

                Catch ex As Exception
                    If ex.GetType.ToString = "System.ObjectDisposedException" Then
                        'MsgBox(ex.Message)
                        'Dim FenStat As New frmState
                        'FenStat.Show()
                    End If
                End Try

                Dim Valeur As Double
                Valeur = Math.Round((CurVal / Som + 0.0001) * 100, 0) 'FenStat.ProgBar.PerformStep()
                If Valeur >= 0 And Valeur <= 100 Then
                    FenStat.ProgBar.Value = Valeur
                    System.Windows.Forms.Application.DoEvents()
                End If

                If ModalAff = True Then System.Windows.Forms.Application.DoEvents()
            End If
        End Sub

        ''' <summary>
        ''' Revo Log enregistre toutes les actions et erreur de l'utilisateurs
        ''' </summary>
        ''' <param name="Infos">Informations de log</param>
        ''' <remarks>NomCMD + True,False + Commentaires + Fichier concerné </remarks>
        ''' 
        Public Function RevoLog(Optional ByVal Infos As String = "")

            Dim Txts() As String
            Dim FichierLog As String
            Txts = Split(Infos, vbCr)
            FichierLog = Ass.LogPath
            If Revo.RevoFiler.EcritureFichier(FichierLog, Txts) Then
                Return True
            Else
                Return False
            End If

            'EXEMPLE
            'connect.RevoLog(connect.DateLog & "Start Script" & vbTab & False & vbTab & "Erreur licence: " & "non valide" & " (" & LicR & ")")


        End Function

        ''' <summary>
        ''' Revo Date/time pour le log
        ''' </summary>
        ''' <remarks>Year-Month-Day Hour:Minute:Second </remarks>
        ''' 
        Public Function DateLog()
            Return DateTime.Now.ToString("yyyy-MM-dd") & " " & TimeOfDay & vbTab
        End Function



        Public Function UpdateProfile(PrinterConfigPath As String, PrinterDescPath As String, PrinterStyleSheetPath As String, _
                                    PlotLogFilePath As String, TemplateDwgPath As String, QNewTemplateFile As String, _
                                    PersoSupport As String, ActionsPath As String, ToolPalettePath As String)


            'Cast the current autocad application from .net to COM interop
            Dim oApp As Autodesk.AutoCAD.Interop.AcadApplication = Autodesk.AutoCAD.ApplicationServices.Application.AcadApplication
            'Get the user preferences on the current profile
            Dim profile As Autodesk.AutoCAD.Interop.AcadPreferences = oApp.Preferences
            'Configure the default preferences
            Dim LogTxt As String = ""
            Dim Usr As String = oApp.Name

            Try
                'Dim server As String = ReadSupportLocation()

                'Printer configuration path
                If PrinterConfigPath <> "" And IO.Directory.Exists(PrinterConfigPath) Then
                    LogTxt += DateLog() & Environment.UserName & " / Plotters :" & vbTab & oApp.Preferences.Files.PrinterConfigPath & " => "
                    oApp.Preferences.Files.PrinterConfigPath = PrinterConfigPath ' "\Plotters\Configuration" ' *.pc3
                    LogTxt += oApp.Preferences.Files.PrinterConfigPath & vbCr
                End If

                'Printer description path
                If PrinterDescPath <> "" And IO.Directory.Exists(PrinterDescPath) Then
                    LogTxt += DateLog() & Environment.UserName & " / PrinterDesc :" & vbTab & oApp.Preferences.Files.PrinterDescPath & " => "
                    oApp.Preferences.Files.PrinterDescPath = PrinterDescPath ' "\Plotters\PMP Files" ' *.pmp
                    LogTxt += oApp.Preferences.Files.PrinterDescPath & vbCr
                End If

                'Printer style sheet path
                If PrinterStyleSheetPath <> "" And IO.Directory.Exists(PrinterStyleSheetPath) Then
                    LogTxt += DateLog() & Environment.UserName & " / PrinterStyleSheet :" & vbTab & oApp.Preferences.Files.PrinterStyleSheetPath & " => "
                    oApp.Preferences.Files.PrinterStyleSheetPath = PrinterStyleSheetPath '"\Plotters\Styles" ' *.ctb
                    LogTxt += oApp.Preferences.Files.PrinterStyleSheetPath & vbCr
                End If

                'Plot log file location
                If PlotLogFilePath <> "" And IO.Directory.Exists(PlotLogFilePath) Then
                    LogTxt += DateLog() & Environment.UserName & " / PlotLog :" & vbTab & oApp.Preferences.Files.PlotLogFilePath & " => "
                    oApp.Preferences.Files.PlotLogFilePath = PlotLogFilePath ' "\Plotters\Logs" ' *.log
                    LogTxt += oApp.Preferences.Files.PlotLogFilePath & vbCr
                End If

                'Template settings
                If TemplateDwgPath <> "" And IO.Directory.Exists(TemplateDwgPath) Then
                    LogTxt += DateLog() & Environment.UserName & " / Template :" & vbTab & oApp.Preferences.Files.TemplateDwgPath & " => "
                    oApp.Preferences.Files.TemplateDwgPath = TemplateDwgPath ' "\Templates" ' *.dwt
                    LogTxt += oApp.Preferences.Files.TemplateDwgPath & vbCr
                End If

                'Template : RapNouv
                If QNewTemplateFile <> "" And IO.File.Exists(QNewTemplateFile) Then
                    LogTxt += DateLog() & Environment.UserName & " / QNewTemplate :" & vbTab & oApp.Preferences.Files.QNewTemplateFile & " => "
                    oApp.Preferences.Files.QNewTemplateFile = QNewTemplateFile ' "\acad.dwt" ' *.dwt
                    LogTxt += oApp.Preferences.Files.QNewTemplateFile & vbCr
                End If

                'Actions (Autres Actions) : _ACTPATH : "C:\Revo\Actions;C:\Test2"
                If ActionsPath <> "" And IO.Directory.Exists(ActionsPath) Then
                    LogTxt += DateLog() & Environment.UserName & " / Actions :" & vbTab & Application.GetSystemVariable("ACTPATH") & " => "
                    Application.SetSystemVariable("ACTPATH", ActionsPath) ' C:\Revo\Actions;C:\Test2
                    LogTxt += Application.GetSystemVariable("ACTPATH") & vbCr
                End If

                'Tools Palettes (ToolPalette) : 
                If ToolPalettePath <> "" And IO.Directory.Exists(ToolPalettePath) Then
                    LogTxt += DateLog() & Environment.UserName & " / ToolPalette :" & vbTab & oApp.Preferences.Files.ToolPalettePath & " => "
                    oApp.Preferences.Files.ToolPalettePath = ToolPalettePath
                    LogTxt += oApp.Preferences.Files.ToolPalettePath & vbCr
                End If

                'Library : 
                'If Library <> "" And IO.Directory.Exists(Library) Then
                '     oApp.Preferences.Files. =Library
                'End If

                'Support Paths
                If PersoSupport <> "" And IO.Directory.Exists(PersoSupport) Then
                    Dim supportpaths() As String = Split(oApp.Preferences.Files.SupportPath, ";")
                    Dim supportlist As New List(Of String)
                    For i As Integer = 0 To UBound(supportpaths)
                        supportlist.Add(supportpaths(i))
                    Next

                    'Add Support directory
                    If Not supportlist.Contains(PersoSupport) Then
                        supportlist.Add(PersoSupport)
                    End If

                    ''Add applications directory
                    'If Not supportlist.Contains(server + "\Applications") Then
                    '    supportlist.Add(server + "\Applications")
                    'End If
                    ''Add images directories
                    'If Not supportlist.Contains(server + "\Images\Internal") Then
                    '    supportlist.Add(server + "\Images\Internal")
                    'End If
                    'If Not supportlist.Contains(server + "\Images\External") Then
                    '    supportlist.Add(server + "\Images\External")
                    'End If

                    'Format new support path string
                    Dim supportpath As String = Nothing
                    For Each item As String In supportlist
                        supportpath = supportpath + item + ";"
                    Next
                    'Add new support
                    oApp.Preferences.Files.SupportPath = supportpath
                    LogTxt += DateLog() & Environment.UserName & " / Perso Support :" & vbTab & PersoSupport & vbCr
                End If

            Catch ex As Exception

            End Try

            '  RevoLog(LogTxt)
            Return (True)

        End Function

        Public Function RevoUpdateProfile()
            Dim info As New Revo.RevoInfo

            'Remplacement uniquement SI la valeur est pas par défaut
            Dim PrinterConfigPath As String = ""
            If info.Plotters <> info.Plotters(True) Then PrinterConfigPath = info.Plotters ' plotters

            Dim PrinterDescPath As String = ""
            If info.Plotters <> info.Plotters(True) Then PrinterDescPath = IO.Path.Combine(info.Plotters, "pmp files") ' plotters\pmp files

            Dim PrinterStyleSheetPath As String = ""
            If info.Plotters <> info.Plotters(True) Then PrinterStyleSheetPath = IO.Path.Combine(info.Plotters, "plot styles") ' plotters\plot styles

            Dim PlotLogFilePath As String = ""
            If info.PlotLogFilePath <> info.PlotLogFilePath(True) Then PlotLogFilePath = info.PlotLogFilePath

            Dim TemplateDwgPath As String = ""
            If info.Template <> info.Template(True) Then TemplateDwgPath = IO.Path.GetDirectoryName(info.Template) ' Template

            Dim QNewTemplateFile As String = ""
            If info.Template <> info.Template(True) Then QNewTemplateFile = info.Template 'Template\Revo11.dwt

            Dim PersoSupport As String = ""
            If info.SupportPath <> info.SupportPath(True) Then PersoSupport = info.SupportPath 'Support

            Dim ActionsPath As String = ""
            If info.ActionsPath <> info.ActionsPath(True) Then ActionsPath = info.ActionsPath 'Actions

            Dim ToolPalettePath = ""
            If info.ToolPalettePath <> info.ToolPalettePath(True) Then ToolPalettePath = info.ToolPalettePath 'ToolPalette

            UpdateProfile(PrinterConfigPath, PrinterDescPath, PrinterStyleSheetPath, PlotLogFilePath, TemplateDwgPath, QNewTemplateFile, PersoSupport, ActionsPath, ToolPalettePath)

            Return True

        End Function

        Public Function PrefUser()

            Dim XMLpath As String
            XMLpath = Ass.PluginPersoXML
            'Test de l'existance du fichier XML Revo
            CreateXML(XMLpath)

            Dim ListPref(0 To 5) As String
            ActLog = False : ActDWG = True : ActDXF = False : ActPTS = False : ActITF = False
            ListPref.SetValue("0", 0) : ListPref.SetValue("1", 1) : ListPref.SetValue("0", 2) : ListPref.SetValue("0", 3) : ListPref.SetValue("0", 4)

            Dim x As New RevoXML(XMLpath)
            If x.SelectValue("/" & Ass.xProduct.ToLower & "/config/ActLog") = "1" Then
                ActLog = True
                ListPref.SetValue("1", 0)
            End If
            If x.SelectValue("/" & Ass.xProduct.ToLower & "/config/ActDWG") <> "1" Then
                ActDWG = False
                ListPref.SetValue("0", 1)
            End If
            If x.SelectValue("/" & Ass.xProduct.ToLower & "/config/ActDXF") = "1" Then
                ActDXF = True
                ListPref.SetValue("1", 2)
            End If
            If x.SelectValue("/" & Ass.xProduct.ToLower & "/config/ActPTS") = "1" Then
                ActPTS = True
                ListPref.SetValue("1", 3)
            End If
            If x.SelectValue("/" & Ass.xProduct.ToLower & "/config/ActITF") = "1" Then
                ActITF = True
                ListPref.SetValue("1", 4)
            End If
            ListPref.SetValue(x.SelectValue("/" & Ass.xProduct.ToLower & "/config/size"), 5)

            Return ListPref
        End Function

        Public Function ReadXMLformat(ByVal docXML As System.Xml.XmlDocument) As List(Of String)

            Dim FormInfo As New List(Of String)
            FormInfo.Add("") '0 ) Number of Format
            FormInfo.Add("") '1 ) Format Name
            FormInfo.Add("") '2 ) Format ext
            FormInfo.Add("") '3 ) Format Code

            Dim NodeFormats As System.Xml.XmlNodeList
            NodeFormats = docXML.SelectNodes("/data-format/format")
            FormInfo(0) = NodeFormats.Count

            Try
                For Each NodeFormat As System.Xml.XmlNode In NodeFormats
                    FormInfo(1) += "|" & NodeFormat.SelectSingleNode("name").InnerText & " (*." & NodeFormat.SelectSingleNode("extension").InnerText & ")" & "|*." & NodeFormat.SelectSingleNode("extension").InnerText
                    FormInfo(2) += ";*." & NodeFormat.SelectSingleNode("extension").InnerText
                    FormInfo(3) += "|" & NodeFormat.Attributes.ItemOf(0).InnerText()
                Next
                If Left(FormInfo(1), 1) = "|" Then FormInfo(1) = Right(FormInfo(1), FormInfo(1).Length - 1)
                If Left(FormInfo(2), 1) = ";" Then FormInfo(2) = Right(FormInfo(2), FormInfo(2).Length - 1)
                If Left(FormInfo(3), 1) = "|" Then FormInfo(3) = Right(FormInfo(3), FormInfo(3).Length - 1)
            Catch ex As Exception

            End Try

            Return FormInfo

        End Function

        Public Sub CreateXML(ByVal XMLpath As String, Optional DefaultConfig As Boolean = True)
            Dim Txts() As String

            Txts = Split(Replace("<?xml version='1.0' encoding='ISO-8859-1'?>|<" & Ass.xProduct.ToLower & ">|  <config>|  </config>|  <cmdFiles>|  </cmdFiles>|</" & Ass.xProduct.ToLower & ">", "'", Chr(34)), "|")
            If File.Exists(XMLpath) = False Then
                Revo.RevoFiler.EcritureFichier(XMLpath, Txts, False)

                XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActDWG", "1", 1) ' <ActDWG>1</ActDWG>
                XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActDXF", "1", 1) ' <ActDXF>1</ActDXF>
                XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActITF", "1", 1) ' <ActITF>1</ActITF>
                XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActPTS", "1", 1) ' <ActPTS>1</ActPTS>

                If DefaultConfig = True Then
                    XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", "ExportData.csv", 1) ' <url>ExportData.csv</url>
                    XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", "ImportData.csv", 2) ' <url>ImportData.csv</url>
                    XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", "Hach_RF2007-CH.csv", 3) ' <url>Hach_RF2007-CH.csv</url>
                    XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", "Hach_couleur_complet.csv", 4) ' <url>Hach_couleur_complet.csv</url>
                    XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", "Hach_Bat_grises.csv", 5) ' <url>Hach_Bat_grises.csv</url>
                    XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", "Hach_Standard.csv", 6) ' <url>Hach_Standard.csv</url>
                    XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", "Hach_Enquete.csv", 7) ' <url>Hach_Enquete.csv</url>
                End If

            End If

        End Sub

        Public Function XMLdelete(ByVal Xpath As String, ByVal Nom As String, Optional ByVal Index As Integer = 1, Optional ByVal XMLPath As String = "")

            Dim XMLperso As New RevoXML(Ass.PluginPersoXML)
            Dim XMLpersoCloud As New RevoXML(Ass.PluginSharedXML)

            If XMLPath = "" Then
                'Test si paramètre est en réseau
                Dim ConfigCloud As Boolean = False
                If Ass.PluginSharedXML(True).ToUpper <> Ass.PluginSharedXML(False).ToUpper Then
                    If InStr("<LogPath><XMLformat><ActionsPath><Plotters><SharedPath><Library><SupportPath><Template><ToolPalettePath><PlotLogFilePath><SyncProfilAcad>".ToUpper, "<" & Nom.ToUpper & ">", ) <> 0 Then
                        XMLPath = Ass.PluginSharedXML
                        ConfigCloud = True
                    End If
                End If
                If ConfigCloud Then 'Supp Element XML à distance
                    XMLpersoCloud.deleteElement(Xpath, Nom, Index)
                Else 'Supp Element XML en local
                    XMLperso.deleteElement(Xpath, Nom, Index)
                End If
            Else
                Dim XMLDdel As New RevoXML(XMLPath)
                XMLDdel.deleteElement(Xpath, Nom, Index)
            End If

            Return "ok"
        End Function

        Public Function XMLwriteConfig(ByVal Xpath As String, ByVal Nom As String, ByVal Valeur As String, Optional ByVal Index As Integer = 1, Optional DefaultConfig As Boolean = True)


            Dim Paths() As String
            Dim XMLPath As String = Ass.PluginPersoXML
            Dim ConfigCloud As Boolean = False
            'Test si paramètre est en réseau
            If Ass.PluginSharedXML(True).ToUpper <> Ass.PluginSharedXML(False).ToUpper Then
                'Modification à distance : LogPath,XMLformat,ActionsPath,Plotters,SharedPath,Library,SupportPath,Template,ToolPalettePath,PlotLogFilePath,SyncProfilAcad
                If InStr("<LogPath><XMLformat><ActionsPath><Plotters><SharedPath><Library><SupportPath><Template><ToolPalettePath><PlotLogFilePath><SyncProfilAcad>".ToUpper, "<" & Nom.ToUpper & ">", ) <> 0 Then
                    XMLPath = Ass.PluginSharedXML
                    ConfigCloud = True
                    '"url" Les actions sont modifiable que en local
                End If
            End If

            'Test de l'existance du fichier XML Revo
            CreateXML(XMLPath, DefaultConfig)

            Dim XMLperso As New RevoXML(Ass.PluginPersoXML)
            Dim XMLpersoCloud As New RevoXML(Ass.PluginSharedXML)

            'Test de l'existance du chemin Xpath
            Dim PathTest As String = ""
            Dim PathBase As String = ""
            Paths = Split(Xpath, "/")

            'Travail à distance
            If ConfigCloud Then 'XMLpersoCloud
                For i = 0 To Paths.Count - 1
                    If Paths(i) <> "" Then
                        PathTest = PathTest & "/" & Paths(i)
                        If XMLpersoCloud.SelectValue(Xpath) = "ERRORXML" Then
                            'MsgBox("création : " & PathTest)
                            If PathBase <> "" Then XMLpersoCloud.addElement(PathBase, Paths(i), "")
                        End If
                        PathBase = PathBase & "/" & Paths(i)
                    End If
                Next
                'Test la présence de la clé
                If XMLpersoCloud.SelectValue(Xpath & "/" & Nom) = "ERRORXML" Or Index > 1 Then
                    'Ecriture de le l'élement XPath
                    If Nom <> "" And Xpath <> "" Then XMLpersoCloud.addElement(Xpath, Nom, Valeur)
                Else
                    'Ecriture de l'élement XPath
                    If Nom <> "" And Xpath <> "" Then XMLpersoCloud.setElementValue(Xpath & "/" & Nom, Valeur, Index)
                End If


            Else 'Travail en local
                For i = 0 To Paths.Count - 1
                    If Paths(i) <> "" Then
                        PathTest = PathTest & "/" & Paths(i)
                        If XMLperso.SelectValue(Xpath) = "ERRORXML" Then
                            'MsgBox("création : " & PathTest)
                            If PathBase <> "" Then XMLperso.addElement(PathBase, Paths(i), "")
                        End If
                        PathBase = PathBase & "/" & Paths(i)
                    End If
                Next
                'Test la présence de la clé
                If XMLperso.SelectValue(Xpath & "/" & Nom) = "ERRORXML" Or Index > 1 Then
                    'Ecriture de le l'élement XPath
                    If Nom <> "" And Xpath <> "" Then XMLperso.addElement(Xpath, Nom, Valeur)
                Else
                    'Ecriture de l'élement XPath
                    If Nom <> "" And Xpath <> "" Then XMLperso.setElementValue(Xpath & "/" & Nom, Valeur, Index)
                End If
            End If

            Return True

        End Function
        Public Function XMLwriteX(ByVal XMLPath As String, ByVal Xpath As String, ByVal Nom As String, ByVal Valeur As String, Optional ByVal Index As Integer = 1)

            If File.Exists(XMLPath) Then

                Dim Paths() As String
                Dim XMLw As New RevoXML(XMLPath)
                
                'Test de l'existance du chemin Xpath
                Dim PathTest As String = ""
                Dim PathBase As String = ""
                Paths = Split(Xpath, "/")
                For i = 0 To Paths.Count - 1
                    If Paths(i) <> "" Then
                        PathTest = PathTest & "/" & Paths(i)
                        If XMLw.SelectValue(Xpath) = "ERRORXML" Then
                            'MsgBox("création : " & PathTest)
                            If PathBase <> "" Then XMLw.addElement(PathBase, Paths(i), "")
                        End If
                        PathBase = PathBase & "/" & Paths(i)
                    End If
                Next
                'Test la présence de la clé
                If XMLw.SelectValue(Xpath & "/" & Nom) = "ERRORXML" Or Index > 1 Then
                    'Ecriture de le l'élement XPath
                    If Nom <> "" And Xpath <> "" Then XMLw.addElement(Xpath, Nom, Valeur)
                Else
                    'Ecriture de l'élement XPath
                    If Nom <> "" And Xpath <> "" Then XMLw.setElementValue(Xpath & "/" & Nom, Valeur, Index)
                End If

                Return True
            Else
                Return False
            End If


        End Function



        'Encodage MD5
        Public Function encMD5(ByVal Varstr As String)

            Dim MD5 As New MD5CryptoServiceProvider
            Dim Data As Byte()
            Dim Result As Byte()
            Dim Res As String = ""
            Dim Tmp As String = ""

            Data = Encoding.ASCII.GetBytes(Varstr)
            Result = MD5.ComputeHash(Data)
            For i As Integer = 0 To Result.Length - 1
                Tmp = Hex(Result(i))
                If Len(Tmp) = 1 Then Tmp = "0" & Tmp
                Res += Tmp
            Next
            Return Res.ToLower

        End Function



    End Class
End Namespace