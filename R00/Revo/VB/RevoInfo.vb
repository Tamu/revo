
Imports System.Resources
Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.IO

Namespace Revo

    Public Class RevoInfo

        Private myType As Type
        'Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

        Public Sub New()
            myType = GetType(frmState)
        End Sub
   
       

        'Information/variable Revo
        '          
        ' ! chargable via XML Perso !    : recherche la variable dans revo-config.xml
        ' ! chargable via XML shared !    : recherche la variable dans revo-shared.xml
        '     -- locked --         : Pas de lecture dans un fichier XML




        ' __________ --------- **********  X M L   P e r s o   ********** ---------  __________



        Public ReadOnly Property PluginSharedXML(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!! New => CLOUD
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "CloudConfig".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return PluginPersoXML
                End If
            End Get

        End Property


        Public ReadOnly Property SystemPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "SystemPath".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, "System\")
                End If
            End Get
        End Property


        Public ReadOnly Property MenuType(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "Menu".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return "revo"
                End If
            End Get
        End Property


        Public ReadOnly Property EDTFolder(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "EDTFolder".ToLower)
                Dim DesktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return DesktopPath
                End If
            End Get
        End Property
        Public ReadOnly Property EDTtemplate(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "EDTtemplate".ToLower)
                Dim PathTemplates As String = IO.Path.Combine(SharedPath, "EDT_Template.htm")
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return PathTemplates
                End If
            End Get
        End Property
        Public ReadOnly Property EDTCheckSurf(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "EDTCheckSurf".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return "0"
                End If
            End Get
        End Property
        Public ReadOnly Property EDTIgnoreRPL(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "EDTIgnoreRPL".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return "0"
                End If
            End Get
        End Property

        Public ReadOnly Property PERFolder(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "PERFolder".ToLower)
                Dim DesktopPath As String = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return DesktopPath
                End If
            End Get
        End Property

        Public ReadOnly Property FormatExport(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "FormatExport".ToLower)
                Dim DesktopPath As String = "1"
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return DesktopPath
                End If
            End Get
        End Property

        Public ReadOnly Property FormatImport(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginPersoXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "FormatImport".ToLower)
                Dim DesktopPath As String = "1"
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return DesktopPath
                End If
            End Get
        End Property









        ' __________ --------- **********  X M L   S h a r e d   ********** ---------  __________




        Public ReadOnly Property LogPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "LogPath".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, xProduct.ToLower & "-log.txt")
                End If
            End Get
        End Property

        Public ReadOnly Property XMLformatPerso(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "XMLformatPerso".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(SharedPath, "format-perso.xml")
                End If
            End Get
        End Property


        Public ReadOnly Property ActionsPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "ActionsPath".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, "Actions\")
                End If
            End Get
        End Property


        Public ReadOnly Property Plotters(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "Plotters".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, "Plotters\")
                End If
            End Get
        End Property

        Public ReadOnly Property SharedPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "SharedPath".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, "Shared\")
                End If
            End Get
        End Property
        Public ReadOnly Property Library(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "Library".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, "Library\")
                End If
            End Get
        End Property
        Public ReadOnly Property SupportPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "SupportPath".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, "Support\")
                End If
            End Get
        End Property
        Public ReadOnly Property Template(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "Template".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, "Template\Revo13.dwt")
                End If
            End Get
        End Property


        Public ReadOnly Property ToolPalettePath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "ToolPalettePath".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return Path.Combine(PluginPath, "ToolPalette\")
                End If
            End Get
        End Property
        Public ReadOnly Property PlotLogFilePath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "PlotLogFilePath".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return PluginPath
                End If
            End Get
        End Property

        Public ReadOnly Property SyncProfilAcad(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
            Get
                Dim XMLperso As New RevoXML(PluginSharedXML)
                Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "SyncProfilAcad".ToLower)
                If vari = "" Then DefVar = True
                If DefVar = False Then
                    Return vari
                Else
                    Return "0"
                End If
            End Get
        End Property












        ' __________ --------- **********  L O C K E D   ********** ---------  __________




        Public ReadOnly Property urlUpdate() As String '-- locked --
            Get
                Return "https://raw.githubusercontent.com/Tamu/revo/master/install/revo-update.xml" ' pour les packages REVO
            End Get
        End Property

        Public ReadOnly Property xTitle() As String  '-- locked --
            Get
                Dim at As Type = GetType(AssemblyTitleAttribute)
                Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
                Dim ta As AssemblyTitleAttribute = CType(r(0), AssemblyTitleAttribute)
                Return ta.Title
            End Get
        End Property
        Public ReadOnly Property xProduct() As String  '-- locked --
            Get
                Dim at As Type = GetType(AssemblyProductAttribute)
                Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
                Dim pt As AssemblyProductAttribute = CType(r(0), AssemblyProductAttribute)
                Return pt.Product
            End Get
        End Property

        Public ReadOnly Property xVersion() As String  '-- locked --
            Get
                Return myType.Assembly.GetName.Version.ToString()
            End Get
        End Property

        Public ReadOnly Property xCopyright() As String  '-- locked --
            Get
                Dim at As Type = GetType(AssemblyCopyrightAttribute)
                Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
                Dim ct As AssemblyCopyrightAttribute = CType(r(0), AssemblyCopyrightAttribute)
                Return ct.Copyright
            End Get
        End Property

        Public ReadOnly Property xCompany() As String  '-- locked --
            Get
                Dim at As Type = GetType(AssemblyCompanyAttribute)
                Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
                Dim ct As AssemblyCompanyAttribute = CType(r(0), AssemblyCompanyAttribute)
                Return ct.Company
            End Get
        End Property

        Public ReadOnly Property xDescription() As String  '-- locked --
            Get
                Dim at As Type = GetType(AssemblyDescriptionAttribute)
                Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
                Dim da As AssemblyDescriptionAttribute = CType(r(0), AssemblyDescriptionAttribute)
                Return da.Description
            End Get
        End Property



        Public ReadOnly Property XMLformatBase() As String ' -- locked --
            Get
                Return Path.Combine(PluginPath, "System\format-base.xml")
            End Get
        End Property


        Public ReadOnly Property PluginPath() As String '-- locked --
            Get
                Dim UpDLLPath As String = Path.GetDirectoryName(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location))
                Return UpDLLPath & "\"
            End Get
        End Property

        Public ReadOnly Property PluginPersoXML() As String '-- locked --
            Get
                Return Path.Combine(PluginPath, xProduct.ToLower & "-perso.xml")
            End Get
        End Property


        Public ReadOnly Property TemplateVersion() As Double '-- locked -- 'ATTENTION INFLUANCE LA COMPATIBILITE DU MODEL DE DESSIN
            Get
                Return 1.3
            End Get
        End Property


        Public ReadOnly Property XMLflow() As String '-- locked --
            Get
                Dim ProcID As String = ""
                Try
                    ProcID = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Catch
                End Try
                Return Path.Combine(PluginPath, ProcID & "-" & xProduct.ToLower & "-flow.xml")
            End Get
        End Property

        Public ReadOnly Property DB3interlis() As String '-- locked --
            Get
                Dim ProcID As String = ""
                Try
                    ProcID = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Catch
                End Try
                Return ProcID & "-Interlis.db3"
            End Get
        End Property

        Public ReadOnly Property DB3interlisOrig() As String '-- locked --
            Get
                Return "Interlis.db3"
            End Get
        End Property




        Public ReadOnly Property DB3system() As String '-- locked --
            Get
                Dim ProcID As String = ""
                Try
                    ProcID = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
                Catch
                End Try
                Return ProcID & "-System.db3"
            End Get
        End Property

        Public ReadOnly Property DB3systemOrig() As String '-- locked --
            Get
                Return "System.db3"
            End Get
        End Property


        Public ReadOnly Property PalletteXML() As String '-- locked --
            Get
                Return IO.Path.Combine(SharedPath, "insertiontxt.xml")
            End Get
        End Property

        Public ReadOnly Property ConfigMOVD() As String '-- locked --
            Get
                Return IO.Path.Combine(SharedPath, "Config-MOVD.xml")
            End Get
        End Property


        Public ReadOnly Property urlLicRight() As String '-- locked --
            Get
                Return "https://github.com/tamu/revo" ' pour les packages REVO
            End Get
        End Property
        Public ReadOnly Property urlDEV() As String '-- locked --
            Get
                Return "https://github.com/tamu/revo" ' pour les packages REVO
            End Get
        End Property


        Public ReadOnly Property Icon16() As System.Drawing.Bitmap '-- locked --
            Get
                Return My.Resources.ribbon_plug16 ' icon par défault des menus
            End Get
        End Property
        Public ReadOnly Property Icon32() As System.Drawing.Bitmap '-- locked --
            Get
                Return My.Resources.ribbon_plug32 ' icon par défault des menus
            End Get
        End Property
        Public ReadOnly Property Icon() As System.Drawing.Icon '-- locked --
            Get
                Return My.Resources.plug ' icon par défault des menus
            End Get
        End Property

    End Class

End Namespace