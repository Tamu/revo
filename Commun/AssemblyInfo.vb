Imports System.Resources

Imports System
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.IO

' General Information about an assembly is controlled through the following 
' set of attributes. Change these attribute values to modify the information
' associated with an assembly.

<Assembly: AssemblyTitle("REVO")> 
<Assembly: AssemblyDescription("plugin Autocad")> 
<Assembly: AssemblyCompany("platform5")> 
<Assembly: AssemblyProduct("revo")> 
<Assembly: AssemblyCopyright("Copyright © platform5 2011")> 
<Assembly: AssemblyTrademark("platform5")> 

'<Assembly: AssemblyTitle("BS+R tools")> 
'<Assembly: AssemblyDescription("plug-in Autocad")> 
'<Assembly: AssemblyCompany("platform5 R&D Sàrl")> 
'<Assembly: AssemblyProduct("bsr")> 
'<Assembly: AssemblyCopyright("Copyright © platform5 2011")> 
'<Assembly: AssemblyTrademark("platform5 R&D Sàrl")> 

'   >>>        <<<<
' !!  ATTENTION NE PAS OUBLIER DE MODIFIER DANS :                   <<<<<<<<<<<<<------------ !!!!!
'         ---   MyCommands.vb   LA CONSTANTE NomCMD
'         ---   urlUpdate


' Version information for an assembly consists of the following four values:
'    Major  .  Minor  . Build Number .  Revision
' <Assembly: AssemblyVersion>     <<<<------- C'est la variable de référence
<Assembly: AssemblyVersion("0.9.7.1")> 
<Assembly: AssemblyFileVersion("0.9.7.1")> 

' In order to sign your assembly you must specify a key to use. Refer to the 
' Microsoft .NET Framework documentation for more information on assembly signing.
'
' Use the attributes below to control which key is used for signing. 
'
' Notes: 
'   (*) If no key is specified, the assembly is not signed.
'   (*) KeyName refers to a key that has been installed in the Crypto Service
'       Provider (CSP) on your machine. KeyFile refers to a file which contains
'       a key.
'   (*) If the KeyFile and the KeyName values are both specified, the 
'       following processing occurs:
'       (1) If the KeyName can be found in the CSP, that key is used.
'       (2) If the KeyName does not exist and the KeyFile does exist, the key 
'           in the KeyFile is installed into the CSP and used.
'   (*) In order to create a KeyFile, you can use the sn.exe (Strong Name) utility.
'       When specifying the KeyFile, the location of the KeyFile should be
'       relative to the project output directory which is
'       %Project Directory%\obj\<configuration>. For example, if your KeyFile is
'       located in the project directory, you would specify the AssemblyKeyFile 
'       attribute as [assembly: AssemblyKeyFile("..\\..\\mykey.snk")]
'   (*) Delay Signing is an advanced option - see the Microsoft .NET Framework
'       documentation for more information on this.
<Assembly: AssemblyDelaySign(False)> 
<Assembly: AssemblyKeyFile("")> 
<Assembly: AssemblyKeyName("")> 

' Setting ComVisible to false makes the types in this assembly not visible 
' to COM components.  If you need to access a type in this assembly from 
' COM, set the ComVisible attribute to true on that type.
<Assembly: ComVisible(False)> 

' The following GUID is for the ID of the typelib if this project is exposed to COM
<Assembly: Guid("fb0a2132-5ca7-4830-8773-03bfb36e65d9")> 

#Region " Classe pour récupérer les informations pour toute les variables"
Public Class AssemblyInfo
    Private myType As Type

    Private AppPath As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)
    Private AppShared As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData)

    Public Sub New()
        myType = GetType(frmState)
    End Sub
    'Public ReadOnly Property AsmName() As String
    '    Get
    '        Return myType.Assembly.GetName.Name.ToString()
    '    End Get
    'End Property
    'Public ReadOnly Property AsmFQName() As String
    '    Get
    '        Return myType.Assembly.GetName.FullName.ToString()
    '    End Get
    'End Property
    'Public ReadOnly Property CodeBase() As String
    '    Get
    '        Return myType.Assembly.CodeBase
    '    End Get
    'End Property 

    Public ReadOnly Property urlUpdate() As String '-- locked --
        Get
            Return "http://platform5rd.com/autocad-pc/revo/revo-update.xml" ' REVO url pour les updates
            'Return "http://platform5rd.com/autocad-pc/revo/bsr-update.xml" ' BS+R url pour les updates
        End Get
    End Property

    Public ReadOnly Property xTitle() As String
        Get
            Dim at As Type = GetType(AssemblyTitleAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim ta As AssemblyTitleAttribute = CType(r(0), AssemblyTitleAttribute)
            Return ta.Title
        End Get
    End Property
    Public ReadOnly Property xProduct() As String
        Get
            Dim at As Type = GetType(AssemblyProductAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim pt As AssemblyProductAttribute = CType(r(0), AssemblyProductAttribute)
            Return pt.Product
        End Get
    End Property

    Public ReadOnly Property xVersion() As String
        Get
            Return myType.Assembly.GetName.Version.ToString()
        End Get
    End Property
    Public ReadOnly Property xCopyright() As String
        Get
            Dim at As Type = GetType(AssemblyCopyrightAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim ct As AssemblyCopyrightAttribute = CType(r(0), AssemblyCopyrightAttribute)
            Return ct.Copyright
        End Get
    End Property
    Public ReadOnly Property xCompany() As String
        Get
            Dim at As Type = GetType(AssemblyCompanyAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim ct As AssemblyCompanyAttribute = CType(r(0), AssemblyCompanyAttribute)
            Return ct.Company
        End Get
    End Property
    Public ReadOnly Property xDescription() As String
        Get
            Dim at As Type = GetType(AssemblyDescriptionAttribute)
            Dim r() As Object = myType.Assembly.GetCustomAttributes(at, False)
            Dim da As AssemblyDescriptionAttribute = CType(r(0), AssemblyDescriptionAttribute)
            Return da.Description
        End Get
    End Property




    'Information Revo
    '          
    '     -- locked --         : Pas de lecture dans un fichier XML
    ' ! chargable via XML !    : recherche la variable dans revo-config.xml



    Public ReadOnly Property RevoPath() As String '-- locked --
        Get
            Return Path.Combine(AppPath, xProduct.ToLower & "\")
        End Get
    End Property
    Public ReadOnly Property XMLPath() As String '-- locked --
        Get
            Return Path.Combine(AppPath, xProduct.ToLower & "\" & xProduct.ToLower & "-perso.xml")
        End Get
    End Property
    Public ReadOnly Property XMLflow() As String '-- locked --
        Get
            Dim ProcID As String = ""
            Try
                ProcID = System.Diagnostics.Process.GetCurrentProcess.Id.ToString()
            Catch
            End Try
            Return Path.Combine(AppPath, xProduct.ToLower & "\" & ProcID & "-" & xProduct.ToLower & "-flow.xml")
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

    Public ReadOnly Property LogPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "LogPath".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return Path.Combine(AppPath, xProduct.ToLower & "\" & xProduct.ToLower & "-log.txt")
            End If
        End Get
    End Property
    Public ReadOnly Property XMLformat(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "XMLformat".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return Path.Combine(AppPath, xProduct.ToLower & "\Shared\" & "format.xml")
            End If
        End Get
    End Property
    Public ReadOnly Property SystemPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "SystemPath".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return Path.Combine(AppPath, xProduct.ToLower & "\System\")
            End If
        End Get
    End Property
    Public ReadOnly Property ActionsPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "ActionsPath".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return Path.Combine(AppShared, xProduct.ToLower & "\Actions\")
            End If
        End Get
    End Property
    'Public ReadOnly Property CTB(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
    '    Get
    '        Dim XMLperso As New RevoXML(XMLPath)
    '        Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "CTB".ToLower)
    '        If vari = "" Then DefVar = True
    '        If DefVar = False Then
    '            Return vari
    '        Else
    '            Return Path.Combine(AppShared, xProduct.ToLower & "\Plotters\Plot Styles\JourCAD.ctb")
    '        End If
    '    End Get
    'End Property
    Public ReadOnly Property SharedPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "SharedPath".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return Path.Combine(AppShared, xProduct.ToLower & "\Shared\")
            End If
        End Get
    End Property
    Public ReadOnly Property Library(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "Library".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return Path.Combine(AppShared, xProduct.ToLower & "\Library\")
            End If
        End Get
    End Property
    Public ReadOnly Property SupportPath(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "SupportPath".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return Path.Combine(AppShared, xProduct.ToLower & "\Support\")
            End If
        End Get
    End Property
    Public ReadOnly Property Template(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "Template".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return Path.Combine(AppShared, xProduct.ToLower & "\Template\Revo10.dwt")
            End If
        End Get
    End Property

    Public ReadOnly Property EDTFolder(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
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
    Public ReadOnly Property EDTCheckSurf(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
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
            Dim XMLperso As New RevoXML(XMLPath)
            Dim vari As String = XMLperso.getElementValue("/" & xProduct.ToLower & "/config/" & "EDTIgnoreRPL".ToLower)
            If vari = "" Then DefVar = True
            If DefVar = False Then
                Return vari
            Else
                Return "0"
            End If
        End Get
    End Property
    'Dim strVal As String = LireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "PERFolder")

    Public ReadOnly Property PERFolder(Optional ByVal DefVar As Boolean = False) As String ' !!!!!  chargable via XML  !!!!!!
        Get
            Dim XMLperso As New RevoXML(XMLPath)
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
    Public ReadOnly Property urlLicRight() As String '-- locked --
        Get
            Return "http://platform5rd.com/"
        End Get
    End Property
    Public ReadOnly Property urlDEV() As String '-- locked --
        Get
            Return "http://www.platform5rd.com/revo/docs/"
        End Get
    End Property
    Public ReadOnly Property VKEY() As String '-- locked --
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\platform5\" & xProduct.ToLower & "\", "VKEY", "")
        End Get
    End Property
    Public ReadOnly Property VLIC() As String '-- locked --
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\platform5\" & xProduct.ToLower & "\", "VLIC", "")
        End Get
    End Property
    Public ReadOnly Property VMAIL() As String '-- locked --
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\platform5\" & xProduct.ToLower & "\", "VMAIL", "")
        End Get
    End Property
    Public ReadOnly Property VDATE() As String '-- locked --
        Get
            Return My.Computer.Registry.GetValue("HKEY_CURRENT_USER\Software\platform5\" & xProduct.ToLower & "\", "VDATE", "")
        End Get
    End Property
    Public ReadOnly Property VLICTYPE() As String '-- locked --
        Get
            Dim NumVer As String = Left(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("ACADVER"), 4)
            Return My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\Software\Autodesk\AutoCAD\R" & NumVer & "\ACAD-8001:40C\", "StandaloneNetworkType", "")
        End Get
    End Property
    Public ReadOnly Property VPass() As String '-- locked --
        Get
            Return "revo4fun" 'pass
        End Get
    End Property
    Public ReadOnly Property Vlogin() As String '-- locked --
        Get
            Return "http://mum.platform5rd.com/revosig" 'form nouveau login
        End Get
    End Property
    Public ReadOnly Property urlLogin() As String '-- locked --
        Get
            Return "http://mum.platform5rd.com/loginrevo" ' form login
        End Get
    End Property
    Public ReadOnly Property urluser() As String '-- locked --
        Get
            Return "http://mum.platform5rd.com/usersig.htm" 'form user info
        End Get
    End Property
    Public ReadOnly Property urlfeedback() As String '-- locked --
        Get
            Return "http://platform5rd.com/autocad-pc/revo/feedback.html" ' url du formulaire de feedback
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

#End Region
