Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.Collections.Generic
Imports System.Runtime.InteropServices
Module CommandLineFunc
    Public Class CommandLine
        ' Methods
        Shared Sub New()
            ResTypes.Item(GetType(String)) = &H138D
            ResTypes.Item(GetType(Double)) = &H1389
            ResTypes.Item(GetType(Point3d)) = &H1391
            ResTypes.Item(GetType(ObjectId)) = &H138E
            ResTypes.Item(GetType(Integer)) = &H1392
            ResTypes.Item(GetType(Short)) = &H138B
            ResTypes.Item(GetType(Point2d)) = &H138A
            ResTypes.Item(GetType(Byte)) = &H138B
        End Sub


#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Private Const DLLApp As String = "acad.exe"    
#Else 'Versio AutoCAD 2013 et +
        Private Const DLLApp As String = "accore.dll"
#End If

        <DllImport(DLLApp, CallingConvention:=CallingConvention.Cdecl)> _
        Private Shared Function acedCmd(ByVal resbuf As IntPtr) As Integer
        End Function


        Public Shared Function CommandC(ByVal ParamArray args As Object()) As Integer

#If _AcadVer_ < 20 Then ' Ancienne Version 2010-2011-2012-2013-2014
            Commandold(args)
#Else 'Versio AutoCAD 2015 et +
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            ed.Command(args)
#End If

        End Function

        Public Shared Function CmdsC(ByVal CmdList As String) As Integer

#If _AcadVer_ < 20 Then ' Ancienne Version 2010-2011-2012-2013-2014
            Dim Args() As String
            Args = Split(CmdList, "|")
            CmdsOld(Args)

#Else 'Version AutoCAD 2015 et + 
            Dim doc = Application.DocumentManager.MdiActiveDocument
            Dim ed = doc.Editor
            Dim Args() As Object
            Args = Split(CmdList, "|")

            Try
                'For i = 0 To CmdList.Count - 1
                '    Dim ArgObj As Object = CmdList(i)
                '    '   Args.Add(ArgObj)
                'Next

                ed.Command(Args)

            Catch ex As Exception
                If ex.Message.ToLower = "einvalidinput" Then

                    Try
                        doc.SendStringToExecute(Replace(CmdList, "|", vbCr) & vbCr, False, False, False)

                    Catch ex2 As System.Exception
                        ed.WriteMessage(vbLf & "Exception: {0}" & vbLf, ex2.Message)
                    End Try



                Else
                    MsgBox("Erreur dans la commande #cmd : " & CmdList & "  (" & ex.Message & ")")
                End If

            End Try

#End If

        End Function



        'Private Shared Async Sub CmdAsync(Args() As Object)
        '    Dim doc = Application.DocumentManager.MdiActiveDocument
        '    Dim ed = doc.Editor

        '    ' Let's ask the user to select the insertion point
        '    Try
        '        Await ed.CommandAsync(Args)
        '        ed.WriteMessage(vbLf & "The cmd Async is completed :" & Args(0))
        '    Catch
        '        ed.WriteMessage(vbLf & "The cmd Async has an Error")
        '    End Try

        'End Sub

        Public Shared Function CmdsOld(ByVal CmdList() As String) As Integer 'Commandes ajouter par THA 2011


            Dim Args As New Collection
            For i = 0 To CmdList.Count - 1
                Dim ArgObj As Object = CmdList(i)
                Args.Add(ArgObj)
            Next
            If Application.DocumentManager.IsApplicationContext Then
                Return 0
            End If

            Dim num As Integer = 0
            Dim num2 As Integer = 0
            Using buffer As ResultBuffer = New ResultBuffer
                Dim obj2 As Object
                For Each obj2 In Args
                    num2 += 1
                    buffer.Add(CommandLine.TypedValueFromObject(obj2))
                Next
                If (num2 > 0) Then
                    Dim strA As String = CStr(Application.GetSystemVariable("USERS1"))
                    Dim flag As Boolean = (String.Compare(strA, "DEBUG", True) = 0)
                    Dim num3 As Integer = Convert.ToInt32(IIf(flag, 1, 0))
                    Dim systemVariable As Object = Application.GetSystemVariable("CMDECHO")
                    Dim num4 As Short = CShort(systemVariable)
                    If ((num4 <> 0) OrElse flag) Then
                        Application.SetSystemVariable("CMDECHO", num3)
                    End If
                    num = CommandLine.acedCmd(buffer.UnmanagedObject)
                    If ((num4 <> 0) OrElse flag) Then
                        Application.SetSystemVariable("CMDECHO", systemVariable)
                    End If
                End If
            End Using
            Return num


        End Function

        Public Shared Function Cmd(ByVal args As IList) As Integer
            If Application.DocumentManager.IsApplicationContext Then
                Return 0
            End If
            Dim num As Integer = 0
            Dim num2 As Integer = 0
            Using buffer As ResultBuffer = New ResultBuffer
                Dim obj2 As Object
                For Each obj2 In args
                    num2 += 1
                    buffer.Add(CommandLine.TypedValueFromObject(obj2))
                Next
                If (num2 > 0) Then
                    Dim strA As String = CStr(Application.GetSystemVariable("USERS1"))
                    Dim flag As Boolean = (String.Compare(strA, "DEBUG", True) = 0)
                    Dim num3 As Integer = Convert.ToInt32(IIf(flag, 1, 0))
                    Dim systemVariable As Object = Application.GetSystemVariable("CMDECHO")
                    Dim num4 As Short = CShort(systemVariable)
                    If ((num4 <> 0) OrElse flag) Then
                        Application.SetSystemVariable("CMDECHO", num3)
                    End If
                    num = CommandLine.acedCmd(buffer.UnmanagedObject)
                    If ((num4 <> 0) OrElse flag) Then
                        Application.SetSystemVariable("CMDECHO", systemVariable)
                    End If
                End If
            End Using
            Return num
        End Function


        Public Shared Function Commandold(ByVal ParamArray args As Object()) As Integer
            If Application.DocumentManager.IsApplicationContext Then
                Return 0
            End If
            Dim stat As Integer = 0
            If args Is Nothing Or args.Length = 0 Then
                Using rb As ResultBuffer = New ResultBuffer()
                    rb.Add(New TypedValue(5000))
                    Return acedCmd(rb.UnmanagedObject)
                End Using
            End If
            Dim cnt As Integer = 0
            Using buffer As ResultBuffer = New ResultBuffer
                Dim o As Object
                For Each o In args
                    cnt += 1
                    buffer.Add(CommandLine.TypedValueFromObject(o))
                Next
                If (cnt > 0) Then
                    Dim s As String = CStr(Application.GetSystemVariable("USERS1"))
                    Dim debug As Boolean = (String.Compare(s, "DEBUG", True) = 0)
                    Dim val As Integer = Convert.ToInt32(IIf(debug, 1, 0))
                    Dim cmdecho As Object = Application.GetSystemVariable("CMDECHO")
                    Dim c As Short = CShort(cmdecho)
                    If ((c <> 0) OrElse debug) Then
                        Application.SetSystemVariable("CMDECHO", val)
                    End If
                    stat = CommandLine.acedCmd(buffer.UnmanagedObject)
                    If ((c <> 0) OrElse debug) Then
                        Application.SetSystemVariable("CMDECHO", cmdecho)
                    End If
                End If
            End Using
            Return stat
        End Function





        Private Shared Function TypedValueFromObject(ByVal val As Object) As TypedValue
            Dim restype As Integer
            If Not ResTypes.TryGetValue(val.GetType, restype) Then
                Throw New InvalidOperationException("Unsupported type in Command() method")
            End If
            Return New TypedValue(restype, val)
        End Function

        ' Fields
        Private Const ACAD_EXE As String = "acad.exe"
        Private Shared ResTypes As Dictionary(Of Type, Integer) = New Dictionary(Of Type, Integer)
        Private Const RT3DPOINT As Integer = &H1391
        Private Const RTENAME As Integer = &H138E
        Private Const RTLONG As Integer = &H1392
        Private Const RTNONE As Integer = &H1388
        Private Const RTNORM As Integer = &H13EC
        Private Const RTPOINT As Integer = &H138A
        Private Const RTREAL As Integer = &H1389
        Private Const RTSHORT As Integer = &H138B
        Private Const RTSTR As Integer = &H138D
    End Class


End Module

Class CommandsAsync


    'Public Async Sub CmdsAsync(Args() As Object)
    '    Try
    '        'For Each cmd In Args
    '        '    Dim tl = CmdAsync("FLATTEN")
    '        '    Console.WriteLine("---> commande FLATTEN")
    '        'Next

    '        '  Dim dt As DataTable = Await CmdAsync(Args)
    '        ' Dim cmdResult As Autodesk.AutoCAD.EditorInput.Editor.CommandResult = CmdAsync(Args)
    '        ' cmdResult.OnCompleted(AddressOf Continuation)


    '        Console.WriteLine("End CmdsAsync")
    '    Catch ex As Exception
    '        Throw ex
    '    End Try
    'End Sub

    'Private Async Function CmdAsync(Cmd As Object) As Threading.Tasks.Task 'As Threading.Tasks.Task(Of Autodesk.AutoCAD.EditorInput.Editor.CommandResult)

    '    ' Dim doc = Application.DocumentManager.MdiActiveDocument
    '    Dim ed As Autodesk.AutoCAD.EditorInput.Editor = Application.DocumentManager.MdiActiveDocument.Editor


    '    Console.WriteLine("2 Async Started")

    '    '  Dim cmdResult As Autodesk.AutoCAD.EditorInput.Editor.CommandResult = ed.CommandAsync("FLATTEN")
    '    '  cmdResult.OnCompleted(AddressOf Continuation)

    '    '  Await ed.CommandAsync("FLATTEN")

    '    Console.WriteLine("3 Async endend")

    '    ' Return cmdResult
    'End Function


    Private Sub Continuation()
        'add some code here if you want
    End Sub





    Private Shared ZoomExit As Boolean = False
    'declare the callback delegation
    Friend Delegate Sub Del()
    Private Shared Function _actionCompletedDelegate() As Del

        Return New Del(AddressOf CreateZoomAsyncCallback)
    End Function


    ' Exit function，check if Zoom command is esc\cancelled
    Shared Sub MdiActiveDocument_CommandCancelled(ByVal sender As Object, ByVal e As CommandEventArgs)
        ZoomExit = True
    End Sub

    Public Sub TestZoom()
        Dim ed = Application.DocumentManager.MdiActiveDocument.Editor
        Dim doc = Application.DocumentManager.MdiActiveDocument
        'esc event
        AddHandler doc.CommandCancelled, AddressOf MdiActiveDocument_CommandCancelled
        ' start Zoom command
        Dim cmdResult1 As Autodesk.AutoCAD.EditorInput.Editor.CommandResult = ed.CommandAsync(New Object() {"_.ZOOM", Autodesk.AutoCAD.EditorInput.Editor.PauseToken, Autodesk.AutoCAD.EditorInput.Editor.PauseToken})
        ' delegate callback function, wait for interaction ends
        ' _actionCompletedDelegate = New Del(AddressOf CreateZoomAsyncCallback)
        cmdResult1.OnCompleted(New Action(AddressOf _actionCompletedDelegate))
        ZoomExit = False
    End Sub
    ' callback function
    Public Shared Sub CreateZoomAsyncCallback()
        Dim ed = Application.DocumentManager.MdiActiveDocument.Editor
        'if Zoom command is running
        If Not ZoomExit Then
            ' AutoCAD hands over to the callback function
            Dim cmdResult1 As Autodesk.AutoCAD.EditorInput.Editor.CommandResult = ed.CommandAsync(New Object() {"_.ZOOM", Autodesk.AutoCAD.EditorInput.Editor.PauseToken, Autodesk.AutoCAD.EditorInput.Editor.PauseToken})
            ' delegate callback function, wait for interaction ends
            '_actionCompletedDelegate = New Del(AddressOf CreateZoomAsyncCallback)
            cmdResult1.OnCompleted(New Action(AddressOf _actionCompletedDelegate))
        Else
            ed.WriteMessage("Zoom Exit")
            Return
        End If
    End Sub





End Class


'Namespace ContextMenuApplication

'    Public Class Commands
'        Implements Autodesk.AutoCAD.Runtime.IExtensionApplication

'        Public Sub Initialize()
'            ScaleMenu.Attach()
'        End Sub

'        Public Sub Terminate()
'            ScaleMenu.Detach()
'        End Sub
'    End Class

'    Public Class ScaleMenu

'        Private Shared cme As Autodesk.AutoCAD.Windows.ContextMenuExtension

'        Public Shared Sub Attach()
'            If (cme Is Nothing) Then
'                cme = New Autodesk.AutoCAD.Windows.ContextMenuExtension
'                Dim mi As Autodesk.AutoCAD.Windows.MenuItem = New Autodesk.AutoCAD.Windows.MenuItem("Scale by 5")
'                AddHandler mi.Click, AddressOf Me.OnScale
'                cme.MenuItems.Add(mi)
'            End If

'            Dim rxc As Autodesk.AutoCAD.Runtime.RXClass = Entity.GetClass(GetType(Entity))
'            Application.AddObjectContextMenuExtension(rxc, cme)
'        End Sub

'        Public Shared Sub Detach()
'            Dim rxc As Autodesk.AutoCAD.Runtime.RXClass = Entity.GetClass(GetType(Entity))
'            Application.RemoveObjectContextMenuExtension(rxc, cme)
'        End Sub

'        Private Shared Async Sub OnScale(ByVal o As Object, ByVal e As EventArgs)
'            Dim dm = Application.DocumentManager
'            Dim doc = dm.MdiActiveDocument
'            Dim ed = doc.Editor
'            ' Get the selected objects
'            Dim psr = ed.GetSelection
'            If (psr.Status <> Autodesk.AutoCAD.EditorInput.PromptStatus.OK) Then
'                Return
'            End If

'            Try

'                Await dm.ExecuteInCommandContextAsync(Function(obj)
'                                                          ' Scale the selected objects by 5 relative to 0,0,0

'                                                          Await ed.CommandAsync("._scale", psr.Value, "", Point3d.Origin, 5)
'End Function, Nothing)


'            Catch ex As System.Exception
'                ed.WriteMessage("" & vbLf & "Exception: {0}" & vbLf, ex.Message)
'            End Try


'            Try


'                ' Ask AutoCAD to execute our command in the right context
'                ' Scale the selected objects by 5 relative to 0,0,0
'                '' ed.CommandAsync("._scale", psr.Value, "", Point3d.Origin, 5)
'                'Warning!!! Lambda constructs are not supported
'                ''     dm.ExecuteInCommandContextAsync(() >= {}, Nothing)

'            Catch ex As System.Exception
'                ed.WriteMessage("" & vbLf & "Exception: {0}" & vbLf, ex.Message)
'            End Try

'        End Sub
'    End Class
'End Namespace