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
            Try
                'For i = 0 To CmdList.Count - 1
                '    Dim ArgObj As Object = CmdList(i)
                '    '   Args.Add(ArgObj)
                'Next
                Args = Split(CmdList, "|")
                ed.Command(Args)

            Catch ex As Exception
                MsgBox("Erreur dans la commande #cmd : " & CmdList & "  (" & ex.Message & ")")
            End Try

#End If

        End Function


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
