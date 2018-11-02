
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.ApplicationServices.Application
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports System.IO


Namespace Revo

    Public Class RevoFiler

       
        Public Shared Function NameCurrentDraw()

            If Application.DocumentManager.Count <> 0 Then
                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                Dim acCurDb As Database = acDoc.Database
                Return acDoc.Name
                'AutoCAD.Application.ActiveDocument.
            Else
                Return ""
            End If



        End Function
        Public Shared Function UpdateDraw()

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
           Dim curDwg As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim curDWG As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
            Dim Ass As New RevoInfo
            Dim AcDocName As String = ""

            '' Specify the template to use, if the template is not found
            '' the default settings are used.
            Dim strTemplatePath As String = Ass.Template
            Dim acDocMgr As DocumentCollection = Application.DocumentManager

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
            Dim acDoc As Document = acDocMgr.Add(strTemplatePath)
#Else 'Versio AutoCAD 2013 et +
            Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.DocumentCollectionExtension.Add(acDocMgr, strTemplatePath)
#End If
            ' acDocMgr.MdiActiveDocument = acDoc
            AcDocName = acDoc.name
            ' acDoc.dispose
            Return AcDocName
        End Function


        Public Shared Function DisableDraw(Optional ByVal FileName As String = "")

            If Application.DocumentManager.Count = 0 Then ' si 0 doc ouvert OK
                Return True
            Else                                         ' sinon cherche le fichier est Save+Close
                Dim acDocMgr As DocumentCollection = Application.DocumentManager
                Dim TestOpen As Boolean = True

                'Return False

                For Each doc As Document In DocumentManager
                    If FileName.ToUpper = doc.Name.ToUpper Then
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                        doc.CloseAndSave(doc.Name)
#Else 'Versio AutoCAD 2013 et +
                        Autodesk.AutoCAD.ApplicationServices.DocumentExtension.CloseAndSave(doc, FileName)
#End If
                        Return True
                        Exit For
                    End If
                Next

                For Each doc As Document In DocumentManager
                    If FileName.ToUpper = doc.Name.ToUpper Then
                        TestOpen = False
                        Exit For
                    End If
                Next

                If TestOpen = True Then Return True

                Return True


            End If

        End Function

        Public Shared Function RenFile(ByVal FileName As String, ByVal NewName As String)

            'Dim Nombak As String = Mid(FileName, 1, FileName.Length - 4) & "_Revo.dwg"

            Try
                If IO.File.Exists(FileName) Then
                    Dim Fi As IO.FileInfo = New IO.FileInfo(FileName)
                    Fi.IsReadOnly = False
                    If IO.File.Exists(NewName) Then IO.File.Delete(NewName)
                    IO.File.Move(FileName, NewName)
                End If

                Return NewName
            Catch ExC As IOException
                Return "" 'ExC.Message
            Catch ExB As System.UnauthorizedAccessException
                Return "" 'ExB.Message
            Catch ExC As Exception
                Return "" 'ExC.Message
            End Try

        End Function

        Public Shared Function CopyBak(ByVal FileName As String)

            Dim Nombak As String = Mid(FileName, 1, FileName.Length - 4) & "_Revo.dwg"
            
            Try
                If IO.File.Exists(Nombak) Then
                    Dim Fi As IO.FileInfo = New IO.FileInfo(Nombak)
                    Fi.IsReadOnly = False
                    IO.File.Delete(Nombak)
                End If

                IO.File.Copy(FileName, Nombak)
                Return Nombak
            Catch ExB As System.UnauthorizedAccessException
                Return "" 'ExB.Message
            Catch ExC As Exception
                Return "" 'ExC.Message
            End Try

        End Function
        Public Shared Function ActiveDraw(ByVal FileName As String)

            Dim TestManager As Boolean = False

            For Each doc As Document In DocumentManager
                If FileName = doc.Name Then
                    TestManager = True ' Fichier trouvé
                    If doc.IsActive = False Then
                        Try
                            Application.DocumentManager.MdiActiveDocument = doc
                        Catch Err As Exception ' eDocumentSwitchDisabled
                            Return False
                        End Try
                    End If

                    Return True
                    Exit Function
                End If
            Next


            If TestManager = False Then
                If File.Exists(FileName) Then
                    Dim Fi As IO.FileInfo = New IO.FileInfo(FileName)
                    Fi.IsReadOnly = False
                    Dim acDocMgr As DocumentCollection = Application.DocumentManager
                    Try
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                        acDocMgr.Open(FileName, False)
#Else 'Versio AutoCAD 2013 et +
                        Dim acDoc As Document
                        acDoc = Autodesk.AutoCAD.ApplicationServices.DocumentCollectionExtension.Open(acDocMgr, FileName, False)
#End If
                        Return True
                    Catch 'ex As Exception
                        Return False
                    End Try
                Else
                    Return False
                End If
            Else
                Return False
            End If

            'Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            'Dim strDWGName As String = acDoc.Name
            'Dim obj As Object = Application.GetSystemVariable("DWGTITLED

        End Function

        Public Shared Function NewDraw(TemplateName As String) As Document

            Dim acDocMgr As DocumentCollection = Application.DocumentManager
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                        Dim acDoc3 As Document = acDocMgr.Add(TemplateName)
#Else 'Versio AutoCAD 2013 et +
            Dim acDoc3 As Document = Autodesk.AutoCAD.ApplicationServices.DocumentCollectionExtension.Add(acDocMgr, TemplateName)
#End If

            Return acDoc3


        End Function


        Public Shared Sub SaveAsDraw(DrawName As String, Optional Version As DwgVersion = DwgVersion.Current)
            Try
                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument

                Using myTrans As Transaction = acDoc.TransactionManager.StartTransaction
                    '' Save the active drawing
                    acDoc.Database.SaveAs(DrawName, True, Version, acDoc.Database.SecurityParameters)
                End Using

            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try

        End Sub

        Public Shared Function NameOpenDraw()

            If Application.DocumentManager.Count <> 0 Then
                Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                ' Dim acCurDb As Database = acDoc.Database
                Dim acDocMgr As DocumentCollection = Application.DocumentManager
                Dim StrFiles As String 'Dim ListingName

                StrFiles = ""

                For Each doc As Document In DocumentManager
                    StrFiles = StrFiles & "|" & doc.Name
                Next

                Return StrFiles
            Else
                Return ""
            End If

        End Function

        Public Shared Function NameDescScript(ByVal FileName As String)
            Dim Title As Boolean = False
            Dim Description As Boolean = False
            Dim Groupe As Boolean = False
            Dim StrFileDesc(0 To 2) As String

            Dim FichierALire As String = FileName

            'Verification de l'existance du FichierALire
            If System.IO.File.Exists(FichierALire) Then
                Try
                    Dim sr As StreamReader = New StreamReader(FichierALire, System.Text.Encoding.Default)
                    Dim ligne As String
                    '--- Traitement du fichier ligne par ligne
                    While Not sr.EndOfStream()
                        ligne = sr.ReadLine()
                        ' MsgBox(ligne)
                        'TODO : CODE TRAITEMENT FICHIER

                        ' DINO CODE
                        ' title ; Geobat
                        ' description;Mise en forme selon norme Geobat au format Cadastral
                        ' groupe;Revo
                        If Mid(ligne.ToUpper, 1, 7) = "#TITLE;" Then StrFileDesc(0) = ligne : Title = True
                        If Mid(ligne.ToUpper, 1, 13) = "#DESCRIPTION;" Then StrFileDesc(1) = ligne : Description = True
                        If Mid(ligne.ToUpper, 1, 8) = "#GROUPE;" Then StrFileDesc(2) = ligne : Groupe = True

                        If Title = True And Description = True And Groupe = True Then Exit While

                    End While
                    '--- Referme StreamReader
                    sr.Close()

                Catch ex As Exception
                    'Traitement de l'exception sinon :
                    Throw ex
                End Try
            Else
                ' MsgBox("fichier " & FichierALire & " inexistant", MsgBoxStyle.Critical, " -- ! -- ")
                Return StrFileDesc(0) = "ErrFile"
            End If

            Return StrFileDesc

        End Function


        Public Shared Function EcritureFichier(ByVal Fichier As String, ByVal Txts() As String, Optional ByVal Mode As Boolean = True)

            Try
                'Instanciation du StreamWriter avec passage du nom du fichier 
                Dim monStreamWriter As StreamWriter = New StreamWriter(Fichier, append:=Mode)

                For Each Txt As String In Txts
                    'Ecriture du texte dans votre fichier
                    monStreamWriter.WriteLine(Txt)
                Next

                'Fermeture du StreamWriter (Trés important)
                monStreamWriter.Close()

                Return True
            Catch 'ex As Exception
                'Code exécuté en cas d'exception
                'MsgBox("Impossible d'écrire dans le fichier:" & vbCrLf & Fichier, MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Erreur d'écriture")
                Return False
            End Try

        End Function


        Public Shared Function OpenExe(ByVal FileName As String)

            Dim Processus As New System.Diagnostics.Process

            If System.IO.File.Exists(FileName) Then

                Processus.StartInfo.FileName = FileName
                'Processus.StartInfo.Arguments = "monargument"
                Processus.Start()
                Processus.Close()
                Return True
            Else
                Return False
            End If


        End Function

        Public Shared Sub CopyDirectory(ByVal sourcePath As String, ByVal destinationPath As String, ByVal OverWrite As Boolean)
            CopyDirectory(New DirectoryInfo(sourcePath), New DirectoryInfo(destinationPath), OverWrite)
        End Sub

        Private Shared Sub CopyDirectory(ByVal source As DirectoryInfo, ByVal destination As DirectoryInfo, ByVal OverWrite As Boolean)
            destination.Create()

            For Each file As FileInfo In source.GetFiles()
                Try
                    file.CopyTo(Path.Combine(destination.FullName, file.Name), OverWrite)
                Catch
                End Try
            Next

            For Each subDirectory As DirectoryInfo In source.GetDirectories()
                CopyDirectory(subDirectory, destination.CreateSubdirectory(subDirectory.Name), OverWrite)
            Next
        End Sub

        Public Shared Function CheckRunApp(ByVal AppName As String)

            Dim Count As Integer = 0
            Try
                For Each Processus As Process In System.Diagnostics.Process.GetProcesses ' Process.GetProcesses()
                    '  MsgBox(Processus.ProcessName)
                    If Processus.ProcessName.EndsWith(AppName) Then
                        Count += 1
                        ' MsgBox(Processus.Id.ToString)
                        'With Processus
                        '.Kill()
                        '.Close() 'Libération des ressources
                        'End With
                    End If
                Next
            Catch
                'Log.Text &= "[" & My.Computer.Clock.GmtTime.ToString & "] Erreur ArreterApplication: " & ex.Message & vbCrLf
            End Try

            Return Count

        End Function


    End Class
End Namespace