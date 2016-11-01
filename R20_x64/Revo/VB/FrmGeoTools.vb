Imports System.Runtime
Imports Autodesk.AutoCAD.Runtime
Imports VB = Microsoft.VisualBasic
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry


Public Class FrmGeoTools
    'auto-enable our toolpalette for AutoCAD
    Implements Autodesk.AutoCAD.Runtime.IExtensionApplication
    Private ActiveZoom As Boolean = True

    Public Sub Initialize() Implements Autodesk.AutoCAD.Runtime.IExtensionApplication.Initialize
        txtBFfound.Focus()
    End Sub
    Public Sub Terminate() Implements Autodesk.AutoCAD.Runtime.IExtensionApplication.Terminate
        txtBFfound.Focus()
    End Sub

    Public intInsTxtTheme, intInsTxtCat, intInsTxtType As Short
    Public Sub LoadVariable()

        txtBFfound.Focus()

    End Sub
  
    
    Private Sub XdataBat_Click(sender As System.Object, e As System.EventArgs) Handles XdataBat.Click

        Dim MyCmd As New Revo.MyCommands
        MyCmd.RevoBatXdata()

    End Sub

    Private Sub AnalyseBF_Click(sender As System.Object, e As System.EventArgs) Handles AnalyseBF.Click

        Dim MyCmd As New Revo.MyCommands
        MyCmd.RevoCheckMO()

    End Sub


    Private Sub txtBFfound_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles txtBFfound.KeyDown
        If e.KeyCode = System.Windows.Forms.Keys.Space Then

            FindBF()
            txtBFfound.Text = Trim(txtBFfound.Text)
        End If
    End Sub
   
    Private Sub btnBFfound_Click(sender As System.Object, e As System.EventArgs) Handles btnBFfound.Click

        If InStr(txtBFfound.Text, "{") = 0 Then
            FindBF()
            txtBFfound.Text = Trim(txtBFfound.Text)
        Else

            Dim db As Database = HostApplicationServices.WorkingDatabase
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Using trans As Transaction = db.TransactionManager.StartTransaction
                Dim Num As String = Trim(txtBFfound.Text)
                doc.LockDocument()

                If InStr(txtBFfound.Text, "(") > 0 Then ' OBJECT ID

                    ''Num = Replace(Replace(Trim(txtBFfound.Text), "(", ""), ")", "")
                    'Dim ID As ObjectId = New ObjectId(Num)  GetHandle(ObjectId blockID)
                    'Zooming.ZoomToObject(ID, 10)
                    'Dim obj As Entity = trans.GetObject(ID, OpenMode.ForWrite)
                    'obj.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.None, 1)
                Else ' HANDLE ID
                    Num = Replace(Replace(Trim(txtBFfound.Text), "{", ""), "}", "")
                    Dim ID As ObjectId = db.GetObjectId(False, New Handle(Convert.ToInt64(Num, 16)), 0)
                    Zooming.ZoomToObject(ID, 10)
                    Dim obj As Entity = trans.GetObject(ID, OpenMode.ForWrite)
                    obj.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.None, 1)
                End If

                trans.Commit()
            End Using
            doc.LockDocument.Dispose()
        End If

    End Sub

    Private Function FindBF()

        ActiveZoom = False

        Dim FindStr As String = Trim(txtBFfound.Text)

        If FindStr <> "" Then
            FindStr = Replace(FindStr, " ", "")
            FindStr = Replace(FindStr, "DP", "")
            FindStr = Replace(FindStr, "(", "")
            FindStr = Replace(FindStr, ")", "")
            If IsNumeric(Trim(FindStr)) Then
            Else
                Return False
            End If
        Else
            Return False
        End If

        FindStr = Trim(FindStr)
        ListFind.Items.Clear()

        lblFindBF.Text = "Recherche en cours ... "

        'Recherche parcelle
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim NbrePoly As Double = 0
        Dim NbreTrouve As Double = 0
        Dim BFColl As New Collection
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim BestIndex As Double = 0

        Try
            doc.LockDocument()
            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.BlockName, "*IMMEUBLE_NUM*")}
                Dim sFilter As New SelectionFilter(Values)
                Dim pres As PromptSelectionResult = ed.SelectAll(sFilter)

                If pres.Status = PromptStatus.OK Then
                    Dim ids() As ObjectId = pres.Value.GetObjectIds
                    NbrePoly = ids.Length

                    lblFindBF.Text = "Recherche en cours ...  " & ids.Length

                    For i = 0 To ids.Length - 1

                        System.Windows.Forms.Application.DoEvents() ' Donne le contrôle à d'autres processus.

                        Dim NumParc As String = ""
                        Dim IDParc As String = ""

                        'Start a transaction
                        'Get the blockreference
                        Dim blkRef As BlockReference = CType(tr.GetObject(ids(i), OpenMode.ForRead), BlockReference)
                        'Get the attributecollection
                        Dim attCol As AttributeCollection = blkRef.AttributeCollection
                        'Loop through the attributes
                        For Each attid As ObjectId In attCol
                            'Get the attribute reference
                            Dim attRef As AttributeReference = CType(tr.GetObject(attid, OpenMode.ForRead), AttributeReference)
                            'Check which attribute this is
                            Select Case attRef.Tag.ToUpper
                                Case "NUMERO"
                                    'Store this value in the Num propety of the return class
                                    NumParc = attRef.TextString
                                    IDParc = blkRef.Handle.ToString
                                    Exit For
                            End Select
                        Next

                        blkRef.Dispose()


                        'Recherche Attribut si numéro

                        If Trim(NumParc.ToUpper) Like "*" & Trim(FindStr.ToUpper) & "*" Then
                            ListFind.Items.Add(NumParc & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & IDParc)
                            If Trim(NumParc.ToUpper) = Trim(FindStr.ToUpper) Then BestIndex = NbreTrouve
                            NbreTrouve += 1
                        End If

                    Next

                End If


            End Using

            doc.LockDocument.Dispose()

        Catch ex As System.Exception
            MsgBox("Erreur d'analyse : " & ex.Message)
        End Try


        lblFindBF.Text = "Recherche de parcelle (BF)" & vbCrLf & "(Espace pour lancer)"

        If NbreTrouve = 0 Then
            ListFind.Items.Add("Aucune parcelle trouvée")
            Return False
        Else
            'zoom
          
            If ListFind.Items.Count >= 1 Then
                ListFind.SelectedIndex = BestIndex
                ZoomNum()
            End If

        End If

        ActiveZoom = True

        Return True
    End Function

    Private Sub ListFind_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListFind.MouseClick

        '  ZoomNum()

    End Sub

    Private Sub ListFind_MouseDoubleClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles ListFind.MouseDoubleClick

        '    ZoomNum()

    End Sub

    Private Sub ListFind_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles ListFind.SelectedIndexChanged

        If ActiveZoom = True Then
            ZoomNum()
        End If


    End Sub

    Private Function ZoomNum()

        Dim Num As String = ""
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument

        Try
            doc.LockDocument()
            Num = ListFind.Items(ListFind.SelectedIndex).ToString

            If Num <> "" Then

                Num = Replace(Num, vbTab & vbTab & vbTab & vbTab & vbTab & vbTab, vbTab)
                Dim NumID() As String = Split(Num, vbTab)

                If NumID.Length = 2 Then

                    Dim ID As ObjectId = db.GetObjectId(False, New Handle(Convert.ToInt64(NumID(1), 16)), 0)
                    Zooming.ZoomToObject(ID, 40)

                End If

            End If

            doc.LockDocument.Dispose()

        Catch ex As System.Exception
            MsgBox("Erreur de recherche : " & ex.Message)
        End Try


        Return True

    End Function



    Private Sub Label3_Click(sender As System.Object, e As System.EventArgs)
        Dim Ass As New Revo.RevoInfo
        System.Diagnostics.Process.Start(Ass.urlDEV)
    End Sub

    Private Sub Label4_Click(sender As System.Object, e As System.EventArgs)
        Dim Ass As New Revo.RevoInfo
        System.Diagnostics.Process.Start(Ass.urlDEV)
    End Sub

End Class
