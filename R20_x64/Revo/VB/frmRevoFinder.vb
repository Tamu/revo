Imports System.IO
Imports System.Windows.Forms
Imports System.Drawing
Imports System
Imports System.Xml
Imports Microsoft

Public Class frmRevoFinder

    Dim Connect As New Revo.connect
    Dim Ass As New Revo.RevoInfo
    Dim PurgeFiles As Boolean
    Public MinFormWidth As Double = 329
    Dim MaxFormWidth As Double = 750
    Public Preflist() As String
   
    Private Sub frm_RevoFinder_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.AcceptButton = btnStart

        'Active l'auto purge des fichiers a traiter
        PurgeFiles = True
        Revo.MyCommands.ReloadMenu = False
       
        'Paramètre personalisé
        If Preflist.Count < 5 Then Preflist = Connect.PrefUser()
        If Preflist(0) = "1" Then btnLog.BackgroundImage = My.Resources.bouton_revo_app_active
        If Preflist(1) = "1" Then btnDWG.BackgroundImage = My.Resources.bouton_revo_app_active
        If Preflist(2) = "1" Then btnDXF.BackgroundImage = My.Resources.bouton_revo_app_active
        If Preflist(3) = "1" Then btnPTS.BackgroundImage = My.Resources.bouton_revo_app_active
        If Preflist(4) = "1" Then btnITF.BackgroundImage = My.Resources.bouton_revo_app_active


        If btnDWG.Visible = False Then Revo.connect.ActDWG = False
        If btnDXF.Visible = False Then Revo.connect.ActDXF = False
        If btnPTS.Visible = False Then Revo.connect.ActPTS = False
        If btnITF.Visible = False Then Revo.connect.ActITF = False

        If Revo.connect.ActDWG = False And Revo.connect.ActDXF = False And _
            Revo.connect.ActPTS = False And Revo.connect.ActITF = False Then
            '  Les formats peuvent être activés via les boutons en bas à gauche de l'outil Revo.
            '  Activer le filtre DWG ?
            btnDWG.BackgroundImage = My.Resources.bouton_revo_app_active
            Revo.connect.ActDWG = True
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActDWG", "1")
        End If

        Me.GridData.AllowDrop = True
        Me.GridCmd.AllowDrop = True
        Me.GridImport.AllowDrop = True

        '  AddRowGridData(Revo.RevoFiler.NameCurrentDraw())
        UpdateGridDataCount()
        UpdateGridImportCount()


        Me.Text = Ass.xTitle
        Me.Text = Ass.xTitle & "         " & Mid(Trim(Ass.xVersion), 1, 7)
        Me.Icon = Ass.Icon


        'Chargement des fichier script CSV : revo-perso.xml à distance (lecture seul)
        If Ass.PluginSharedXML(True).ToUpper <> Ass.PluginSharedXML(False).ToUpper Then
            Dim xcld As New RevoXML(Ass.PluginSharedXML)
            For i = 1 To xcld.countElements("/" & Ass.xProduct.ToLower & "/cmdFiles/url")
                AddRowGridCmd(xcld.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url", i), True)
            Next
        End If
        'Chargement des fichier script CSV : revo-perso.xml en local
        Dim x As New RevoXML(Ass.PluginPersoXML)
        For i = 1 To x.countElements("/" & Ass.xProduct.ToLower & "/cmdFiles/url")
            AddRowGridCmd(x.getElementValue("/" & Ass.xProduct.ToLower & "/cmdFiles/url", i), False)
        Next


        'Hide importation browser
        If btnPTS.Visible = False And btnITF.Visible = False Then btnDataBrows.Visible = False

        TableLayoutFlipBrowser.RowStyles.Item(0).SizeType = SizeType.Percent
        TableLayoutFlipBrowser.RowStyles.Item(0).Height = 100
        TableLayoutFlipBrowser.RowStyles.Item(1).SizeType = SizeType.Absolute
        TableLayoutFlipBrowser.RowStyles.Item(1).Height = 36


    End Sub

    Private Sub frm_RevoFinder_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        ' e.Cancel = True
        '  MsgBox("Super")
        If e.CloseReason = 3 Then Revo.MyCommands.StatFrmRevoFinder = False
    End Sub

    Private Sub DataGridView_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridData.DragEnter

        'Si le drop en question est un drop de fichiers,
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' alors on accepte le drop sous forme de copyDrop
            e.Effect = DragDropEffects.Copy
        Else 'sinon
            'on accepte pas
            e.Effect = DragDropEffects.None
        End If
    End Sub

    
    Private Sub btnStart_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStart.Click

        'Test si aucun fichier => Ajoute
        If GridData.RowCount = 0 Then AddRowGridData(Revo.RevoFiler.NameCurrentDraw())

        Revo.MyCommands.StatFrmRevoFinder = True
        Me.Hide() '    StartScript()
    End Sub


    Private Sub GridCmd_CellContentDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridCmd.CellContentDoubleClick

        'Test si aucun fichier => Ajoute
        If GridData.RowCount = 0 Then AddRowGridData(Revo.RevoFiler.NameCurrentDraw())

        Revo.MyCommands.StatFrmRevoFinder = True
        Me.Hide() '  StartScript()
    End Sub

    Private Sub DataGridView_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridData.DragDrop
        Dim strFiles() As String
        'Variable qui contiendra un tableau contenant les fichiers
        Dim i As Long 'Variable boucle

        'on recupere le drop dans le tableau
        strFiles = e.Data.GetData(DataFormats.FileDrop)

        For i = 0 To strFiles.GetUpperBound(0)
            AddRowGridData(strFiles(i))
        Next

        UpdateGridDataCount()
    End Sub

    Private Sub GridData_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GridData.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContextMenuStrip1.Show(CType(sender, Control), e.Location)
        End If

    End Sub

    Private Sub SelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAll.Click

        For i = 0 To Me.GridData.Rows.Count - 1
            Me.GridData.Rows(i).Selected = (True)
        Next
    End Sub

    Private Sub EffacerLesFichiersSélectionnerToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EffacerLesFichiersSélectionnerToolStripMenuItem.Click
        Dim TotalSelect As Double = 0

        For i = 1 To Me.GridData.SelectedRows.Count 'TotalSelect
            DeleteRow()
        Next

    End Sub
    Private Function DeleteRow()

        For i = 0 To Me.GridData.Rows.Count - 1
            If Me.GridData.Rows(i).Selected = True Then
                Me.GridData.Rows.RemoveAt(i)
                Exit For
            End If
        Next

        Return "ok"
    End Function

    Private Sub btnFiles_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFiles.Click
        Dim strFiles() As String
        Dim XMLformat As String = Ass.XMLformatBase
        Dim NbreFormat As Double = 0
        'Variable qui contiendra un tableau contenant les fichiers
        Dim i As Long 'Variable boucle
        Dim ConfirmOK As Double
        Dim Filters As String = ""

        ' Attention en fonction des boutons décalage dans la listes
        ' (1) Formats compatibles
        ' (2) Dessin (*.dwg)
        ' (3) DXF (*.dxf)
        ' (4) Interlis (*.itf)
        ' (x) Points ...

        If Revo.connect.ActDWG = True Then Filters += ";*.dwg"
        If Revo.connect.ActDXF = True Then Filters += ";*.dxf"
        'If Revo.connect.ActITF = True Then Filters += ";*.itf"

        'If File.Exists(XMLformat) And Revo.connect.ActPTS Then 'Test de l'existance du fichier XML Revo
        'Dim x As New RevoXML(XMLformat) 'lecture des extentions
        'NbreFormat = x.countElements("/data-format/format")
        'For i = 1 To NbreFormat
        'Filters += ";*." & x.getElementValue("/data-format/format/extension", i)
        'Next
        'End If
        If Mid(Filters, 1, 1) = ";" Then Filters = Mid(Filters, 2, Len(Filters) - 1)
        If Filters = "" Then Filters = "-/-"
        Dim FormatSelect As Double = 1

        Filters = "Formats compatibles|" & Filters
        If Revo.connect.ActDWG = True Then Filters += "|Dessin (*.dwg)|*.dwg" : FormatSelect += 1
        If Revo.connect.ActDXF = True Then Filters += "|DXF (*.dxf)|*.dxf" : FormatSelect += 1
        'If Revo.connect.ActITF = True Then Filters += "|Interlis (*.itf)|*.itf" : FormatSelect += 1

        'Importation des formats points : format.xml **********
        'If File.Exists(XMLformat) And Revo.connect.ActPTS Then 'Test de l'existance du fichier XML Revo
        'Dim x As New RevoXML(XMLformat) 'lecture des extentions
        'Dim VarExt As String = ""
        'For i = 1 To NbreFormat
        'VarExt = x.getElementValue("/data-format/format/extension", i)
        'Filters += "|" & x.getElementValue("/data-format/format/name", i) & " (*." & VarExt & ")" & "|*." & VarExt
        'Next
        'End If

        ' If PurgeFiles = True Then DeleteRow()

        OpenFileDialog1.Filter = Filters
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.FileName = ""
        ConfirmOK = OpenFileDialog1.ShowDialog()

        If ConfirmOK = 1 Then
            Dim NumFormat As Double
            NumFormat = OpenFileDialog1.FilterIndex - FormatSelect
            If NumFormat < 0 Then NumFormat = 0
            strFiles = OpenFileDialog1.FileNames
            For i = 0 To strFiles.GetUpperBound(0)
                AddRowGridData(strFiles(i), NumFormat)
            Next
        End If

        UpdateGridDataCount()
        Me.GridData.Focus()
    End Sub



    Private Sub btnFolders_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFolders.Click
        ' Dim strFiles() As String
        'Variable qui contiendra un tableau contenant les fichiers
        Dim ConfirmOK As Double
        'Dim instance As UnauthorizedAccessException


        Windows.Forms.Cursor.Current = Cursors.WaitCursor ' Curseur en attente
        'Windows.Forms.Cursor.Current = Cursors.Default ' Curseur en attente

        FolderBrowserDialog1.Description = "Sélectionner un dossier"
        ConfirmOK = (FolderBrowserDialog1.ShowDialog())
        '        MsgBox(FolderBrowserDialog1.)

        If ConfirmOK = 1 Then


            Dim dir As String = FolderBrowserDialog1.SelectedPath

            If Directory.Exists(dir) = True Then
                Try

                    Dim FilterList As New List(Of String)
                    If Revo.connect.ActDWG Then FilterList.Add("*.DWG")
                    If Revo.connect.ActDXF Then FilterList.Add("*.DXF")
                    Dim files As List(Of FileInfo) = GetMultiFiles(dir, FilterList)
                    For Each file As FileInfo In files
                        AddRowGridData(file.FullName)
                    Next

                Catch Ex As UnauthorizedAccessException
                    MsgBox(Ex.Message)
                Finally
                End Try

            End If
        End If

        UpdateGridDataCount()
        Me.GridData.Focus()

        Windows.Forms.Cursor.Current = Cursors.Default ' Curseur en attente
    End Sub

    Private Function GetMultiFiles(ByVal Path As String, ByVal FilterList As List(Of String)) As List(Of FileInfo)

        Dim d As New DirectoryInfo(Path)
        Dim files As List(Of FileInfo) = New List(Of FileInfo)

        'Iterate through the FilterList
        For Each Filter As String In FilterList

            'the files are appended to the file array
            files.AddRange(d.GetFiles(Filter, SearchOption.AllDirectories))
        Next

        Return files

    End Function

    Private Sub GridData_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridData.CellContentClick
        UpdateGridDataCount()
    End Sub

    Private Sub GridData_RowStateChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowStateChangedEventArgs) Handles GridData.RowStateChanged
        UpdateGridDataCount()
    End Sub

    Private Function AddRowGridData(ByVal FileName As String, Optional ByVal FormatPts As Double = 0)

        Windows.Forms.Cursor.Current = Cursors.WaitCursor ' Curseur en attente
        If FileName <> "" Then
            Dim DWG As Boolean = Revo.connect.ActDWG
            Dim DXF As Boolean = Revo.connect.ActDXF
            Dim NameFile As String
            Dim NamePath As String

            NameFile = Path.GetFileName(FileName)
            NamePath = Path.GetDirectoryName(FileName)

            If Directory.Exists(FileName) = True Then
                Try
                    'charge les sous fichier
                    Dim FilterList As New List(Of String)
                    If DWG Then FilterList.Add("*.DWG")
                    If DXF Then FilterList.Add("*.DXF")
                    Dim files As List(Of FileInfo) = GetMultiFiles(FileName, FilterList)
                    For Each file As FileInfo In files
                        AddRowGridData(file.FullName)
                    Next

                Catch Ex As UnauthorizedAccessException
                    MsgBox(Ex.Message)
                End Try

            ElseIf VisualBasic.Right(NameFile.ToUpper, 4) = ".DWG" And DWG = True Then
                Me.GridData.Rows.Add(My.Resources.icon_dwg, NameFile, NamePath, FormatPts)
                Me.GridData.Rows(Me.GridData.Rows.Count - 1).Height = 36
            ElseIf VisualBasic.Right(NameFile.ToUpper, 4) = ".DXF" And DXF = True Then
                Me.GridData.Rows.Add(My.Resources.icon_dxf, NameFile, NamePath, FormatPts)
                Me.GridData.Rows(Me.GridData.Rows.Count - 1).Height = 36
            End If
        End If
        Return "ok"

        Windows.Forms.Cursor.Current = Cursors.Default ' Curseur en attente
    End Function

    Private Function UpdateGridDataCount()

        Dim NomTitre As String
        NomTitre = "Traitement multi-fichiers"

        If GridData.Rows.Count > 0 Then
            btnFileBrows.Text = NomTitre & "   (" & GridData.Rows.Count & " élément(s) )"
        Else
            btnFileBrows.Text = NomTitre
        End If

        Return (GridData.Rows.Count)

    End Function

    Private Function AddRowGridImport(ByVal FileName As String, Optional ByVal FormatPts As Double = 0)

        Windows.Forms.Cursor.Current = Cursors.WaitCursor ' Curseur en attente
        If FileName <> "" Then
            Dim DWG As Boolean = Revo.connect.ActDWG
            Dim DXF As Boolean = Revo.connect.ActDXF
            Dim ITF As Boolean = Revo.connect.ActITF
            Dim PTS As Boolean = Revo.connect.ActPTS
            Dim NameFile As String
            Dim NamePath As String

            NameFile = Path.GetFileName(FileName)
            NamePath = Path.GetDirectoryName(FileName)

            If Directory.Exists(FileName) = True Then
                Try

                    Dim FilterList As New List(Of String)
                    If DWG Then FilterList.Add("*.DWG")
                    If DXF Then FilterList.Add("*.DXF")
                    If ITF Then FilterList.Add("*.ITF")
                    If PTS Then
                        Dim XMLformat As String = Ass.XMLformatBase
                        If File.Exists(XMLformat) And Revo.connect.ActPTS Then 'Test de l'existance du fichier XML Revo
                            Dim x As New RevoXML(XMLformat) 'lecture des extentions
                            Dim NbreFormat As Double = x.countElements("/data-format/format")
                            For i = 1 To NbreFormat
                                FilterList.Add("*." & x.getElementValue("/data-format/format/extension", i).ToUpper)
                            Next
                        End If
                    End If

                    'charge les sous fichier
                    Dim files As List(Of FileInfo) = GetMultiFiles(FileName, FilterList)
                    For Each file As FileInfo In files
                        AddRowGridImport(file.FullName)
                    Next

                Catch Ex As UnauthorizedAccessException
                    MsgBox(Ex.Message)
                Finally
                End Try

            ElseIf VisualBasic.Right(NameFile.ToUpper, 4) = ".DWG" And DWG = True Then
                Me.GridImport.Rows.Add(My.Resources.icon_dwg, NameFile, NamePath, FormatPts)
                Me.GridImport.Rows(Me.GridImport.Rows.Count - 1).Height = 36
            ElseIf VisualBasic.Right(NameFile.ToUpper, 4) = ".DXF" And DXF = True Then
                Me.GridImport.Rows.Add(My.Resources.icon_dxf, NameFile, NamePath, FormatPts)
                Me.GridImport.Rows(Me.GridImport.Rows.Count - 1).Height = 36
            ElseIf VisualBasic.Right(NameFile.ToUpper, 4) = ".ITF" And ITF = True Then
                Me.GridImport.Rows.Add(My.Resources.icon_itf, NameFile, NamePath, FormatPts)
                Me.GridImport.Rows(Me.GridImport.Rows.Count - 1).Height = 36
            Else                                                       'Autres formats

                'Boucle sur les formats de points
                If PTS = True Then
                    'Dim ListExtPts() As String = Split(".ptp|.gco|.yxz|.txt|.gsi", "|") ' Attention test mêmes extension
                    'For i = 0 To ListExtPts.Count - 1
                    'If VisualBasic.Right(NameFile.ToUpper, 4) = ListExtPts(i).ToUpper Then '".PTP"
                    Me.GridImport.Rows.Add(My.Resources.icon_pts, NameFile, NamePath, FormatPts)
                    Me.GridImport.Rows(Me.GridImport.Rows.Count - 1).Height = 36
                    'Exit For
                    'End If
                    'Next
                End If

            End If
        End If
        Return "ok"

        Windows.Forms.Cursor.Current = Cursors.Default ' Curseur en attente
    End Function

    Private Function UpdateGridImportCount()

        Dim NomTitre As String
        NomTitre = "Importation de données"

        If GridImport.Rows.Count > 0 Then
            btnDataBrows.Text = NomTitre & "   (" & GridImport.Rows.Count & " élément(s) )"
        Else
            btnDataBrows.Text = NomTitre
        End If

        Return (GridImport.Rows.Count)

    End Function
    Private Sub btnActif_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnActif.Click

        AddRowGridData(Revo.RevoFiler.NameCurrentDraw())

        UpdateGridDataCount()
        Me.GridData.Focus()
    End Sub

    Private Sub btnOpen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOpen.Click

        Dim strFiles() As String
        'Variable qui contiendra un tableau contenant les fichiers
        '        Dim i As Long 'Variable boucle

        '        strFiles = New Autodesk.AutoCAD.Revo ' Autodesk.AutoCAD.Revo.RevoFiler.NameOpenDraw()

        '  For Each FilesName As String In strFiles

        strFiles = Split(Revo.RevoFiler.NameOpenDraw(), "|")
        ' Next 
        For Each strFile As String In strFiles
            If strFile <> "" Then
                AddRowGridData(strFile)
            End If
        Next

        UpdateGridDataCount()
        Me.GridData.Focus()
    End Sub



    Private Function AddRowGridCmd(ByVal FileName As String, Optional SharedXML As Boolean = False)

        Windows.Forms.Cursor.Current = Cursors.WaitCursor ' Curseur en attente
        If FileName <> "" Then
            Dim NameFile As String
            Dim NamePath As String
            Dim Filetype As String = "*.CSV"
            Dim NameInfo() As String
            Dim Nom() As String = Nothing
            Dim Description() As String = Nothing
            Dim Groupe() As String = Nothing

            NameFile = Path.GetFileName(FileName)
            NamePath = Ass.ActionsPath
            If NamePath.ToUpper <> Path.GetDirectoryName(FileName).ToUpper & "\" Then
                If InStr(FileName, "\") <> 0 Then
                    NamePath = Path.GetDirectoryName(FileName).ToUpper & "\"
                End If
            End If
            FileName = Path.Combine(NamePath, NameFile)

            If Directory.Exists(FileName) = True Then
                Try
                    'charge les sous fichier
                    For Each file As String In System.IO.Directory.GetFiles(FileName, Filetype, SearchOption.AllDirectories)
                        AddRowGridCmd(file)
                    Next
                Catch Ex As UnauthorizedAccessException
                    MsgBox(Ex.Message)
                Finally
                End Try
            ElseIf VisualBasic.Right(NameFile.ToUpper, 4) = ".CSV" Then
                If File.Exists(FileName) Then
                    NameInfo = Revo.RevoFiler.NameDescScript(FileName)
                    If NameInfo(0) <> "ErrFile" Then
                        Nom = Split(NameInfo(0), ";")
                        Description = Split(NameInfo(1), ";")
                        Groupe = Split(NameInfo(2), ";")
                        If Nom.Length >= 2 Then Nom(0) = Nom(1)
                        If Description.Length >= 2 Then Description(0) = Description(1)
                        If Groupe.Length >= 2 Then Groupe(0) = Groupe(1)

                    End If
                Else
                    Nom = Split("! Flux introuvable !", ";")
                    Groupe = Split("-", ";")
                    Description = Split("Fichier introuvable vérifie son chemin", ";")
                End If

                Dim IconRv As System.Drawing.Bitmap = Ass.Icon32
                If SharedXML Then IconRv = My.Resources.ribbon_plug_lock
                Me.GridCmd.Rows.Add(IconRv, Nom(0), Groupe(0), NameFile, NamePath, SharedXML)
                Me.GridCmd.Rows(Me.GridCmd.Rows.Count - 1).Cells(1).ToolTipText = Description(0)
                Me.GridCmd.Rows(Me.GridCmd.Rows.Count - 1).Height = 36

            End If
        End If

        Return "ok"

        Windows.Forms.Cursor.Current = Cursors.Default ' Curseur en attente

    End Function




    Private Sub SelectAllToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SelectAllToolStripMenuItem.Click

        For i = 0 To Me.GridCmd.Rows.Count - 1
            Me.GridCmd.Rows(i).Selected = (True)
        Next

    End Sub

    Private Sub DeleteToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DeleteToolStripMenuItem.Click
        Dim TotalSelect As Double = 0

        For i = 1 To Me.GridCmd.SelectedRows.Count
            DeleteRowCmd()
        Next

        'Activation du rechargement du menu
        Revo.MyCommands.ReloadMenu = True

    End Sub
    Private Function DeleteRowCmd()
        Dim XMLpath As String
        XMLpath = Ass.PluginPersoXML
        Dim x As New RevoXML(XMLpath)
        Dim NbreCmdCloud As Double = 0
        For i = 0 To Me.GridCmd.Rows.Count - 1 ' 0 to Count-1
            If Me.GridCmd.Rows(i).Cells(5).Value = True Then NbreCmdCloud += 1
        Next
        If NbreCmdCloud <> 0 Then NbreCmdCloud -= 1

        For i = NbreCmdCloud To Me.GridCmd.Rows.Count - 1 ' 0 to Count-1
            If Me.GridCmd.Rows(i).Selected = True Then
                Me.GridCmd.Rows.RemoveAt(i)

                'Mise à jour XML
                Dim NbreCmd As Double = NbreCmdCloud
                If NbreCmdCloud > 0 Then NbreCmd += 1
                x.deleteElement("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", i + 1 - NbreCmd)
                Exit For
            End If
        Next

        Return "ok"
    End Function

    Private Sub EditToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles EditToolStripMenuItem.Click

        If Revo.RevoFiler.OpenExe _
        (Path.Combine(Me.GridCmd.CurrentRow.Cells(4).Value, Me.GridCmd.CurrentRow.Cells(3).Value)) = False Then

            MsgBox("Le lancement du fichier n'as pu être executé")
        End If

        'System.IO.File.Open("C:\Program Files\DinO\local\geobat.csv", FileMode.Open) 'Me.GridCmd.Rows(0).Cells(2))
    End Sub




    Private Sub btnLog_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLog.Click

        If Revo.connect.ActLog = True Then
            btnLog.BackgroundImage = My.Resources.bouton_revo_app
            Revo.connect.ActLog = False
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActLog", "0")
        Else
            btnLog.BackgroundImage = My.Resources.bouton_revo_app_active
            Revo.connect.ActLog = True
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActLog", "1")
        End If
    End Sub
    Private Sub OuvrirLhistoriqueToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OuvrirLhistoriqueToolStripMenuItem.Click
        System.Diagnostics.Process.Start(Ass.LogPath)
    End Sub

    Private Sub btnLog_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles btnLog.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContextMenuStrip3.Show(CType(sender, Control), e.Location)
        End If

    End Sub

    Private Sub btnLog_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLog.MouseHover
        If Revo.connect.ActLog = False Then btnLog.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub
    Private Sub btnLog_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnLog.MouseLeave
        If Revo.connect.ActLog = False Then btnLog.BackgroundImage = My.Resources.bouton_revo_app
    End Sub
    Private Sub btnDWG_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDWG.Click

        If Revo.connect.ActDWG = True Then
            btnDWG.BackgroundImage = My.Resources.bouton_revo_app
            Revo.connect.ActDWG = False
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActDWG", "0")
        Else
            btnDWG.BackgroundImage = My.Resources.bouton_revo_app_active
            Revo.connect.ActDWG = True
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActDWG", "1")
        End If
    End Sub
    Private Sub btnDWG_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDWG.MouseHover
        If Revo.connect.ActDWG = False Then btnDWG.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub
    Private Sub btnDWG_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDWG.MouseLeave
        If Revo.connect.ActDWG = False Then btnDWG.BackgroundImage = My.Resources.bouton_revo_app
    End Sub
    Private Sub btnDXF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDXF.Click

        If Revo.connect.ActDXF = True Then
            btnDXF.BackgroundImage = My.Resources.bouton_revo_app
            Revo.connect.ActDXF = False
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActDXF", "0")
        Else
            btnDXF.BackgroundImage = My.Resources.bouton_revo_app_active
            Revo.connect.ActDXF = True
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActDXF", "1")
        End If
    End Sub
    Private Sub btnDXF_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDXF.MouseHover
        If Revo.connect.ActDXF = False Then btnDXF.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub
    Private Sub btnDXF_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDXF.MouseLeave
        If Revo.connect.ActDXF = False Then btnDXF.BackgroundImage = My.Resources.bouton_revo_app
    End Sub
    Private Sub btnPTS_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPTS.Click

        If Revo.connect.ActPTS = True Then
            btnPTS.BackgroundImage = My.Resources.bouton_revo_app
            Revo.connect.ActPTS = False
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActPTS", "0")
        Else
            btnPTS.BackgroundImage = My.Resources.bouton_revo_app_active
            Revo.connect.ActPTS = True
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActPTS", "1")
        End If
    End Sub
    Private Sub btnPTS_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPTS.MouseHover
        If Revo.connect.ActPTS = False Then btnPTS.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub
    Private Sub btnPTS_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPTS.MouseLeave
        If Revo.connect.ActPTS = False Then btnPTS.BackgroundImage = My.Resources.bouton_revo_app
    End Sub
    Private Sub btnITF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnITF.Click

        If Revo.connect.ActITF = True Then
            btnITF.BackgroundImage = My.Resources.bouton_revo_app
            Revo.connect.ActITF = False
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActITF", "0")
        Else
            btnITF.BackgroundImage = My.Resources.bouton_revo_app_active
            Revo.connect.ActITF = True
            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "ActITF", "1")
        End If

    End Sub
    Private Sub btnITF_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnITF.MouseHover
        If Revo.connect.ActITF = False Then btnITF.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub
    Private Sub btnITF_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnITF.MouseLeave
        If Revo.connect.ActITF = False Then btnITF.BackgroundImage = My.Resources.bouton_revo_app
    End Sub

    Private Sub btnAddCmd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCmd.Click

        Dim strFiles() As String
        'Variable qui contiendra un tableau contenant les fichiers
        Dim i As Long 'Variable boucle
        Dim ConfirmOK As Double

        OpenFileDialog1.Filter = "Flux " & Ass.xTitle & " (*.csv)|*.csv" '|DXF (*.dxf)|*.dxf"
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.FileName = ""
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.InitialDirectory = Ass.ActionsPath
        ConfirmOK = OpenFileDialog1.ShowDialog()

        If ConfirmOK = 1 Then
            strFiles = OpenFileDialog1.FileNames
            For i = 0 To strFiles.GetUpperBound(0)

                'Activation du rechargement du menu
                Revo.MyCommands.ReloadMenu = True

                AddRowGridCmd(strFiles(i))
                'Ecriture dans le XML
                Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", Replace(strFiles(i), Ass.ActionsPath, ""), Me.GridCmd.Rows.Count)
            Next
        End If

        Me.GridCmd.Focus()


        'UpdateGridDataCount()

    End Sub



    Private Sub GridCmd_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridCmd.DragDrop

        Dim strFiles() As String
        'Variable qui contiendra un tableau contenant les fichiers
        Dim i As Long 'Variable boucle

        'on recupere le drop dans le tableau
        strFiles = e.Data.GetData(DataFormats.FileDrop)

        For i = 0 To strFiles.GetUpperBound(0)
            If Strings.Right(strFiles(i).ToUpper, 4) = ".CSV" Then
                AddRowGridCmd(strFiles(i))
                'Ecriture dans le XML
                Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/cmdFiles", "url", Replace(strFiles(i), Ass.ActionsPath, ""), Me.GridCmd.Rows.Count)

                'Activation du rechargement du menu
                Revo.MyCommands.ReloadMenu = True
            End If
        Next

        'Mise à jour du bouton lancement rapide
        ' --->>> A programmer ...

    End Sub

    Private Sub GridCmd_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridCmd.DragEnter

        'Si le drop en question est un drop de fichiers,
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' alors on accepte le drop sous forme de copyDrop
            e.Effect = DragDropEffects.Copy
        Else 'sinon
            'on accepte pas
            e.Effect = DragDropEffects.None
        End If

    End Sub

    Private Sub GridCmd_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GridCmd.MouseDown
        If e.Button = Windows.Forms.MouseButtons.Right Then

            Dim Locked As Boolean = False
            For i = 1 To Me.GridCmd.SelectedRows.Count
                If Me.GridCmd.Item(5, GridCmd.SelectedCells(i).RowIndex).Value = True Then
                    Locked = True
                    ' Exit For
                Else
                    ' Locked = False
                End If
            Next

            If Locked Then
                For Each Bouton As ToolStripMenuItem In ContextMenuStrip2.Items
                    If Bouton.Text.ToLower Like "*effacer*" Then
                        If Bouton.Text Like "*Fonction verrouillée*" Then
                        Else
                            Bouton.Enabled = False
                            Bouton.Text = "Fonction verrouillée : " & Bouton.Text
                        End If
                    End If
                Next
            Else
                For Each Bouton As ToolStripMenuItem In ContextMenuStrip2.Items
                    If Bouton.Text.ToLower Like "*fonction verrouillée*" Then
                        Bouton.Enabled = True
                        Bouton.Text = Replace(Bouton.Text, "Fonction verrouillée : ", "")
                    End If
                Next
            End If

            ContextMenuStrip2.Show(CType(sender, Control), e.Location)

        End If
    End Sub

    Private Sub btnOptions_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOptions.Click
        'lab_log.Text = Connect.EnvoyerMail()

        Me.Hide()
        Dim Fenetre As New frmOptions

        Fenetre.ShowDialog()

        Me.Show()
    End Sub

    Private Sub btnStart_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs)
        btnStart.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub

    Private Sub btnStart_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs)
        btnStart.BackgroundImage = My.Resources.bouton_revo_app
    End Sub


    Private Sub frm_RevoFinder_ResizeEnd(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.ResizeEnd

        Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "size", Me.Height & "/" & Me.Width)

        Try
            If Me.Width <= MinFormWidth Then
                btnHide.Text = ">"
                TableLayoutFileData.ColumnStyles.Item(0).Width = 100
            Else
                btnHide.Text = "<"
                TableLayoutFileData.ColumnStyles.Item(0).Width = 0
            End If
        Catch
        End Try

    End Sub

    Private Sub btnHide_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHide.Click

        Try

            If Me.Width <= MinFormWidth Then
                'Déplier
                TableLayoutFileData.ColumnStyles.Item(0).Width = 100
                Me.Width = MaxFormWidth
                btnHide.Text = "<"
            Else
                'Plier
                TableLayoutFileData.ColumnStyles.Item(0).Width = 0
                Me.Width = MinFormWidth
                btnHide.Text = ">"
            End If

            Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "size", Me.Height & "/" & Me.Width)

        Catch

        End Try



    End Sub
    Private Sub btnHide_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnHide.MouseHover
        btnHide.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub

    Private Sub btnHide_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnHide.MouseLeave
        btnHide.BackgroundImage = My.Resources.bouton_revo_app
    End Sub
    Private Sub btnDataBrows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDataBrows.Click

        TableLayoutFlipBrowser.RowStyles.Item(0).SizeType = SizeType.Absolute
        TableLayoutFlipBrowser.RowStyles.Item(0).Height = 36
        TableLayoutFlipBrowser.RowStyles.Item(1).SizeType = SizeType.Percent
        TableLayoutFlipBrowser.RowStyles.Item(1).Height = 100


        '  btnDataBrows.BackgroundImage = Nothing
        btnDataBrows.FlatAppearance.BorderSize = 0
        btnFileBrows.FlatAppearance.BorderSize = 1

    End Sub
    Private Sub btnDataBrows_MouseHover(sender As Object, e As System.EventArgs) Handles btnDataBrows.MouseHover
        btnDataBrows.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub

    Private Sub btnFileBrows_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnFileBrows.Click


        TableLayoutFlipBrowser.RowStyles.Item(0).SizeType = SizeType.Percent
        TableLayoutFlipBrowser.RowStyles.Item(0).Height = 100
        TableLayoutFlipBrowser.RowStyles.Item(1).SizeType = SizeType.Absolute
        TableLayoutFlipBrowser.RowStyles.Item(1).Height = 36

        btnDataBrows.FlatAppearance.BorderSize = 1
        '  btnFileBrows.BackgroundImage = 
        btnFileBrows.FlatAppearance.BorderSize = 0

    End Sub
    Private Sub btnFileBrows_MouseHover(sender As Object, e As System.EventArgs) Handles btnFileBrows.MouseHover
        btnFileBrows.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub
    Private Sub btnImportFilesFile_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImportFilesFile.Click
        Dim strFiles() As String
        Dim XMLformat As String = Ass.XMLformatBase
        Dim NbreFormat As Double = 0
        'Variable qui contiendra un tableau contenant les fichiers
        Dim i As Long 'Variable boucle
        Dim ConfirmOK As Double
        Dim Filters As String = ""

        ' Attention en fonction des boutons décalage dans la listes
        ' (1) Formats compatibles
        ' (2) Dessin (*.dwg)
        ' (3) DXF (*.dxf)
        ' (4) Interlis (*.itf)
        ' (x) Points ...

        If Revo.connect.ActDWG = True Then Filters += ";*.dwg"
        If Revo.connect.ActDXF = True Then Filters += ";*.dxf"
        If Revo.connect.ActITF = True Then Filters += ";*.itf"

        If File.Exists(XMLformat) And Revo.connect.ActPTS Then 'Test de l'existance du fichier XML Revo
            Dim x As New RevoXML(XMLformat) 'lecture des extentions
            NbreFormat = x.countElements("/data-format/format")
            For i = 1 To NbreFormat
                Filters += ";*." & x.getElementValue("/data-format/format/extension", i)
            Next
        End If
        If Mid(Filters, 1, 1) = ";" Then Filters = Mid(Filters, 2, Len(Filters) - 1)
        If Filters = "" Then Filters = "-/-"
        Dim FormatSelect As Double = 1

        Filters = "Formats compatibles|" & Filters
        If Revo.connect.ActDWG = True Then Filters += "|Dessin (*.dwg)|*.dwg" : FormatSelect += 1
        If Revo.connect.ActDXF = True Then Filters += "|DXF (*.dxf)|*.dxf" : FormatSelect += 1
        If Revo.connect.ActITF = True Then Filters += "|Interlis (*.itf)|*.itf" : FormatSelect += 1

        'Importation des formats points : format.xml **********
        If File.Exists(XMLformat) And Revo.connect.ActPTS Then 'Test de l'existance du fichier XML Revo
            Dim x As New RevoXML(XMLformat) 'lecture des extentions
            Dim VarExt As String = ""
            For i = 1 To NbreFormat
                VarExt = x.getElementValue("/data-format/format/extension", i)
                Filters += "|" & x.getElementValue("/data-format/format/name", i) & " (*." & VarExt & ")" & "|*." & VarExt
            Next
        End If

        ' If PurgeFiles = True Then DeleteRow()

        OpenFileDialog1.Filter = Filters
        OpenFileDialog1.Multiselect = True
        OpenFileDialog1.FileName = ""
        ConfirmOK = OpenFileDialog1.ShowDialog()

        If ConfirmOK = 1 Then
            Dim NumFormat As Double
            NumFormat = OpenFileDialog1.FilterIndex - FormatSelect
            If NumFormat < 0 Then NumFormat = 0
            strFiles = OpenFileDialog1.FileNames
            For i = 0 To strFiles.GetUpperBound(0)
                AddRowGridImport(strFiles(i), NumFormat)
            Next
        End If

        UpdateGridImportCount()
        Me.GridImport.Focus()
    End Sub

    Private Sub btnImportFilesFolder_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImportFilesFolder.Click
        ' Dim strFiles() As String
        'Variable qui contiendra un tableau contenant les fichiers
        Dim ConfirmOK As Double
        'Dim instance As UnauthorizedAccessException


        Windows.Forms.Cursor.Current = Cursors.WaitCursor ' Curseur en attente
        'Windows.Forms.Cursor.Current = Cursors.Default ' Curseur en attente

        FolderBrowserDialog1.Description = "Sélectionner un dossier"
        ConfirmOK = (FolderBrowserDialog1.ShowDialog())
        '        MsgBox(FolderBrowserDialog1.)

        If ConfirmOK = 1 Then


            Dim dir As String = FolderBrowserDialog1.SelectedPath

            If Directory.Exists(dir) = True Then
                Try

                    Dim FilterList As New List(Of String)
                    If Revo.connect.ActDWG Then FilterList.Add("*.DWG")
                    If Revo.connect.ActDXF Then FilterList.Add("*.DXF")
                    If Revo.connect.ActITF Then FilterList.Add("*.ITF")
                    If Revo.connect.ActPTS Then
                        Dim XMLformat As String = Ass.XMLformatBase
                        If File.Exists(XMLformat) And Revo.connect.ActPTS Then 'Test de l'existance du fichier XML Revo
                            Dim x As New RevoXML(XMLformat) 'lecture des extentions
                            Dim NbreFormat As Double = x.countElements("/data-format/format")
                            For i = 1 To NbreFormat
                                FilterList.Add("*." & x.getElementValue("/data-format/format/extension", i).ToUpper)
                            Next
                        End If
                    End If

                    'charge les sous fichier
                    Dim files As List(Of FileInfo) = GetMultiFiles(dir, FilterList)
                    For Each file As FileInfo In files
                        AddRowGridImport(file.FullName)
                    Next

                Catch Ex As UnauthorizedAccessException
                    MsgBox(Ex.Message)
                Finally
                End Try

            End If
        End If

        UpdateGridImportCount()
        Me.GridImport.Focus()

        Windows.Forms.Cursor.Current = Cursors.Default ' Curseur en attente
    End Sub

    Private Sub GridImport_CellContentClick(ByVal sender As System.Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridImport.CellContentClick
        UpdateGridImportCount()
    End Sub
    Private Sub GridImport_RowStateChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewRowStateChangedEventArgs) Handles GridImport.RowStateChanged
        UpdateGridImportCount()
    End Sub

    Private Sub GridImport_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridImport.DragDrop
        Dim strFiles() As String
        'Variable qui contiendra un tableau contenant les fichiers
        Dim i As Long 'Variable boucle

        'on recupere le drop dans le tableau
        strFiles = e.Data.GetData(DataFormats.FileDrop)

        For i = 0 To strFiles.GetUpperBound(0)
            AddRowGridImport(strFiles(i))
        Next

        UpdateGridImportCount()
    End Sub

    Private Sub GridImport_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles GridImport.DragEnter
        'Si le drop en question est un drop de fichiers,
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            ' alors on accepte le drop sous forme de copyDrop
            e.Effect = DragDropEffects.Copy
        Else 'sinon
            'on accepte pas
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub GridImport_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GridImport.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContextMenuStrip4.Show(CType(sender, Control), e.Location)
        End If

    End Sub


    Private Sub ImportSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportSelectAll.Click
        For i = 0 To Me.GridImport.Rows.Count - 1
            Me.GridImport.Rows(i).Selected = (True)
        Next
    End Sub

    Private Sub ImportImportSelectAll_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ImportImportSelectAll.Click

        Dim TotalSelect As Double = 0

        For i = 1 To Me.GridImport.SelectedRows.Count 'TotalSelect
            DeleteRowImport()
        Next

    End Sub
    Private Function DeleteRowImport()

        For i = 0 To Me.GridImport.Rows.Count - 1
            If Me.GridImport.Rows(i).Selected = True Then
                Me.GridImport.Rows.RemoveAt(i)
                Exit For
            End If
        Next
        Return "ok"

    End Function

End Class