Imports System.IO
Imports Microsoft
Imports System.Windows.Forms

Public Class frmOptions
    Dim Connect As New Revo.connect
    Dim Ass As New Revo.RevoInfo

    Private Sub frmOptions_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        LoadingVariable()

    End Sub
    Private Sub LoadingVariable()

        Dim TestAdd As Boolean
        Me.Text = Ass.xTitle & " Options"
        Me.AcceptButton = btnOK
        Me.CancelButton = btnEscape

        'Test si dossier distant : XMLPathCloud > XMLformat,ActionsPath,Plotters,SharedPath,Library,SupportPath,Template
        If Ass.PluginSharedXML(True).ToUpper <> Ass.PluginSharedXML(False).ToUpper Then
            btnNetwork.Enabled = False
            btnLock.Visible = True
            Dim InfoBtn As Boolean = True
            'For Each Bouton As ToolStripMenuItem In ContextMenuStrip1.Items
            '    If Bouton.Text <> "" And Bouton.Text <> "" Then
            '        Bouton.Enabled = False
            '        Bouton.Text = "Fonction verrouillée : " & Bouton.Text
            '    End If
            'Next
        End If

        '                         Nom Param,Acad déf,   Var param enregistré     , Nom variable ,  Type ,     Var param défaut(var,déf=oui)
        TestAdd = AddRowGridParam("EDT Folder", False, Ass.EDTFolder, "EDTfolder", "PATH", Ass.EDTFolder(True))
        TestAdd = AddRowGridParam("PER Folder", False, Ass.PERFolder, "PERfolder", "PATH", Ass.PERFolder(True))


        TestAdd = AddRowGridParam("Actions", False, Ass.ActionsPath, "ActionsPath", "PATH", Ass.ActionsPath(True))
        TestAdd = AddRowGridParam("Plotters", False, Ass.Plotters, "Plotters", "PATH", Ass.Plotters(True))
        TestAdd = AddRowGridParam(Ass.xTitle & " Shared", False, Ass.SharedPath, "SharedPath", "PATH", Ass.SharedPath(True))
        TestAdd = AddRowGridParam("Support", False, Ass.SupportPath, "SupportPath", "PATH", Ass.SupportPath(True))
        TestAdd = AddRowGridParam("Template", False, Ass.Template, "Template", "Gabarit (*.dwt)|*.dwt", Ass.Template(True))
        TestAdd = AddRowGridParam("Block library", False, Ass.Library, "Library", "PATH", Ass.Library(True))

       
        'PROVISOIRE    ' TestAdd = AddRowGridParam("Format de points", False, Ass.XMLformatPerso, "XMLformatPerso", "Fichier des formats des points (*.xml)|*.xml", Ass.XMLformatPerso(True))

        '   TestAdd = AddRowGridParam(Ass.xTitle & " System", False, Ass.SystemPath, "SystemPath", "PATH", Ass.SystemPath(True))
        '  TestAdd = AddRowGridParam(Ass.xTitle & " Log", False, Ass.LogPath, "LogPath", "Fichier Log(*.txt)|*.txt", Ass.LogPath(True))



        'ATTENTION SI FICHIER/DOSSIER NON EXISTANT : PAS D'AFFICHAGE DE L'OPTION


        '*Actions\					ActionsPath 	add Acad/Actions
        '*Plotters\Plot Styles\JourCAD.ctb	CTB			link Acad/Plot Styles
        '*Shared\					SharedPath		Revo
        '*Support\					SupportPath	add Acad/Support
        '*Template\Revo10.dwt			Template		add Acad/Support
        'System\					SystemPath	Revo


        'Chargement du logo
        PicLogo.Image = Ass.Icon32


    End Sub
    Private Sub btnLock_Click(sender As System.Object, e As System.EventArgs) Handles btnLock.Click
        MsgBox("Les paramètres sont configurés en réseau, contact l'administrateur pour modifier ces paramètres", vbInformation + vbOKOnly, "Paramètre verrouillé")
    End Sub

    Private Function AddRowGridParam(ByVal ParamName As String, ByVal Acad As Boolean, ByVal FileName As String, ByVal VariableParam As String, ByVal TypeParam As String, ByVal DefaultVar As String)

        Windows.Forms.Cursor.Current = Windows.Forms.Cursors.WaitCursor ' Curseur en attente

        Dim TestExist As Boolean = False
        Dim AcadTrue As Integer = 0
        'Dim Params As List(Of String)

        If Acad = True Then AcadTrue = 1
        If FileName <> "" Then
            If TypeParam = "PATH" Then
                TestExist = IO.Directory.Exists(FileName)
            Else
                TestExist = IO.File.Exists(FileName)
            End If
        End If

        If TestExist = True Then
            ' Params = Split(FileName, vbCrLf)
            Me.GridParam.Rows.Add(ParamName, FileName, VariableParam, TypeParam, DefaultVar)
            ' Me.GridParam.Rows(Me.GridParam.Rows.Count - 1).Cells(1).Style = FlatStyle.Flat
            'Me.GridParam.Rows(Me.GridParam.Rows.Count - 1).Cells(1).ToolTipText = "Cocher pour utiliser les paramètres par défauts d'Autocad"
            Me.GridParam.Rows(Me.GridParam.Rows.Count - 1).Height = 36
        End If

        Return True

        Windows.Forms.Cursor.Current = Windows.Forms.Cursors.Default ' Curseur en attente

    End Function

    Private Sub GridParam_CellDoubleClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles GridParam.CellDoubleClick
        If GridParam.CurrentCell.ColumnIndex = 1 Then
            ModifPathCmd(GridParam.CurrentCell.RowIndex)
        End If
    End Sub

  
    Private Sub GridParam_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles GridParam.MouseDown

        If e.Button = Windows.Forms.MouseButtons.Right Then
            ContextMenuStrip1.Show(CType(sender, Control), e.Location)
        End If
    End Sub

    Private Sub ModifPath_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ModifPath.Click
        ModifPathCmd(GridParam.CurrentCell.RowIndex)
    End Sub

    Private Sub ModifPathCmd(ByVal RowNum As Double)
        Dim strFile As String = ""
        Dim ConfirmOK As Double = 0

        'Vérifier si paramètre partagée et verrouillé
        If btnLock.Visible = True Then
            If GridParam.Rows(RowNum).Cells(0).Value <> "EDT Folder" And GridParam.Rows(RowNum).Cells(0).Value <> "PER Folder" Then

                MsgBox("Les paramètres partagés ne peuvent pas être editer par l'utilisateur" & vbCrLf & "Contacter votre administrateur", _
                        vbInformation + vbOKOnly, "Impossible de modifier ce paramètre")
                Return
            End If


        End If

        If GridParam.Rows(RowNum).Cells(3).Value = "PATH" Then
            FolderBrowserDialog1.ShowNewFolderButton = True
            FolderBrowserDialog1.Description = "Sélectionner un dossier pour définir le paramètre : " & GridParam.Rows(RowNum).Cells(0).Value
            FolderBrowserDialog1.SelectedPath = GridParam.Rows(RowNum).Cells(1).Value
            ConfirmOK = FolderBrowserDialog1.ShowDialog
            strFile = FolderBrowserDialog1.SelectedPath
            If VisualBasic.Right(strFile, 1) <> "\" Then strFile = strFile & "\"
        Else
            OpenFileDialog1.Filter = GridParam.Rows(RowNum).Cells(3).Value '|DXF (*.dxf)|*.dxf"
            OpenFileDialog1.Multiselect = False
            OpenFileDialog1.CheckFileExists = True
            OpenFileDialog1.Title = "Sélectionner le fichier : " & GridParam.Rows(RowNum).Cells(0).Value
            OpenFileDialog1.InitialDirectory = GridParam.Rows(RowNum).Cells(1).Value
            OpenFileDialog1.FileName = GridParam.Rows(RowNum).Cells(1).Value
            ConfirmOK = OpenFileDialog1.ShowDialog()
            strFile = OpenFileDialog1.FileName
        End If


            If ConfirmOK = 1 Then
            If strFile <> "" Then GridParam.Rows(RowNum).Cells(1).Value = strFile
                'AddRowGridCmd(strFiles(i))
            End If

            GridParam.Focus()

    End Sub

    Private Sub ChargerParamDefautToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ChargerParamDefautToolStripMenuItem.Click
        DefautPathCmd(GridParam.CurrentCell.RowIndex)
    End Sub

    Private Sub DefautPathCmd(ByVal RowNum As Double)

        'Vérifier si paramètre partagée et verrouillé
        If btnLock.Visible = True Then
            If GridParam.Rows(RowNum).Cells(0).Value <> "EDT Folder" And GridParam.Rows(RowNum).Cells(0).Value <> "PER Folder" Then
                MsgBox("Les paramètres partagés ne peuvent pas être editer par l'utilisateur" & vbCrLf & "Contacter votre administrateur", _
                        vbInformation + vbOKOnly, "Impossible de modifier ce paramètre")
                Return
            End If
        End If

        Dim strFile As String = ""

        strFile = GridParam.Rows(RowNum).Cells(4).Value
        If strFile <> "" Then GridParam.Rows(RowNum).Cells(1).Value = strFile
        '  GridParam.Refresh()
        GridParam.Focus()

    End Sub

    Private Sub btnOK_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnOK.Click

        For i = 0 To GridParam.RowCount - 1
            If (GridParam.Rows(i).Cells(4).Value).ToString.ToUpper <> (GridParam.Rows(i).Cells(1).Value).ToString.ToUpper Then
                'Ecriture dans le XML
                Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", (GridParam.Rows(i).Cells(2).Value).ToString.ToLower, GridParam.Rows(i).Cells(1).Value)
            Else
                Connect.XMLdelete("/" & Ass.xProduct.ToLower & "/config", (GridParam.Rows(i).Cells(2).Value).ToString.ToLower)
            End If
        Next
        Me.Close()

    End Sub

    Private Sub btnOK_MouseHover(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.MouseHover
        btnOK.BackgroundImage = My.Resources.bouton_revo_app_active()
    End Sub

    Private Sub btnOK_MouseLeave(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOK.MouseLeave
        btnOK.BackgroundImage = My.Resources.bouton_revo_app
    End Sub

    Private Sub btnEscape_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnEscape.Click
        Me.Close()
    End Sub

    ''' <summary>
    ''' Share folders
    ''' </summary>
    ''' <param name="sender">Objet bouton</param>
    '''<param name="e">Event bouton</param>
    ''' <remarks>Copy folders and files</remarks>
    Private Sub btnNetwork_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnNetwork.Click
        'Choisir le dossier partagé en réseau
        Dim ConfirmOK As Integer = 0
        Dim FolderName As String = ""
        FolderBrowserDialog1.ShowNewFolderButton = True
        FolderBrowserDialog1.Description = "Sélectionner un dossier qui sera accessible en réseau"
        ConfirmOK = FolderBrowserDialog1.ShowDialog
        FolderName = FolderBrowserDialog1.SelectedPath
        If VisualBasic.Right(FolderName, 1) <> "\" Then FolderName = FolderName & "\"

        If ConfirmOK = 1 Then

            'Création du fichier xml distant
            Dim PathCloud As String = Path.Combine(FolderName, "revo-shared.xml")
            conn.CreateXML(PathCloud, False)
            'Inscrit dans le perso.xml local le dossier distant
            If IO.File.Exists(PathCloud) Then
                conn.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "cloudconfig", PathCloud, 1, False) '<cloudconfig>revo-shared.xml</cloudconfig>
            End If

            'Copier la structure de dossier du local en réseaux
            'Actions + Plotters + Shared + Support + Template
            'ActionsPath + SharedPath + SupportPath + Template
            Dim Variable As String = ""
            Dim DefaultPath As String = Ass.PluginPersoXML
            Dim NewPath As String = ""
            For i = 0 To GridParam.RowCount - 1
                Variable = GridParam.Rows(i).Cells(2).Value.ToString.ToLower ' ATTENTION Minuscule
                Connect.Message(Ass.xProduct.ToLower & "-perso", "En cours de copie" & vbCrLf & Variable, False, i, GridParam.RowCount - 1)
                'Création des parametres partagé
                If "actionspath" = Variable Or _
                    "sharedpath" = Variable Or _
                    "supportpath" = Variable Or _
                    "template" = Variable Or _
                    "xmlformatperso" = Variable Or _
                    "library" = Variable Or _
                    "plotters" = Variable Then

                    Dim VariablePath As String = GridParam.Rows(i).Cells(4).Value.ToString
                    Try
                        If GridParam.Rows(i).Cells(3).Value.ToString = "PATH" Then
                            NewPath = Path.Combine(FolderName, Replace(VariablePath, DefaultPath, ""))
                            Revo.RevoFiler.CopyDirectory(VariablePath, NewPath, False)
                        Else
                            NewPath = Path.Combine(FolderName, Replace(VariablePath, DefaultPath, ""))
                            Directory.CreateDirectory(Path.GetDirectoryName(NewPath)) 'Crée un repertoire
                            If File.Exists(NewPath) = False Then File.Copy(VariablePath, NewPath, False)
                        End If

                        'Modifier le fichier de config
                        If NewPath <> "" Then GridParam.Rows(i).Cells(1).Value = NewPath

                        'Modifie le fichier XML distant
                        conn.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", Variable, NewPath, 1, False) '<cloudconfig>revo-shared.xml</cloudconfig>

                    Catch ex As Exception
                        Connect.RevoLog(Connect.DateLog & "Shared folder" & vbTab & False & vbTab & "Copy: " & ex.Message & vbTab & FolderName)
                        Connect.Message(Ass.xProduct.ToLower & "-perso", "Impossible de crée le dossier partagé." & vbCrLf & _
                                        "Contrôler les droits d'écriture !", False, 98, 100)
                        Connect.Message(Ass.xProduct.ToLower & "-perso", "", True, 100, 100)
                        NewPath = ""
                        Exit For
                    End Try
                End If
            Next

            If NewPath <> "" Then

                Me.GridParam.Rows.Clear()
                LoadingVariable()

                Connect.RevoLog(Connect.DateLog & "Shared folder" & vbTab & True & vbTab & "Copy: " & True & vbTab & FolderName)
                Connect.Message(Ass.xProduct.ToLower & "-perso", "Le dossier partagé est crée:" & vbCrLf & FolderName, False, 98, 100)
                Connect.Message(Ass.xProduct.ToLower & "-perso", "", True, 100, 100)
            End If

        End If


    End Sub

    ''' <summary>
    ''' Load config-revo.xml file
    ''' </summary>
    ''' <param name="sender">Objet bouton</param>
    '''<param name="e">Event bouton</param>
    ''' <remarks>Erase the file</remarks>
    Private Sub btnImportXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnImportXML.Click
        'Load config-revo.xml

        Dim ConfirmOK As Integer = 0
        Dim FileName As String = ""
        OpenFileDialog1.Filter = Ass.xProduct.ToLower & "-perso (*.xml)|*.xml"
        OpenFileDialog1.Multiselect = False
        OpenFileDialog1.CheckFileExists = True
        OpenFileDialog1.Title = "Sélectionner le fichier de personalisation " & Ass.xTitle
        OpenFileDialog1.InitialDirectory = Ass.PluginPath
        OpenFileDialog1.FileName = Ass.PluginPersoXML
        ConfirmOK = OpenFileDialog1.ShowDialog()
        FileName = OpenFileDialog1.FileName

        If ConfirmOK = 1 Then 'Copy in XML defaut
            Try
                File.Copy(FileName, Ass.PluginPersoXML, True)
                Me.Close()
                Connect.RevoLog(Connect.DateLog & "Import XML perso" & vbTab & True & vbTab & "Saved: " & True & vbTab & FileName)
                Connect.Message(Ass.xProduct.ToLower & "-perso", "Importation du fichier XML terminé", False, 99, 100)
                Connect.Message(Ass.xProduct.ToLower & "-perso", "Importation du fichier XML terminé", True, 100, 100)
            Catch ex As Exception
                Connect.RevoLog(Connect.DateLog & "Import XML perso" & vbTab & False & vbTab & "Saved: " & False & vbTab & FileName)
                If FileName = Ass.PluginPersoXML Then
                    Connect.Message(Ass.xProduct.ToLower & "-perso", "Impossible d'importer le fichier déjà chargé !", False, 99, 100, "critical")
                Else
                    Connect.Message(Ass.xProduct.ToLower & "-perso", "Impossible d'importer le fichier XML", False, 99, 100, "critical")
                End If
                Connect.Message(Ass.xProduct.ToLower & "-perso", "", True, 100, 100)
            End Try
        End If

    End Sub

    ''' <summary>
    ''' Save-as config-revo.xml file
    ''' </summary>
    ''' <param name="sender">Objet bouton</param>
    '''<param name="e">Event bouton</param>
    ''' <remarks>Copy the file</remarks>
    Private Sub btnExportXML_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExportXML.Click
        Dim ConfirmOK As Integer = 0
        Dim FileName As String = ""
        SaveFileDialog1.Filter = Ass.xProduct.ToLower & "-perso (*.xml)|*.xml"
        SaveFileDialog1.Title = "Sélectionner le fichier de personalisation " & Ass.xTitle
        'SaveFileDialog1.InitialDirectory = Ass.RevoPath
        SaveFileDialog1.FileName = Ass.xProduct.ToLower & "-perso.xml" 'Ass.XMLPath
        ConfirmOK = SaveFileDialog1.ShowDialog()
        FileName = SaveFileDialog1.FileName

        If ConfirmOK = 1 Then 'Save in XML defaut
            For i = 0 To GridParam.RowCount - 1
                If (GridParam.Rows(i).Cells(4).Value).ToString.ToUpper <> (GridParam.Rows(i).Cells(1).Value).ToString.ToUpper Then
                    'Ecriture dans le XML
                    Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", (GridParam.Rows(i).Cells(2).Value).ToString.ToLower, GridParam.Rows(i).Cells(1).Value)
                Else
                    Connect.XMLdelete("/" & Ass.xProduct.ToLower & "/config", (GridParam.Rows(i).Cells(2).Value).ToString.ToLower)
                End If
            Next

            'Copy the saved file
            Try
                File.Copy(Ass.PluginPersoXML, FileName, True)
                Me.Close()
                Connect.RevoLog(Connect.DateLog & "Export XML perso" & vbTab & True & vbTab & "Saved: " & True & vbTab & FileName)
                Connect.Message(Ass.xProduct.ToLower & "-perso", "Exportation du fichier XML terminé", False, 98, 100)
                Connect.Message(Ass.xProduct.ToLower & "-perso", "Exportation du fichier XML terminé", True, 100, 100)
            Catch ex As Exception
                Connect.RevoLog(Connect.DateLog & "Export XML perso" & vbTab & False & vbTab & "Saved: " & False & vbTab & FileName)
                Connect.Message(Ass.xProduct.ToLower & "-perso", "Impossible d'exporter le fichier XML", False, 99, 100, "critical")
                Connect.Message(Ass.xProduct.ToLower & "-perso", "", True, 100, 100)
            End Try
        End If

    End Sub

   
End Class