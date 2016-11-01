Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors
Imports System.IO
Imports frms = System.Windows.Forms

Module modImportMD01MO

    ' Public Prog As frmProgress
    ' Public Connect As Revo.connect
    Public conn As New Revo.connect ' include SQLite in Revo THA
    Dim Ass As New Revo.RevoInfo

    Public Structure PolylineVertexes
        Public Verts As Point2dCollection
        Public lngVertCount As Long
        Public AcadObject As ObjectId
        Public WithArcs As Boolean
    End Structure
    'PolylineVertex class to hold collection of PolylineVertexes objects
    Class PlVertCol(Of PolylineVertexes)
        Inherits CollectionBase
        'Sub to add to the collection
        Public Sub Add(ByVal p As PolylineVertexes)
            InnerList.Add(p)
        End Sub
        'Sub to remove from the collection
        Public Sub Remove(ByVal index As Integer)
            InnerList.RemoveAt(index)
        End Sub
        'Function to return an item from the collection
        Public Function Item(ByVal index As Integer) As PolylineVertexes
            Return DirectCast(InnerList.Item(index), PolylineVertexes)
        End Function
    End Class

    Public booCancelProcess As Boolean = False
    Private lngLine As Int64 = 0
    Private lngLinesCount As Int64 = 0
    Public ITFTopic As String = ""
    Public ITFTable As String = ""
    Private ITFObject As String = ""
    Private ITFObjectID As String = ""
    Public ITFTablesHorsModele As String = ""

    ''' <summary>
    ''' Imports the supplied ITF file into AutoCAD
    ''' </summary>
    ''' <param name="strSourceFile">The filename of the ITF file to import</param>
    ''' <param name="CreerPartTerr"></param>
    ''' <returns>True if successful.</returns>
    ''' <remarks></remarks>
    Public Function ImportITFFile(ByVal strSourceFile As String, Optional ByVal CreerPartTerr As Boolean = True) As Boolean

        conn.Message(Ass.xTitle, "Importation du fichier ITF en cours ...", False, 20, 100) ' include SQLite in Revo THA
        frms.Application.DoEvents()

        'Initialisation de l'opération
        Dim booSuccess As Boolean = False
        booCancelProcess = False
        ITFTablesHorsModele = ""
        'Prog = New frmProgress
        'Prog.lblMsg.Text = "Reading ITF File..."
        'Prog.Show()
        'Connect.FenStat.ProgBar.Value = 20
        Try


            'Traitement sur tout les dessin (Revo13 ou plus ou moins)
            Dim RVscript As New Revo.RevoScript
            RVscript.cmdLA("#LA;>MO_ODL_PISCINE_SUPERFICIE;[[TrueColor]]7;[[Lineweight]]0.20;")



            'Call the function to import the ITF file into the SQLite database
            booSuccess = ImportITF(strSourceFile)


            If booSuccess Then
                'Call the function to generate the entities from the information in SQLite
                UCSOn(False)
                booSuccess = DessinITF(CreerPartTerr)
                UCSOn(True)

                'Ajout des Echelles 1:500 et 1:1000
                '#Cmd;ANNOAUTOSCALE|4|_CANNOSCALE|1:500|_CANNOSCALE|1:1000|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1
               

                'Type hachures : Plan RF  'Ajouter pour Belloti 2014    (ATTENTION REMPLACER LES "]];" > "]]"

                RVscript.cmdHA("#HA;MO_ODS_BATSOUT;[[Layer]]>MO_ODS_BATSOUT_HACH;[[StyleName]]SOLID;")
                RVscript.cmdHA("#HA;MO_CS_BAT;[[Layer]]>MO_CS_BAT_HACH;[[StyleName]]SOLID;")
                RVscript.cmdHA("#HA;MO_CS_EAU_COURS;[[Layer]]>MO_CS_EAU_COURS_HACH;[[StyleName]]SOLID;")
                RVscript.cmdHA("#HA;MO_CS_EAU_ROSELIERE;[[Layer]]>MO_CS_EAU_ROSELIERE_HACH;[[StyleName]]ROSELIERE;[[ScaleFactor]]1;")
                RVscript.cmdHA("#HA;MO_CS_SSV_ROCHER;[[Layer]]>MO_CS_SSV_ROCHER_HACH;[[StyleName]]FELS;[[ScaleFactor]]3.45000004768371;")
                RVscript.cmdHA("#HA;MO_CS_SSV_EBOULIS;[[Layer]]>MO_CS_SSV_EBOULIS_HACH;[[StyleName]]GEROELL;[[ScaleFactor]]3.45000004768371;")
                RVscript.cmdHA("#HA;MO_CS_SV_CI_VIGNE;[[Layer]]>MO_CS_SV_CI_VIGNE_HACH;[[StyleName]]VIGNE;[[ScaleFactor]]1;")
                RVscript.cmdHA("#HA;MO_CS_SV_TOURB;[[Layer]]>MO_CS_SV_TOURB_HACH;[[StyleName]]TOURBIERE_MARAIS;[[ScaleFactor]]1;")
                RVscript.cmdHA("#HA;MO_CS_SB_PAT_OUVERT;[[Layer]]>MO_CS_SB_PAT_OUVERT_HACH;[[StyleName]]DOTS;[[ScaleFactor]]5.03999996185302;")
                RVscript.cmdHA("#HA;MO_CS_SB_PAT_DENSE;[[Layer]]>MO_CS_SB_PAT_DENSE_HACH;[[StyleName]]DOTS;[[ScaleFactor]]2.51999998092651;")
                RVscript.cmdHA("#HA;MO_CS_SB_FORET;[[Layer]]>MO_CS_SB_FORET_HACH;[[StyleName]]DOTS;[[ScaleFactor]]1.25999999046325;")

                RVscript.cmdLA("#LA;MO_ODL_PISCINE_SUPERFICIE;[[LayerOn]]1;[[Freeze]]0")

                'RVscript.cmdHA("#HA;MO_ODS_BATSOUT;[[Layer]]>MO_ODS_BATSOUT_HACH;[[StyleName]]SOLID;")
                'RVscript.cmdHA("#HA;MO_CS_BAT;[[Layer]]>MO_CS_BAT_HACH;[[StyleName]]SOLID;")
                'RVscript.cmdHA("#HA;MO_CS_SB_FORET;[[Layer]]>MO_CS_SB_FORET_HACH;[[StyleName]]SOLID;")

                'RVscript.cmdCmd("#Cmd;ANNOAUTOSCALE|4|_CANNOSCALE|1:500|_CANNOSCALE|1:1000|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1")
                'RVscript.cmdCmd("#Cmd;regen")
                ChangeScale("1000")

                Zooming.ZoomExtents()

                'Mise à jour des attributs
                RVscript.cmdCmd("#Cmd;_AttSync|_N|*")
                '+---++++++++++++-+------+-------
            End If



            If ITFTablesHorsModele <> "" Then ' MD01MOVD
                frms.MessageBox.Show("Attention, une ou plusieurs tables n'ont pas été importées:" _
                                     & vbCrLf & ITFTablesHorsModele & vbCrLf & _
                                     "Cela peut-être dû à une problème d'intégrité à la table ou au fait qu'il s'agit d'une table particulière non prévue" _
                                     & vbCrLf & "dans le modèle de données MD.01-MO-VD." _
                                     & vbCrLf & "Si vous n'êtes pas sûr de la procédure à suivre, commencez par vérifier votre fichier ITF avec un checker Interlis.", _
                                     "Problème(s) à l'import ITF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)

            End If
        Catch ex As Exception
        Finally
            'Prog.Close()
            'Prog.Dispose()
        End Try

        conn.Message(Ass.xTitle, "Mise en forme graphique ...", False, 90, 100) ' include SQLite in Revo THA
        conn.Message(Ass.xTitle, "", True, 100, 100)

        Return booSuccess
    End Function

    Private Function ImportITF(ByVal strSourceFile As String) As Boolean

        Dim sRead As StreamReader = Nothing
        ReinitLineCounter()
        'Call the function to count the number of lines in the ITF file
        lngLinesCount = GetFileLinesCount(strSourceFile)
        'Prog.ProgressBar1.Maximum = Convert.ToInt32(lngLinesCount)
        'Check that we read the number of lines in the file without error
        If lngLinesCount <> -1 Then
            'Open the database
            'OpenDatabase()
            'and clear any existing data in the tables
            ClearTables()

            Try
                'Read a line from the ITF file 
                sRead = New StreamReader(strSourceFile, System.Text.Encoding.UTF7)
                Dim strLine As String = ""
                'Loop until we find the "MODL"
                Do Until strLine.StartsWith("MODL", StringComparison.Ordinal)
                    strLine = ReadITFFileLine(sRead)
                    If strLine Like "MODL *" Then
                        If strLine Like "MODL MD01MOVD*" Or strLine Like "MODL MD01MOCH24F*" Then
                            '  MsgBox("MD01MOVD model ok")
                            '  ElseIf strLine Like "MODL MD01MOFR24F*" Then
                            '     MsgBox("L'importation du modèle Interlis en MD01MOFR24F est un version beta." & vbCrLf _
                            '            & "Merci de contacter le support pour communiquer des remarques.", vbInformation, "Interruption de l'importation")
                        Else
                            MsgBox("Le modèle Interlis n'est pas compatible." & vbCrLf _
                                   & "Les modèles validés sont : MD01MOVD" & vbCrLf & vbCrLf _
                                   & "Merci de contacter le support pour avoir plus d'informations.", vbInformation, "Interruption de l'importation")
                            Exit Function
                        End If
                    End If
                Loop
                Dim strStart As String = ""
                'Loop around reading lines from the file
                Do
                    'Read a line from the file
                    strLine = ReadITFFileLine(sRead)
                    'Get the first four characters from the line
                    strStart = strLine.Substring(0, 4)
                    Select Case strStart
                        'If this is a TOPI line
                        Case "TOPI"
                            'Get the topic from the line (from the fifth character onwards)
                            ITFTopic = strLine.Substring(5)
                            ITFTable = "?"
                            ITFObject = "0"
                            ITFObjectID = "0"

                            'Call the function to read the topic
                            If Not ITFReadTopic(sRead) Then
                                If booCancelProcess Then
                                    frms.MessageBox.Show("Import du fichier INTERLIS '" & strSourceFile & "' annulé !", "Import ITF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                                    Return False
                                Else
                                    frms.MessageBox.Show("Erreur à la lecture du fichier INTERLIS" & vbCrLf & "'" & strSourceFile & "':" & vbCrLf & vbCrLf _
                                    & "Topic: " & ITFTopic & vbCrLf & "Table: " & ITFTable & vbCrLf & "Object: " & ITFObject, "Import ITF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                                    Return False
                                End If
                            End If
                        Case Else
                            Exit Do
                    End Select
                Loop
                If Not UpdateData() Then
                    Return False
                End If
            Catch ex1 As Exception
                frms.MessageBox.Show(ex1.Message, "Import ITF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)

            Catch ex2 As System.IndexOutOfRangeException
                frms.MessageBox.Show("Erreur à la lecture du fichier INTERLIS" & vbCrLf & ITFTopic & " / " & ITFTable & vbCrLf & ex2.Message, "Import ITF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)

            Finally
                If Not sRead Is Nothing Then
                    sRead.Close()
                    sRead.Dispose()
                End If
            End Try
            Return True
        End If
    End Function

    Private Sub ReinitLineCounter()
        lngLine = 0
    End Sub

    ''' <summary>
    ''' Gets the number of lines in the ITF file being imported
    ''' </summary>
    ''' <param name="strFileName">The filename of the ITF file being imported</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetFileLinesCount(ByVal strFileName As String) As Int64
        'Check that the file exists
        If File.Exists(strFileName) Then
            Dim sRead As StreamReader = Nothing
            Try
                'Open the file using Stream Reader
                sRead = New StreamReader(strFileName)
                'Read the contents of the file
                Dim strFileContents As String = sRead.ReadToEnd
                'Split the file by chr(10)
                Dim Lines() As String = strFileContents.Split(Chr(10))
                'Return the number of elements in the Lines() array -2 to remove the last empty line and the "ENDE" line 
                Return Lines.Length - 2
            Catch ex As Exception
                'Something went wrong so inform the user
                frms.MessageBox.Show(ex.Message.ToString, ex.TargetSite.ToString, Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                Return -1
            Finally
                'Close and dispose of the reader
                If Not sRead Is Nothing Then
                    sRead.Close()
                    sRead.Dispose()
                End If
            End Try
        Else
            'File does not exist!
            Return -1
        End If
    End Function

    ''' <summary>
    ''' Reads a line from the itf file
    ''' </summary>
    ''' <param name="sRead">The stream reader containing the itf file</param>
    ''' <returns>The line read from the itf file</returns>
    ''' <remarks>Combines logical lines that are split over multiple physical lines into a single line</remarks>
    Public Function ReadITFFileLine(ByRef sRead As StreamReader) As String
        Dim booFixFmt As Boolean = False
        Dim ContFix As Boolean = False
        Dim strLine As String = ""
        'Keep reading lines until the line doesn't end with "\" (denoting that this line continues on the next line)
        Do
            strLine = strLine & sRead.ReadLine()
            'Remove carriage return from the end of the line
            'If strLine.EndsWith(Chr(13)) Then strLine = strLine.Substring(0, strLine.Length - 1)
            'Increment the line counter
            lngLine += 1
            'Prog.ProgressBar1.Value += 1
            frms.Application.DoEvents()
            'Trim any spaces from the ends
            strLine = strLine.Trim
            If strLine.EndsWith("\", StringComparison.Ordinal) Then
                'This line continues on the next line
                booFixFmt = True
                ContFix = True
                'Remove the "\" from the end
                strLine.Substring(0, strLine.Length - 1)
            Else
                booFixFmt = False
            End If
        Loop While booFixFmt = True
        'If the line was on more than one physical line then remove "CONT" from the string
        ' If booFixFmt Then strLine = strLine.Replace("CONT", "") 'Ligne rempacé par ci-dessous / THA ERR Corrigée le 06.01.2014
        If ContFix Then strLine = strLine.Replace("\CONT ", "") ' THA ERR Corrigée le 06.01.2014

        'Call the function to remove duplicate spaces
        strLine = KeepOnly1Space(strLine)
        'Return the line
        Return strLine
    End Function

    ''' <summary>
    ''' Reads the TOPIC portion of the file
    ''' </summary>
    ''' <param name="sRead">The stream reader containing the itf file</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Private Function ITFReadTopic(ByRef sRead As StreamReader) As Boolean
        Dim strLine As String = ""
        Do
            'Read another line from the ITF file
            strLine = ReadITFFileLine(sRead)
            Select Case strLine.Substring(0, 4)
                'If the first four characters are "TABL" then
                Case "TABL"
                    'Get the name of the table
                    ITFTable = strLine.Substring(5)
                    ITFObject = "0"
                    ITFObjectID = "0"

                    'Call the function to read the information for this table
                    If Not ITFReadTable(sRead) Then Return False
                Case Else
                    Return True
            End Select
        Loop
        Return True
    End Function

    ''' <summary>
    ''' Reads the data into the table
    ''' </summary>
    ''' <param name="sRead">The stream reader containing the itf file</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Private Function ITFReadTable(ByRef sRead As StreamReader) As Boolean
        Dim arrAttrObj() As String
        Dim NbreDo As Double = 0
        Dim booObjLineaire As Boolean = False
        Dim intFieldNr As Integer = 0
        Dim lngSeqId As Int16 = 0
        Dim strLine As String = ""
        Dim SkipGeomLoop As Boolean = False
        Dim FieldNameString As String = ""
        Dim FieldValString As String = ""
        Dim strSql As String = ""
        'Read another line from the file
        strLine = ReadITFFileLine(sRead)
        Dim dt As System.Data.DataTable = Nothing
        Dim strStart As String = strLine.Substring(0, 4)
        'Loop until we find the end of the table


        Do
            NbreDo += 1
            conn.Message(Ass.xTitle, "Importation du fichier ITF en cours : " & ITFTable & " (" & NbreDo & ")", False, 20, 100) ' include SQLite in Revo THA

            'frms.Application.DoEvents()
            If booCancelProcess Then Return False
            Select Case strStart

                Case "OBJE" 'If the line begins with "OBJE" 

                    SkipGeomLoop = False
                    'Split the line up 
                    arrAttrObj = Split(strLine.Substring(5), " ")
                    'Store the ITFObject
                    ITFObject = arrAttrObj(0)

                    If arrAttrObj.Length > 1 Then
                        ITFObjectID = arrAttrObj(1)
                    End If

                    '        Il faudrait ajouter un ID2 à la table "Objets_divers_Element_surfacique"
                    '        Modifier la requête SQL pour faire une recherche en fonction de l'ID2
                    '        If ITFTable = "Element_surfacique_Geometrie" Then ' Modif BUG OD THA 27.01.2012
                    '        ITFObject = arrAttrObj(1)   ' JE PENSE QUE C'EST UNE CONNERIE !??! 
                    '        End If   ' A cause de la posibilité d'avoir de multiple ID

                    Try ' Impossible de trouver la colonne ...
                        If dt Is Nothing OrElse dt.TableName <> ITFTopic & "_" & ITFTable Then
                            dt = GetTable(ITFTopic & "_" & ITFTable)
                        End If

                    Catch 'ex As Exception
                        MsgBox("modImportMD01MO / ITFReadTable / OBJE ") ' & ex.Message)
                    End Try


                    'Check that we have the recordset
                    If dt Is Nothing Then
                        'If not then add the table name to the ITFTablesHorsModele string
                        ITFTablesHorsModele = ITFTablesHorsModele & ITFTopic & "_" & ITFTable & vbCrLf
                        Exit Do
                    End If
                    If Not ITFTable.EndsWith("_Geometrie", StringComparison.Ordinal) And Not ITFTable.EndsWith("_Perimetre", StringComparison.Ordinal) _
                                                           And Not ITFTable.EndsWith("_Ligne_auxiliaire", StringComparison.Ordinal) Then
                        'Add new row to the recordset
                        Dim dr As DataRow = dt.NewRow()
                        intFieldNr = 0
                        'Loop through the attribute array filling the columns with the values
                        For i As Integer = 0 To arrAttrObj.Length - 1

                            Try

                                If dt.Columns(intFieldNr).ColumnName = "Perimetre" Or _
                                          dt.Columns(intFieldNr).ColumnName = "Geometrie" Or _
                                          dt.Columns(intFieldNr).ColumnName = "Ligne_auxiliaire" Then
                                    intFieldNr += 1
                                End If
                                If arrAttrObj(i) <> "@" Then
                                    'SQLite does not have a field type of Date
                                    'If rs.Fields(intFieldNr).Type = dao.DataTypeEnum.dbDate Then
                                    '    arrAttrObj(i) = arrAttrObj(i).Substring(0, 4) & "-" & _
                                    '                    arrAttrObj(i).Substring(4, 2) & "-" & _
                                    '                    arrAttrObj(i).Substring(arrAttrObj(i).Length - 2)
                                    'End If
                                    If arrAttrObj(i) <> "_" Then
                                        dr(intFieldNr) = arrAttrObj(i)
                                    End If
                                End If
                                intFieldNr += 1

                            Catch ex As System.IndexOutOfRangeException
                                MsgBox(ex.Message & " (" & strLine & " / " & ITFTopic & " / " & ITFTable & ")")
                            End Try
                        Next
                        dt.Rows.Add(dr)
                        'Check table name for specific value
                        If ITFTable.Equals("Partie_limite_nationale") Or ITFTable.Equals("Element_lineaire") Or ITFTable.Equals("Arete") _
                                                                Or ITFTable.Equals("Ligne_coordonnees") Or ITFTable.Equals("Partie_limite_canton") _
                                                                Or ITFTable.Equals("Limite_communeProj") Or ITFTable.Equals("Partie_limite_district") _
                                                                Or ITFTable.Equals("Troncon_rue") Or ITFTable.Equals("Error") Then
                            'Append "_Geometrie" to the table name
                            ITFTable = ITFTable & "_Geometrie"
                            'Set the booObjLineaire flag to true
                            booObjLineaire = True
                            'and set the SkipGeomLoop flag to true so that we perform the top of this loop again for this line in the file
                            SkipGeomLoop = True
                        End If
                    End If

                    If Not SkipGeomLoop Then
                        Do
                            'Read another line from the file
                            strLine = ReadITFFileLine(sRead)
                            'Get the first four characters of this line
                            strStart = strLine.Substring(0, 4)
                            'Set the values in the table depending on the first four characters of the line
                            Select Case strStart
                                Case "STPT"
                                    arrAttrObj = Split(strLine.Substring(5), " ")
                                    lngSeqId = 0
                                    If ITFTable.StartsWith("PosImmeuble", StringComparison.Ordinal) Then
                                        dt = GetTable(ITFTopic & "_" & ITFTable & "_Ligne_Auxiliaire")
                                    End If
                                    SetFieldValues(dt, arrAttrObj, Convert.ToInt16(1), lngSeqId)
                                Case "LIPT"
                                    arrAttrObj = Split(strLine.Substring(5), " ")
                                    lngSeqId += Convert.ToInt16(1)
                                    SetFieldValues(dt, arrAttrObj, Convert.ToInt16(2), lngSeqId)
                                Case "ARCP"
                                    arrAttrObj = Split(strLine.Substring(5), " ")
                                    lngSeqId += Convert.ToInt16(1)
                                    SetFieldValues(dt, arrAttrObj, Convert.ToInt16(4), lngSeqId)
                                Case "ELIN"
                                    If booObjLineaire Then
                                        booObjLineaire = False
                                        ITFTable = ITFTable.Substring(0, ITFTable.Length - 10)
                                        dt = Nothing
                                    End If
                                    strLine = ReadITFFileLine(sRead)
                                    strStart = strLine.Substring(0, 4)
                                    Exit Do
                            End Select
                        Loop Until strStart = "ETAB" Or strStart = "OBJE"
                    End If

                Case Else
                    Return True
            End Select

        Loop

        Return True
    End Function

    '''' <summary>
    '''' Reads the data into the table
    '''' </summary>
    '''' <param name="sRead">The stream reader containing the itf file</param>
    '''' <returns>True if successful</returns>
    '''' <remarks></remarks>
    'Private Function ITFReadTable(ByRef sRead As StreamReader) As Boolean
    '    Dim arrAttrObj() As String
    '    Dim booObjLineaire As Boolean = False
    '    Dim intFieldNr As Integer = 0
    '    Dim lngSeqId As Int16 = 0
    '    Dim strLine As String = ""
    '    Dim SkipGeomLoop As Boolean = False
    '    Dim FieldNameString As String = ""
    '    Dim FieldValString As String = ""
    '    Dim strSql As String = ""
    '    'Read another line from the file
    '    strLine = ReadITFFileLine(sRead)
    '    Dim dt As System.Data.DataTable = Nothing
    '    Dim strStart As String = strLine.Substring(0, 4)
    '    'Loop until we find the end of the table
    '    Do
    '        'frms.Application.DoEvents()
    '        If booCancelProcess Then Return False
    '        Select Case strStart
    '            Case "OBJE" 'If the line begins with "OBJE" 
    '                SkipGeomLoop = False
    '                'Split the line up 
    '                arrAttrObj = Split(strLine.Substring(5), " ")
    '                'Store the ITFObject
    '                ITFObject = Convert.ToDouble(arrAttrObj(0))
    '                Try
    '                    'Check if the recordset we have is not the same as the one we're about to open
    '                    '(No need to open it again if it's already open)
    '                    If rs.Name <> ITFTopic & "_" & ITFTable Then
    '                        rs = Nothing
    '                        rs = objDb.OpenRecordset(ITFTopic & "_" & ITFTable, dao.RecordsetTypeEnum.dbOpenTable)
    '                    End If
    '                Catch ex As Exception
    '                    'The attempt to get the name caused an error as the recordset was not set to anything                    
    '                    rs = Nothing
    '                    rs = objDb.OpenRecordset(ITFTopic & "_" & ITFTable, dao.RecordsetTypeEnum.dbOpenTable)
    '                End Try
    '                'Check that we have the recordset
    '                If rs Is Nothing Then
    '                    'If not then add the table name to the ITFTablesHorsModele string
    '                    ITFTablesHorsModele = ITFTablesHorsModele & ITFTopic & "_" & ITFTable & vbCrLf
    '                    Exit Do
    '                End If
    '                If Not ITFTable.EndsWith("_Geometrie", StringComparison.Ordinal) And Not ITFTable.EndsWith("_Perimetre", StringComparison.Ordinal) _
    '                                                       And Not ITFTable.EndsWith("_Ligne_auxiliaire", StringComparison.Ordinal) Then
    '                    'Add new row to the recordset
    '                    rs.AddNew()
    '                    intFieldNr = 0
    '                    'Loop through the attribute array filling the columns with the values
    '                    For i As Integer = 0 To arrAttrObj.Length - 1
    '                        If rs.Fields(intFieldNr).Name = "Perimetre" Or _
    '                                  rs.Fields(intFieldNr).Name = "Geometrie" Or _
    '                                  rs.Fields(intFieldNr).Name = "Ligne_auxiliaire" Then
    '                            intFieldNr += 1
    '                        End If
    '                        If arrAttrObj(i) <> "@" Then
    '                            If rs.Fields(intFieldNr).Type = dao.DataTypeEnum.dbDate Then
    '                                arrAttrObj(i) = arrAttrObj(i).Substring(0, 4) & "-" & _
    '                                                arrAttrObj(i).Substring(4, 2) & "-" & _
    '                                                arrAttrObj(i).Substring(arrAttrObj(i).Length - 2)
    '                            End If
    '                            If arrAttrObj(i) <> "_" Then
    '                                rs.Fields(intFieldNr).Value = arrAttrObj(i)
    '                            End If
    '                        End If
    '                        intFieldNr += 1
    '                    Next
    '                    rs.Update()
    '                    'Check table name for specific value
    '                    If ITFTable.Equals("Partie_limite_nationale") Or ITFTable.Equals("Element_lineaire") Or ITFTable.Equals("Arete") _
    '                                                            Or ITFTable.Equals("Ligne_coordonnees") Or ITFTable.Equals("Partie_limite_canton") _
    '                                                            Or ITFTable.Equals("Limite_communeProj") Or ITFTable.Equals("Partie_limite_district") _
    '                                                            Or ITFTable.Equals("Troncon_rue") Or ITFTable.Equals("Error") Then
    '                        'Append "_Geometrie" to the table name
    '                        ITFTable = ITFTable & "_Geometrie"
    '                        'Set the booObjLineaire flag to true
    '                        booObjLineaire = True
    '                        'and set the SkipGeomLoop flag to true so that we perform the top of this loop again for this line in the file
    '                        SkipGeomLoop = True
    '                    End If
    '                End If
    '                If Not SkipGeomLoop Then
    '                    Do
    '                        'Read another line from the file
    '                        strLine = ReadITFFileLine(sRead)
    '                        'Get the first four characters of this line
    '                        strStart = strLine.Substring(0, 4)
    '                        'Set the values in the table depending on the first four characters of the line
    '                        Select Case strStart
    '                            Case "STPT"
    '                                arrAttrObj = Split(strLine.Substring(5), " ")
    '                                lngSeqId = 0
    '                                If ITFTable.StartsWith("PosImmeuble", StringComparison.Ordinal) Then
    '                                    rs = objDb.OpenRecordset(ITFTopic & "_" & ITFTable & "_Ligne_Auxiliaire")
    '                                End If
    '                                SetFieldValues(rs, arrAttrObj, Convert.ToInt16(1), lngSeqId)
    '                            Case "LIPT"
    '                                arrAttrObj = Split(strLine.Substring(5), " ")
    '                                lngSeqId += Convert.ToInt16(1)
    '                                SetFieldValues(rs, arrAttrObj, Convert.ToInt16(2), lngSeqId)
    '                            Case "ARCP"
    '                                arrAttrObj = Split(strLine.Substring(5), " ")
    '                                lngSeqId += Convert.ToInt16(1)
    '                                SetFieldValues(rs, arrAttrObj, Convert.ToInt16(4), lngSeqId)
    '                            Case "ELIN"
    '                                If booObjLineaire Then
    '                                    booObjLineaire = False
    '                                    ITFTable = ITFTable.Substring(0, ITFTable.Length - 10)
    '                                    rs.Close()
    '                                End If
    '                                strLine = ReadITFFileLine(sRead)
    '                                strStart = strLine.Substring(0, 4)
    '                                Exit Do
    '                        End Select
    '                    Loop Until strStart = "ETAB" Or strStart = "OBJE"
    '                End If
    '            Case Else
    '                Try
    '                    rs.Close()
    '                Catch ex As Exception
    '                End Try
    '                Return True
    '        End Select
    '    Loop
    '    Try
    '        rs.Close()
    '    Catch ex As Exception
    '    End Try
    '    Return True
    'End Function

    ''' <summary>
    ''' Adds the values in the array into the recordset
    ''' </summary>
    ''' <param name="dt">The datatable to add the values into</param>
    ''' <param name="arrAttrObj">The array of values to add</param>
    ''' <param name="entType">The type of entity</param>
    ''' <param name="lngSeqID">The sequence id</param>
    ''' <remarks></remarks>
    Private Sub SetFieldValues(ByRef dt As System.Data.DataTable, ByRef arrAttrObj() As String, ByVal entType As Int16, ByVal lngSeqID As Int16)
        Dim dr As DataRow = dt.NewRow
        dr(0) = ITFObject
        dr(1) = 1
        dr(2) = entType
        dr(3) = lngSeqID
        For i As Integer = 0 To arrAttrObj.Length - 1
            dr(i + 4) = arrAttrObj(i)
        Next

        If dt.TableName = "Objets_divers_Element_surfacique_Geometrie" Then
            dr(8) = ITFObjectID 'uniquement pour la table : Objets_divers_Element_surfacique_Geometrie
        End If

        dt.Rows.Add(dr)
        'InsertRow(dt, dr)
        'UpdateData(dt)
    End Sub
    '''' <summary>
    '''' Adds the values in the array into the recordset
    '''' </summary>
    '''' <param name="rs">The recordset to add the values into</param>
    '''' <param name="arrAttrObj">The array of values to add</param>
    '''' <param name="entType">The type of entity</param>
    '''' <param name="lngSeqID">The sequence id</param>
    '''' <remarks></remarks>
    'Private Sub SetFieldValues(ByRef rs As dao.Recordset, ByRef arrAttrObj() As String, ByVal entType As Int16, ByVal lngSeqID As Int16)
    '    rs.AddNew()
    '    rs.Fields(0).Value = ITFObject
    '    rs.Fields(1).Value = 1
    '    rs.Fields(2).Value = entType
    '    rs.Fields(3).Value = lngSeqID
    '    For i As Integer = 0 To arrAttrObj.Length - 1
    '        rs.Fields(i + 4).Value = arrAttrObj(i)
    '    Next
    '    rs.Update()
    'End Sub

    ''' <summary>
    ''' Removes instances of two consecutive spaces within the string
    ''' </summary>
    ''' <param name="strLine">The string to remove the double spaces from</param>
    ''' <returns>The string with the double spaces removed</returns>
    ''' <remarks></remarks>
    Public Function KeepOnly1Space(ByVal strLine As String) As String
        'Loop while the string contains two consecutive spaces
        Do While strLine.Contains("  ")
            'Replace the double space with a single space
            strLine = strLine.Replace("  ", " ")
        Loop
        'Return the processed string
        Return strLine
    End Function

    ''' <summary>
    ''' Creates the ITF File
    ''' </summary>
    ''' <param name="CreerPartTerr"></param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Private Function DessinITF(ByVal CreerPartTerr As Boolean) As Boolean


        'Check variable DrawOrder : Disable 
        Application.SetSystemVariable("DRAWORDERCTL", 0) ' THA 19.01.2016 Suppression de l'ordre de tracé


        conn.Message(Ass.xTitle, "Création des objets graphiques", False, 30, 100) ' include SQLite in Revo THA

        Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument
            Try
                Dim dblFacteEch As Double = 0
                'Get the current scale of the drawing
                Dim strEch As String = GetCurrentScale()
                'Ensure the scale was set
                If Not strEch = "" And IsNumeric(strEch) Then

                    'Set the linetypescale
                    dblFacteEch = CDbl(strEch) / 1000
                    HostApplicationServices.WorkingDatabase.Ltscale = dblFacteEch
                    'Thaw all layers
                    FreezeAllLayers(False)
                    'Switch on layers
                    ShowLayer("MO_BF_INCOMPLET", True)
                    ShowLayer("MO_CS_INCOMPLET", True)
                    ShowLayer("MO_RNT_INCOMPLET", True)
                    ShowLayer("MO_RPL_INCOMPLET", True)
                    'Call the sub to generate the autocad objects from the access information
                    conn.Message(Ass.xTitle, "Lecture de la base de donnée", False, 50, 100) ' include SQLite in Revo THA
                    'If AccessITFtoAcad(dblFacteEch) = False Then Return False
                    AccessITFtoAcad(dblFacteEch) '
                    Zooming.ZoomExtents()


                    'Call the function to build the area entities
                    BuildAreaEntities(CreerPartTerr)


                    If Not booCancelProcess Then
                        RestoreDefaultLayerState()
                    End If


                    ManageTerritoryBoundaries()
                    PlaceObjectsOnTop("INSERT")

                    Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
                    ed.Regen()

                    If booCancelProcess Then
                        frms.MessageBox.Show("Opération de contruction et dessin des objets dans AutoCAD annulée !", _
                                             "Traitement interrompu", Windows.Forms.MessageBoxButtons.OK, _
                                             Windows.Forms.MessageBoxIcon.Exclamation)
                    End If
                    Return True
                Else
                    Return False
                End If
            Catch ex As Exception
                frms.MessageBox.Show(ex.Message, "DessinITF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                Return False
            Finally
                'objDb.Close()
                'objDb = Nothing
            End Try
            Return True
        End Using



        'Check variable DrawOrder : Enable
        Application.SetSystemVariable("DRAWORDERCTL", 3) ' THA 19.01.2016 Suppression de l'ordre de tracé


    End Function

    ''' <summary>
    ''' Generates objects in AutoCAD from the information stored in Access
    ''' </summary>
    ''' <param name="dblFactEch">The current scale</param>
    ''' <remarks></remarks>
    Private Function AccessITFtoAcad(ByVal dblFactEch As Double)



            Dim xDataID As Integer = 0
            Dim strJCLayer As String = ""
            Dim strJCBlock As String = ""
            Dim strQueryName As String = ""
            'Open the recordset containing the block layer information
            Dim dtTemp As System.Data.DataTable = GetTable("_JCBlock_JCLayer_Corresp")
            Dim dv As DataView = Nothing
            If Not dtTemp Is Nothing Then
                dv = New DataView(dtTemp)
                dv.Sort = "JC_Block"
            End If
            Dim dtQs As System.Data.DataTable = GetQueries()
            Dim dt As System.Data.DataTable
            'Set the elevation of drawing to 0
            DefineElevation(0)
            Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction



            'Loop through the queries in the Access database
            'Prog.ProgressBar1.Maximum = dtQs.Rows.Count
            Dim MaxCount As Double = dtQs.Rows.Count
            Dim iCount As Double = 0
            'Prog.ProgressBar1.Value = 0
            'Prog.lblMsg.Text = "Creating Features..."
            'Connect.Message(Ass.xTitle, "Création des objets", False, 40, 100) ' include SQLite in Revo THA
            'frms.Application.DoEvents()
            For Each Query As DataRow In dtQs.Rows


                'Prog.ProgressBar1.Value += 1
                iCount += 1

                'Get the name of the query
                strQueryName = Query("QueryName").ToString

                conn.Message("Requête : " & strQueryName, "Analyse de la base de données : " & iCount & "/" & dtQs.Rows.Count, False, iCount, dtQs.Rows.Count) ' include SQLite in Revo THA

                'Linear objects and surfaces
                If strQueryName.StartsWith("_", StringComparison.Ordinal) Then



                    'Get the layername
                    strJCLayer = strQueryName.Substring(1)
                    'Remove the "+" from the end if there is one
                    If strJCLayer.EndsWith("+", StringComparison.Ordinal) Then strJCLayer = strJCLayer.Substring(0, strJCLayer.Length - 1)
                    'Open the recordset for this query
                    dt = GetQueryResult(strQueryName, Query("SQL").ToString)



                        'If we have any records
                        If dt Is Nothing Then
                            Dim bob As String = ""
                        Else

                            If dt.Rows.Count > 0 Then
                                Dim lngRecCount As Int64 = dt.Rows.Count
                                Dim lngRecOk As Int16 = 0
                                'Set the layer for this block to current
                                ActivateLayer(strJCLayer)
                                Dim Verts As New Point2dCollection 'Vertices of linear ojects
                                'Dim ArcVerts As New Point2dCollection 'Vertices of arcs
                                Dim intIndex As Integer = 0
                                Dim dicVerts As New Dictionary(Of Integer, Point2d) 'To hold the vertices against the indexes
                                Dim dicArcVerts As New Dictionary(Of Integer, Point2d) 'To hold the arc vertices against the indexes
                                'Store the GID
                                Dim dblGID As Double = Convert.ToDouble(dt.Rows(0)("GID"))
                                'Loop around the records in the recordset
                            For i As Integer = 0 To dt.Rows.Count

                                If i = dt.Rows.Count OrElse dblGID <> Convert.ToDouble(dt.Rows(i)("GID")) Then



                                    'New object
                                    'Create the polyline using the vertices 
                                    Dim pLine As Polyline = InsertLWPolyline(dicVerts, strJCLayer)
                                    Dim Keys As System.Collections.IEnumerator = dicArcVerts.Keys.GetEnumerator
                                    'Loop through the vertices in the arcvertex hash table
                                    'so we can set the bulge factor

                                    While Keys.MoveNext
                                        'Call the function to get the bulge amount
                                        Dim blge As Double = GetBulgeValue(CType(dicVerts(Convert.ToInt32(Keys.Current)), Point2d), _
                                                                           CType(dicArcVerts(Convert.ToInt32(Keys.Current)), Point2d), _
                                                                           CType(dicVerts(Convert.ToInt32(Keys.Current) + 1), Point2d))


                                        'DirectCast : remplacé par CType (ou TryCast) TH 12.12.2011

                                        'Set the bulge in the polyline
                                        SetPolyBulge(pLine.ObjectId, Convert.ToInt32(Keys.Current) - 1, blge)
                                    End While


                                    'Clear the vertex collection and the vertex hashtable
                                    Verts.Clear()
                                    dicVerts.Clear()
                                    'Reset the index counter
                                    intIndex = 0
                                    'Clear the arc hashtable
                                    dicArcVerts.Clear()
                                    If i = dt.Rows.Count Then Exit For
                                    dblGID = Convert.ToDouble(dt.Rows(i)("GID"))

                                    If strQueryName = "_MO_ODS_BATSOUT" Or strQueryName = "_MO_ODS_COUVINDEP" Then
                                        'rs.MovePrevious()
                                        AddRegAppTableRecord(Ass.xProduct.ToUpper) ' Ass.xProduct.ToUpper "REVO"
                                        Dim acResBuf As New ResultBuffer
                                        acResBuf.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, Ass.xProduct.ToUpper)) 'Ass.xProduct.ToUpper)) "REVO"
                                        acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "ID" & xDataID + 1))
                                        acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, dt.Rows(i - 1)("NUMERO").ToString & ""))
                                        acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, dt.Rows(i - 1)("DESIGNATION").ToString & ""))
                                        SetObjXRecord(pLine.ObjectId, Ass.xProduct.ToUpper, acResBuf) ' Ass.xProduct.ToUpper "REVO"
                                        pLine.Closed = True
                                        'pLine.XData = acResBuf
                                        'rs.MoveNext()
                                    End If


                                End If

                                'Create a point from the values in the recordset
                                Dim V As New Point2d(Convert.ToDouble(dt.Rows(i)("x0")), Convert.ToDouble(dt.Rows(i)("y0")))

                                'If the etype value is 4 (i.e. its an arc) then add the vertex to the ArcVerts hashtable
                                If Convert.ToInt32(dt.Rows(i)("etype")) = 4 Then

                                    Try ' Ajout de la condition 
                                        dicArcVerts.Add(intIndex, V)

                                    Catch ex As System.ArgumentException ' THA - ajout d'une expetion - 13.06.2014
                                        '  MsgBox("Merci de contacter le développeur car un objet est traité de manière incorrect" _
                                        '    & vbCrLf & i & " : " & V.X & "," & V.Y & "  (" & ex.Message & ")")
                                        conn.RevoLog(conn.DateLog & "AccessITFtoAcad / Arc " & i & " : " & V.X & "," & V.Y & "  (" & ex.Message & ")")

                                    End Try
                                Else
                                    'It's just a vertex so add it to the verts collection and the verts hash table
                                    intIndex += 1
                                    'Verts.Add(V)
                                    dicVerts.Add(intIndex, V)

                                End If
                                'rs.MoveNext()



                            Next
                            End If
                    End If

                 
                    'rs.Close()
                    dt = Nothing


                ElseIf strQueryName.StartsWith("+", StringComparison.Ordinal) Then
                    '*******************
                    'Not implemented yet
                    '*******************
                    'strJCLayer = strQueryName.Replace("+", "")
                    'ActivateLayer(strJCLayer)
                    'rs = objDb.OpenRecordset(strQueryName, dao.RecordsetTypeEnum.dbOpenSnapshot)
                    'If rs.RecordCount > 0 Then
                    '    Do Until rs.EOF
                    '        Dim insPt As New Point3d(Convert.ToDouble(rs.Fields("Geometrie_re").Value), Convert.ToDouble(rs.Fields("Geometrie_ho").Value), 0)
                    '        Dim blkID As ObjectId = InsertBlock(insPt, "_NONE", 100)
                    '        'Set block colour to red
                    '        If Not blkID.IsNull Then
                    '            SetBlockColour(blkID, Color.FromColor(Drawing.Color.Red))
                    '        End If
                    '        rs.MoveNext()
                    '    Loop
                    'End If
                Else
                    'Point objects
                    strJCBlock = strQueryName
                    dt = GetQueryResult(strJCBlock, Query("SQL").ToString)
                    Dim insertPt As Point3d = Nothing
                    'Loop through the records
                    Try
                        Dim test As String = dt.TableName.ToString
                    Catch
                        MsgBox("La base de donnée est illisible" & vbCrLf & "Vérifier votre configuration !", MsgBoxStyle.Critical)
                        Return False
                    End Try

                    If dt.Rows.Count > 0 Then
                        Dim dblOrient As Double = 0
                        For i As Integer = 0 To dt.Rows.Count - 1
                            dblOrient = 100
                            Dim alt As Double = 0
                            'Get the altitude of the point (if available)
                            Try
                                If Not IsNothing(dt.Rows(i)("ALTITUDE")) Then
                                    alt = Convert.ToDouble(dt.Rows(i)("ALTITUDE"))
                                End If
                            Catch ex2 As SystemException
                            Catch ex As Exception
                            End Try

                            'Get the position at which to insert the point
                            Try 'Erreur du point d'insertion  modif THA 23.10.2015
                                insertPt = New Point3d(Convert.ToDouble(dt.Rows(i)("X")), Convert.ToDouble(dt.Rows(i)("Y")), alt)
                                'get the rotation for the point (if available)
                                Try
                                    dblOrient = Convert.ToDouble(dt.Rows(i)("ORI"))
                                Catch ex2 As SystemException
                                Catch ex As Exception
                                End Try
                                'See if the block should be on a specific layer
                                Dim Found As Integer = dv.Find(strJCBlock)
                                If Found = -1 Then
                                    strJCLayer = "0"
                                Else
                                    strJCLayer = dv.Item(Found)("JC_Layer").ToString
                                End If
                                ActivateLayer(strJCLayer)
                                'Get the ID of the layer to put the block on
                                Dim layID As ObjectId = GetLayerIDFromLayerName(strJCLayer)
                                'Call the function to insert the block

                                Dim oID As ObjectId = InsertBlock(insertPt, strJCBlock, dblOrient, layID, dblFactEch)
                                'If the block was inserted successfully then call the function to add its attributes
                                If Not oID.IsNull Then
                                    SetBlockAttributeValues(dt.Rows(i), oID)
                                End If
                                'next record
                                'rs.MoveNext()


                            Catch ex1 As System.InvalidCastException
                                ' If dt.Rows(i)("X").ToString = "" Then MsgBox(dt.Rows(i)("UFID"))
                                ' MsgBox("Erreur de format de coordonnée : " & dt.TableName.ToString & vbCrLf & ex1.Message, vbExclamation + vbOKOnly, "Erreur du point d'insertion")
                                conn.RevoLog(conn.DateLog & "AccessITFtoAcad / Pts : " & dt.TableName.ToString & vbTab & "UFID" & vbTab & dt.Rows(i)("UFID").ToString)
                            Catch
                                MsgBox("Erreur de format de coordonnée : " & dt.TableName.ToString, vbExclamation + vbOKOnly, "Erreur du point d'insertion")
                            End Try

                            conn.Message("Point objects : " & strJCBlock, "Enregistrement des données : " & strJCLayer, False, i, dt.Rows.Count) ' include SQLite in Revo THA

                        Next
                    End If
                    'rs.Close()
                    dt = Nothing
                End If




            Next
            'Save the changes
            trans.Commit()

            Try
                'rs.Close()
                dt = Nothing
            Catch
            End Try

        End Using


        Return True


    End Function
    '''' <summary>
    '''' Generates objects in AutoCAD from the information stored in Access
    '''' </summary>
    '''' <param name="dblFactEch">The current scale</param>
    '''' <remarks></remarks>
    'Private Sub AccessITFtoAcad(ByVal dblFactEch As Double)
    '    Dim xDataID As Integer = 0
    '    Dim strJCLayer As String = ""
    '    Dim strJCBlock As String = ""
    '    Dim strQueryName As String = ""
    '    'Open the recordset containing the block layer information
    '    Dim rsTemp As dao.Recordset = objDb.OpenRecordset("_JCBlock_JCLayer_Corresp", dao.RecordsetTypeEnum.dbOpenTable)
    '    rsTemp.Index = "JC_Block"
    '    'Set the elevation of drawing to 0
    '    DefineElevation(0)
    '    Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction
    '        'Loop through the queries in the Access database
    '        Dim intQueryCOunt As Int16 = objDb.QueryDefs.Count
    '        For i As Int16 = 0 To intQueryCOunt - Convert.ToInt16(1)
    '            'Get the name of the query
    '            strQueryName = objDb.QueryDefs(i).Name.ToString
    '            'Linear objects and surfaces
    '            If strQueryName.StartsWith("_", StringComparison.Ordinal) Then
    '                'Get the layername
    '                strJCLayer = strQueryName.Substring(1)
    '                'Remove the "+" form the end if there is one
    '                If strJCLayer.EndsWith("+", StringComparison.Ordinal) Then strJCLayer = strJCLayer.Substring(0, strJCLayer.Length - 1)
    '                'Open the recordset for this query
    '                rs = objDb.OpenRecordset(strQueryName, dao.RecordsetTypeEnum.dbOpenSnapshot)
    '                'If we have any records
    '                If rs.RecordCount > 0 Then
    '                    Dim lngRecCount As Int64 = rs.RecordCount
    '                    Dim lngRecOk As Int16 = 0
    '                    'Set the layer for this block to current
    '                    ActivateLayer(strJCLayer)
    '                    Dim Verts As New Point2dCollection 'Vertices of linear ojects
    '                    'Dim ArcVerts As New Point2dCollection 'Vertices of arcs
    '                    Dim lngIndex As Int64 = 0
    '                    Dim htVerts As New Hashtable 'To hold the vertices against the indexes
    '                    Dim htArcVerts As New Hashtable 'To hold the arc vertices against the indexes
    '                    'Store the GID
    '                    Dim dblGID As Double = Convert.ToDouble(rs.Fields("GID").Value)
    '                    'Loop around the records in the recordset
    '                    Do
    '                        If rs.EOF OrElse dblGID <> Convert.ToDouble(rs.Fields("GID").Value) Then
    '                            'New object
    '                            'Create the polyline using the vertices 
    '                            Dim pLine As Polyline = InsertLWPolyline(Verts, strJCLayer)
    '                            Dim Keys As System.Collections.IEnumerator = htArcVerts.Keys.GetEnumerator
    '                            'Loop through the vertices in the arcvertex hash table
    '                            'so we can set the bulge factor
    '                            While Keys.MoveNext
    '                                'Call the function to get the bulge amount
    '                                Dim blge As Double = GetBulgeValue(DirectCast(htVerts(Keys.Current), Point2d), _
    '                                                                   DirectCast(htArcVerts(Keys.Current), Point2d), _
    '                                                                   DirectCast(htVerts(Convert.ToInt64(Keys.Current) + 1), Point2d))
    '                                'Set the bulge in the polyline
    '                                SetPolyBulge(pLine.ObjectId, Convert.ToInt32(Keys.Current) - 1, blge)
    '                            End While
    '                            'Clear the vertex collection and the vertex hashtable
    '                            Verts.Clear()
    '                            htVerts.Clear()
    '                            'Reset the index counter
    '                            lngIndex = 0
    '                            'Clear the arc hashtable
    '                            htArcVerts.Clear()
    '                            If rs.EOF Then Exit Do
    '                            dblGID = Convert.ToDouble(rs.Fields("GID").Value)
    '                            If strQueryName = "_MO_ODS_BATSOUT" Or strQueryName = "_MO_ODS_COUVINDEP" Then
    '                                rs.MovePrevious()
    '                                AddRegAppTableRecord("REVO")
    '                                Dim acResBuf As New ResultBuffer
    '                                acResBuf.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, "REVO"))
    '                                acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "ID" & xDataID + 1))
    '                                acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, rs.Fields("NUMERO").Value.ToString & ""))
    '                                acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, rs.Fields("DESIGNATION").Value.ToString & ""))
    '                                SetObjXRecord(pLine.ObjectId, "REVO", acResBuf)
    '                                'pLine.XData = acResBuf
    '                                rs.MoveNext()                                    
    '                            End If
    '                        End If
    '                        'Create a point from the values in the recordset
    '                        Dim V As New Point2d(Convert.ToDouble(rs.Fields("x0").Value), Convert.ToDouble(rs.Fields("y0").Value))
    '                        'If the etype value is 4 (i.e. its an arc) then add the vertex to the ArcVerts hashtable
    '                        If Convert.ToInt32(rs.Fields("etype").Value) = 4 Then
    '                            htArcVerts.Add(lngIndex, V)
    '                        Else
    '                            'It's just a vertex so add it to the verts cololection and the verts hash table
    '                            lngIndex += 1
    '                            Verts.Add(V)
    '                            htVerts.Add(lngIndex, V)
    '                        End If
    '                        rs.MoveNext()
    '                    Loop
    '                End If
    '                rs.Close()
    '            ElseIf strQueryName.StartsWith("+", StringComparison.Ordinal) Then
    '                '*******************
    '                'Not implemented yet
    '                '*******************
    '                'strJCLayer = strQueryName.Replace("+", "")
    '                'ActivateLayer(strJCLayer)
    '                'rs = objDb.OpenRecordset(strQueryName, dao.RecordsetTypeEnum.dbOpenSnapshot)
    '                'If rs.RecordCount > 0 Then
    '                '    Do Until rs.EOF
    '                '        Dim insPt As New Point3d(Convert.ToDouble(rs.Fields("Geometrie_re").Value), Convert.ToDouble(rs.Fields("Geometrie_ho").Value), 0)
    '                '        Dim blkID As ObjectId = InsertBlock(insPt, "_NONE", 100)
    '                '        'Set block colour to red
    '                '        If Not blkID.IsNull Then
    '                '            SetBlockColour(blkID, Color.FromColor(Drawing.Color.Red))
    '                '        End If
    '                '        rs.MoveNext()
    '                '    Loop
    '                'End If
    '            Else
    '                'Point objects
    '                strJCBlock = strQueryName
    '                rs = objDb.OpenRecordset(strJCBlock, dao.RecordsetTypeEnum.dbOpenSnapshot)
    '                Dim insertPt As Point3d = Nothing
    '                'Loop through the records
    '                If rs.RecordCount > 0 Then
    '                    Dim dblOrient As Double = 0
    '                    Do Until rs.EOF
    '                        dblOrient = 100
    '                        Dim alt As Double = 0
    '                        'Get the altitude of the point (if available)
    '                        Try
    '                            If Not IsNothing(rs.Fields("ALTITUDE").Value) Then
    '                                alt = Convert.ToDouble(rs.Fields("ALTITUDE").Value)
    '                            End If
    '                        Catch ex As Exception
    '                        End Try
    '                        'Get the position at which to insert the point
    '                        insertPt = New Point3d(Convert.ToDouble(rs.Fields("X").Value), Convert.ToDouble(rs.Fields("Y").Value), alt)
    '                        'get the rotation for the point (if available)
    '                        Try
    '                            dblOrient = Convert.ToDouble(rs.Fields("ORI").Value)
    '                        Catch ex As Exception
    '                        End Try
    '                        'See if the block should be on a specific layer
    '                        rsTemp.Seek("=", strJCBlock)
    '                        If rsTemp.NoMatch = False Then
    '                            strJCLayer = rsTemp.Fields("JC_Layer").Value.ToString
    '                        Else
    '                            strJCLayer = "0"
    '                        End If
    '                        ActivateLayer(strJCLayer)
    '                        'Get the ID of the layer to put the block on
    '                        Dim layID As ObjectId = GetLayerIDFromLayerName(strJCLayer)
    '                        'Call the function to insert the block
    '                        Dim oID As ObjectId = InsertBlock(insertPt, strJCBlock, dblOrient, layID, dblFactEch)
    '                        'If the block was inserted successfully then call the function to add its attributes
    '                        If Not oID.IsNull Then
    '                            SetBlockAttributeValues(rs, oID)
    '                        End If
    '                        'next record
    '                        rs.MoveNext()
    '                    Loop
    '                End If
    '                rs.Close()
    '            End If
    '        Next
    '        'Save the changes
    '        trans.Commit()
    '        Try
    '            rs.Close()
    '        Catch ex As Exception
    '        End Try
    '    End Using

    'End Sub

    ''' <summary>
    ''' Builds the area entities
    ''' </summary>
    ''' <param name="CreerPartTerr"></param>
    ''' <remarks></remarks>
    Private Sub BuildAreaEntities(ByVal CreerPartTerr As Boolean)  ' PROBLEME AVEC AUTOCAD 2015  !!!!!!!!!!!!!!
        'Thaw layer 0
        FreezeLayer("0", False)
        'And set it current
        ActivateLayer("0")
        'Freeze all other layers
        FreezeAllLayers(True)
        'Prog.ProgressBar1.Maximum = 6
        'Prog.lblMsg.Text = "Building Area Entities..."
        'Prog.ProgressBar1.Value = 0
        '        conn.Message("Revo", "Création des entités surfaciques", True, 55, 100) ' include SQLite in Revo THA
        conn.Message(Ass.xTitle, "Création des entités surfaciques", False, 60, 100) ' include SQLite in Revo THA
        '   frms.Application.DoEvents()

        'Create the area partitions
        If CreerPartTerr Then
            CreateAreaPartition("MO_CS")
            conn.Message(Ass.xTitle, "Création des entités surfaciques: MO_CS", False, 62, 100) ' include SQLite in Revo THA

            If booCancelProcess Then Exit Sub
            CreateAreaPartition("MO_BF", "MO_CS")
            conn.Message(Ass.xTitle, "Création des entités surfaciques: MO_BF", False, 64, 100) ' include SQLite in Revo THA
            If booCancelProcess Then Exit Sub

        End If
        CreateAreaPartition("MO_RPL", "MO_BF")
        conn.Message(Ass.xTitle, "Création des entités surfaciques: MO_RPL", False, 65, 100) ' include SQLite in Revo THA
        If booCancelProcess Then Exit Sub
        CreateAreaPartition("MO_COM", "MO_RPL")
        conn.Message(Ass.xTitle, "Création des entités surfaciques: MO_COM", False, 66, 100) ' include SQLite in Revo THA
        If booCancelProcess Then Exit Sub
        CreateAreaPartition("MO_RNT", "MO_COM")
        conn.Message(Ass.xTitle, "Création des entités surfaciques: MO_RNT", False, 68, 100) ' include SQLite in Revo THA
    End Sub

    ''' <summary>
    ''' CReates the area partitions
    ''' </summary>
    ''' <param name="strJCLayer">The layer containing the features to create the area partitions from</param>
    ''' <param name="strJCPrevLayer">Optional previous layer to freeze</param>
    ''' <param name="booFastSearch"></param>
    ''' <remarks></remarks>
    Public Sub CreateAreaPartition(ByVal strJCLayer As String, Optional ByVal strJCPrevLayer As String = "", Optional ByVal booFastSearch As Boolean = True)
        'Prog.ProgressBar1.Value += 1
        ''frms.Application.DoEvents()
        'Get the database for the drawing
        'Start a transaction with the database
        'Thaw the layer to work on
        FreezeLayer(strJCLayer & "_INCOMPLET", False)
        'And set it active
        ActivateLayer(strJCLayer & "_INCOMPLET")
        'If we have been passed a previous layer then freeze it
        If strJCPrevLayer <> "" Then FreezeLayer(strJCPrevLayer & "_INCOMPLET", True)


     
            'CReate regions for all the polylines on this layer
        ' CreateRegionsForAll2(strJCLayer & "_INCOMPLET")
        CreateRegionsForAll()
       

        ''frms.Application.DoEvents()

        If booCancelProcess Then Exit Sub
        ''''''''SKIPPED NEXT SECTION'''''''''
        ''Generation of surfaces for the different sub groups
        'Dim lngQueryCount As Int64 = objDb.QueryDefs.Count
        'For i As Int64 = 0 To lngQueryCount - 1
        '    Dim strQueryName As String = objDb.QueryDefs(i).Name
        '    If strQueryName.StartsWith("+" & strJCLayer, StringComparison.Ordinal) Then
        '    Dim Centroids As CoordCollection(Of Coord) = GetCentroids(strQueryName)       
        '        If Centroids.Count > 0 Then
        '            For j As Integer = 0 To Centroids.Count - 1
        '                Dim c As Coord = DirectCast(Centroids(j), Coord)
        '                Dim pLine As Polyline = Boundary(c.X, c.Y)
        '                If Not pLine Is Nothing Then
        '                    SetObjectLayer(pLine.ObjectId, strQueryName.Replace("+", ""))
        '                End If
        '            Next
        '        End If
        '    End If
        'Next
        '''''''END OF SKIPPED SECTION''''''''''
        'Conversion of regions to polylines
        ConvertRegionsToPolylines()


        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction


            frms.Application.DoEvents()
            If booCancelProcess Then Exit Sub
            'Call the function to retrieve the polylines
            Dim plCol As PlVertCol(Of PolylineVertexes) = GetPolylines(strJCLayer & "_INCOMPLET")
            Dim xDataID As Integer = 0
            'If there are any polylines
            If plCol.Count > 0 Then
                Dim dtQs As System.Data.DataTable = GetQueries()
                If Not dtQs Is Nothing Then
                    'Loop through the queries ' ???????????????????????????????????????????????????????????
                    For i As Integer = 0 To dtQs.Rows.Count - 1 'Prend énormement de temps !!!! C'est lent ! A optimiser ! Surtout lorsqu'il ne trouve apparament rien !???
                        'Get the query name
                        Dim strQueryName As String = dtQs.Rows(i)("QueryName").ToString
                        'If the query name starts with "+"
                        If strQueryName.StartsWith("+" & strJCLayer, StringComparison.Ordinal) Then
                            'Call the function to get the centroids from this query
                            Dim Centroids As CoordCollection(Of Coord) = GetCentroids(dtQs.Rows(i))
                            'If there are any centroids returned
                            If Centroids.Count > 0 Then

                                conn.Message(Ass.xTitle & "  (" & strQueryName & " - " & strJCLayer & " - " & i & ")", "Recherche des entités surfaciques: " & i & " / " & dtQs.Rows.Count - 1 _
                                                  & "     (centroïdes : " & Centroids.Count & ")", False, i, dtQs.Rows.Count - 1)

                                'Loop through the centroids 
                                Dim NbreCentr As Double = 0
                                For Each c As Coord In Centroids
                                    NbreCentr += 1
                                    ' conn.Message(Ass.xTitle, "Recherche des entités surfaciques: " & i & " / " & dtQs.Rows.Count - 1 _
                                    '                    & "     (centroïdes : " & NbreCentr & " / " & Centroids.Count & ")", False, i, dtQs.Rows.Count - 1)

                                    'Collection to hold the polyline(s) that the centroid is within
                                    Dim FoundPolys As New PlVertCol(Of PolylineVertexes)
                                    Dim intLoc As Integer 'not used


                                    '1er recherche

                                    'Loop through the polylines checking if the centroid is within them
                                    For Each pl As PolylineVertexes In plCol ' *** ???? Faire un filtre avec une distance approch. ???

                                        'If the polyline doesn't have arc segements
                                        If pl.WithArcs = False Then
                                            'Use the mathematical algorithm to determine if the centroid is within it
                                            If PointInsidePolygon(pl.Verts, New Point2d(c.X, c.Y), intLoc) Then
                                                'If it is, then add it to the FoundPolys collection
                                                FoundPolys.Add(pl)
                                            End If
                                        Else 'The polyline contains arc segments so use the Ray in AutoCAD method
                                            'to determine if the centroid is within  it
                                            If PtIsInsideLWP(New Point3d(c.X, c.Y, 0), pl.AcadObject) Then
                                                'If it is, then add it to the FoundPolys collection
                                                FoundPolys.Add(pl)
                                            End If   'ATTENTION TEST CONDITION ---- !!!!! THA 30.03.2011: re-Active 8.10.2011

                                        End If
                                    Next


                                    '2ème recherche si trouvé => FoundPolys.Count = 0 





                                    'Process the found polys
                                    If FoundPolys.Count = 0 Then
                                        'Nothing found (shouldn't happen)
                                    ElseIf FoundPolys.Count = 1 Then 'The centroid was within only one polyline
                                        Dim pLine As Polyline = DirectCast(trans.GetObject(FoundPolys.Item(0).AcadObject, OpenMode.ForRead), Polyline)
                                        'Call the function to close the polyline and set its layer to the name of the query without the "+"
                                        ClosePolyAndSetLayer(pLine.ObjectId, Replace(strQueryName, "+", ""))
                                        'IF the query name is "+MO_CS_BAT" then add the xdata
                                        If strQueryName = "+MO_CS_BAT" Then
                                            xDataID += 1
                                            'Call the function to get the xdata information from the centroid
                                            Dim resBuf As ResultBuffer = GetXDataInfoFromCentroid(c, xDataID)
                                            'Set the xdata on the polyline
                                            'SetObjXRecord(pLine.ObjectId, "REVO", resBuf)
                                            SetPolyXData(pLine.ObjectId, resBuf)
                                        End If
                                    Else 'The centroid was within more than one polyline
                                        Dim dblMinArea As Double = 0
                                        Dim SmallestPolyIndx = 0
                                        Dim pLine As Polyline
                                        'Loop through the found polylines looking for one with the smallest area
                                        For j As Integer = 0 To FoundPolys.Count - 1
                                            pLine = DirectCast(trans.GetObject(FoundPolys.Item(j).AcadObject, OpenMode.ForRead), Polyline)
                                            'If we haven't already store the minimum
                                            If dblMinArea = 0 Then
                                                'Set the minimum to the area of the polyline
                                                dblMinArea = pLine.Area
                                                'Set the smallest poly index to 0
                                                SmallestPolyIndx = 0
                                            Else
                                                'Test if this polyline is smaller than the previously stored smallest polyline
                                                If pLine.Area < dblMinArea Then
                                                    'If so, then store this poly area as the smallest one
                                                    dblMinArea = pLine.Area
                                                    'Set the smallest poly index to j
                                                    SmallestPolyIndx = j
                                                End If
                                            End If
                                        Next
                                        'Close the smallest poly and set the layer
                                        ClosePolyAndSetLayer(FoundPolys.Item(SmallestPolyIndx).AcadObject, Replace(strQueryName, "+", ""))
                                        'If the query name is "+MO_CS_BAT" then add the xdata
                                        If strQueryName = "+MO_CS_BAT" Then
                                            xDataID += 1
                                            'Call the function to get the xdata information from the centroid
                                            Dim resBuf As ResultBuffer = GetXDataInfoFromCentroid(c, xDataID)
                                            'Set the xdata on the polyline
                                            SetPolyXData(FoundPolys.Item(SmallestPolyIndx).AcadObject, resBuf)
                                            'SetObjXRecord(FoundPolys.Item(SmallestPolyIndx).AcadObject, "REVO", resBuf)                                    
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                End If
            End If
            trans.Commit()
            plCol = Nothing
        End Using
    End Sub

    ''' <summary>
    ''' Gets the information for the xdata from the centroid
    ''' </summary>
    ''' <param name="c">The centroid to get the data from</param>
    ''' <param name="xDataID">The xdata id</param>
    ''' <returns>A resultbuffer containing the data</returns>
    ''' <remarks></remarks>
    Public Function GetXDataInfoFromCentroid(ByVal c As Coord, ByVal xDataID As Integer) As ResultBuffer
        AddRegAppTableRecord(Ass.xProduct.ToUpper) '("REVO")
        'Declare a new result buffer
        Dim acResBuf As New ResultBuffer
        acResBuf.Add(New TypedValue(DxfCode.ExtendedDataRegAppName, Ass.xProduct.ToUpper)) '"REVO"))
        'The ID is where the data is held
        acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, "ID" & xDataID))
        If c.ID.Contains("|") Then
            'Split the data held in the ID property of the centroid
            Dim dataBits() As String = c.ID.Split(Convert.ToChar("|"))
            'Add the data parts to the resultbuffer
            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, dataBits(0).Replace("_", " ") & ""))
            acResBuf.Add(New TypedValue(DxfCode.ExtendedDataAsciiString, dataBits(1).Replace("_", " ") & ""))
        End If
        Return acResBuf
    End Function

    Public Sub CreateRegionsForAll()
        'CommandLine.Command("_REGION", "ALL", "")
        CommandLine.CommandC("_.REGION", "_ALL", "") 'TOUT   ' Modif. 2015 ???
    End Sub

    Public Sub CreateRegionsForAll_2(Optional LayerName As String = "")

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Using docLock As DocumentLock = acDoc.LockDocument


            Try
                Dim db = HostApplicationServices.WorkingDatabase
                Dim curveClass = RXClass.GetClass(GetType(Curve))
                Using tr = db.TransactionManager.StartTransaction()
                    Dim curSpace = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite), BlockTableRecord)
                    Dim curves = New DBObjectCollection()
                    For Each id As ObjectId In curSpace
                        If id.ObjectClass.IsDerivedFrom(curveClass) Then
                            Dim curve = DirectCast(tr.GetObject(id, OpenMode.ForRead), Curve)

                            If LayerName = "" Then
                                curves.Add(curve)
                            Else
                                If curve.Layer = LayerName Then curves.Add(curve)
                            End If

                        End If
                    Next

                    If curves.Count <> 0 Then
                        For Each dbObj As DBObject In Region.CreateFromCurves(curves)
                            Dim region1 = DirectCast(dbObj, Region)
                            region1.SetDatabaseDefaults()
                            curSpace.AppendEntity(region1)

                            'curSpace.AppendEntity(region1) 'Add THA 2015
                            tr.AddNewlyCreatedDBObject(region1, True)
                        Next
                    End If

                    tr.Commit()
                End Using

            Catch ex As Exception
                MsgBox("CreateRegionsForAll : " & ex.Message)
            End Try

        End Using
    End Sub


    ''' <summary>
    ''' Returns a collection of coord objects containing the centroids read from the database
    ''' </summary>
    ''' <param name="strQueryRow">The Query datarow form the Queries table containing the Query Name and the query SQL statement</param>
    ''' <returns>A collection of coord objects holding the centroids</returns>
    ''' <remarks></remarks>
    Private Function GetCentroids(ByVal strQueryRow As DataRow) As CoordCollection(Of Coord)
        'Create a new coordcollection to hold the centroids
        Dim Centroids As New CoordCollection(Of Coord)
        'Open the database
        'OpenDatabase()
        'Open the query
        Dim dt As System.Data.DataTable = GetTable(strQueryRow("QueryName").ToString, strQueryRow("SQL").ToString)
        'If there are any records in the recordset
        If dt.Rows.Count > 0 Then
            'Loop until the end of the recordset
            For i As Integer = 0 To dt.Rows.Count - 1
                'Create a new coord object
                Dim C As New Coord
                'Store the x coordinate
                C.X = Convert.ToDouble(dt.Rows(i)("Geometrie_re"))
                'Store the y coordinate
                C.Y = Convert.ToDouble(dt.Rows(i)("Geometrie_ho"))
                'Particular case for buildings. Store the number and the type
                If strQueryRow("QueryName").ToString = "+MO_CS_BAT" Then
                    C.ID = dt.Rows(i)("NUMERO").ToString & "|" & dt.Rows(i)("DESIGNATION").ToString
                End If
                'Add this centroid to the collection
                Centroids.Add(C)
                'rs.MoveNext()
            Next
        End If
        'Return the centroid collection
        Return Centroids
    End Function

    '''' <summary>
    '''' Returns a collection of coord objects containing the centroids read from the database
    '''' </summary>
    '''' <param name="strQueryName">The name of the query to retrieve the information from the Access database</param>
    '''' <returns>A collection of coord objects holding the centroids</returns>
    '''' <remarks></remarks>
    'Private Function GetCentroids(ByVal strQueryName As String) As CoordCollection(Of Coord)
    '    'Create a new coordcollection to hold the centroids
    '    Dim Centroids As New CoordCollection(Of Coord)
    '    'Open the database
    '    OpenDatabase()
    '    'Open the query
    '    Dim rs As dao.Recordset = objDb.OpenRecordset(strQueryName, dao.RecordsetTypeEnum.dbOpenSnapshot)
    '    'If there are any records in the recordset
    '    If rs.RecordCount > 0 Then
    '        'Loop until the end of the recordset
    '        Do Until rs.EOF
    '            'Create a new coord object
    '            Dim C As New Coord
    '            'Store the x coordinate
    '            C.X = Convert.ToDouble(rs.Fields("Geometrie_re").Value)
    '            'Store the y coordinate
    '            C.Y = Convert.ToDouble(rs.Fields("Geometrie_ho").Value)
    '            'Particular case for buildings. Store the number and the type
    '            If strQueryName = "+MO_CS_BAT" Then
    '                C.ID = rs.Fields("NUMERO").Value.ToString & "|" & rs.Fields("DESIGNATION").Value.ToString
    '            End If
    '            'Add this centroid to the collection
    '            Centroids.Add(C)
    '            rs.MoveNext()
    '        Loop
    '    End If
    '    'Return the centroid collection
    '    Return Centroids
    'End Function
    ''' <summary>
    ''' Calls the AutoCAD boundary command passing the centroid of the area
    ''' and attempts to return the polyline created by the boundary command 
    ''' </summary>
    ''' <param name="dblX">The X coordinate of the centroid</param>
    ''' <param name="dblY">The Y coordinate of the centroid</param>
    ''' <returns>The polyline created by the boundary command if successful, otherwise, nothing</returns>
    ''' <remarks></remarks>
    Public Function Boundary(ByVal dblX As Double, ByVal dblY As Double) As Polyline
        Dim PrevCount As Integer = 0
        Dim Pline As Polyline = Nothing
        'Set the system variable to not show the message display in the command window
        Application.SetSystemVariable("NOMUTT", 1)
        'Get the active document
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        'Get the database for the current document
        Dim db As Database = doc.Database
        'Start as transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the blocktable
            Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord
            'Get the current layoutmanager
            Dim lm As LayoutManager = LayoutManager.Current
            'Get the blocktablerecord for the current space
            If lm.CurrentLayout.ToUpper = "MODEL" Then
                btr = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
            Else
                btr = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.PaperSpace), OpenMode.ForRead), BlockTableRecord)
            End If
            'Get the entities in the current space
            Dim bRefs As ObjectIdCollection = btr.GetBlockReferenceIds(True, True)
            'Store the count of those entities
            PrevCount = bRefs.Count
            'Call the boundary command passing the centroid
            CommandLine.CommandC("_-boundary", New Point2d(dblX, dblY))
            'Get the count of entities now
            bRefs = btr.GetBlockReferenceIds(True, True)
            'If the count is greater than before running the boundary command 
            If bRefs.Count > PrevCount Then
                Pline = DirectCast(trans.GetObject(bRefs(bRefs.Count - 1), OpenMode.ForRead), Polyline)
            End If
        End Using
        Application.SetSystemVariable("NOMUTT", 0)
        Return Pline
    End Function

    ''' <summary>
    ''' Converts the region features to polylines
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ConvertRegionsToPolylines()
        'Get the current layer
        Dim strCurLayer As String = GetCurrentLayer()
        'Set up typedvalues to find Regions on the current layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "REGION"), New TypedValue(DxfCode.LayerName, strCurLayer)}
        'Call the function to select all items using the typedvalues
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then 'If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Loop through the regions we found
                For Each oid As ObjectId In pres.Value.GetObjectIds
                    Dim reg As Region = DirectCast(trans.GetObject(oid, OpenMode.ForRead), Region)
                    'Call the function to explode the region
                    Dim entCol As ObjectIdCollection = ExplodeRegion(reg)
                    'Erase the region
                    reg.Erase()
                    ''GOTO MethodePedit
                    'Dim oLWP1 As Polyline = Nothing
                    'Dim oLWP2 As Polyline = Nothing
                    'For i As Integer = 0 To Segments.Count - 1
                    '    If i = 0 Then
                    '        oLWP1 = ConvertEntityToLWPolyLine(DirectCast(Segments(i), Entity))
                    '    End If
                    '    oLWP2 = ConvertEntityToLWPolyLine(DirectCast(Segments(i + 1), Entity))
                    '    If oLWP1 Is Nothing Or oLWP2 Is Nothing Then Exit For
                    '    If Join2Polylines(oLWP1, oLWP2, 0.001) Then Exit For
                    'Next
                    'MethodePedit
                    'Call the autocad command to join the individul bits of the region together 


                    ' --- A TESTER pour supprimer le message d'erreur : Ignorer Ordre de tracé ---------------------   3 juin 2013  THA
                    ' -----------------------------------------------------------------------------------------------

                    If strCurLayer = "MO_RNT_INCOMPLET" Or strCurLayer = "MO_COM_INCOMPLET" Then
                        CommandLine.CommandC("_.PEDIT", "m", "_all", "", "_yes", "_joi", "", "", "") ' Modif. 2015
                    Else
                        CommandLine.CommandC("_.PEDIT", entCol(0), "_YES", "_JOI", "_ALL", "", "", "") ' Modif. 2015
                    End If
                Next
                'Save the changes
                trans.Commit()
            End Using
        End If
    End Sub

    ''' <summary>
    ''' Explodes the region
    ''' </summary>
    ''' <param name="reg">The region to explode</param>
    ''' <returns>An objectIDCollection of the region parts</returns>
    ''' <remarks></remarks>
    Public Function ExplodeRegion(ByVal reg As Region) As ObjectIdCollection
        'Get the current layer
        Dim strCurLayer As String = GetCurrentLayer()
        'Get the ID for the current layer so we can set the exploded region bits to the correct layer
        Dim LayID As ObjectId = GetLayerIDFromLayerName(strCurLayer)
        Dim retCol As New ObjectIdCollection 'Collection to return
        Dim Segments As New DBObjectCollection 'Collection to hold the exploded region bits
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Make sure the region is editable
            If Not reg.IsWriteEnabled Then reg.UpgradeOpen()
            'Explode it
            reg.Explode(Segments)
            'Get the block table record for the modelspace
            Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
            'Loop through the entities in the exploded region adding them to the modelspace and setting their layer
            For Each id As DBObject In Segments
                'Get the entity
                Dim ent As Entity = DirectCast(id, Entity)
                If Not TypeOf ent Is Region Then
                    'Add the entity to modelspace
                    Dim newId As ObjectId = btr.AppendEntity(DirectCast(id, Entity))
                    'Set the entity's layer
                    ent.LayerId = LayID
                    'Let the transaction knwo about the entity
                    trans.AddNewlyCreatedDBObject(DirectCast(id, Entity), True)
                    'Add the id  of the entity to the return collection
                    retCol.Add(newId)
                Else
                    reg.Erase()
                End If
            Next
            'Save the changes
            trans.Commit()
        End Using
        'Return the collection
        Return retCol
    End Function

    ''' <summary>
    ''' Gets a collection of PolylineVertexes objects containing the polylines on the required layer
    ''' </summary>
    ''' <param name="strLayer">The layer to get the polylines from</param>
    ''' <returns>A collection of PolylineVertexes objects</returns>
    ''' <remarks></remarks>
    Public Function GetPolylines(ByVal strLayer As String) As PlVertCol(Of PolylineVertexes)
        'CReate a new PolylineVertexes collection
        Dim plCol As New PlVertCol(Of PolylineVertexes)
        Dim plCount As Int64 = 0
        Dim VertCount As Int64 = 0
        'Create the typedvalues to select Polylines on the required layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, strLayer)}
        'Select all the polylines on the required layer
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then 'If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Loop through the polylines
                For Each oid As ObjectId In pres.Value.GetObjectIds()
                    Dim pLine As Polyline = DirectCast(trans.GetObject(oid, OpenMode.ForRead), Polyline)
                    'Make the polyline writeable
                    pLine.UpgradeOpen()
                    'Ensure it's closed
                    pLine.Closed = True
                    'Call the function to get the polyline vertices
                    Dim verts As Point2dCollection = GetPolyCoordinates(pLine)
                    'Create a new PolylineVertexes object to hold this polyline
                    Dim plV As New PolylineVertexes
                    'Store the objectid
                    plV.AcadObject = pLine.ObjectId
                    plV.WithArcs = False
                    'Store the vertices
                    plV.Verts = verts
                    'Loop through the vertices checking the bulge value at each one                    
                    For i As Integer = 0 To verts.Count - 1
                        'If the bulge value is greater then 0 then the polyline contains arc segments
                        If pLine.GetBulgeAt(i) > 0 Then
                            'Set the WithArcs property of the PolylineVertxes object to True
                            plV.WithArcs = True
                            Exit For
                        End If
                    Next
                    'Add it to the collection
                    plCol.Add(plV)
                Next
                'Save any changes
                trans.Commit()
            End Using
        End If
        'Return the collection
        Return plCol
    End Function

    ''' <summary>
    ''' Tests if the supplied point is within the supplied polyline
    ''' </summary>
    ''' <param name="startPt">The point to check</param>
    ''' <param name="oId">The objectId of the polyline to check</param>
    ''' <returns>True if the point is inside, otherwise false</returns>
    ''' <remarks></remarks>
    Public Function PtIsInsideLWP(ByVal startPt As Point3d, ByVal oId As ObjectId) As Boolean


        ' Fonction a optimiser
        ' Dans les gros fichiers ili c'est très long :  THA 30.03.2011

        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the polyline
            Dim LWP As Polyline = DirectCast(trans.GetObject(oId, OpenMode.ForRead), Polyline)
            Dim IntersectPts As Point3dCollection = Nothing
            Dim Inside As Boolean = True
            Dim InCount As Integer = 0
            Dim OutCount As Integer = 0
            'Loop four times firing rays in four different directions
            For i As Int16 = 1 To 4
                'Create a new Ray
                Dim objRay As New Ray
                'Set the ray start point
                objRay.BasePoint = startPt
                Select Case i
                    Case 1
                        'and give it a direction
                        objRay.UnitDir = New Vector3d(1, 0, 0)
                    Case 2
                        'and give it a direction
                        objRay.UnitDir = New Vector3d(0, 1, 0)
                    Case 3
                        'and give it a direction
                        objRay.UnitDir = New Vector3d(-1, 0, 0)
                    Case 4
                        'and give it a direction
                        objRay.UnitDir = New Vector3d(0, -1, 0)
                End Select
                'Declare a collection to hold the intersection points
                IntersectPts = New Point3dCollection
                'Call the IntersectWith method of the ray to store any intersections between it and the polyline
                'objRay.IntersectWith(LWP, Intersect.OnBothOperands, IntersectPts, CType(0, System.IntPtr), CType(0, System.IntPtr))
                'OBSOLETE    objRay.IntersectWith(LWP, Intersect.OnBothOperands, IntersectPts, 0, 0)

                'New Function : THA 07.02.2012
                objRay.IntersectWith(LWP, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)
                'Example : List_Entities(i).IntersectWith(List_Entities(i + 1), Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)


                objRay.Dispose()
                'Point is inside if intersectpoints count is odd
                Inside = (IntersectPts.Count And 1) = 1
                If Inside Then
                    InCount += 1
                Else
                    OutCount += 1
                End If
            Next
            Return InCount > OutCount
        End Using
    End Function

    Public Function ConvertDictionary(Of key, value)(ByVal dic As Dictionary(Of key, value)) As List(Of KeyValuePair(Of key, value))
        Dim lst As New List(Of KeyValuePair(Of key, value))
        For Each p As KeyValuePair(Of key, value) In dic
            lst.Add(p)
        Next
        Return lst
    End Function

    Public Function CompareIntKeys(ByVal a As KeyValuePair(Of Integer, Point2d), ByVal b As KeyValuePair(Of Integer, Point2d)) As Integer
        Return a.Key - b.Key
    End Function


    '''' <summary>
    '''' Selects polylines within the crossing window defined by the supplied points
    '''' </summary>
    '''' <param name="strLayer">The layer that the selected polylines should be on</param>
    '''' <param name="p1">The first corner of the selection window</param>
    '''' <param name="p2">The opposite corner of the selection window</param>
    '''' <returns>A list of PolylineVertexes</returns>
    '''' <remarks></remarks>
    'Public Function GetPolylinesByRect(ByVal strLayer As String, ByVal p1 As Point3d, _
    '                                   ByVal p2 As Point3d) As PlVertCol(Of PolylineVertexes)
    '    Dim colPLVerts As New PlVertCol(Of PolylineVertexes)
    '    'Set up the typedvalues to select polylines on the desired layer
    '    Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, strLayer)}
    '    'Create the selectionfilter from the typedvalues
    '    Dim sFilter As New SelectionFilter(Values)
    '    'Get the editor
    '    Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
    '    'Select polylines within the crossing window
    '    Dim pres As PromptSelectionResult = ed.SelectCrossingWindow(p1, p2, sFilter, False)
    '    'Check that we found something
    '    If pres.Status = PromptStatus.OK Then
    '        'Get the database
    '        Dim db As Database = HostApplicationServices.WorkingDatabase
    '        'And start a transaction
    '        Using trans As Transaction = db.TransactionManager.StartTransaction
    '            'Loop through the selected polylines
    '            For Each id As ObjectId In pres.Value.GetObjectIds
    '                'get the polyline from the database
    '                Dim pLine As Polyline = CType(trans.GetObject(id, OpenMode.ForRead), Polyline)
    '                pLine.UpgradeOpen()
    '                'Close the polyline
    '                pLine.Closed = True
    '                'Call the function to get the polyline vertices
    '                Dim verts As Point2dCollection = GetPolyCoordinates(pLine)
    '                'Create a new PolylineVertexes object to hold this polyline
    '                Dim plV As New PolylineVertexes
    '                'Store the objectid
    '                plV.AcadObject = pLine.ObjectId
    '                plV.WithArcs = False
    '                'Store the vertices
    '                plV.Verts = verts
    '                'Loop through the vertices checking the bulge value at each one                    
    '                For i As Integer = 0 To verts.Count - 1
    '                    'If the bulge value is greater then 0 then the polyline contains arc segments
    '                    If pLine.GetBulgeAt(i) > 0 Then
    '                        'Set the WithArcs property of the PolylineVertxes object to True
    '                        plV.WithArcs = True
    '                        Exit For
    '                    End If
    '                Next
    '                'Add this PolylineVertex object to the return list
    '                colPLVerts.Add(plV)
    '            Next
    '        End Using
    '    End If
    '    Return colPLVerts
    'End Function



    ''' <summary>
    ''' Selects polylines within the crossing window defined by the supplied points
    ''' </summary>
    ''' <param name="strLayer">The layer that the selected polylines should be on</param>
    ''' <param name="p1">The first corner of the selection window</param>
    ''' <param name="p2">The opposite corner of the selection window</param>
    ''' <returns>A list of PolylineVertexes</returns>
    ''' <remarks></remarks>
    Public Function GetPolylinesByRect(ByVal strLayer As String, ByVal p1 As Point3d, _
                                       ByVal p2 As Point3d) As PlVertCol(Of PolylineVertexes)
        Dim colPLVerts As New PlVertCol(Of PolylineVertexes)
        'Set up the typedvalues to select polylines on the desired layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, strLayer)}
        'Create the selectionfilter from the typedvalues
        Dim sFilter As New SelectionFilter(Values)
        'Get the editor
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        'Select polylines within the crossing window
        Dim pres As PromptSelectionResult = ed.SelectCrossingWindow(p1, p2, sFilter, False)
        'Check that we found something
        If pres.Status = PromptStatus.OK Then
            'Get the database
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'And start a transaction
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Loop through the selected polylines
                For Each id As ObjectId In pres.Value.GetObjectIds
                    'get the polyline from the database
                    Dim pLine As Polyline = CType(trans.GetObject(id, OpenMode.ForRead), Polyline)
                    pLine.UpgradeOpen()
                    'Close the polyline
                    pLine.Closed = True
                    'Call the function to get the polyline vertices
                    Dim verts As Point2dCollection = GetPolyCoordinates(pLine)
                    'Create a new PolylineVertexes object to hold this polyline
                    Dim plV As New PolylineVertexes
                    'Store the objectid
                    plV.AcadObject = pLine.ObjectId
                    plV.WithArcs = False
                    'Store the vertices
                    plV.Verts = verts
                    'Loop through the vertices checking the bulge value at each one                    
                    For i As Integer = 0 To verts.Count - 1
                        'If the bulge value is greater then 0 then the polyline contains arc segments
                        If pLine.GetBulgeAt(i) > 0 Then
                            'Set the WithArcs property of the PolylineVertxes object to True
                            plV.WithArcs = True
                            Exit For
                        End If
                    Next
                    'Add this PolylineVertex object to the return list
                    colPLVerts.Add(plV)
                Next
            End Using
        End If
        Return colPLVerts
    End Function

End Module
