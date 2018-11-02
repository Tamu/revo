Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports System.IO
Imports frms = System.Windows.Forms
Module modPER

    Dim Ass As New Revo.RevoInfo

    Public Class PER_PL
        Private _Theme As String
        Private _IdentDN As String
        Private _Numero As String
        Private _Signe As String
        Private _E As Double
        Private _N As Double
        Private _PrecPlan As String
        Private _FiabPlan As String
        Private _ObjID As ObjectId
        Private _RayonArc As Double
        Private _Index As Integer

        Public Property Theme() As String
            Get
                Return _Theme
            End Get
            Set(ByVal value As String)
                _Theme = value
            End Set
        End Property
        Public Property IdentDN() As String
            Get
                Return _IdentDN
            End Get
            Set(ByVal value As String)
                _IdentDN = value
            End Set
        End Property
        Public Property Numero() As String
            Get
                Return _Numero
            End Get
            Set(ByVal value As String)
                _Numero = value
            End Set
        End Property
        Public Property Signe() As String
            Get
                Return _Signe
            End Get
            Set(ByVal value As String)
                _Signe = value
            End Set
        End Property
        Public Property E() As Double
            Get
                Return _E
            End Get
            Set(ByVal value As Double)
                _E = value
            End Set
        End Property
        Public Property N() As Double
            Get
                Return _N
            End Get
            Set(ByVal value As Double)
                _N = value
            End Set
        End Property
        Public Property PrecPlan() As String
            Get
                Return _PrecPlan
            End Get
            Set(ByVal value As String)
                _PrecPlan = value
            End Set
        End Property
        Public Property FiabPlan() As String
            Get
                Return _FiabPlan
            End Get
            Set(ByVal value As String)
                _FiabPlan = value
            End Set
        End Property
        Public Property ObjID() As ObjectId
            Get
                Return _ObjID
            End Get
            Set(ByVal value As ObjectId)
                _ObjID = value
            End Set
        End Property
        Public Property RayonArc() As Double
            Get
                Return _RayonArc
            End Get
            Set(ByVal value As Double)
                _RayonArc = value
            End Set
        End Property
        Public Property Index() As Integer
            Get
                Return _Index
            End Get
            Set(ByVal value As Integer)
                _Index = value
            End Set
        End Property
    End Class
    Private lstPL As List(Of PER_PL)
    Private Wait As frmWait = Nothing

    Public Function CalculPER(ByVal lst As List(Of PerListItem)) As Boolean
        'call the function that sets up the progress form
        SetupWaitForm()
        Application.ShowModelessDialog(Wait)
        'Save the current layer state
        SaveLayerState("REVO_Temp")
        'List of layers to switch on and thaw
        Dim Layers As String() = New String() {"MO_BF", "MO_BF_PTS", "MO_BF_IMMEUBLE", _
                                                     "MO_CS_BAT", "MO_CS_BAT_NUM", "MO_CS_PTS", _
                                                     "MUT_BF", "MUT_BF_PTS", "MUT_BF_IMMEUBLE", _
                                                     "MUT_CS_BAT", "MUT_CS_BAT_NUM", "MUT_CS_PTS", _
                                               "MO_ODS_BATSOUT", "MO_ODS_BATS_NUM"}
        'Switch the layers on 
        ShowLayers(Layers)
        'And thaw them
        FreezeLayers(Layers, False)
        Zooming.ZoomExtents()
        'Name and location of report file
        Dim strReport As String = Ass.PERFolder ' LireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "PERFolder")
        If Not strReport.EndsWith("\") Then strReport = strReport & "\"
        strReport = strReport & "PER-" & Replace(DateTime.Now.ToString("yyMMdd") & TimeOfDay, ":", "") & ".htm"

        Try

            Dim sWriter As New StreamWriter(strReport)
            'Write the header of the report
            WritePERHeader(sWriter)
            For Each PEL As PerListItem In lst
                'List to hold the result objects
                lstPL = New List(Of PER_PL)
                'Progress update
                If PEL.PolyLayer.Contains("_BF") Then
                    Wait.lblStatus1.Text = "Traiment du bien-fonds " & PEL.ComNum
                Else
                    Wait.lblStatus1.Text = "Traiment du bâtiment " & PEL.ComNum
                End If
                Wait.progBar1.Value = 10
                frms.Application.DoEvents()
                'BF Object (polyline)
                Dim polyID As ObjectId = PEL.PolyID
                'Analysis of BF / PER extraction
                If ExtractPER(polyID) Then
                    Wait.progBar1.Value = 90
                    frms.Application.DoEvents()
                    Dim strObj As String = ""
                    If PEL.PolyLayer.Contains("_BF") Then
                        strObj = "bien-fonds"
                    ElseIf PEL.PolyLayer.Contains("_CS_BAT") Then
                        strObj = "b&acirc;timent"
                    End If
                    Dim dblSurf As Double = GetPolyArea(PEL.PolyID)
                    If PEL.PolyLayer.Contains("_BF") Then
                        dblSurf = dblSurf - SearchForHoles(PEL.PolyID, "BF")
                    Else
                        dblSurf = dblSurf - SearchForHoles(PEL.PolyID, "CS")
                    End If
                    'Report
                    'Definition of the object
                    sWriter.WriteLine("<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width: 170mm;"">")
                    sWriter.WriteLine("<tr>")
                    sWriter.WriteLine("  <td colspan=""2"" align=""center"">Objet: <b>" & strObj & "<br>&nbsp;</b></td>")
                    sWriter.WriteLine("</tr>")
                    sWriter.WriteLine("<tr>")
                    sWriter.WriteLine("  <td width=""25%"">Commune: <b>" & PEL.Com & "</b></td>")
                    sWriter.WriteLine("  <td>Surface technique: " & Format(dblSurf, "0.0000") & " m2</td>")
                    sWriter.WriteLine("</tr>")
                    sWriter.WriteLine("<tr>")
                    sWriter.WriteLine("  <td width=""25%"">Num&eacute;ro: <b>" & PEL.Num & "</b></td>")
                    sWriter.WriteLine("  <td>Surface technique arrondie: " & Math.Round(dblSurf, 0) & " m2</td>")
                    sWriter.WriteLine("</tr>")
                    sWriter.WriteLine("<tr>")
                    sWriter.WriteLine("  <td colspan=""2"">&nbsp;</td>") 'Space
                    sWriter.WriteLine("</tr>")
                    sWriter.WriteLine("</table>")
                    'List of vertices
                    sWriter.WriteLine("<table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width: 170mm;"">")
                    sWriter.WriteLine("<tr>")
                    sWriter.WriteLine("<td>THEME</td><td>COM</td><td>PLAN</td><td>NUM</td><td>SIGNE</td><td>PRECP</td><td>FIABP</td><td>Y</td><td>X</td><td>Rayon</td>")
                    sWriter.WriteLine("</tr>")
                    lstPL = lstPL.OrderBy(Function(x) x.Index).ToList
                    For j As Integer = 0 To lstPL.Count - 1
                        Dim PERPL As PER_PL = lstPL.Item(j)




                        If PERPL.FiabPlan = Nothing Then PERPL.FiabPlan = "-"
                        If PERPL.IdentDN = Nothing Then PERPL.IdentDN = "VD0000000000"
                        If PERPL.Numero = Nothing Then PERPL.Numero = "-"
                        If PERPL.PrecPlan = Nothing Then PERPL.PrecPlan = "0"
                        If PERPL.Signe = Nothing Then PERPL.Signe = "-"
                        If PERPL.Theme = Nothing Then PERPL.Theme = "-"


                        sWriter.WriteLine("<tr>")
                        sWriter.WriteLine("<td>" & PERPL.Theme.ToString & "</td><td>" & PERPL.IdentDN.Substring(3, 3).ToString & "</td><td>" & PERPL.IdentDN.Substring(PERPL.IdentDN.Length - 4).ToString & "</td>")
                        sWriter.WriteLine("<td>" & PERPL.Numero.ToString & "</td><td>" & PERPL.Signe.ToString & "</td><td>" & FormatNumber(PERPL.PrecPlan, 1) & "</td><td>" & PERPL.FiabPlan.ToString & "</td>")
                        sWriter.WriteLine("<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>")
                        sWriter.WriteLine("</tr>")

                        sWriter.WriteLine("<tr>")
                        sWriter.WriteLine("<td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td>")
                        sWriter.WriteLine("<td>" & Format(PERPL.E, "0.000") & "</td><td>" & Format(PERPL.N, "0.000") & "</td>")

                        If PERPL.RayonArc <> 0 Then
                            'Calculation of the radius (signed)
                            Dim dblR As Double
                            If j < lstPL.Count - 1 Then dblR = GetRadiusFromBulge(PERPL.E, PERPL.N, lstPL.Item(j + 1).E, lstPL.Item(j + 1).N, PERPL.RayonArc)
                            sWriter.WriteLine("<td>" & Format(Math.Abs(dblR), "0.000") & "</td>")
                        Else
                            sWriter.WriteLine("<td>&nbsp;</td>")
                        End If
                        sWriter.WriteLine("</tr>")
                    Next
                    sWriter.WriteLine("</table>")
                    'Separator with the following PER (if applicable)
                    If lst.IndexOf(PEL) < lstPL.Count Then WritePERSeparator(sWriter)
                End If
                Wait.progBar1.Value = 100
                Dim indx As Integer = lst.IndexOf(PEL)
                If indx = 0 Then
                    Wait.lblStatus1.Text = indx & " / " & lst.Count & " PER calculé (" & CInt((indx / lst.Count) * 100) & "%)"
                Else
                    Wait.lblStatus1.Text = indx & " / " & lst.Count & " PER calculés (" & CInt((indx / lst.Count) * 100) & "%)"
                End If
                Wait.progBar2.Value = Convert.ToInt32((indx / lst.Count) * 100)
            Next
            Wait.lblStatus1.Text = "Création du rapport..."
            Wait.progBar1.Visible = False
            Wait.lblStatus2.Text = ""
            Wait.progBar2.Visible = False
            frms.Application.DoEvents()
            Zooming.ZoomPrevious()
            Wait.Close()
            Wait.Dispose()
            WritePERFooter(sWriter)
            sWriter.Close()
            sWriter.Dispose()
            System.Diagnostics.Process.Start(strReport)


        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Function

    Private Sub SetupWaitForm()
        Wait = New frmWait
        With Wait
            .lblTitle.Text = "Génération de contour périmétrique"
            .lblSubTitle.Text = "Patienter pendant l'analyse des limites"
            .lblInfo1.Text = "Analyse du contour et extraction des points du périmètre"
            .lblInfo1.Visible = True
            .lblStatus1.Text = "Initialisation..."
            .progBar1.Visible = True
            .progBar1.Value = 0
            .progBar1.Minimum = 0
            .progBar1.Maximum = 100
            .progBar2.Visible = True
            .progBar2.Value = 0
            .progBar2.Minimum = 0
            .progBar2.Maximum = 100
        End With

    End Sub

    Private Sub ShowLayers(ByVal Layers() As String)
        For Each l As String In Layers
            ShowLayer(l, True)
        Next
    End Sub
    Private Sub FreezeLayers(ByVal Layers() As String, ByVal booFreeze As Boolean)
        For Each l As String In Layers
            FreezeLayer(l, booFreeze)
        Next
    End Sub

    Private Function ExtractPER(ByVal objPLID As ObjectId) As Boolean
        Try
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                Dim objPL As Polyline = CType(trans.GetObject(objPLID, OpenMode.ForRead), Polyline)
                Dim strLayer As String = objPL.Layer
                'Get the vertices of the polyline and loop through them 
                Dim Verts As Point3dCollection = GetPolyVertices(objPLID)
                For i As Integer = 0 To Verts.Count - 1
                    Dim PERPL As New PER_PL
                    PERPL.E = Verts(i).X
                    PERPL.N = Verts(i).Y
                    PERPL.RayonArc = objPL.GetBulgeAt(i)
                    'Find the corresponding PL
                    PERPL.ObjID = GetPLatCoord(Verts(i), strLayer)
                    'If PL is found, extract the attributes
                    If Not PERPL.ObjID.IsNull Then
                        Dim blkRef As BlockReference = CType(trans.GetObject(PERPL.ObjID, OpenMode.ForRead), BlockReference)
                        'Get the attributes from the block
                        PERPL.Theme = blkRef.Layer.Replace("_PTS", "")
                        PERPL.Signe = GetBlockAttribute(PERPL.ObjID, "SIGNE")
                        PERPL.IdentDN = GetBlockAttribute(PERPL.ObjID, "IDENTDN")
                        PERPL.Numero = GetBlockAttribute(PERPL.ObjID, "NUMERO")
                        If PERPL.Signe = "" Then PERPL.Signe = "non materialise"
                        PERPL.PrecPlan = GetBlockAttribute(PERPL.ObjID, "PRECPLAN")
                        PERPL.FiabPlan = GetBlockAttribute(PERPL.ObjID, "FIABPLAN")
                        blkRef.Dispose()
                    End If
                    Debug.WriteLine(PERPL.Numero & "-" & i)
                    PERPL.Index = i
                    lstPL.Add(PERPL)
                    Wait.progBar1.Value = Convert.ToInt32((i / Verts.Count) * 80)
                    frms.Application.DoEvents()
                Next
                objPL.Dispose()
            End Using
            Return True
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "ExtractPER", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Return False
        End Try

    End Function

    Public Function GetPLatCoord(ByVal p As Point3d, ByVal strLayer As String) As ObjectId
        Dim objBlockID As ObjectId = SelectBlockAtPoint(p, "MO_BF_PTS")
        'Search on layers MO_PFP1/2/3
        If objBlockID.IsNull Then objBlockID = SelectBlockAtPoint(p, "MO_PFP3")
        If objBlockID.IsNull Then objBlockID = SelectBlockAtPoint(p, "MO_PFP2")
        If objBlockID.IsNull Then objBlockID = SelectBlockAtPoint(p, "MO_PFP1")
        'If still not found, search MUT category
        If objBlockID.IsNull Then objBlockID = SelectBlockAtPoint(p, "MUT_BF_PTS")
        If objBlockID.IsNull Then objBlockID = SelectBlockAtPoint(p, "MUT_BF_PFP3")
        If objBlockID.IsNull Then objBlockID = SelectBlockAtPoint(p, "MUT_BF_PFP2")
        If objBlockID.IsNull Then objBlockID = SelectBlockAtPoint(p, "MUT_BF_PFP1")
        If objBlockID.IsNull Then
            If strLayer.Contains("_CS_BAT") Then
                objBlockID = SelectBlockAtPoint(p, "MO_CS_PTS")
                If objBlockID.IsNull Then objBlockID = SelectBlockAtPoint(p, "MUT_CS_PTS")
            End If
        End If
        If objBlockID.IsNull Then
            Return Nothing
        Else
            Return objBlockID
        End If
    End Function


    Public Function SelectBlockAtPoint(ByVal p As Point3d, ByVal strLayer As String) As ObjectId
        Try
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayerName, strLayer)}
            Dim verts As New Point3dCollection
            verts.Add(New Point3d(p.X - 0.1, p.Y - 0.1, 0))
            verts.Add(New Point3d(p.X + 0.1, p.Y - 0.1, 0))
            verts.Add(New Point3d(p.X + 0.1, p.Y + 0.1, 0))
            verts.Add(New Point3d(p.X - 0.1, p.Y + 0.1, 0))
            Dim ids() As ObjectId = SelectCrossingPoly(Values, verts)
            If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                If ids.Count = 1 Then
                    SelectBlockAtPoint = ids(0)
                ElseIf ids.Count > 1 Then
                    Dim db As Database = HostApplicationServices.WorkingDatabase
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        Dim intResCount As Integer = 0
                        For Each id As ObjectId In ids
                            Dim blk As BlockReference = CType(trans.GetObject(id, OpenMode.ForRead), BlockReference)
                            If Math.Round(blk.Position.X, 3) = Math.Round(p.X, 3) And Math.Round(blk.Position.Y, 3) = Math.Round(p.Y, 3) Then
                                intResCount += 1
                                SelectBlockAtPoint = id
                            End If
                        Next
                        If intResCount > 1 Then
                            SelectBlockAtPoint = Nothing
                        End If
                    End Using
                Else
                    SelectBlockAtPoint = Nothing
                End If
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "SelectBlockAtPoint", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Return Nothing
        End Try
    End Function

    Private Function WritePERHeader(ByRef sWriter As StreamWriter, Optional ByVal strCanton As String = "Vaud") As Boolean
      
        sWriter.WriteLine("")
        sWriter.WriteLine("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">")
        sWriter.WriteLine("<html>")
        sWriter.WriteLine("")
        sWriter.WriteLine("<head>")
        sWriter.WriteLine("  <title>PER</title>")
        sWriter.WriteLine("  <meta http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"">")
        sWriter.WriteLine("  <style type=text/css>")
        sWriter.WriteLine("    body {")
        sWriter.WriteLine("    background-color: #FFFFFF;")
        sWriter.WriteLine("    font-family: Arial, Helvetica, sans-serif;")
        sWriter.WriteLine("    font-size: 9pt;")
        sWriter.WriteLine("    margin-left: 0mm;")
        sWriter.WriteLine("    margin-top: 0mm;")
        sWriter.WriteLine("    margin-right: 0mm;")
        sWriter.WriteLine("    margin-bottom: 0mm;")
        sWriter.WriteLine("}")
        sWriter.WriteLine("    table { border-collapse: collapse; }")
        sWriter.WriteLine("    td { font-family: Arial, Helvetica, sans-serif; font-size: 9pt; }")
        sWriter.WriteLine("    .Titre { font-size: 12pt; font-weight: bold; }")
        sWriter.WriteLine("  </style>")
        sWriter.WriteLine("</head>")
        sWriter.WriteLine("")
        sWriter.WriteLine("<body>")
        sWriter.WriteLine("  <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width: 170mm;"">")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
        sWriter.WriteLine("        <tr>")
        'Ignore call to CheckAppState as binary values for License are not in registry 
        'CheckAppState()
        'If booDemo = False Then
        ' Dim strLicence As String
        'strLicence = LireBaseRegistre(0, "SoftWare\SWWS\JourCAD", "LicenceName") & "<br>" & LireBaseRegistre(0, "SoftWare\SWWS\JourCAD", "LicenceSite")
        'sWriter.WriteLine("          <td width=""40%""><div align=""left"">" & strLicence & "</div></td>")
        sWriter.WriteLine("          <td width=""60%""><div align=""left"">REVO " & Ass.xVersion & "<br>&nbsp;</div></td>")
        'Else
        '    sWriter.WriteLine("          <td width=""60%""><div align=""left"">JourCAD " & Format(Ass.xVersion, "0.0") & "<br>Version d'&eacute;valuation</div></td>")
        'End If
        ''Date / time
        Dim strInfos As String
        If Trim(GetCurrentWinLogin) <> "" Then
            strInfos = "Etabli par " & GetCurrentWinLogin() & "<br>le " & Date.Today.ToLongDateString & "&nbsp;&agrave;&nbsp;" & TimeOfDay.ToShortTimeString
        Else
            strInfos = "Etabli le " & Date.Today.ToLongDateString & "&nbsp;&agrave;&nbsp;" & TimeOfDay.ToShortTimeString & "<br>&nbsp;"
        End If
        sWriter.WriteLine("          <td width=""40%""><div align=""right"">" & strInfos & "</div></td>")
        sWriter.WriteLine("        </tr>")
        sWriter.WriteLine("      </table></td>")
        sWriter.WriteLine("    </tr>")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td height=""30"">&nbsp;</td>")
        sWriter.WriteLine("    </tr>")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td><div align=""center""><span class=""Titre"">CONTOUR PERIMETRIQUE</span></div></td>")
        sWriter.WriteLine("    </tr>")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td>&nbsp;</td>")
        sWriter.WriteLine("    </tr>")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td height=""2px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000;""></td>")
        sWriter.WriteLine("    </tr>")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td height=""30"">&nbsp;</td>")
        sWriter.WriteLine("    </tr>")
    End Function

    Private Function WritePERSeparator(ByRef sWriter As StreamWriter) As Boolean
        sWriter.WriteLine("  <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width: 170mm;"">")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td height=""30"">&nbsp;</td>")
        sWriter.WriteLine("    </tr>")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td height=""2px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000;""></td>")
        sWriter.WriteLine("    </tr>")
        sWriter.WriteLine("    <tr>")
        sWriter.WriteLine("      <td height=""30"">&nbsp;</td>")
        sWriter.WriteLine("    </tr>")
        sWriter.WriteLine("  </table>")
    End Function

    Private Function WritePERFooter(ByRef sWriter As StreamWriter) As Boolean
        sWriter.WriteLine("</body>")
        sWriter.WriteLine("")
        sWriter.WriteLine("</html>")
    End Function
    ''' <summary>
    ''' Returns the total area of the polylines found within the supplied polyline
    ''' </summary>
    ''' <param name="objPolyID">The polyline to search within </param>
    ''' <param name="strTheme">The layer to filter on</param>
    ''' <returns>The total area of all of the polylines found within the supplied polyline</returns>
    ''' <remarks></remarks>
    Private Function SearchForHoles(ByVal objPolyID As ObjectId, ByVal strTheme As String) As Double
        Dim dblSurfTrous As Double = 0
        Dim objIlot As ObjectId = Nothing
        'Find all items within the supplied polyline that are on the desired layers
        'Get the vertices of the polyline
        Dim Verts As Point3dCollection = GetPolyVertices(objPolyID)
        'Set up the typedvalues for the selection filter
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, String.Format("MO_{0},MUT_{0}", strTheme))}
        'Make the selection (window polygon)
        Dim ids() As ObjectId = SelectWindowPoly(Values, Verts)
        'Check that we found something
        If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
            'Loop through the found items adding up their areas
            For Each objIlot In ids
                dblSurfTrous += GetPolyArea(objIlot)
            Next
        End If
        'Return the sum of the internal areas
        Return dblSurfTrous
    End Function





    Public Function PERstart()

        ' Sélectionner les numéros de parcelles
        Dim CollSelect As New Collection
        CollSelect = SelectObj(False, "Sélectionner les numéros de parcelles et bâtiments (cliquer sur Enter pour terminer)", False, False, True)

        'Filtre objets
        'BF + DDP :
   
        ' Calque : MO_BF_DDP_IMMEUBLE, MO_BF_IMMEUBLE
        '          MUT_BF_DDP_IMMEUBLE, MUT_BF_IMMEUBLE
        '          RAD_BF_DDP_IMMEUBLE, RAD_BF_IMMEUBLE



        Dim lst As New List(Of PerListItem)

        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction

            Try
                For Each BL As BlockReference In CollSelect
                   
                    Dim EDTItem As EDTListItem = GetInfoFromBlock(BL.ObjectId)
                    Dim PolyLayer As String = "MO_BF"
                    Dim TypeObj As String = "BF"

                    Dim StartLayer As String = "MO_" ' MUT_ RAD_
                    If BL.Layer.StartsWith("MUT_") Then
                        StartLayer = "MUT_"
                    ElseIf BL.Layer.StartsWith("RAD_") Then
                        StartLayer = "RAD_"
                    End If

                    If BL.Layer Like "*BF_*IMMEUBLE" Then 'BF
                        PolyLayer = Replace(BL.Layer, "_IMMEUBLE", "")
                        TypeObj = "BF"
                    ElseIf BL.Layer Like "*CS_*" Then 'CS
                        PolyLayer = Replace(BL.Layer, "_NUM", "")
                        TypeObj = "CS"
                    ElseIf BL.Layer Like "*ODS_BAT*_NUM" Then 'ODS
                        PolyLayer = Replace(BL.Layer, "_NUM", "")
                        PolyLayer = Replace(PolyLayer, "ODS_BATS", "ODS_BATSOUT") 'StartLayer & "ODS_BATSOUT"
                        TypeObj = "CS"
                    End If


                    Dim PolyId As ObjectId = PerimSearch(BL.Position, PolyLayer)
                    Dim ItemPer As New PerListItem

                    If PolyId <> Nothing Then

                        If TypeObj = "BF" Then
                            ItemPer.BF = Nothing
                            ItemPer.BlockID = BL.ObjectId
                            ItemPer.Com = EDTItem.Com
                            ItemPer.DisplayVal = PolyLayer & " " & EDTItem.Com & " / " & EDTItem.Num
                            ItemPer.ID = "BF" & EDTItem.Com & " / " &   EDTItem.Num
                            ItemPer.Num = EDTItem.Num
                            ItemPer.PolyID = PolyId
                            ItemPer.PolyLayer = PolyLayer

                        ElseIf TypeObj = "CS" Then
                            'Trouver le périmètre de la parcelle
                            Dim BFPolyId As ObjectId = PerimSearch(BL.Position, StartLayer & "BF")
                            'Find the number of the building block inside the polyline
                            Dim blkBFid As ObjectId = GetBATBlockByPolyline(BFPolyId, StartLayer & "BF_IMMEUBLE")
                            Dim blkBFinfo As EDTListItem = GetInfoFromBlock(blkBFid)

                            ItemPer.BF = blkBFinfo.Num
                            ItemPer.BlockID = BL.ObjectId
                            ItemPer.Com = blkBFinfo.ComNum
                            ItemPer.DisplayVal = PolyLayer & " " & blkBFinfo.Com & " / " & blkBFinfo.Num
                            ItemPer.ID = Nothing '"CS" & ComNum & " / " & BFnum
                            ItemPer.Num = EDTItem.Num
                            ItemPer.PolyID = PolyId
                            ItemPer.PolyLayer = PolyLayer
                        Else

                            'Trouver le périmètre de la parcelle
                            Dim BFPolyId As ObjectId = PerimSearch(BL.Position, StartLayer & "BF")
                            'Find the number of the building block inside the polyline
                            Dim blkBFid As ObjectId = GetBATBlockByPolyline(BFPolyId, StartLayer & "BF_IMMEUBLE")
                            Dim blkBFinfo As EDTListItem = GetInfoFromBlock(blkBFid)

                            ItemPer.BF = blkBFinfo.Num
                            ItemPer.BlockID = BL.ObjectId
                            ItemPer.Com = blkBFinfo.ComNum
                            ItemPer.DisplayVal = PolyLayer & " " & blkBFinfo.Com & " / " & blkBFinfo.Num
                            ItemPer.ID = Nothing '"ODS" & ComNum & " / " & BFnum
                            ItemPer.Num = EDTItem.Num
                            ItemPer.PolyID = PolyId
                            ItemPer.PolyLayer = PolyLayer
                        End If


                        lst.Add(ItemPer)

                    Else

                    End If


                    'If Replace(Replace(BL.Name, "MUT_", ""), "RAD_", "") Like "IMMEUBLE_NUM" Then 'Pas de prise en compte des DDP
                    '    Dim BLn As New BLnum
                    '    BLn.ID = BL.ObjectId
                    '    BLn.Name = BL.Name
                    '    BLn.Pt = BL.Position
                    '    CollNumParcelle.Add(BLn)
                    'End If
                Next
            Catch
            End Try

            If lst.Count = 0 Then
                MsgBox("Aucun numéro n'as été séléctionné", vbInformation + vbOKOnly, "Pas d'objet séléctionné")
                Return False
            End If

        End Using


        ' Sélectionner le dossier de destination
        Dim RVinfo As New Revo.RevoInfo
        Dim strFolder As String = RVinfo.PERFolder

        If System.IO.Directory.Exists(strFolder) = False Then
            Dim fb As New Windows.Forms.FolderBrowserDialog
            fb.Description = "Sélectionner le dossier dans lequel enregistrer le fichier PER résultat"
            If fb.ShowDialog = Windows.Forms.DialogResult.OK Then
                strFolder = fb.SelectedPath
            Else
                strFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            End If

            If Not strFolder.EndsWith("\") Then strFolder = strFolder & "\"
        End If


        ' Calcul du PER + Mise en forme PER

        If lst.Count > 0 Then
            'run the process on the list
            CalculPER(lst)
            Return True
        Else
            Return False
        End If

    End Function


    Private Function PerimSearch(Pt3D As Point3d, LayerName As String) As ObjectId

        Dim plID As ObjectId = GetLWPByCentroidAndRect(Pt3D, 1000, LayerName, True)

        'Check whether we found the polyline
        If plID.IsNull Then
            'No polyline found so show message giving the user the option to re-select
            If frms.MessageBox.Show("Erreur de la recherche", "Impossible de trouver un périmètre de BF" & vbCrLf & _
                                    "Merci de signaler le problème avec un jeu de donnée", Windows.Forms.MessageBoxButtons.OK, _
                Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.OK Then
            End If

            Return Nothing

        Else
            Return plID
        End If

    End Function
End Module
