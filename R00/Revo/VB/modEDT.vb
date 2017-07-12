Imports frms = System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors



Module modEDT
    Public EDTFolder As String = ""
    Public EDTCheckSurf As String = "False"
    Public EDTIgnoreRPL As String = "False"
    Private HTMLcontent As String = ""

    Dim booDemo As Boolean = False
    Dim Ass As New Revo.RevoInfo
    Private Class BFbyRPL
        Private _BFID As ObjectId 'Polyline
        Private _RPLID As ObjectId  'Block
        ''' <summary>
        ''' The objectid of the BF polyline
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property BF() As ObjectId
            Get
                Return _BFID
            End Get
            Set(ByVal value As ObjectId)
                _BFID = value
            End Set
        End Property
        ''' <summary>
        ''' The objectid of the RPL block
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RPL() As ObjectId
            Get
                Return _RPLID
            End Get
            Set(ByVal value As ObjectId)
                _RPLID = value
            End Set
        End Property
    End Class
    Private Class DDP
        Private _objPolyID As ObjectId
        Private _Surface As Double
        Private _Partiel As Boolean
        Private _Numero As String
        Public Property objPolyID() As ObjectId
            Get
                Return _objPolyID
            End Get
            Set(ByVal value As ObjectId)
                _objPolyID = value
            End Set
        End Property
        Public Property Surface() As Double
            Get
                Return _Surface
            End Get
            Set(ByVal value As Double)
                _Surface = value
            End Set
        End Property
        Public Property Partiel() As Boolean
            Get
                Return _Partiel
            End Get
            Set(ByVal value As Boolean)
                _Partiel = value
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
    End Class
    Private Class CS
        Private _Type As String
        Private _Surface As Double
        Private _RegCS As ObjectId
        Private _Description As String
        Private _ID As String
        Private _Divers As String
        Private _BF_Ilot As Boolean
        Public Property Type() As String
            Get
                Return _Type
            End Get
            Set(ByVal value As String)
                _Type = value
            End Set
        End Property
        Public Property Surface() As Double
            Get
                Return _Surface
            End Get
            Set(ByVal value As Double)
                _Surface = value
            End Set
        End Property
        Public Property RegCS() As ObjectId
            Get
                Return _RegCS
            End Get
            Set(ByVal value As ObjectId)
                _RegCS = value
            End Set
        End Property
        Public Property Description() As String
            Get
                Return _Description
            End Get
            Set(ByVal value As String)
                _Description = value
            End Set
        End Property
        Public Property ID() As String
            Get
                Return _ID
            End Get
            Set(ByVal value As String)
                _ID = value
            End Set
        End Property
        Public Property Divers() As String
            Get
                Return _Divers
            End Get
            Set(ByVal value As String)
                _Divers = value
            End Set
        End Property
        Public Property BF_Ilot() As Boolean
            Get
                Return _BF_Ilot
            End Get
            Set(ByVal value As Boolean)
                _BF_Ilot = value
            End Set
        End Property
    End Class
    Private Class ODS
        Private _Description As String
        Private _ID As String
        Private _Area As Double
        Private _Divers As String
        Private _Type As String
        Private _Surface As Double
        Private _ODSRegID As ObjectId
        Public Property Description() As String
            Get
                Return _Description
            End Get
            Set(ByVal value As String)
                _Description = value
            End Set
        End Property
        Public Property ID() As String
            Get
                Return _ID
            End Get
            Set(ByVal value As String)
                _ID = value
            End Set
        End Property
        Public Property Area() As Double
            Get
                Return _Area
            End Get
            Set(ByVal value As Double)
                _Area = value
            End Set
        End Property
        Public Property Divers() As String
            Get
                Return _Divers
            End Get
            Set(ByVal value As String)
                _Divers = value
            End Set
        End Property
        Public Property Type() As String
            Get
                Return _Type
            End Get
            Set(ByVal value As String)
                _Type = value
            End Set
        End Property
        Public Property Surface() As Double
            Get
                Return _Surface
            End Get
            Set(ByVal value As Double)
                _Surface = value
            End Set
        End Property
        Public Property ODSRegID() As ObjectId
            Get
                Return _ODSRegID
            End Get
            Set(ByVal value As ObjectId)
                _ODSRegID = value
            End Set
        End Property
    End Class

    Private lstBFbyRPL As List(Of BFbyRPL)
    Private lstDDP As List(Of DDP)
    Private lstCS As List(Of CS)
    Private lstODS As List(Of ODS)
    Private Wait As frmWait = Nothing
    ''' <summary>
    ''' Performs the EDT calculation
    ''' </summary>
    ''' <param name="lst">The list of items to process from the EDT form</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Public Function CalculEDT(ByVal lst As List(Of EDTListItem), ByVal IgnoreRPL As Boolean, ByVal ControleRF As Boolean) As Boolean

        Dim RVinfo As New Revo.RevoInfo
        System.IO.File.Copy(System.IO.Path.Combine(RVinfo.SystemPath, RVinfo.DB3systemOrig), System.IO.Path.Combine(RVinfo.SystemPath, RVinfo.DB3system), True)

        'Clear the EDT_donnees table in the database
        ClearTable("EDT_donnees")
        'Store the number of items to process
        Dim intObjCount As Integer = lst.Count
        'Create the progress form
        Wait = New frmWait
        'Set its labels
        SetUpWaitForm(Wait, intObjCount)
        'Show the progress form
        Wait.Show()
        frms.Application.DoEvents()
        'Zoom to the extents of the drawing
        Zooming.ZoomExtents()
        'Save the current layer state
        SaveLayerState("REVO_Temp")
        'Thaw all layers
        FreezeAllLayers(False)
        'And switch them all on
        ShowAllLayers(True)
        'Regen (required after thawing layers)
        Application.DocumentManager.MdiActiveDocument.Editor.Regen()
        'Set the active layer
        ActivateLayer("MO_BF_INCOMPLET")
        'Call the function to process the items in the list
        ProcessListItems(lst, IgnoreRPL, ControleRF)
        Wait.lblStatus1.Text = "Régénération du dessin..."
        Wait.progBar1.Visible = False
        Wait.lblStatus2.Text = ""
        Wait.progBar2.Visible = False
        frms.Application.DoEvents()
        RestoreLayerState("REVO_Temp")
        RemoveLayerState("REVO_Temp")
        Zooming.ZoomPrevious()
        Wait.Close()
    End Function

    Private Sub ProcessListItems(ByVal lst As List(Of EDTListItem), ByVal IgnoreRPL As Boolean, ByVal ControleRF As Boolean)
        Dim ItemCounter As Integer = 0
        Dim dblSurfTech As Double = 0
        For Each edt As EDTListItem In lst
            ItemCounter += 1
            Dim polyID As ObjectId = edt.PolyID
            Dim blkId As ObjectId = edt.BlockID
            Dim strID As String = edt.ComNum
            Dim strCom As String = edt.Com
            Dim strNum As String = edt.Num
            Wait.lblStatus1.Text = "Analyse de la répartition des plans..."
            Wait.progBar1.Value = 0
            frms.Application.DoEvents()
            'Reset the list to hold the objects
            lstBFbyRPL = New List(Of BFbyRPL)
            IntersectBFwithRPL(polyID, IgnoreRPL)
            Wait.lblStatus1.Text = "Recherche des DDP..."
            Wait.progBar1.Value = 5
            frms.Application.DoEvents()
            GetDDPInsideBF(polyID)
            Dim Counter As Integer = -1

            For Each oBFRPL In lstBFbyRPL
                lstCS = New List(Of CS)
                lstODS = New List(Of ODS)
                Dim intPlan As Integer = 0
                Dim strCodePlan As String = ""
                Counter += 1
                If Not oBFRPL.RPL.IsNull Then
                    Try
                        intPlan = Convert.ToInt32(GetBlockAttribute(oBFRPL.RPL, "NUMERO"))
                    Catch ex As Exception
                        'returned value is not an integer
                    End Try
                    If intPlan = 0 Then
                        intPlan = 9000 + Counter
                    Else
                        strCodePlan = GetBlockAttribute(oBFRPL.RPL, "CODEPLAN")
                    End If
                End If
                Wait.lblStatus1.Text = "Recherche et analyse des îlots (BF)"
                Wait.progBar1.Value += Convert.ToInt32(Math.Abs((5 / lstBFbyRPL.Count)))
                frms.Application.DoEvents()
                'Search for BF islands storing the sum of the areas for later use
                dblSurfTech = (SearchForHolesInBF(oBFRPL.BF, strCom, intPlan, strNum, strID, strCodePlan) * -1)
                Wait.lblStatus1.Text = "Calcul des intersections spatiales (BF / CS)"
                Wait.progBar1.Value = 10
                frms.Application.DoEvents()
                Dim strCurrentLayer As String = ""
                If IntersectBFwithCS(oBFRPL.BF, , lstBFbyRPL.Count, strCurrentLayer) Then
                    Dim dt As System.Data.DataTable = GetTable("EDT_Donnees", dbName:=Ass.DB3system)
                    WriteCSObjectsToTable(dt, strCom, intPlan, strNum, strID, strCodePlan)
                    Wait.lblStatus1.Text = "Calcul des intersections spatiales (BF / ODS)"
                    frms.Application.DoEvents()
                    IntersectBFwithODS(oBFRPL.BF, , lstBFbyRPL.Count)
                    WriteODSObjectsToTable(dt, strCom, intPlan, strNum, strID, strCodePlan)
                    Wait.progBar1.Value = Convert.ToInt32(Math.Abs(100 * (Counter + 1) / lstBFbyRPL.Count))
                    frms.Application.DoEvents()
                Else
                    Wait.Hide()
                    frms.MessageBox.Show("L'état descriptif n'a pas pu être calculé !" & vbCrLf & _
                                         "Cela est peut être dû à un problème de topologie ou à une erreur interne à AutoCAD." _
                                         & vbCrLf & "Contrôler la géométrie et la topologie des couches CS (" & strCurrentLayer & _
                                         ")." & vbCrLf & vbCrLf & _
                                         "L'objet posant problème a été sélectionné: vérifiez qu'il s'agit d'une polyligne fermée et" _
                                         & vbCrLf & "qu'il se trouve dans le bon calque.", "Calcul impossible", _
                                          Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                    Exit Sub
                End If
            Next


            Wait.lblStatus1.Text = "Création du rapport..."
            Wait.lblStatus2.Text = ItemCounter & " / " & lst.Count & " EDT calculé" & (IIf(ItemCounter = 1, "s", "")).ToString & " (" & Convert.ToInt32((ItemCounter / lst.Count) * 100) & "%)"
            If ControleRF Then
                Dim dblSurf As Double = 0
                Double.TryParse((GetBlockAttribute(edt.BlockID, "SUPERFICIE_TOTALE")), dblSurf)
                dblSurfTech += GetPolyArea(polyID)
                If dblSurf = 0 Then
                    Wait.Hide()
                    frms.MessageBox.Show("La surface calculée pour le bien-fonds " & strID & _
                                         " n'a pas pu être comparée à la surface RF," & vbCrLf & _
                                         "car celle-ci n'a pas été importée/définie dans le fichier dessin !", _
                                         "Aucun contrôle", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                    Wait.Show()
                Else
                    If Math.Abs(dblSurf - dblSurfTech) > 1 Then
                        Wait.Hide()
                        frms.MessageBox.Show("La surface technique calculée pour le bien-fonds " & strID & _
                                             " ne correspond pas" & vbCrLf & _
                                             "à la surface RF importée/définie dans le fichier dessin !" & vbCrLf & vbCrLf _
                                             & "Contrôler que la topologie est correcte et que la valeur SUPERFICIE_TOTALE" _
                                             & vbCrLf & "du bloc IMMEUBLE est correcte.", "Aucun contrôle", _
                                              Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                        Wait.Show()
                    End If
                End If
            End If
            CreateEDTHtml(strID, True)
        Next


    End Sub

    ''' <summary>
    ''' Initialises the progress form setting the label contents
    ''' </summary>
    ''' <param name="frm">The progress form</param>
    ''' <param name="intObjCount">The number of items to process</param>
    ''' <remarks></remarks>
    Private Sub SetUpWaitForm(ByRef frm As frmWait, ByVal intObjCount As Integer)
        With frm
            .lblTitle.Text = "Génération d'état descriptif"
            .lblSubTitle.Text = "Patienter pendant l'analyse des objets"
            .lblInfo1.Text = "Analyse de la couverture du sol et calcul des intersections"
            .lblInfo1.Visible = True
            .lblStatus1.Text = "Initialisation..."
            .progBar1.Visible = True
            .progBar1.Value = 0
            .progBar1.Minimum = 0
            .progBar1.Maximum = 100
            .lblStatus2.Text = "0 / " & intObjCount & " EDT calculé (0%)"
            .progBar2.Visible = True
            .progBar2.Value = 0
            .progBar2.Minimum = 0
            .progBar2.Maximum = 100
        End With
    End Sub


    Private Function IntersectBFwithRPL(ByVal polyid As ObjectId, ByVal IgnoreRPL As Boolean) As Boolean
        Try
            Dim Verts As Point3dCollection = GetPolyVertices(polyid)
            'Set up the typedvalues to filter the selection.
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_RPL")}
            Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)
            If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                'Get the database and start a transaction
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Loop through the found objects
                    For Each id As ObjectId In ids
                        'Convert this entity to a region
                        Dim objRPLreg As Region = ConvertPolylineToRegion(id)
                        'Convert the supplied polyline to a region
                        Dim objBFReg As Region = ConvertPolylineToRegion(polyid)
                        objBFReg = CType(trans.GetObject(objBFReg.ObjectId, OpenMode.ForWrite), Region)
                        'Subtract this region from the BF region 
                        objBFReg.BooleanOperation(BooleanOperationType.BoolIntersect, objRPLreg)
                        'If the area of the resulting region is greater than 0
                        If objBFReg.Area > 0.01 Then
                            'Create a new BFbyRPL object
                            Dim BFbyRPL As New BFbyRPL
                            'Set the BF property to the polyline from the region
                            BFbyRPL.BF = ConvertRegionToPolyline(objBFReg, False)
                            'If the user has not ticked the Ignore RPL on the EDT form 
                            If Not IgnoreRPL Then
                                'Call the function to get the RPL from the BF
                                BFbyRPL.RPL = GetRPLFromBF(BFbyRPL.BF)
                            Else
                                'Otherwise set the RPL property to nothing
                                BFbyRPL.RPL = Nothing
                            End If
                            'Add the object to the global list
                            lstBFbyRPL.Add(BFbyRPL)
                        End If
                    Next
                    trans.Commit()
                End Using
            Else
                'Nothing found within then polyline
                'So create a new BFbyRPL object
                Dim BFbyRPL As New BFbyRPL
                'Set the BF property to the id of the BF polyline
                BFbyRPL.BF = polyid
                'If the user has not ticked the Ignore RPL on the EDT form 
                If Not IgnoreRPL Then
                    'Call the function to get the RPL from the BF
                    BFbyRPL.RPL = GetRPLFromBF(polyid)
                Else
                    'Otherwise set the RPL property to nothing
                    BFbyRPL.RPL = Nothing
                End If
                'Add the object to the global list
                lstBFbyRPL.Add(BFbyRPL)
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetRPLFromBF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    ''' <summary>
    ''' Gets the RPL from the BF
    ''' </summary>
    ''' <param name="polyID">The objectid of the BF polyline</param>
    ''' <returns>The objectid of the RPL block</returns>
    ''' <remarks></remarks>
    Private Function GetRPLFromBF(ByVal polyID As ObjectId) As ObjectId
        Dim objBlockPlan As ObjectId = Nothing
        Try
            'get the vertices of the polyline
            Dim Verts As Point3dCollection = GetPolyVertices(polyID)
            'Set up typedvalues to select polylines on the desired layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_RPL")}
            Dim IDs() As ObjectId = SelectCrossingPoly(Values, Verts)
            'Convert the BF polyline to a region
            Dim objRegBF As Region = ConvertPolylineToRegion(polyID)
            Dim objRegPL As Region = Nothing
            If Not IDs Is Nothing AndAlso Not IDs.Length = 0 Then
                'Loop through the found polylines converting them to regions and subtracting the BF polyline from them
                For Each id As ObjectId In IDs
                    objRegPL = ConvertPolylineToRegion(id)
                    'Store the area of the BF poly
                    Dim dblSurfInit As Double = objRegPL.Area
                    'Subtract the BF from this region
                    objRegPL.BooleanOperation(BooleanOperationType.BoolSubtract, CType(objRegBF.Clone, Region))
                    'Check the size of the resulting region
                    If Math.Abs(dblSurfInit - objRegPL.Area) > 0.01 Then
                        'Get the block from this polyline
                        objBlockPlan = GetBlockfromLWP1(id, "MO_RPL_PLAN", True)
                    End If
                Next
                If Not objRegPL Is Nothing Then objRegPL.Dispose()
                objRegBF.Dispose()
            Else
                'Select everything that passes the filter
                Dim pres As PromptSelectionResult = SelectAllItems(Values)
                'Check that we found something
                If pres.Status = PromptStatus.OK Then
                    'Get the database and start a transaction
                    Dim db As Database = HostApplicationServices.WorkingDatabase
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        'Loop through the found objects
                        For Each id As ObjectId In pres.Value.GetObjectIds
                            'Get the object from the database
                            Dim pline As Polyline = CType(trans.GetObject(polyID, OpenMode.ForRead), Polyline) 'Modif TH 27.02.2016
                            'Dim pline As Polyline = CType(trans.GetObject(id, OpenMode.ForRead), Polyline) 

                            'Check whether this object is within the BF polyline
                            If IsObjectInsidePolyline(pline, id) Then 'Modif TH 27.02.2016
                                'If IsObjectInsidePolyline(pline, polyID) Then
                                'Get the block from the polyline
                                objBlockPlan = GetBlockfromLWP1(id, "MO_RPL_PLAN", True)
                                pline.Dispose()
                                Exit For
                            End If
                        Next
                    End Using
                Else
                    objBlockPlan = Nothing
                End If
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetRPLFromBF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Return objBlockPlan
    End Function

    ''' <summary>
    ''' Selects the block from the polyline
    ''' </summary>
    ''' <param name="polyID">The objectid of the polyline</param>
    ''' <param name="strLayer">The layer that the block should be on</param>
    ''' <param name="booWaitFormLoaded">Optional parameter to say whether the progress form is loaded</param>
    ''' <returns>The objectid of the found block</returns>
    ''' <remarks></remarks>
    Public Function GetBlockfromLWP1(ByVal polyID As ObjectId, ByVal strLayer As String, Optional ByVal booWaitFormLoaded As Boolean = False) As ObjectId
        Dim retVal As ObjectId = Nothing 'Return value
        Try
            'Get the vertices of the polyline
            Dim Verts As Point3dCollection = GetPolyVertices(polyID)
            'Set up typedvalues to select blocks on the correct layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayerName, strLayer)}
            Dim IDs() As ObjectId = SelectCrossingPoly(Values, Verts)
            If Not IDs Is Nothing AndAlso IDs.Length = 1 Then
                'Return the objectid
                Return IDs(0)
            Else
                'Found no objects or more than one
                'So get the database
                Dim db As Database = HostApplicationServices.WorkingDatabase
                'Start a transaction
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Get the polyline from the database
                    Dim pline As Polyline = CType(trans.GetObject(polyID, OpenMode.ForRead, True), Polyline)
                    'Zoom to the extents of the polyline
                    Zooming.ZoomToObject(polyID)
                    'And highlight the polyline
                    pline.Highlight()
                    'If the wait form (progress form) is loaded then hide it
                    If booWaitFormLoaded Then Wait.Hide() : frms.Application.DoEvents()
                    'Call the function to ask the user to select the block manually
                    retVal = GetBlockBySelection(strLayer)
                    'If the wait form is loaded then show it again
                    If booWaitFormLoaded Then Wait.Show()
                    'unhighlight the polyline
                    pline.Unhighlight()
                    'If a block has been selected manually then we assume that the zoom has been modified and therefore need to zoom again
                    If Not retVal.IsNull Then
                        Zooming.ZoomExtents()
                    Else
                        Zooming.ZoomPrevious() 'otherwise only one previous zoom
                    End If
                End Using
            End If
            frms.Application.DoEvents()
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetBlockfromLWP", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Return retVal 'Return the selected id
    End Function

    Private Function GetDDPInsideBF(ByVal polyID As ObjectId) As Boolean
        lstDDP = New List(Of DDP)
        Dim Verts As Point3dCollection = GetPolyVertices(polyID)
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_BF_DDP,MUT_BF_DDP")}
        Dim Ids() As ObjectId = SelectCrossingPoly(Values, Verts)
        If Not Ids Is Nothing AndAlso Ids.Length > 0 Then
            'Get the database and start a transaction
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Loop through the found objects
                For Each id As ObjectId In Ids
                    Dim pLine As Polyline = CType(trans.GetObject(id, OpenMode.ForRead), Polyline)
                    'Convert this entity to a region
                    Dim objDDPreg As Region = ConvertPolylineToRegion(id)
                    'Convert the supplied polyline to a region
                    Dim objBFReg As Region = ConvertPolylineToRegion(polyID)
                    'Subtract this region from the BF region 
                    objDDPreg.BooleanOperation(BooleanOperationType.BoolIntersect, objBFReg)
                    'If the area of the resulting region is greater than 0
                    If objDDPreg.Area > 0.01 Then
                        'Create a new DDP object
                        Dim objDDP As New DDP
                        'Set the PolyID property 
                        objDDP.objPolyID = id
                        'And the area property 
                        objDDP.Surface = objDDPreg.Area
                        'Set the Partiel flag
                        objDDP.Partiel = (Math.Abs(objDDPreg.Area - pLine.Area) > 0.01)
                        'try to get the block from the polyline
                        Dim ddpBlockID As ObjectId = GetBlockfromLWP1(id, pLine.Layer & "_IMMEUBLE", True)
                        'If we found a block 
                        If Not ddpBlockID.IsNull Then
                            'then get the value of the NUMERO attribute and store it
                            objDDP.Numero = GetBlockAttribute(ddpBlockID, "NUMERO")
                        Else
                            'Otherwise set set the Numero property to a non breaking space
                            objDDP.Numero = "&nbsp;"
                        End If
                    End If
                Next
            End Using
        End If
    End Function

    ''' <summary>
    ''' Find holes within the BF
    ''' </summary>
    ''' <param name="BFID"></param>
    ''' <param name="strCom"></param>
    ''' <param name="intPlan"></param>
    ''' <param name="strNum"></param>
    ''' <param name="strID"></param>
    ''' <param name="strCodePlan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SearchForHolesInBF(ByVal BFID As ObjectId, ByVal strCom As String, ByVal intPlan As Integer, ByVal strNum As String, _
                                        ByVal strID As String, ByVal strCodePlan As String) As Double
        Dim dblSurfRFTotale As Double = 0
        'Get the vertices of the polyline
        Dim Verts As Point3dCollection = GetPolyVertices(BFID)
        'Set up typedvalues to select polylines on the required layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_BF,MUT_BF")}
        Dim IDs() As ObjectId = SelectWindowPoly(Values, Verts)
        If Not IDs Is Nothing AndAlso Not IDs.Length = 0 Then
            'Store how many objects we found
            Dim intHolesCount As Integer = IDs.Length
            'Open the EDT_Donnees table
            Dim dt As System.Data.DataTable = GetTable("EDT_Donnees", dbName:=Ass.DB3system)
            'Loop through the found objects (holes)
            For Each IlotBF As ObjectId In IDs
                IntersectBFwithCS(IlotBF, True)
                'Store the CS objects in the database
                WriteCSObjectsToTable(dt, strCom, intPlan, strNum, strID, strCodePlan)
                lstCS.Clear()
                dblSurfRFTotale += GetPolyArea(IlotBF)
                'Calculation of the ODS surfaces BF island
                IntersectBFwithODS(IlotBF, True)
                'Store the ODS objects in the database
                WriteODSObjectsToTable(dt, strCom, intPlan, strNum, strID, strCodePlan)
                lstODS.Clear()
            Next
        End If
        Return dblSurfRFTotale
    End Function

    Private Function WriteCSObjectsToTable(ByRef dt As System.Data.DataTable, ByVal strCom As String, ByVal intPlan As Integer, _
                                           ByVal strNum As String, ByVal strId As String, ByVal strCodePlan As String) As Boolean
        For Each oCS As CS In lstCS
            Dim dr As DataRow = dt.NewRow
            dr("Theme") = 2
            dr("NoCom") = strCom
            dr("NoPlan") = intPlan
            dr("NoBF") = strNum.Replace("DP", "").Trim
            dr("DP") = strNum.Contains("DP")
            dr("Immeuble") = strId
            dr("CodePlan") = strCodePlan
            dr("JC_Layer") = oCS.Type
            If oCS.Type = "MO_CS_BAT" Or oCS.Type = "MUT_CS_BAT" Then
                dr("ID_objet") = oCS.ID
                dr("Designation_objet") = oCS.Description.Replace(" ", "_")
                dr("Info_objet") = oCS.Divers
            End If
            dr("Surface") = oCS.Surface
            dt.Rows.Add(dr)
        Next
        UpdateTable(dt, Ass.DB3system)
    End Function

    Private Function WriteODSObjectsToTable(ByRef dt As System.Data.DataTable, ByVal strCom As String, ByVal intPlan As Integer, _
                                       ByVal strNum As String, ByVal strId As String, ByVal strCodePlan As String) As Boolean
        For Each oODS As ODS In lstODS
            Dim dr As DataRow = dt.NewRow
            dr("Theme") = 3
            dr("NoCom") = strCom
            dr("NoPlan") = intPlan
            dr("NoBF") = strNum.Replace("DP", "").Trim
            dr("DP") = strNum.Contains("DP")
            dr("Immeuble") = strId
            dr("CodePlan") = strCodePlan
            dr("JC_Layer") = oODS.Type
            dr("ID_objet") = oODS.ID
            dr("Designation_objet") = oODS.Description.Replace(" ", "_")
            dr("Info_objet") = oODS.Divers
            dr("Surface") = oODS.Surface
            dt.Rows.Add(dr)
        Next
        UpdateTable(dt, Ass.DB3system)
    End Function

    Private Function IntersectBFwithCS(ByVal BF As ObjectId, Optional ByVal booBFhole As Boolean = False, _
                                       Optional ByVal intProgressDivider As Integer = 1, _
                                       Optional ByRef strCurrentLayer As String = "") As Boolean
        Dim strFmtSurf As String = "#,##0.0000"
        Dim ErrID As ObjectId = Nothing
        Dim retVal As Boolean = False
        Try
            'Get the vertices of the BF polyline
            Dim Verts As Point3dCollection = GetPolyVertices(BF)
            'Convert it to a region
            Dim objRegBF As Region = ConvertPolylineToRegion(BF)
            Dim objRegCS As Region = Nothing
            Dim objContPLID As ObjectId = Nothing
            Dim dblSurfTech As Double = 0
            'Set up typedvalues to select POLYLINES
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE")}
            'Make the selection
            Dim IDs() As ObjectId = SelectCrossingPoly(Values, Verts)
            'Check that we found something
            If Not IDs Is Nothing AndAlso Not IDs.Count = 0 Then
                Dim CommitTrans As Boolean = False
                Dim db As Database = HostApplicationServices.WorkingDatabase
                'loop through the found objects inside the BF
                For Each id As ObjectId In IDs
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        'Get the BF polyline
                        Dim BFPl As Polyline = CType(trans.GetObject(BF, OpenMode.ForRead), Polyline)
                        'get the polyline for this object
                        Dim pline As Polyline = CType(trans.GetObject(id, OpenMode.ForRead), Polyline)
                        'Get its layer
                        Dim Lay As String = pline.Layer
                        'Check the layer
                        If (Lay.StartsWith("MO_CS_") Or Lay.StartsWith("MUT_CS_")) _
                                    And Not Lay.Contains("_CS_PTS") And Not Lay.Contains("HACH") Then
                            'Store the layer for use if there is an error
                            strCurrentLayer = Lay
                            'Convert it to a region
                            objRegCS = ConvertPolylineToRegion(id)
                            'Copy the BF region as we don't want to modify the original
                            Dim objRegBFCopy As Region = CType(objRegBF.Clone, Region)
                            'Check that we converted the CS and BF region successfully
                            If objRegCS Is Nothing Then
                                ErrID = id
                            ElseIf objRegBFCopy Is Nothing Then
                                ErrID = objRegBF.ObjectId
                            Else
                                ErrID = Nothing
                            End If
                            Try
                                'Intersect the BF with the CS
                                objRegCS.BooleanOperation(BooleanOperationType.BoolIntersect, objRegBFCopy)
                            Catch ex As Exception
                                If ErrID.IsNull Then
                                    pline.Highlight()
                                    objRegBF.Highlight()
                                Else
                                    Zooming.ZoomToObject(ErrID)
                                    Dim ent As Entity = CType(trans.GetObject(ErrID, OpenMode.ForRead), Entity)
                                    ent.Highlight()
                                End If
                                DeleteGreenTempObjects()
                                Return False
                            End Try
                            'If the CS object intersects the BF
                            If objRegCS.Area > 0 And Not Lay.EndsWith("CS_INCOMPLET") Then
                                'get the CS region from the database
                                objRegCS = CType(trans.GetObject(objRegCS.ObjectId, OpenMode.ForWrite), Region)
                                'And change its colour to green
                                objRegCS.ColorIndex = 3
                                objRegCS.DowngradeOpen()
                                'Create a new CS object and set its properties
                                Dim oCS As New CS
                                oCS.Type = Lay
                                oCS.Surface = objRegCS.Area
                                oCS.RegCS = objRegCS.ObjectId
                                'If Buiding
                                If Lay.StartsWith("MO_CS_BAT") Or Lay.StartsWith("MUT_CS_BAT") Then
                                    'Get the XData from the polyline
                                    Dim tv() As TypedValue = pline.XData.AsArray
                                    Try
                                        'Store the desctiption and ID from the XData
                                        oCS.Description = tv(3).Value.ToString
                                        oCS.ID = tv(2).Value.ToString
                                    Catch ex As Exception
                                        If tv.Length <= 3 Then ' MODIFICATION 08.05.2014 THA : si erreur > modif 3 en 2 
                                            oCS.Description = tv(2).Value.ToString
                                            oCS.ID = tv(1).Value.ToString
                                        Else
                                            oCS.Description = "?"
                                            oCS.ID = "-1"
                                        End If
                                    End Try
                                    If pline.Area - objRegCS.Area > 0.01 Then
                                        oCS.Divers = Format(pline.Area, strFmtSurf)
                                    End If
                                End If
                                oCS.BF_Ilot = booBFhole
                                lstCS.Add(oCS)
                            Else

                            End If
                        End If
                        If Not booBFhole Then
                            Wait.progBar1.Value += Convert.ToInt32(Math.Abs((30 / IDs.Count) / intProgressDivider))
                            frms.Application.DoEvents()
                        End If
                        pline.Dispose()
                        trans.Commit()
                    End Using
                Next
                If Not booBFhole Then
                    Wait.lblStatus1.Text = "Recherche et analyse des îlots (CS)..."
                    frms.Application.DoEvents()
                End If
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Dim BFPl As Polyline = CType(trans.GetObject(BF, OpenMode.ForRead), Polyline)
                    If SearchForHolesInCS(intProgressDivider) Then
                        'Remove temporary objects
                        DeleteGreenTempObjects()
                        If Not booBFhole Then
                            Wait.lblStatus1.Text = "Contrôle des surfaces..."
                            frms.Application.DoEvents()
                        End If
                        'Calculation of Total / Control
                        dblSurfTech = 0
                        For Each CS In lstCS
                            dblSurfTech += CS.Surface
                        Next
                        'Entire surface of the object (without holes)
                        Dim dblSurfTot As Double = objRegBF.Area
                        Dim dblDelta As Double = dblSurfTech - dblSurfTot
                        If Math.Abs(dblDelta) >= 0.5 Then
                            If dblSurfTech < dblSurfTot Then
                                'Try to find the CS polyline that contains this BF by scaling up the BF polyline and selecting within it
                                'We loop around trying to find the polyline each time scaling the BF polyline by larger and larger amounts
                                'until we find it
                                For j As Integer = 10 To 100 Step 10
                                    objContPLID = GetCSthatContainsBF(BFPl, j)
                                    If Not objContPLID.IsNull Then Exit For
                                Next
                                If Not objContPLID.IsNull Then
                                    Dim objContPl As Polyline = CType(trans.GetObject(objContPLID, OpenMode.ForRead), Polyline)
                                    Dim oCS As New CS
                                    oCS.Type = objContPl.Layer
                                    oCS.Surface = BFPl.Area - dblSurfTech
                                    lstCS.Add(oCS)
                                    'Recalculation of the total
                                    dblSurfTech = 0
                                    For Each CS In lstCS
                                        dblSurfTech += CS.Surface
                                    Next
                                    objContPl.Dispose()
                                    retVal = True
                                Else
                                    retVal = False
                                End If
                            Else
                                Wait.Hide()
                                frms.MessageBox.Show("Erreur à l'établissement de l'état descriptif:" & _
                                                     vbCrLf & "La surface totale de la parcelle selon la limite de bien-fonds est de " _
                                                     & Format(dblSurfTot, "#,##0.00") & "m²" & vbCrLf & _
                                                     "alors que la somme des types de couverture du sol fait " _
                                                     & Format(dblSurfTech, "#,##0.00") & "m² !" & vbCrLf & _
                                                     "Vérifier la topologie des éléments 'CS' et qu'il n'y a pas de doublons ou de superpositions.", _
                                                     "Calcul impossible", Windows.Forms.MessageBoxButtons.OK, _
                                                     Windows.Forms.MessageBoxIcon.Error)
                                retVal = True
                            End If
                        Else
                            retVal = True
                        End If
                    Else

                    End If
                    DeleteGreenTempObjects()
                    trans.Commit()
                End Using
            Else 'Nothing found within the BF polyline
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Dim BFPl As Polyline = CType(trans.GetObject(BF, OpenMode.ForRead), Polyline)
                    'Find the CS object that this BF is within
                    objContPLID = GetCSthatContainsBF(BFPl, 20)
                    If Not objContPLID.IsNull Then
                        Dim objContPL As Polyline = CType(trans.GetObject(objContPLID, OpenMode.ForRead), Polyline)
                        Dim oCS As New CS
                        oCS.Type = objContPL.Layer
                        oCS.Surface = objContPL.Area
                        lstCS.Add(oCS)
                        'Recalculation of the total
                        dblSurfTech = 0
                        For Each CS In lstCS
                            dblSurfTech += CS.Surface
                        Next
                        objContPL.Dispose()
                        BFPl.Dispose()
                        Return True
                    Else
                        Return False
                    End If
                End Using

            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "IntersectBFwithCS", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Return retVal
    End Function

    ''' <summary>
    ''' Delete the temporary green regions
    ''' </summary>
    ''' <returns>True if there were any regions to delete and therefore the outer transaction needs to be committed too</returns>
    ''' <remarks></remarks>
    Private Function DeleteGreenTempObjects() As Boolean
        Try
            Dim retVal As Boolean = False
            'Set up the typedvalues to select green regions on the desired layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "REGION"), _
                                        New TypedValue(DxfCode.LayerName, "MO_BF_INCOMPLET"), _
                                        New TypedValue(DxfCode.Color, 3)}
            'Selct all items that pass the filter
            Dim pres As PromptSelectionResult = SelectAllItems(Values)
            'Check whether we found anything
            If pres.Status = PromptStatus.OK Then
                'Set the flag to say that the outer transaction needs to be committed
                retVal = True
                'Get the databse abnd start a transaction
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Loop through the found ids
                    For Each id As ObjectId In pres.Value.GetObjectIds
                        'Get the entity from the database
                        Dim ent As Entity = CType(trans.GetObject(id, OpenMode.ForWrite), Entity)
                        'and delete it
                        ent.Erase()
                    Next
                    'save the changes (if this function was called from within another transaction
                    'then the outer transaction needs to be committed too)
                    trans.Commit()
                End Using
            End If
            Return retVal
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "DeleteGreenTempObjects", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    ''' <summary>
    ''' Finds holes inside the CS objects and subtracts the area of these holes from the area of the CS object
    ''' </summary>
    ''' <param name="intProgressDivider"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SearchForHolesInCS(ByVal intProgressDivider As Integer) As Boolean
        Try
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'Loop through the CS objects in the list
            For Each oCS In lstCS
                If oCS.BF_Ilot = False Then
                    'Get the CS Region
                    Dim objCSreg As Region = CType(oCS.RegCS.GetObject(OpenMode.ForRead), Region)
                    If objCSreg.Area > 0 Then
                        'Convert the region to a polyline retaining the original region for the following intersections
                        Dim objCSpoly As ObjectId = ConvertRegionToPolyline(objCSreg, True)
                        'Check that the region converted to a polyline OK
                        If objCSpoly.IsNull Then
                            'The object could not be converted
                        Else
                            'Get the vertices of the polyline
                            Dim Verts As Point3dCollection = GetPolyVertices(objCSpoly)
                            If Verts.Count > 2 Then 'Ignore lines
                                'Set up the values to filter green regions
                                Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "REGION"), _
                                                                                      New TypedValue(DxfCode.LayerName, "MO_BF_INCOMPLET"), _
                                                                                      New TypedValue(DxfCode.Color, 3)}
                                'Select the green regions inside the CS polyline
                                Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)
                                'Cehck that we found some
                                If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                                    'Copy the CS region
                                    Dim objRefReg As Region = CType(objCSreg.Clone, Region)
                                    'Loop through the found items sunbtracting each one from the copy of the CS region
                                    For Each id As ObjectId In ids
                                        Dim objSubreg As Region = CType(id.GetObject(OpenMode.ForRead), Region)
                                        'Check that this sub region is not the CS region
                                        If objCSreg.ObjectId <> objSubreg.ObjectId And objSubreg.Area > 0 Then
                                            objRefReg.BooleanOperation(BooleanOperationType.BoolSubtract, CType(objSubreg.Clone, Region))
                                        End If
                                        objSubreg.Dispose()
                                    Next
                                    If objRefReg.Area > 0 Then oCS.Surface = objRefReg.Area
                                    objRefReg.Dispose()
                                End If
                            End If
                        End If
                        objCSreg.Color = Color.FromColorIndex(ColorMethod.None, 3)
                    End If
                    Wait.progBar1.Value = Wait.progBar1.Value + Convert.ToInt32(Math.Abs(((35 / lstCS.Count) / intProgressDivider)))
                End If
            Next
            Return True
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "SearchForHolesInCS", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try

    End Function

    ''' <summary>
    ''' Finds the CS that contains the BF
    ''' </summary>
    ''' <param name="objPoly">The BF polyline</param>
    ''' <param name="dblFacteurRecherche"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCSthatContainsBF(ByVal objPoly As Polyline, ByVal dblFacteurRecherche As Integer) As ObjectId
        Try
            Dim retVal As ObjectId = Nothing 'Return value
            'Get the vertices of the BF polyline after scaling it up 
            Dim Verts As Point3dCollection = GetScaledPolyVertices(objPoly.ObjectId, dblFacteurRecherche, True)
            'Set up typedvalue to select polylines
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_CS*,MUT_CS_*")}
            'Set up typedvalues to select polylines on the BF polyline layer
            Dim Values2() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, objPoly.Layer)}
            'Select all polylines within the scaled up FB polyline
            Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)
            'Check that we found something
            If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                'Get the database and start a transaction 
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Loop through the objectids selecting polylines within each one on the BF layer
                    For Each objPLCSID As ObjectId In ids
                        'Get the polyline
                        Dim objPLCS As Polyline = CType(trans.GetObject(objPLCSID, OpenMode.ForRead), Polyline)
                        'Get the layer that this polyline is on
                        Dim Lay As String = objPLCS.Layer
                        'Check that it's on a layer that we're interested in
                        If (Lay.StartsWith("MO_CS_") And Not Lay.StartsWith("MO_CS_PTS") And Not Lay.Contains("HACH")) Or _
                            (Lay.StartsWith("MUT_CS") And Not Lay.StartsWith("MUT_CS_PTS") And Not Lay.Contains("HACH")) Then
                            'Get the vertices of this polyline
                            Verts = GetPolyVertices(objPLCSID)
                            'Select all polylines on the BF layer within this polyline
                            Dim ids2() As ObjectId = SelectCrossingPoly(Values2, Verts)
                            'Check that we found something
                            If Not ids2 Is Nothing AndAlso Not ids2.Count = 0 Then
                                'Loop through the found objects
                                For Each objPLbf As ObjectId In ids2
                                    'If this polyline is the BF polyline
                                    If objPLbf.Equals(objPoly.ObjectId) Then
                                        'Set the return value to the containing polyline
                                        retVal = objPLCSID
                                        'Dispose of this CS polyline as we're done with it
                                        objPLCS.Dispose()
                                        'Return the objectid of the CS polyline
                                        Return retVal
                                    End If
                                Next
                            End If
                        End If
                    Next
                End Using
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetCSthatContainsBF", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Return Nothing
    End Function

    Private Function IntersectBFwithODS(ByVal BFID As ObjectId, Optional ByVal booBFHole As Boolean = False, _
                                        Optional ByVal intProgressDivider As Integer = 1) As Boolean
        If Not booBFHole Then
            Wait.lblStatus1.Text = "Calcul des intersections spatiales (BF / ODS)"
            frms.Application.DoEvents()
        End If
        Try
            Dim strFmtSurf As String = "#,##0.0000"
            'Get the vertices of the BF polyline
            Dim verts As Point3dCollection = GetPolyVertices(BFID)
            'Convert the BF polyline to a region 
            Dim objRegBF As Region = ConvertPolylineToRegion(BFID)
            'Select all ODS polylines that intersect (or are inside) the BF polyline
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE")}
            Dim ids() As ObjectId = SelectCrossingPoly(Values, verts)
            If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                'Get the database and start a transaction
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Loop through the found objects
                    For Each objLWPID As ObjectId In ids
                        'Get the polyline from the database
                        Dim objLWP As Polyline = CType(trans.GetObject(objLWPID, OpenMode.ForRead), Polyline)
                        'get the polyline's layer
                        Dim lay As String = objLWP.Layer
                        'Check that it's on one of the layers we want
                        If lay.StartsWith("MO_ODS_BATSOUT") Or lay.StartsWith("MO_ODS_COUVINDEP") Or _
                                        lay.StartsWith("MUT_ODS_BATSOUT") Or lay.StartsWith("MUT_ODS_COUVINDEP") Then
                            'Convert ODS polyline to region
                            Dim objRegODS As Region = ConvertPolylineToRegion(objLWPID)
                            'Chekc that the polyline converted to a region ok
                            If Not objRegODS Is Nothing Then
                                'Copy the BF region and intersect this region with the copy  
                                Dim objRegBFCopy = CType(objRegBF.Clone, Region)
                                objRegODS.BooleanOperation(BooleanOperationType.BoolIntersect, objRegBFCopy)
                                'if the resulting region has an area greater than 0
                                If objRegODS.Area > 0 Then
                                    'Create a new ODS object
                                    Dim oODS As New ODS
                                    'And set its properties
                                    oODS.Type = objLWP.Layer
                                    oODS.Surface = objRegODS.Area
                                    oODS.ODSRegID = objRegODS.ObjectId
                                    'Get the XData from the polylne
                                    Try
                                        Dim tv() As TypedValue = objLWP.XData.AsArray
                                        oODS.Description = tv(1).Value.ToString
                                        oODS.ID = tv(0).Value.ToString
                                    Catch ex As Exception
                                        oODS.Description = "?"
                                        oODS.ID = "-1"
                                    End Try
                                    If objLWP.Area - objRegODS.Area > 0.01 Then
                                        oODS.Divers = Format(objLWP.Area, strFmtSurf)
                                    End If
                                    lstODS.Add(oODS)
                                End If
                                objRegODS.Dispose()
                            End If
                            If Not booBFHole Then
                                Wait.progBar1.Value += Convert.ToInt16(Math.Abs((10 / ids.Count) / intProgressDivider))
                                frms.Application.DoEvents()
                            End If
                        End If
                    Next
                End Using
            End If
            objRegBF.Dispose()
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "IntersectBFwithODS", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    ''' <summary>
    ''' Creates the HTML document to display the results
    ''' </summary>
    ''' <param name="strBFID">The ID of the BF polyline</param>
    ''' <param name="booOpenHTML">Optional boolean to open the document once created or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateEDTHtml(ByVal strBFID As String, Optional ByVal booOpenHTML As Boolean = True) As Boolean
        Try
            'MsgBox(My.Application.Culture.ToString) 'affiche 'fr-FR'
            My.Application.ChangeCulture("fr-CH")
            'MsgBox(My.Application.Culture.ToString) 'affiche 'fr-FR'

            Dim dblSurfTotale As Double = 0 ' Int64 MODIF TH / 10.01.2012
            'Dim dt As System.Data.DataTable = GetRecordset("EDT_CS", "SELECT SUM(Surface) FROM EDT_CS WHERE Immeuble='" & strBFID & "'", "System\System.db3")
            Dim dt As System.Data.DataTable = GetRecordsetUsingWrapper("EDT_CS", "SELECT SUM(Surface) FROM EDT_CS WHERE Immeuble='" & strBFID & "'", Ass.DB3system)
            If Not dt Is Nothing AndAlso Not dt.Rows.Count = 0 Then
                dblSurfTotale = Convert.ToDouble(dt.Rows(0)(0)) 'MODIF TH / 10.01.2012
            Else
                Return False
            End If
            Dim strFmtSurf As String = "#,##0.0000"
            Dim strReport As String = Ass.EDTFolder ' LireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "EDTFolder")
            If strReport = String.Empty Then
                strReport = EDTFolder
            End If
            If Not strReport.EndsWith("\") Then strReport = strReport & "\"
            strReport = strReport & strBFID.Replace(" / ", "_") & "_1.htm"

            'Chargement du Templates
            Dim Connect As New Revo.connect
            Dim RVinfo As New Revo.RevoInfo
            Dim HTMLtemplate As String = ""
            If System.IO.File.Exists(RVinfo.EDTtemplate) Then
                Try
                    Dim sr As System.IO.StreamReader = New System.IO.StreamReader(RVinfo.EDTtemplate, System.Text.Encoding.Default)
                    While Not sr.EndOfStream()
                        HTMLtemplate += sr.ReadLine() & vbCrLf   '--- Traitement du fichier ligne par ligne
                    End While
                    sr.Close()
                Catch ex As Exception
                    Throw ex
                    Connect.RevoLog(Connect.DateLog & "Write EDT" & vbTab & False & vbTab & "Erreur lecture: " & ex.Message & vbTab & RVinfo.EDTtemplate)
                End Try
            Else
                Connect.RevoLog(Connect.DateLog & "Write EDT" & vbTab & False & vbTab & "Erreur lecture: " & "fichier inexistant" & vbTab & RVinfo.EDTtemplate)
            End If

            'Replace Variable HTML
            ' [[REVO]]
            HTMLtemplate = Replace(HTMLtemplate, "[[REVO]]", "<a href=""http://platform5rd.com"" target=""_blank"" style=""color:#000; font-style:normal ; font:Arial, Helvetica, sans-serif; font-size:10px"">" & "REVO " & Ass.xVersion & "</a> ")
            HTMLtemplate = Replace(HTMLtemplate, "[[User]]", GetCurrentWinLogin())              ' [[User]]
            HTMLtemplate = Replace(HTMLtemplate, "[[Date]]", Date.Today.ToLongDateString())     ' [[Date]] 
            HTMLtemplate = Replace(HTMLtemplate, "[[Hour]]", TimeOfDay.ToShortTimeString())     ' [[Hour]]

            HTMLcontent = ""         'Reset content 

            'Write the header
            'Voir Templates <<  Header

            'dt = GetRecordset("EDT_plans", "SELECT * FROM EDT_plans WHERE Immeuble='" & strBFID & "'", "System\System.db3")
            dt = GetRecordsetUsingWrapper("EDT_plans", "SELECT * FROM EDT_plans WHERE Immeuble='" & strBFID & "'", Ass.DB3system)
            If Not dt Is Nothing AndAlso Not dt.Rows.Count = 0 Then
                Dim strBF As String = dt.Rows(0)("NoBF").ToString
                If Convert.ToBoolean(dt.Rows(0)("DP")) = True Then strBF = "DP" & strBF
                'District, parcelle, commune, surface
                HTMLwriter("    <tr>")
                HTMLwriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                HTMLwriter("        <tr>")
                HTMLwriter("          <td width=""50%"">District : <b><span class=""Titre2"">" & dt.Rows(0)("NomDistrict").ToString & "</span></b></td>")
                HTMLwriter("          <td width=""25%"" align=""center""><b>Parcelle:</b></td>")
                HTMLwriter("          <td width=""25%""><b><span class=""Titre2"">" & strBF & "</span></b></td>")
                HTMLwriter("        </tr>")
                HTMLwriter("        <tr>")
                HTMLwriter("          <td colsplan=""3"">&nbsp;</td>")
                HTMLwriter("        </tr>")
                HTMLwriter("        <tr>")
                HTMLwriter("          <td width=""50%"">Commune : <b><span class=""Titre2"">" & dt.Rows(0)("NoCom").ToString & " (" & dt.Rows(0)("NoComCH").ToString & ") " & dt.Rows(0)("NomCommune").ToString & "</span></b></td>")
                HTMLwriter("          <td width=""25%"">Surface technique [m&sup2;]:</td>")
                HTMLwriter("          <td width=""25%"">" & Format(dblSurfTotale, strFmtSurf) & "</td>")
                HTMLwriter("        </tr>")
                HTMLwriter("        <tr>")
                HTMLwriter("          <td width=""50%"">&nbsp;</td>")
                HTMLwriter("          <td width=""25%"">Surface arrondie [m&sup2;]:</td>")
                HTMLwriter("          <td width=""25%"">" & Format(dblSurfTotale, "#,##0") & "</td>")
                HTMLwriter("        </tr>")
                HTMLwriter("        <tr>")
                HTMLwriter("          <td colsplan=""3"">&nbsp;</td>")
                HTMLwriter("        </tr>")
                'Local Names
                HTMLwriter("      </table></td>")
                HTMLwriter("    </tr>")
                'Surfaces by plan
                HTMLwriter("    <tr>")
                HTMLwriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                HTMLwriter("        <tr>")
                HTMLwriter("          <td width=""5%"" align=""right""><b>Plan</b></td>")
                HTMLwriter("          <td width=""7%"">&nbsp;</td>")
                HTMLwriter("          <td width=""28%""><b>Type de mensuration</b></td>")
                HTMLwriter("          <td width=""12%"" align=""right""><b>Code plan</b></td>")
                HTMLwriter("          <td width=""28%"">&nbsp;</td>")
                HTMLwriter("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
                HTMLwriter("        </tr>")
                HTMLwriter("        <tr>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td align=""center""><span class=""Legende"">[m&sup2;]</span></td>")
                HTMLwriter("        </tr>")
                'Loop for each plan
                Dim intPlan As Integer, strCodePlan As String, strTypeMens As String
                Dim dblSomme As Double = 0
                For i As Integer = 0 To dt.Rows.Count - 1
                    intPlan = Convert.ToInt32(dt.Rows(i)("NoPlan"))
                    strCodePlan = dt.Rows(i)("CodePlan").ToString
                    strTypeMens = Replace(CodePlanToTypeMens(intPlan, strCodePlan), "é", "&eacute;")
                    HTMLwriter("        <tr>")
                    If intPlan > 9000 Then
                        HTMLwriter("          <td align=""right"">&nbsp;</td>")
                        HTMLwriter("          <td>&nbsp;</td>")
                        HTMLwriter("          <td>&nbsp;</td>")
                        HTMLwriter("          <td align=""right"">&nbsp;</td>")
                    Else
                        HTMLwriter("          <td align=""right"">" & intPlan & "</td>")
                        HTMLwriter("          <td>&nbsp;</td>")
                        HTMLwriter("          <td>" & strTypeMens & "</td>")
                        HTMLwriter("          <td align=""right"">" & strCodePlan & "</td>")
                    End If
                    HTMLwriter("          <td>&nbsp;</td>")
                    HTMLwriter("          <td align=""right"">" & Format(Convert.ToDouble(dt.Rows(i)("Surface")), strFmtSurf) & "</td>")
                    HTMLwriter("        </tr>")
                    dblSomme = dblSomme + Convert.ToDouble(dt.Rows(i)("Surface"))
                Next
                'Total
                HTMLwriter("        <tr>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td><b>Total</b></td>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td>&nbsp;</td>")
                HTMLwriter("          <td align=""right"" style=""border-top-color:#000000; border-top-style:solid; border-top-width:3px;""><b>" & Format(dblSomme, strFmtSurf) & "</b></td>")
                HTMLwriter("        </tr>")
                HTMLwriter("      </table></td>")
                HTMLwriter("    </tr>")

                HTMLwriter("    <tr>")
                HTMLwriter("      <td height=""30"">&nbsp;</td>")
                HTMLwriter("    </tr>")
                'DDP
                If lstDDP.Count > 0 Then
                    'Separator
                    HTMLwriter("    <tr>")
                    HTMLwriter("      <td height=""4px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000; border-top-style:solid; border-top-width:2px; border-top-color:#000000;""></td>")
                    HTMLwriter("    </tr>")
                    HTMLwriter("    <tr>")
                    HTMLwriter("      <td>&nbsp;</td>")
                    HTMLwriter("    </tr>")
                    If lstDDP.Count = 1 Then
                        HTMLwriter("    <tr>")
                        HTMLwriter("      <td>Parcelle contenant le DDP " & lstDDP(0).Numero & "</td>")
                        HTMLwriter("    </tr>")
                    Else
                        HTMLwriter("    <tr>")
                        HTMLwriter("      <td>Parcelle contenant les DDP " & lstDDP(0).Numero)
                        For i = 1 To lstDDP.Count - 1
                            HTMLwriter(", " & lstDDP(i).Numero)
                        Next
                        HTMLwriter("</td>")
                        HTMLwriter("    </tr>")
                    End If
                    HTMLwriter("    <tr>")
                    HTMLwriter("      <td>&nbsp;</td>")
                    HTMLwriter("    </tr >")
                End If
                WriteEDTCSTitle()
                'dt = GetRecordset("EDT_CS", "SELECT * FROM EDT_CS WHERE Immeuble='" & strBFID & "'", "System\system.db3")
                dt = GetRecordsetUsingWrapper("EDT_CS", "SELECT * FROM EDT_CS WHERE Immeuble='" & strBFID & "'", Ass.DB3system)
                If Not dt Is Nothing AndAlso Not dt.Rows.Count = 0 Then
                    Dim strNat As String = ""

                    Dim strGenre As String = CStr(IIf(dt.Rows(0)("Categorie") Is Nothing, "", dt.Rows(0)("Categorie").ToString))
                    intPlan = Convert.ToInt32(dt.Rows(0)("NoPlan"))
                    Dim dblSommeGenre As Double = 0
                    Dim dblSommeNat As Double = 0.0
                    dblSomme = 0
                    WriteEDTCSHeader(strGenre = "Bâtiment")
                    'Start of Table
                    htmlWriter("    <tr>")
                    htmlWriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                    Dim rowCounter As Integer = 0
                    Do
                        If rowCounter = dt.Rows.Count And strGenre = "Bâtiment" Then
                            Exit Do
                        ElseIf rowCounter = dt.Rows.Count And strGenre <> "Bâtiment" Then
                            GoTo TotalGenre
                        End If
                        If rowCounter = dt.Rows.Count OrElse strGenre <> dt.Rows(rowCounter)("Categorie").ToString Then
TotalGenre:

                            If strGenre <> "Bâtiment" Then
                                'Total of previous type
                                htmlWriter("        <tr>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td><b>Total du genre " & strGenre & "</b></td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td align=""right"" style=""border-top-color:#000000; border-top-style:solid; border-top-width:3px;""><b>" & Format(dblSommeGenre, strFmtSurf) & "</b></td>")
                                htmlWriter("        </tr>")
                                htmlWriter("        <tr>")
                                htmlWriter("          <td colspan=""7"">&nbsp;</td>")
                                htmlWriter("        </tr>")
                            Else
                                'Total and separator at the end of the building
                                htmlWriter("      </table></td>")
                                htmlWriter("    </tr>")
                                Call WriteEDTCSHeader(False) 'Untitled "CS", separator and header only
                                htmlWriter("    <tr>")
                                htmlWriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                            End If
                            If rowCounter = dt.Rows.Count Then Exit Do
                            strGenre = dt.Rows(rowCounter)("Categorie").ToString
                            'Title (except for buildings)
                            If strGenre <> "Bâtiment" Then
                                htmlWriter("        <tr>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td><b>" & strGenre & "</b></td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("          <td>&nbsp;</td>")
                                htmlWriter("        </tr>")
                            End If
                            'Reset the totals
                            dblSommeGenre = 0
                            dblSommeNat = 0
                        End If
                        dblSomme += Convert.ToDouble(dt.Rows(rowCounter)("Surface"))
                        dblSommeGenre += Convert.ToDouble(dt.Rows(rowCounter)("Surface"))
                        dblSommeNat += Convert.ToDouble(dt.Rows(rowCounter)("Surface"))
                        If strGenre = "Bâtiment" Then
                            htmlWriter("        <tr>")
                            If Convert.ToInt32(dt.Rows(rowCounter)("NoPlan")) < 9000 Then
                                htmlWriter("          <td width=""5%"" align=""right"">" & dt.Rows(rowCounter)("NoPlan").ToString & "</td>")
                            Else
                                htmlWriter("          <td width=""5%"" align=""right"">&nbsp;</td>")
                            End If
                            htmlWriter("          <td width=""3%"">&nbsp;</td>")
                            htmlWriter("          <td width=""32%""><b>Bâtiment</b></td>")
                            htmlWriter("          <td width=""14%"" align=""right"">" & dt.Rows(rowCounter)("ID_objet").ToString & "</td>")
                            htmlWriter("          <td width=""6%"">&nbsp;</td>")
                            If (dt.Rows(rowCounter)("Info_objet").ToString & "") <> "" Then
                                htmlWriter("          <td width=""20%"" align=""right"">" & dt.Rows(rowCounter)("Info_objet").ToString & "</td>")
                            Else
                                htmlWriter("          <td width=""20%"">&nbsp;</td>")
                            End If
                            htmlWriter("          <td width=""20%"" align=""right"">" & Format(Convert.ToDouble(dt.Rows(rowCounter)("Surface")), strFmtSurf) & "</td>")
                            htmlWriter("        </tr>")
                            htmlWriter("        <tr>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>" & dt.Rows(rowCounter)("Designation_objet").ToString & "</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("        </tr>")
                        Else
                            htmlWriter("        <tr>")
                            If Convert.ToInt32(dt.Rows(rowCounter)("NoPlan")) < 9000 Then
                                htmlWriter("          <td align=""right"">" & dt.Rows(rowCounter)("NoPlan").ToString & "</td>")
                            Else
                                htmlWriter("          <td align=""right"">&nbsp;</td>")
                            End If
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>" & dt.Rows(rowCounter)("EDT").ToString & "</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td align=""right"">" & Format(Convert.ToDouble(dt.Rows(rowCounter)("Surface")), strFmtSurf) & "</td>")
                            strNat = dt.Rows(rowCounter)("EDT").ToString
                            'Check if this is the last record in the table
                            If Not rowCounter + 1 >= dt.Rows.Count Then
                                'This is not the last record then we can check the next record
                                If strNat <> dt.Rows(rowCounter + 1)("EDT").ToString Then
                                    htmlWriter("          <td align=""right"">" & Format(dblSommeNat, strFmtSurf) & "</td>")
                                    dblSommeNat = 0
                                Else
                                    htmlWriter("          <td>&nbsp;</td>")
                                End If
                            Else
                                'This is the last record
                                htmlWriter("          <td align=""right"">" & Format(dblSommeNat, strFmtSurf) & "</td>")
                                dblSommeNat = 0
                            End If
                            htmlWriter("        </tr>")
                        End If
                        rowCounter += 1
                    Loop
                    'Grand total of the plot
                    htmlWriter("        <tr>")
                    htmlWriter("          <td colspan=""7"">&nbsp;</td>")
                    htmlWriter("        </tr>")

                    htmlWriter("        <tr>")
                    htmlWriter("          <td width=""5%"">&nbsp;</td>")
                    htmlWriter("          <td width=""3%"">&nbsp;</td>")
                    htmlWriter("          <td width=""32%""><b>Total de la parcelle</b></td>")
                    htmlWriter("          <td width=""14%"">&nbsp;</td>")
                    htmlWriter("          <td width=""6%"">&nbsp;</td>")
                    htmlWriter("          <td width=""20%"">&nbsp;</td>")
                    htmlWriter("          <td width=""20%"" align=""right""><b>" & Format(dblSomme, strFmtSurf) & "</b></td>")
                    htmlWriter("        </tr>")
                    'Foot section (end of table)
                    WriteEDTCSFooter()
                    'OD
                    'dt = GetRecordset("EDT_CS", "SELECT * FROM EDT_OD WHERE Immeuble='" & strBFID & "'", "System\system.db3")
                    dt = GetRecordsetUsingWrapper("EDT_CS", "SELECT * FROM EDT_OD WHERE Immeuble='" & strBFID & "'", Ass.DB3system)
                    If Not dt Is Nothing AndAlso Not dt.Rows.Count = 0 Then
                        WriteEDTODSHeader(False)
                        'Start of table
                        htmlWriter("    <tr>")
                        htmlWriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                        rowCounter = 0
                        Do Until rowCounter = dt.Rows.Count
                            'Writing the ODS object
                            htmlWriter("        <tr>")
                            If Convert.ToInt32(dt.Rows(rowCounter)("NoPlan")) < 9000 Then
                                htmlWriter("          <td width=""5%"" align=""right"">" & dt.Rows(rowCounter)("NoPlan").ToString & "</td>")
                            Else
                                htmlWriter("          <td width=""5%"" align=""right"">&nbsp;</td>")
                            End If
                            htmlWriter("          <td width=""3%"">&nbsp;</td>")
                            htmlWriter("          <td width=""32%""><b>" & dt.Rows(rowCounter)("Categorie").ToString & "</b></td>")
                            htmlWriter("          <td width=""14%"" align=""right"">" & dt.Rows(rowCounter)("ID_objet").ToString & "</td>")
                            htmlWriter("          <td width=""6%"">&nbsp;</td>")
                            If (dt.Rows(rowCounter)("Info_objet").ToString & "") <> "" Then
                                htmlWriter("          <td width=""20%"" align=""right"">" & dt.Rows(rowCounter)("Info_objet").ToString & "</td>") 'Already formatted
                            Else
                                htmlWriter("          <td width=""20%"">&nbsp;</td>")
                            End If
                            htmlWriter("          <td width=""20%"" align=""right"">" & Format(Convert.ToDouble(dt.Rows(rowCounter)("Surface")), strFmtSurf) & "</td>")
                            htmlWriter("        </tr>")

                            '2nd line for the designation
                            htmlWriter("        <tr>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>" & dt.Rows(rowCounter)("Designation_objet").ToString & "</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("          <td>&nbsp;</td>")
                            htmlWriter("        </tr>")
                            rowCounter += 1
                        Loop
                        'End of section
                        htmlWriter("      </table></td>")
                        htmlWriter("    <tr>")
                        htmlWriter("    <tr>")
                        htmlWriter("      <td>&nbsp;</td>")
                        htmlWriter("    </tr>")
                    End If
                    'Foot of page
                    'Voir Templates <<  Footer : WriteEDTFooter(sWriter)


                    Dim htmlWrite As New System.IO.StreamWriter(strReport, False, System.Text.Encoding.Unicode)
                    HTMLtemplate = Replace(HTMLtemplate, "[[EDT]]", HTMLcontent)
                    htmlWrite.WriteLine(HTMLtemplate)
                    htmlWrite.Close()
                    htmlWrite.Dispose()
                End If
                If booOpenHTML Then System.Diagnostics.Process.Start(strReport)
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "CreateEDTHtml", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Return False
        End Try

    End Function
    Private Function HTMLwriter(Contents As String)

        HTMLcontent += Contents & vbCrLf
        Return True

    End Function
    Private Function WriteEDTHeader(ByRef sWriter As System.IO.StreamWriter) As Boolean
        Try
            sWriter.WriteLine("<!-- START EDTHeader -->")
            sWriter.WriteLine("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">")
            sWriter.WriteLine("<html>")
            sWriter.WriteLine("")
            sWriter.WriteLine("<head>")
            sWriter.WriteLine("  <title>EDT</title>")
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
            sWriter.WriteLine("    .Titre2 { font-size: 10pt; }")
            sWriter.WriteLine("    .Legende { font-size: 8pt; }")
            sWriter.WriteLine("  </style>")
            sWriter.WriteLine("</head>")
            sWriter.WriteLine("")
            sWriter.WriteLine("<body>")
            sWriter.WriteLine("  <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width: 170mm;"">")
            sWriter.WriteLine("    <tr>")
            sWriter.WriteLine("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            sWriter.WriteLine("        <tr>")
            'Ignore call to CheckAppState as binary values for License are not in registry
            'CheckAppState
            'If Not booDemo Then
            ' Dim strLicence As String = LireBaseRegistre(0, "SoftWare\SWWS\JourCAD", "LicenceName") & "<br>" & LireBaseRegistre(0, "SoftWare\SWWS\JourCAD", "LicenceSite")
            ' sWriter.WriteLine("          <td width=""40%""><div align=""left"">" & strLicence & "</div></td>")
            sWriter.WriteLine("<td width=""60%""><div align=""left"">REVO " & Ass.xVersion & "<br>&nbsp;</div></td>")
            'Else
            '    sWriter.WriteLine("          <td width=""60%""><div align=""left"">REVO " & Ass.xVersion & "<br>Version d'&eacute;valuation</div></td>")
            'End If
            Dim strInfos As String = ""
            Dim logon As String = GetCurrentWinLogin()
            If logon <> String.Empty Then
                strInfos = "Etabli par " & logon & "<br>le " & Date.Today.ToLongDateString & "&nbsp;&agrave;&nbsp;" & TimeOfDay.ToShortTimeString
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
            sWriter.WriteLine("      <td><div align=""center""><span class=""Titre"">ETAT DESCRIPTIF TECHNIQUE</span></div></td>")
            sWriter.WriteLine("    </tr>")
            sWriter.WriteLine("    <tr>")
            sWriter.WriteLine("      <td>&nbsp;</td>")
            sWriter.WriteLine("    </tr>")
            sWriter.WriteLine("    <tr>")
            sWriter.WriteLine("      <td height=""4px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000; border-top-style:solid; border-top-width:2px; border-top-color:#000000;""></td>")
            sWriter.WriteLine("    </tr>")
            sWriter.WriteLine("    <tr>")
            sWriter.WriteLine("      <td height=""30"">&nbsp;</td>")
            sWriter.WriteLine("    </tr>")
            sWriter.WriteLine("<!-- END EDTHeader -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTHeader", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try

    End Function

    Private Function WriteEDTCSTitle() As Boolean
        Try
            HTMLwriter("<!-- START EDTCSTitle -->")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td height=""4px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000; border-top-style:solid; border-top-width:2px; border-top-color:#000000;""></td>")
            HTMLwriter("    </tr>")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td>&nbsp;</td>")
            HTMLwriter("    </tr>")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td bgcolor=""#CCCCCC""><div align=""center""><span class=""Titre"">COUVERTURE DU SOL</span></div></td>")
            HTMLwriter("    </tr>")
            HTMLwriter("<!-- END EDTCSTitle -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTCSTitle", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try

    End Function

    Private Function WriteEDTCSHeader(ByVal booBat As Boolean) As Boolean
        Try
            'New Table
            HTMLwriter("<!-- START EDTCSHeader -->")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td>&nbsp;</td>")
            HTMLwriter("    </tr>")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td height=""4px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000; border-top-style:solid; border-top-width:2px; border-top-color:#000000;""></td>")
            HTMLwriter("    </tr>")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td>&nbsp;</td>")
            HTMLwriter("    </tr>")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            HTMLwriter("        <tr>")
            HTMLwriter("          <td width=""5%"" align=""right""><b>Plan</b></td>")
            HTMLwriter("          <td width=""3%"">&nbsp;</td>")
            HTMLwriter("          <td width=""32%""><b>Genre</b></td>")
            If booBat = True Then 'Buildings
                HTMLwriter("          <td width=""14%"" align=""right""><b>No b&acirc;timent</b></td>")
            Else
                HTMLwriter("          <td width=""14%"">&nbsp;</td>")
            End If
            HTMLwriter("          <td width=""6%"">&nbsp;</td>")
            HTMLwriter("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
            HTMLwriter("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
            HTMLwriter("        </tr>")
            HTMLwriter("        <tr>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>D&eacute;signation</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td align=""center""><span class=""Legende"">[m&sup2;]</span></td>")
            HTMLwriter("          <td align=""center""><span class=""Legende"">[m&sup2;]</span></td>")
            HTMLwriter("        </tr>")
            HTMLwriter("        <tr>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            If booBat = True Then
                HTMLwriter("          <td align=""center""><b>B&acirc;timent complet</b></td>")
                HTMLwriter("          <td>&nbsp;</td>")
            Else
                HTMLwriter("          <td align=""center""><b><span class=""Legende"">Objet par plan</span></b></td>")
                HTMLwriter("          <td align=""center""><b><span class=""Legende"">Total par objet / genre</span></b></td>")
            End If
            HTMLwriter("        </tr>")
            HTMLwriter("        <tr>")
            HTMLwriter("          <td colspan=""7"" height=""30"">&nbsp;</td>")
            HTMLwriter("        </tr>")
            HTMLwriter("      </table></td>")
            HTMLwriter("    </tr>")
            HTMLwriter("<!-- END EDTCSHeader -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTCSHeader", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    Private Function WriteEDTCSFooter() As Boolean
        Try
            HTMLwriter("<!-- START EDTCSFooter -->")
            HTMLwriter("      </table></td>")
            HTMLwriter("    </tr>")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td>&nbsp;</td>")
            HTMLwriter("    </tr>")
            HTMLwriter("<!-- END EDTCSFooter -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTCSFooter", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    Private Function WriteEDTODSHeader(ByVal booBat As Boolean) As Boolean
        Try
            HTMLwriter("<!-- START EDTODSHeader -->")
            'Title and separator
            HTMLwriter("    <tr>")
            HTMLwriter("      <td bgcolor=""#CCCCCC""><div align=""center""><span class=""Titre"">OBJETS DIVERS (hors statistique)</span></div></td>")
            HTMLwriter("    </tr>")
            HTMLwriter("    <tr>")
            HTMLwriter("      <td>&nbsp;</td>")
            HTMLwriter("    </tr>")

            'new table
            HTMLwriter("    <tr>")
            HTMLwriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")

            'Separator and title of columns
            HTMLwriter("    <tr>")
            HTMLwriter("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            HTMLwriter("        <tr>")
            HTMLwriter("          <td width=""5%"" align=""right""><b>Plan</b></td>")
            HTMLwriter("          <td width=""3%"">&nbsp;</td>")
            HTMLwriter("          <td width=""32%""><b>Genre</b></td>")
            HTMLwriter("          <td width=""14%"" align=""right""><b>No b&acirc;timent</b></td>")
            HTMLwriter("          <td width=""6%"">&nbsp;</td>")
            HTMLwriter("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
            HTMLwriter("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
            HTMLwriter("        </tr>")
            HTMLwriter("        <tr>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>D&eacute;signation</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td align=""cente)""><span class=""Legende"">[m&sup2;]</span></td>")
            HTMLwriter("          <td align=""center""><span class=""Legende"">[m&sup2;]</span></td>")
            HTMLwriter("        </tr>")
            HTMLwriter("        <tr>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("          <td align=""center""><b>B&acirc;timent complet</b></td>")
            HTMLwriter("          <td>&nbsp;</td>")
            HTMLwriter("        </tr>")
            HTMLwriter("        <tr>")
            HTMLwriter("          <td colspan=""7"" height=""30"">&nbsp;</td>")
            HTMLwriter("        </tr>")
            HTMLwriter("      </table></td>")
            HTMLwriter("    </tr>")
            HTMLwriter("<!-- END EDTODSHeader -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTODSHeader", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    Private Function WriteEDTFooter(ByRef sWriter As System.IO.StreamWriter) As Boolean
        Try
            sWriter.WriteLine("<!-- START EDTFooter -->")
            sWriter.WriteLine("  </table>")
            sWriter.WriteLine("</body>")
            sWriter.WriteLine("")
            sWriter.WriteLine("</html>")
            sWriter.WriteLine("<!-- END EDTFooter -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTFooter", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function
End Module




Module modEDT2


    Private Connect As New Revo.connect
    Private lstBFbyRPL As List(Of BFbyRPL)
    Private lstDDP As List(Of DDP)
    Private lstCS As List(Of CS)
    Private lstODS As List(Of ODS)
    Private HTMLcontent As String = ""
    Private RVinfo As New Revo.RevoInfo
  
    Private Class BFbyRPL
        Private _BFID As ObjectId 'Polyline
        Private _RPLID As ObjectId  'Block
        ''' <summary>
        ''' The objectid of the BF polyline
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property BF() As ObjectId
            Get
                Return _BFID
            End Get
            Set(ByVal value As ObjectId)
                _BFID = value
            End Set
        End Property
        ''' <summary>
        ''' The objectid of the RPL block
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        Public Property RPL() As ObjectId
            Get
                Return _RPLID
            End Get
            Set(ByVal value As ObjectId)
                _RPLID = value
            End Set
        End Property
    End Class
    Private Class DDP
        Private _objPolyID As ObjectId
        Private _Surface As Double
        Private _Partiel As Boolean
        Private _Numero As String
        Public Property objPolyID() As ObjectId
            Get
                Return _objPolyID
            End Get
            Set(ByVal value As ObjectId)
                _objPolyID = value
            End Set
        End Property
        Public Property Surface() As Double
            Get
                Return _Surface
            End Get
            Set(ByVal value As Double)
                _Surface = value
            End Set
        End Property
        Public Property Partiel() As Boolean
            Get
                Return _Partiel
            End Get
            Set(ByVal value As Boolean)
                _Partiel = value
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
    End Class
    Private Class CS
        Private _Type As String
        Private _Surface As Double
        Private _RegCS As ObjectId
        Private _Description As String
        Private _ID As String
        Private _Divers As String
        Private _BF_Ilot As Boolean
        Public Property Type() As String
            Get
                Return _Type
            End Get
            Set(ByVal value As String)
                _Type = value
            End Set
        End Property
        Public Property Surface() As Double
            Get
                Return _Surface
            End Get
            Set(ByVal value As Double)
                _Surface = value
            End Set
        End Property
        Public Property RegCS() As ObjectId
            Get
                Return _RegCS
            End Get
            Set(ByVal value As ObjectId)
                _RegCS = value
            End Set
        End Property
        Public Property Description() As String
            Get
                Return _Description
            End Get
            Set(ByVal value As String)
                _Description = value
            End Set
        End Property
        Public Property ID() As String
            Get
                Return _ID
            End Get
            Set(ByVal value As String)
                _ID = value
            End Set
        End Property
        Public Property Divers() As String
            Get
                Return _Divers
            End Get
            Set(ByVal value As String)
                _Divers = value
            End Set
        End Property
        Public Property BF_Ilot() As Boolean
            Get
                Return _BF_Ilot
            End Get
            Set(ByVal value As Boolean)
                _BF_Ilot = value
            End Set
        End Property
    End Class
    Private Class ODS
        Private _Description As String
        Private _ID As String
        Private _Area As Double
        Private _Divers As String
        Private _Type As String
        Private _Surface As Double
        Private _ODSRegID As ObjectId
        Public Property Description() As String
            Get
                Return _Description
            End Get
            Set(ByVal value As String)
                _Description = value
            End Set
        End Property
        Public Property ID() As String
            Get
                Return _ID
            End Get
            Set(ByVal value As String)
                _ID = value
            End Set
        End Property
        Public Property Area() As Double
            Get
                Return _Area
            End Get
            Set(ByVal value As Double)
                _Area = value
            End Set
        End Property
        Public Property Divers() As String
            Get
                Return _Divers
            End Get
            Set(ByVal value As String)
                _Divers = value
            End Set
        End Property
        Public Property Type() As String
            Get
                Return _Type
            End Get
            Set(ByVal value As String)
                _Type = value
            End Set
        End Property
        Public Property Surface() As Double
            Get
                Return _Surface
            End Get
            Set(ByVal value As Double)
                _Surface = value
            End Set
        End Property
        Public Property ODSRegID() As ObjectId
            Get
                Return _ODSRegID
            End Get
            Set(ByVal value As ObjectId)
                _ODSRegID = value
            End Set
        End Property
    End Class
    Private Class BLnum
        Private _Name As String
        Private _ID As ObjectId
        Private _Pt As Point3d

        Public Property Name() As String
            Get
                Return _Name
            End Get
            Set(ByVal value As String)
                _Name = value
            End Set
        End Property
        Public Property ID() As ObjectId
            Get
                Return _ID
            End Get
            Set(ByVal value As ObjectId)
                _ID = value
            End Set
        End Property
        Public Property Pt() As Point3d
            Get
                Return _Pt
            End Get
            Set(ByVal value As Point3d)
                _Pt = value
            End Set
        End Property
      
    End Class


    Public Function EDTStart()

        ' Sélectionner les numéros de parcelles
        Dim CollSelect As New Collection
        CollSelect = SelectObj(False, "Sélectionner les numéros de parcelles (cliquer sur Enter pour terminer)", False, False, True)

        'Filtre objets
        'BF + DDP :
        ' Bloc :   IMMEUBLE_NUM, IMMEUBLE_NUM_DDP, IMMEUBLE_NUM_DDP_PROJ, IMMEUBLE_NUM_DDP_PROJ
        '          MUT_IMMEUBLE_NUM, MUT_IMMEUBLE_NUM_DDP
        '          RAD_IMMEUBLE_NUM, RAD_IMMEUBLE_NUM_DDP
        '
        ' Calque : MO_BF_DDP_IMMEUBLE, MO_BF_IMMEUBLE
        '          MUT_BF_DDP_IMMEUBLE, MUT_BF_IMMEUBLE
        '          RAD_BF_DDP_IMMEUBLE, RAD_BF_IMMEUBLE


        Dim CollNumParcelle As New Collection


        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction

            Try
                For Each BL As BlockReference In CollSelect
                    If Replace(Replace(BL.Name, "MUT_", ""), "RAD_", "") Like "IMMEUBLE_NUM" Then 'Pas de prise en compte des DDP
                        Dim BLn As New BLnum
                        BLn.ID = BL.ObjectId
                        BLn.Name = BL.Name
                        BLn.Pt = BL.Position
                        CollNumParcelle.Add(BLn)
                    End If
                Next
            Catch
            End Try

            If CollNumParcelle.Count = 0 Then
                MsgBox("Aucun bloc Immeuble_num* n'as été séléctionné", vbInformation + vbOKOnly, "Pas d'objet séléctionné")
                Return False
            End If

        End Using


        ' Sélectionner le dossier de destination
        Dim strFolder As String = RVinfo.EDTFolder

        If System.IO.Directory.Exists(strFolder) = False Then
            Dim fb As New Windows.Forms.FolderBrowserDialog
            fb.Description = "Sélectionner le dossier dans lequel enregistrer le fichier EDT résultat"
            If fb.ShowDialog = Windows.Forms.DialogResult.OK Then
                strFolder = fb.SelectedPath
            Else
                strFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            End If

            If Not strFolder.EndsWith("\") Then strFolder = strFolder & "\"
        End If


        ' Calcul du EDT + Mise en forme EDT

        If CollNumParcelle.Count > 0 Then
            'run the process on the list
            CalculEDT2(CollNumParcelle, False, False)
        Else
            Return False
        End If



        'Test
        'MsgBox(strFolder)


        Return True

    End Function

    Public Function CalculEDT2(ByVal CollNumParcelle As Collection, ByVal IgnoreRPL As Boolean, ByVal ControleRF As Boolean) As Boolean


        'Initialisation copie de la base de données
        ' Try 'Copie la base de donnée interlis.db3 (DB3interlisOrig -> DB3interlis)

        System.IO.File.Copy(System.IO.Path.Combine(RVinfo.SystemPath, RVinfo.DB3systemOrig), System.IO.Path.Combine(RVinfo.SystemPath, RVinfo.DB3system), True)

        'Save the current layer state
        SaveLayerState("REVO_Temp")
        FreezeAllLayers(False)
        'And switch them all on
        ShowAllLayers(True)
        'Geler calque avec des points
        'FreezeLayer("MO_BF_PTS", True)
        'FreezeLayer("MO_COM_PTS", True)
        'FreezeLayer("MO_CS_PTS", True)
        'FreezeLayer("MO_OD_PTS", True)
        'FreezeLayer("MO_VALEUR_PTS", True)
        'FreezeLayer("MO_PFP1", True)
        'FreezeLayer("MO_PFP2", True)
        'FreezeLayer("MO_PFP3", True)
        'FreezeLayer("MO_PFA1", True)
        'FreezeLayer("MO_PFA2", True)
        'FreezeLayer("MO_PFA3", True)



        'Call the function to process the items in the list

        'Store the number of items to process
        Dim ItemCounter As Integer = 0
        Dim FirstBL As ObjectId = Nothing


        For Each BLn As BLnum In CollNumParcelle ' Boucle pour chaque parcelle

            ClearTable("EDT_donnees")


            If ItemCounter = 0 Then FirstBL = BLn.ID
            Dim dblSurfTech As Double = 0
            ItemCounter += 1 'Create the progress form

            'Set its labels
            Connect.Message("Génération d'état descriptif (" & ItemCounter & " / " & CollNumParcelle.Count & ")", "Patienter pendant l'analyse des objets ... ", False, 4, 100)
            frms.Application.DoEvents()


            'Zoom to the extents of the drawing
            Zooming.ZoomExtents()
            'Thaw all layers

            'Regen (required after thawing layers)
            Application.DocumentManager.MdiActiveDocument.Editor.Regen()


            'Set the active layer
            ActivateLayer("MO_BF_INCOMPLET")

            ' ---------------------------------------------------------------------------------------------------------


            '1) Recherche du périmètre de la parcelle 
            '   - Biens Fonds : MO_BF
            '   - Biens Fonds projetés : MO_BF_PROJ
            '   - DDP : MO_BF_DDP
            '   - DDP projetés : MO_BF_DDP_PROJ

            'Filtre objets
            'BF + DDP :
            ' Bloc :   IMMEUBLE_NUM, IMMEUBLE_NUM_DDP, IMMEUBLE_NUM_PROJ, IMMEUBLE_NUM_DDP_PROJ
            '          MUT_IMMEUBLE_NUM, MUT_IMMEUBLE_NUM_DDP
            '          RAD_IMMEUBLE_NUM, RAD_IMMEUBLE_DDP
            '
            ' Calque : MO_BF_DDP_IMMEUBLE, MO_BF_IMMEUBLE
            '          MUT_BF_DDP_IMMEUBLE, MUT_BF_IMMEUBLE
            '          RAD_BF_DDP_IMMEUBLE, RAD_BF_IMMEUBLE

            Dim LName As String = "MO_BF"
            If BLn.Name = "IMMEUBLE_NUM" Then
                LName = "MO_BF"
            ElseIf BLn.Name = "IMMEUBLE_NUM_DDP" Then
                LName = "MO_BF_DDP"
            End If

            '2) Identification du périmètre extérieurs (BF ou DDP)
            Dim PerimColl As New Collection ' Listing des poliligne qui entoure le points du numéro du bien fonds
            Dim PolyId As ObjectId = PerimSearch(BLn.Pt, LName)


            If PolyId <> Nothing Then

                Dim PerimPolyLayer As String = ""
                Dim PerimPolyID As ObjectId = Nothing

                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Dim PerimPoly As Polyline
                    PerimPoly = CType(trans.GetObject(PolyId, OpenMode.ForRead), Polyline)
                    PerimPolyLayer = PerimPoly.Layer
                    PerimPolyID = PerimPoly.ObjectId
                End Using

                Dim EDTItem As EDTListItem = GetInfoFromBlock(BLn.ID)
                EDTItem.PolyLayer = PerimPolyLayer
                EDTItem.PolyID = PerimPolyID
                EDTItem.ID = "BF" & EDTItem.ComNum
                EDTItem.DisplayVal = PerimPolyLayer & " " & EDTItem.ComNum

                lstBFbyRPL = New List(Of BFbyRPL)
                Dim NumParcProv As String = ""
                Try
                    NumParcProv = EDTItem.Num
                Catch ex As Exception
                End Try


                '3) Recherche de plans (RPL)

                Connect.Message("Génération d'état descriptif : BF " & NumParcProv & " (" & ItemCounter & " / " & CollNumParcelle.Count & ")", "Analyse de la répartition des plans.. ", False, 5, 100)
                frms.Application.DoEvents()
                IntersectBFwithRPL2(PerimPolyID, IgnoreRPL)
             
                'Zoom sur parcelle pour optimisation de la recherche !  C'est bien ??????????
                Zooming.ZoomToObject(PolyId, 50) ' Padding en mètre


                'Recherche de DDP
                Connect.Message("Génération d'état descriptif : BF " & NumParcProv & " (" & ItemCounter & " / " & CollNumParcelle.Count & ")", "Recherche des DDP... ", False, 10, 100)
                frms.Application.DoEvents()
                Try
                    GetDDPInsideBF2(PolyId)
                Catch ex As Exception
                    MsgBox("Recherche de DDP : " & ex.Message, vbInformation, "Erreur de la recherche de DDP")
                End Try


                '4) Boucle dans les RPL
                Dim Counter As Integer = -1
                Dim CheckoBFRPL As Boolean = True

                For Each oBFRPL In lstBFbyRPL
                    lstCS = New List(Of CS)
                    lstODS = New List(Of ODS)
                    Dim intPlan As Integer = 0
                    Dim strCodePlan As String = ""
                    Counter += 1

                    Dim NumBF As String = ""
                    NumBF = EDTItem.Num

                    If Not oBFRPL.RPL.IsNull Then
                        Try
                            intPlan = Convert.ToInt32(GetBlockAttribute(oBFRPL.RPL, "NUMERO"))
                        Catch ex As Exception
                            'returned value is not an integer
                        End Try
                        If intPlan = 0 Then
                            intPlan = 9000 + Counter
                        Else
                            strCodePlan = GetBlockAttribute(oBFRPL.RPL, "CODEPLAN")
                        End If
                    End If

                    Connect.Message("Génération d'état descriptif : BF " & NumBF & " (" & ItemCounter & " / " & CollNumParcelle.Count & ")", "Recherche et analyse des îlots (BF)", False, Counter, lstBFbyRPL.Count)
                    frms.Application.DoEvents()


                    '5) Soustraction des périmètres intérieurs (BF ou DDP)

                    'Search for BF islands storing the sum of the areas for later use
                    dblSurfTech = (SearchForHolesInBF2(oBFRPL.BF, EDTItem.Com, intPlan, EDTItem.Num, EDTItem.ComNum, strCodePlan) * -1)


                    Connect.Message("Génération d'état descriptif : BF " & NumBF & " (" & ItemCounter & " / " & CollNumParcelle.Count & ")", "Calcul des intersections spatiales (BF / CS)", False, 15, 100)
                    frms.Application.DoEvents()

                    Dim strCurrentLayer As String = ""
                    If IntersectBFwithCS2(oBFRPL.BF, NumBF, , lstBFbyRPL.Count, strCurrentLayer) Then
                        Dim dt As System.Data.DataTable = GetTable("EDT_Donnees", dbName:=RVinfo.DB3system)
                        WriteCSObjectsToTable2(dt, EDTItem.Com, intPlan, EDTItem.Num, EDTItem.ComNum, strCodePlan)

                        Connect.Message("Génération d'état descriptif : BF " & NumBF & " (" & ItemCounter & " / " & CollNumParcelle.Count & ")", "Calcul des intersections spatiales (BF / ODS)", False, 20, 100)
                        frms.Application.DoEvents()

                        IntersectBFwithODS2(oBFRPL.BF, , lstBFbyRPL.Count)
                        WriteODSObjectsToTable2(dt, EDTItem.Com, intPlan, EDTItem.Num, EDTItem.ComNum, strCodePlan)



                        Connect.Message("Génération d'état descriptif : BF " & NumBF & " (" & ItemCounter & " / " & CollNumParcelle.Count & ")", "Calcul des intersections spatiales (BF / ODS)", False, Counter + 1, lstBFbyRPL.Count)
                        frms.Application.DoEvents()

                    Else

                        ' Connect.Message("Génération d'état descriptif : BF " & NumBF & " (" & ItemCounter & " / " & CollNumParcelle.Count & ")", "Calcul des intersections spatiales (BF / ODS)", True, 90, 100)

                        'frms.MessageBox.Show("L'état descriptif n'a pas pu être calculé !" & vbCrLf & _
                        '                     "Cela est peut être dû à un problème de topologie ou à une erreur interne à AutoCAD." _
                        '                     & vbCrLf & "Contrôler la géométrie et la topologie des couches CS (" & strCurrentLayer & _
                        '                     ")." & vbCrLf & vbCrLf & _
                        '                     "L'objet posant problème a été sélectionné: vérifiez qu'il s'agit d'une polyligne fermée et" _
                        '                     & vbCrLf & "qu'il se trouve dans le bon calque.", "Calcul impossible", _
                        '                      Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)


                        Dim Parc As String = ""
                        If NumBF <> "" Then Parc = "(BF " & NumBF & ")"
                        Connect.Message("L'état descriptif n'a pas pu être calculé " & Parc & ".", _
                                               "Cela est peut être dû à un problème de topologie " & vbCrLf & _
                                              "Contrôler la géométrie et la topologie des couches CS (" & strCurrentLayer & ").", _
                                             False, 99, 100, "critical")

                        CheckoBFRPL = False
                        Exit For
                    End If
                Next

                CheckoBFRPL = True ' TEST -----

                If CheckoBFRPL Then
                    Connect.Message("Création du rapport : BF " & NumParcProv, ItemCounter & " / " & CollNumParcelle.Count & " EDT calculé", False, ItemCounter, CollNumParcelle.Count)
                    CreateEDTHtml2(EDTItem.ComNum, NumParcProv, True)
                End If

            Else
                Connect.Message("Génération d'état descriptif", "Erreur de recherche de parcelle", False, 90, 100, "critical")
                ' Return Nothing
            End If

            '  Zooming.ZoomPrevious()
            'Zooming.ZoomPrevious()

            ' CommandLine.CommandC(Chr(27))

            'Regen (required after thawing layers)
            Application.DocumentManager.MdiActiveDocument.Editor.Regen()
            ' Application.DocumentManager.MdiActiveDocument.Editor.UpdateScreen()

        Next



        ' ---------------------------------------------------------------------------------------------------------
        ' CommandLine.CommandC(Chr(27) & Chr(27))

        Connect.Message("Génération d'état descriptif", "Régénération du dessin... ", False, 97, 100)
        frms.Application.DoEvents()
        RestoreLayerState("REVO_Temp")
        RemoveLayerState("REVO_Temp")
        'Zooming.ZoomPrevious()
        If FirstBL <> Nothing Then Zooming.ZoomToObject(FirstBL)
        Connect.Message("Génération d'état descriptif", "Régénération du dessin... ", True, 100, 100)


        'Suppresion de la base de données

        Return True

    End Function



    Public Function IsPointInPolyLine(Polyline As Polyline, Point As Point3d) As Boolean
        Dim hIsInPLine As Boolean = False
        Dim hPoint1 As Point3d = Point
        'Dim hPoint2 As Point3d = PolarPoints(hPoint1, 0, 1)
        Dim hPoint2 As Point3d = PolarPoints3D(hPoint1, 0, 1)

        Dim hVector As Vector3d = hPoint2 - hPoint1
        Dim hRay As New Ray()

        hRay.BasePoint = New Point3d(Point.X, Point.Y, 0)
        hRay.UnitDir = hVector

        Dim hIntersectionPoints As New Point3dCollection()
        'Autodesk.AutoCAD.DatabaseServices.IntersectWith(hRay, Polyline, Intersect.OnBothOperands, hIntersectionPoints)
        hRay.IntersectWith(Polyline, Intersect.OnBothOperands, hIntersectionPoints, IntPtr.Zero, IntPtr.Zero)

        'hRay.IntersectWith(
        '  Polyline,
        '  Intersect.OnBothOperands,
        '  hIntersectionPoints, 0, 0);

        Dim hMod As Integer = hIntersectionPoints.Count Mod 2
        If hMod = 0 Then
            hIsInPLine = False
        Else
            hIsInPLine = True
        End If
        Return hIsInPLine

    End Function


    Private Function SelectObj(Optional ByVal SingleSelect As Boolean = True, Optional ByVal Message As String = "Sélectionner les objets", Optional ByVal ActivePolyligne As Boolean = True, Optional ByVal ActiveTexte As Boolean = True, Optional ByVal ActiveBlock As Boolean = False) As Collection

        ''Selection of the polylign
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Text As New DBText
        Dim collection_retour As New Collection
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        'Polyligne de périmètre total
        '' Start a transaction

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim acOPrompt As PromptSelectionOptions = New PromptSelectionOptions()

            acOPrompt.MessageForAdding = Message '"Sélectionner le périmètre de la parcelle (poyligne périmètrique)"
            acOPrompt.SingleOnly = SingleSelect

            '' Request for objects to be selected in the drawing area
            Dim acSSPrompt As PromptSelectionResult = acDoc.Editor.GetSelection(acOPrompt)

            '' If the prompt status is OK, objects were selected
            If acSSPrompt.Status = PromptStatus.OK Then
                Dim acSSet As SelectionSet = acSSPrompt.Value
                '' Step through the objects in the selection set

                For Each acSSObj As SelectedObject In acSSet
                    '' Check to make sure a valid SelectedObject object was returned

                    If Not IsDBNull(acSSObj) Then
                        '' Open the selected object for write
                        Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForWrite)
                        If Not IsDBNull(acEnt) Then
                            If TypeName(acEnt) Like "Polyline*" And ActivePolyligne Then
                                collection_retour.Add(acEnt)
                                If SingleSelect = True Then Exit For
                            ElseIf TypeName(acEnt) Like "DBText" And ActiveTexte Then
                                collection_retour.Add(acEnt)
                                If SingleSelect = True Then Exit For
                            ElseIf TypeName(acEnt) Like "*BlockReference*" And ActiveBlock Then
                                collection_retour.Add(acEnt)
                                If SingleSelect = True Then Exit For
                            End If
                        End If
                    End If
                Next
            Else
                If acSSPrompt.Status.ToString = "Cancel" Then
                    collection_retour.Clear()
                Else
                    ed.WriteMessage(vbLf & "Aucun éléments sélectionné ou élément sélectionné non valide")
                    ed.WriteMessage(vbLf & "('ECHAP' pour quitter)")
                    ' collection_retour = SelectPolyline()
                End If
            End If
            ''End of the selection of the polylign 
        End Using
        Return collection_retour
    End Function


    Private Function PerimSearch(Pt3D As Point3d, LayerName As String) As ObjectId


        Dim plID As ObjectId = GetLWPByCentroidAndRect(Pt3D, 1000, LayerName, False)
      
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


    Private Function IntersectBFwithRPL2(ByVal polyid As ObjectId, ByVal IgnoreRPL As Boolean) As Boolean
        Try

            Dim IdsColl As New Collection 'List des ids des RPL (plans)

            Try
                Dim Verts As Point3dCollection = GetPolyVertices(polyid)
                'Set up the typedvalues to filter the selection.
                Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_RPL")}
                Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)



                If ids IsNot Nothing Then
                    For i = 0 To ids.Count - 1
                        IdsColl.Add(ids(i))
                    Next
                End If
            Catch
                Return False
            End Try


            If IdsColl.Count <> 0 Then
                'Get the database and start a transaction
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Loop through the found objects
                    For Each id As ObjectId In IdsColl
                        'Convert this entity to a region
                        Dim objRPLreg As Region = ConvertPolylineToRegion(id)
                        'Convert the supplied polyline to a region
                        Dim objBFReg As Region = ConvertPolylineToRegion(polyid)
                        objBFReg = CType(trans.GetObject(objBFReg.ObjectId, OpenMode.ForWrite), Region)
                        'Subtract this region from the BF region 
                        objBFReg.BooleanOperation(BooleanOperationType.BoolIntersect, objRPLreg)
                        'If the area of the resulting region is greater than 0
                        If objBFReg.Area > 0.01 Then
                            'Create a new BFbyRPL object
                            Dim BFbyRPL As New BFbyRPL
                            'Set the BF property to the polyline from the region
                            BFbyRPL.BF = ConvertRegionToPolyline2(objBFReg, False)
                            'If the user has not ticked the Ignore RPL on the EDT form 
                            If Not IgnoreRPL Then
                                'Call the function to get the RPL from the BF
                                BFbyRPL.RPL = GetRPLFromBF2(BFbyRPL.BF)
                            Else
                                'Otherwise set the RPL property to nothing
                                BFbyRPL.RPL = Nothing
                            End If
                            'Add the object to the global list
                            lstBFbyRPL.Add(BFbyRPL)
                        End If
                    Next
                    trans.Commit()
                End Using
            Else
                'Nothing found within then polyline
                'So create a new BFbyRPL object
                Dim BFbyRPL As New BFbyRPL
                'Set the BF property to the id of the BF polyline
                BFbyRPL.BF = polyid
                'If the user has not ticked the Ignore RPL on the EDT form 
                If Not IgnoreRPL Then
                    'Call the function to get the RPL from the BF
                    BFbyRPL.RPL = GetRPLFromBF2(polyid)
                Else
                    'Otherwise set the RPL property to nothing
                    BFbyRPL.RPL = Nothing
                End If
                'Add the object to the global list
                lstBFbyRPL.Add(BFbyRPL)
            End If


        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("GetRPLFromBF", "Erreur de recherche d'intersection BF - RPL" & vbCrLf & ex.Message, False, 100, 100, "critical")


        End Try
    End Function

    ''' <summary>
    ''' Find holes within the BF
    ''' </summary>
    ''' <param name="BFID"></param>
    ''' <param name="strCom"></param>
    ''' <param name="intPlan"></param>
    ''' <param name="strNum"></param>
    ''' <param name="strID"></param>
    ''' <param name="strCodePlan"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SearchForHolesInBF2(ByVal BFID As ObjectId, ByVal strCom As String, ByVal intPlan As Integer, ByVal strNum As String, _
                                        ByVal strID As String, ByVal strCodePlan As String) As Double
        Dim dblSurfRFTotale As Double = 0
        'Get the vertices of the polyline
        Dim Verts As Point3dCollection = GetPolyVertices(BFID)
        'Set up typedvalues to select polylines on the required layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_BF,MUT_BF")}
        Dim IDs() As ObjectId = SelectWindowPoly(Values, Verts)
        If Not IDs Is Nothing AndAlso Not IDs.Length = 0 Then
            'Store how many objects we found
            Dim intHolesCount As Integer = IDs.Length
            'Open the EDT_Donnees table
            Dim dt As System.Data.DataTable = GetTable("EDT_Donnees", dbName:=rvinfo.DB3system)
            'Loop through the found objects (holes)
            For Each IlotBF As ObjectId In IDs
                IntersectBFwithCS2(IlotBF, strNum, True)
                'Store the CS objects in the database
                WriteCSObjectsToTable2(dt, strCom, intPlan, strNum, strID, strCodePlan)
                lstCS.Clear()
                dblSurfRFTotale += GetPolyArea(IlotBF)
                'Calculation of the ODS surfaces BF island
                IntersectBFwithODS2(IlotBF, True)
                'Store the ODS objects in the database
                WriteODSObjectsToTable2(dt, strCom, intPlan, strNum, strID, strCodePlan)
                lstODS.Clear()
            Next
        End If
        Return dblSurfRFTotale
    End Function

    Private Function IntersectBFwithCS2(ByVal BF As ObjectId, NumBF As String, Optional ByVal booBFhole As Boolean = False, _
                                     Optional ByVal intProgressDivider As Integer = 1, _
                                     Optional ByRef strCurrentLayer As String = "") As Boolean
        Dim strFmtSurf As String = "#,##0.0000"
        Dim ErrID As ObjectId = Nothing
        Dim retVal As Boolean = False
        Try
            'Get the vertices of the BF polyline
            Dim Verts As Point3dCollection = GetPolyVertices(BF)
            'Convert it to a region
            Dim objRegBF As Region = ConvertPolylineToRegion(BF)
            Dim objRegCS As Region = Nothing
            Dim objContPLID As ObjectId = Nothing
            Dim dblSurfTech As Double = 0
            'Set up typedvalues to select POLYLINES
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE")}
            'Make the selection
            Dim IDs() As ObjectId = SelectCrossingPoly(Values, Verts)
            'Check that we found something
            If Not IDs Is Nothing AndAlso Not IDs.Count = 0 Then
                Dim CommitTrans As Boolean = False
                Dim db As Database = HostApplicationServices.WorkingDatabase
                'loop through the found objects inside the BF
                For Each id As ObjectId In IDs
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        'Get the BF polyline
                        Dim BFPl As Polyline = CType(trans.GetObject(BF, OpenMode.ForRead), Polyline)
                        'get the polyline for this object
                        Dim pline As Polyline = CType(trans.GetObject(id, OpenMode.ForRead), Polyline)
                        'Get its layer
                        Dim Lay As String = pline.Layer
                        'Check the layer
                        If (Lay.StartsWith("MO_CS_") Or Lay.StartsWith("MUT_CS_")) _
                                    And Not Lay.Contains("_CS_PTS") And Not Lay.Contains("HACH") Then
                            'Store the layer for use if there is an error
                            strCurrentLayer = Lay
                            'Convert it to a region
                            objRegCS = ConvertPolylineToRegion(id)
                            'Copy the BF region as we don't want to modify the original
                            Dim objRegBFCopy As Region = CType(objRegBF.Clone, Region)
                            'Check that we converted the CS and BF region successfully
                            If objRegCS Is Nothing Then
                                ErrID = id
                            ElseIf objRegBFCopy Is Nothing Then
                                ErrID = objRegBF.ObjectId
                            Else
                                ErrID = Nothing
                            End If
                            Try
                                'Intersect the BF with the CS
                                objRegCS.BooleanOperation(BooleanOperationType.BoolIntersect, objRegBFCopy)
                            Catch ex As Exception
                                If ErrID.IsNull Then
                                    pline.Highlight()
                                    objRegBF.Highlight()
                                Else
                                    Zooming.ZoomToObject(ErrID)
                                    Dim ent As Entity = CType(trans.GetObject(ErrID, OpenMode.ForRead), Entity)
                                    ent.Highlight()
                                End If
                                DeleteGreenTempObjects2()
                                Return False
                            End Try
                            'If the CS object intersects the BF
                            If objRegCS.Area > 0 And Not Lay.EndsWith("CS_INCOMPLET") Then
                                'get the CS region from the database
                                objRegCS = CType(trans.GetObject(objRegCS.ObjectId, OpenMode.ForWrite), Region)
                                'And change its colour to green
                                objRegCS.ColorIndex = 3
                                objRegCS.DowngradeOpen()
                                'Create a new CS object and set its properties
                                Dim oCS As New CS
                                oCS.Type = Lay
                                oCS.Surface = objRegCS.Area
                                oCS.RegCS = objRegCS.ObjectId
                                'If Buiding
                                If Lay.StartsWith("MO_CS_BAT") Or Lay.StartsWith("MUT_CS_BAT") Then
                                    'Get the XData from the polyline
                                    Dim tv() As TypedValue = pline.XData.AsArray
                                    Try
                                        'Store the desctiption and ID from the XData
                                        oCS.Description = tv(3).Value.ToString
                                        oCS.ID = tv(2).Value.ToString
                                    Catch ex As Exception
                                        If tv.Length <= 3 Then ' MODIFICATION 08.05.2014 THA : si erreur > modif 3 en 2 
                                            oCS.Description = tv(2).Value.ToString
                                            oCS.ID = tv(1).Value.ToString
                                        Else
                                            oCS.Description = "?"
                                            oCS.ID = "-1"
                                        End If
                                    End Try
                                    If pline.Area - objRegCS.Area > 0.01 Then
                                        oCS.Divers = Format(pline.Area, strFmtSurf)
                                    End If
                                Else
                                    'Supprimer les surface des bâtiments intérieurs
                                    ' MsgBox("Recherche de :" & Lay)
                                End If
                                oCS.BF_Ilot = booBFhole
                                lstCS.Add(oCS)
                            Else

                            End If
                        End If
                        If Not booBFhole Then
                            Connect.Message("Analyse : BF " & NumBF, "Recherche et analyse des îlots (CS)...", False, 60, 100)
                            frms.Application.DoEvents()
                        End If
                        pline.Dispose()
                        trans.Commit()
                    End Using
                Next
                If Not booBFhole Then
                    Connect.Message("Analyse : BF " & NumBF, "Recherche et analyse des îlots (CS)...", False, 70, 100)
                    frms.Application.DoEvents()
                End If
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Dim BFPl As Polyline = CType(trans.GetObject(BF, OpenMode.ForRead), Polyline)
                    If SearchForHolesInCS2(intProgressDivider, NumBF) Then
                        'Remove temporary objects
                        DeleteGreenTempObjects2()
                        If Not booBFhole Then
                            Connect.Message("Analyse : BF " & NumBF, "Contrôle des surfaces...", False, 75, 100)
                            frms.Application.DoEvents()
                        End If
                        'Calculation of Total / Control
                        dblSurfTech = 0
                        For Each CS In lstCS
                            dblSurfTech += CS.Surface
                        Next
                        'Entire surface of the object (without holes)
                        Dim dblSurfTot As Double = objRegBF.Area
                        Dim dblDelta As Double = dblSurfTech - dblSurfTot
                        If Math.Abs(dblDelta) >= 0.5 Or dblDelta = 0 Then
                            If dblSurfTech < dblSurfTot Then
                                'Try to find the CS polyline that contains this BF by scaling up the BF polyline and selecting within it
                                'We loop around trying to find the polyline each time scaling the BF polyline by larger and larger amounts
                                'until we find it
                                For j As Integer = 10 To 100 Step 10
                                    objContPLID = GetCSthatContainsBF2(BFPl, j)
                                    If Not objContPLID.IsNull Then Exit For
                                Next


                                If Not objContPLID.IsNull Then
                                    Dim objContPl As Polyline = CType(trans.GetObject(objContPLID, OpenMode.ForRead), Polyline)
                                    Dim oCS As New CS
                                    oCS.Type = objContPl.Layer
                                    oCS.Surface = BFPl.Area - dblSurfTech
                                    lstCS.Add(oCS)
                                    'Recalculation of the total
                                    dblSurfTech = 0
                                    For Each CS In lstCS
                                        dblSurfTech += CS.Surface
                                    Next
                                    objContPl.Dispose()
                                    retVal = True
                                Else
                                    retVal = False
                                End If
                            Else

                                Connect.Message("Analyse : BF " & NumBF, "Recherche et analyse des îlots (CS)...", True, 70, 100)

                                'Check surface bâtiment -- ajout TH 09.05.2015
                                Dim dblSurfTechCheck As Double = 0
                                For Each CS In lstCS
                                    dblSurfTechCheck += CS.Surface
                                Next

                                If Math.Round(BFPl.Area, 2) <> Math.Round(dblSurfTechCheck, 2) Then

                                    'Fait le totales des surfaces bâtiment pour soutraire à la surface jardin
                                    Dim SurfBat As Double = 0
                                    Dim SurfAutre As Double = 0
                                    For Each CS In lstCS
                                        If CS.Type = "MO_CS_BAT" Then
                                            SurfBat += CS.Surface
                                        ElseIf CS.Type = "MO_CS_SV_JARDIN" Then
                                        Else 'Autre surface
                                            SurfAutre += CS.Surface
                                        End If
                                    Next
                                    For Each CS In lstCS
                                        If CS.Type = "MO_CS_SV_JARDIN" Then
                                            Dim CSSurfJardin As Double = CS.Surface
                                            CS.Surface = CS.Surface - SurfBat

                                            'Controle si erreur de déduction
                                            dblSurfTechCheck = 0
                                            For Each CS2 In lstCS
                                                dblSurfTechCheck += CS2.Surface
                                            Next
                                            If Math.Round(BFPl.Area, 2) <> Math.Round(dblSurfTechCheck, 2) Then
                                                For Each CS3 In lstCS
                                                    Dim DiffCheck As Double = Math.Round(BFPl.Area - SurfBat - SurfAutre, 2)
                                                    If Math.Round(CSSurfJardin - DiffCheck, 2) = Math.Round(CS3.Surface, 2) And CS3.Type = "MO_CS_BAT" Then                                                        'Batiment à déduire 
                                                        CS.Surface = CSSurfJardin - CS3.Surface
                                                        Exit For
                                                    End If
                                                Next
                                                Exit For
                                            Else
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If


                                'Refait un check finale
                                dblSurfTechCheck = 0
                                For Each CS In lstCS
                                    dblSurfTechCheck += CS.Surface
                                Next

                                '''TEST -------------------------------
                                ''Dim StrSurf As String = ""
                                ''For Each CS In lstCS
                                ''    dblSurfTechCheck += CS.Surface
                                ''    StrSurf += CS.BF_Ilot & " : " & CS.Surface & vbCrLf
                                ''Next
                                ''MsgBox(StrSurf)

                                If Math.Round(BFPl.Area, 2) <> Math.Round(dblSurfTechCheck, 2) Then

                                    Dim ConErr As New Revo.connect
                                    Dim Parc As String = ""
                                    If NumBF <> "" Then Parc = "(BF " & NumBF & ")"
                                    ConErr.Message("Erreur de l'état descriptif " & Parc & ": Vérifier la topologie des éléments 'CS'", _
                                                           "La surface totale du bien-fonds est de " _
                                                         & Format(dblSurfTot, "#,##0.00") & "m²" & vbCrLf & _
                                                         "alors que la somme des couverture du sol fait " _
                                                         & Format(dblSurfTech, "#,##0.00") & "m² !" _
                                                        , False, 99, 100, "critical")

                                    'frms.MessageBox.Show("Erreur à l'établissement de l'état descriptif:" & _
                                    '                     vbCrLf & "La surface totale de la parcelle selon la limite de bien-fonds est de " _
                                    '                     & Format(dblSurfTot, "#,##0.00") & "m²" & vbCrLf & _
                                    '                     "alors que la somme des types de couverture du sol fait " _
                                    '                     & Format(dblSurfTech, "#,##0.00") & "m² !" & vbCrLf & _
                                    '                     "Vérifier la topologie des éléments 'CS' et qu'il n'y a pas de doublons ou de superpositions.", _
                                    '                     "Calcul impossible", Windows.Forms.MessageBoxButtons.OK, _
                                    '                     Windows.Forms.MessageBoxIcon.Error)

                                    'retVal = True 'uniquement pour les test
                                    retVal = False
                                Else
                                    retVal = True
                                End If

                            End If
                        Else
                            retVal = True
                        End If
                    Else

                    End If
                    DeleteGreenTempObjects2()
                    trans.Commit()
                End Using
            Else 'Nothing found within the BF polyline
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    Dim BFPl As Polyline = CType(trans.GetObject(BF, OpenMode.ForRead), Polyline)
                    'Find the CS object that this BF is within
                    objContPLID = GetCSthatContainsBF2(BFPl, 20)
                    If Not objContPLID.IsNull Then
                        Dim objContPL As Polyline = CType(trans.GetObject(objContPLID, OpenMode.ForRead), Polyline)
                        Dim oCS As New CS
                        oCS.Type = objContPL.Layer
                        oCS.Surface = objContPL.Area
                        lstCS.Add(oCS)
                        'Recalculation of the total
                        dblSurfTech = 0
                        For Each CS In lstCS
                            dblSurfTech += CS.Surface
                        Next
                        objContPL.Dispose()
                        BFPl.Dispose()
                        Return True
                    Else
                        Return False
                    End If
                End Using

            End If
        Catch ex As Exception


            Dim ConnErr As New Revo.connect
            ConnErr.Message("IntersectBFwithCS : BF " & NumBF, "Erreur de recherche d'intersection BF - CS" & vbCrLf & ex.Message, False, 100, 100, "critical")

        End Try
        Return retVal
    End Function

    Private Function WriteCSObjectsToTable2(ByRef dt As System.Data.DataTable, ByVal strCom As String, ByVal intPlan As Integer, _
                                         ByVal strNum As String, ByVal strId As String, ByVal strCodePlan As String) As Boolean
        Try

        For Each oCS As CS In lstCS
            Dim dr As DataRow = dt.NewRow
            dr("Theme") = 2
            dr("NoCom") = strCom
            dr("NoPlan") = intPlan
            dr("NoBF") = strNum.Replace("DP", "").Trim
            dr("DP") = strNum.Contains("DP")
            dr("Immeuble") = strId
            dr("CodePlan") = strCodePlan
            dr("JC_Layer") = oCS.Type
            If oCS.Type = "MO_CS_BAT" Or oCS.Type = "MUT_CS_BAT" Then
                dr("ID_objet") = oCS.ID
                dr("Designation_objet") = oCS.Description.Replace(" ", "_")
                dr("Info_objet") = oCS.Divers
            End If
            dr("Surface") = oCS.Surface
            dt.Rows.Add(dr)
            Next

        Catch 'ex As Exception
            ' MsgBox("Erreur d'enregistrment de la BD")
        End Try

        UpdateTable(dt, RVinfo.DB3system)
    End Function

    Private Function IntersectBFwithODS2(ByVal BFID As ObjectId, Optional ByVal booBFHole As Boolean = False, _
                                      Optional ByVal intProgressDivider As Integer = 1) As Boolean
        If Not booBFHole Then
            Connect.Message("Analyse", "Calcul des intersections spatiales (BF / ODS)", False, 80, 100)
            frms.Application.DoEvents()
        End If
        Try
            Dim strFmtSurf As String = "#,##0.0000"
            'Get the vertices of the BF polyline
            Dim verts As Point3dCollection = GetPolyVertices(BFID)
            'Convert the BF polyline to a region 
            Dim objRegBF As Region = ConvertPolylineToRegion(BFID)
            'Select all ODS polylines that intersect (or are inside) the BF polyline
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE")}
            Dim ids() As ObjectId = SelectCrossingPoly(Values, verts)
            If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                'Get the database and start a transaction
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Loop through the found objects
                    For Each objLWPID As ObjectId In ids
                        'Get the polyline from the database
                        Dim objLWP As Polyline = CType(trans.GetObject(objLWPID, OpenMode.ForRead), Polyline)
                        'get the polyline's layer
                        Dim lay As String = objLWP.Layer
                        'Check that it's on one of the layers we want
                        If lay.StartsWith("MO_ODS_BATSOUT") Or lay.StartsWith("MO_ODS_COUVINDEP") Or _
                                        lay.StartsWith("MUT_ODS_BATSOUT") Or lay.StartsWith("MUT_ODS_COUVINDEP") Then
                            'Convert ODS polyline to region
                            Dim objRegODS As Region = ConvertPolylineToRegion(objLWPID)
                            'Chekc that the polyline converted to a region ok
                            If Not objRegODS Is Nothing Then
                                'Copy the BF region and intersect this region with the copy  
                                Dim objRegBFCopy = CType(objRegBF.Clone, Region)
                                objRegODS.BooleanOperation(BooleanOperationType.BoolIntersect, objRegBFCopy)
                                'if the resulting region has an area greater than 0
                                If objRegODS.Area > 0 Then
                                    'Create a new ODS object
                                    Dim oODS As New ODS
                                    'And set its properties
                                    oODS.Type = objLWP.Layer
                                    oODS.Surface = objRegODS.Area
                                    oODS.ODSRegID = objRegODS.ObjectId
                                    'Get the XData from the polylne
                                    Try
                                        Dim tv() As TypedValue = objLWP.XData.AsArray

                                        If tv.Length = 4 Then
                                            oODS.Description = tv(3).Value.ToString
                                            oODS.ID = tv(2).Value.ToString
                                        Else
                                            oODS.Description = tv(1).Value.ToString
                                            oODS.ID = tv(0).Value.ToString
                                        End If
                                      
                                    Catch ex As Exception

                                        Try 'Recherche le numéro de bat sout

                                            'Layer : MO_ODS_BATS_NUM
                                            Dim Verts2 As Point3dCollection = GetPolyVertices(objLWP.ObjectId)
                                            Dim Values2() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayerName, "MO_ODS_BATS_NUM")}
                                            Dim IDs2() As ObjectId = SelectWindowPoly(Values2, Verts2)

                                            If IDs2 IsNot Nothing Then
                                                oODS.ID = GetBlockAttributeTrans(IDs2(0), "NUMERO", trans)
                                                oODS.Description = GetBlockAttributeTrans(IDs2(0), "DESIGNATION", trans)
                                            Else
                                                oODS.Description = "-"
                                                oODS.ID = "-"
                                            End If

                                        Catch ex1 As Exception
                                            oODS.Description = "-"
                                            oODS.ID = "-"
                                        End Try
                                    End Try
                                    If objLWP.Area - objRegODS.Area > 0.01 Then
                                        oODS.Divers = Format(objLWP.Area, strFmtSurf)
                                    End If
                                    lstODS.Add(oODS)
                                End If
                                objRegODS.Dispose()
                            End If
                            If Not booBFHole Then
                                Connect.Message("Analyse", "Calcul des intersections spatiales (BF / ODS)", False, intProgressDivider, ids.Count)
                                frms.Application.DoEvents()
                            End If
                        End If
                    Next
                End Using
            End If
            objRegBF.Dispose()
        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("IntersectBFwithODS2", "Erreur de recherche d'intersection BF - ODS" & vbCrLf & ex.Message, False, 100, 100, "critical")

        End Try
    End Function

    ''' <summary>
    ''' Gets a value from the attribute in the block
    ''' </summary>
    ''' <param name="blkId">The ID of the block to get the attribute from</param>
    ''' <param name="strAttName">The name of the attribute to get the value of</param>
    ''' <returns>The value from the attribute</returns>
    ''' <remarks></remarks>
    Private Function GetBlockAttributeTrans(ByVal blkId As ObjectId, ByVal strAttName As String, trans As Transaction) As String
        Dim retVal As String = String.Empty
        Try
            'Get the database and start a transaction
            'Get the block reference from the batabase
            Dim blkRef As BlockReference = CType(trans.GetObject(blkId, OpenMode.ForRead), BlockReference)
            'Get the attribute collection 
            Dim attCol As AttributeCollection = blkRef.AttributeCollection
            'Loop through the attributes in the collection
            For Each attid As ObjectId In attCol
                'Get the attribute reference from the database
                Dim attRef As AttributeReference = CType(trans.GetObject(attid, OpenMode.ForRead), AttributeReference)
                'Check if this is the attribute we're looking for
                If attRef.Tag.ToUpper = strAttName.ToUpper Then
                    'Store the value from the attribute
                    retVal = attRef.TextString
                    attRef.Dispose()
                    Exit For
                End If
            Next
            'Dispose the blockreference as we don't need it anymore
            blkRef.Dispose()

        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("GetBlockAttributeTrans", "Erreur de recherche d'attribut" & vbCrLf & ex.Message, False, 100, 100, "critical")


        End Try
        Return retVal
    End Function


    ''' <summary>
    ''' Finds the CS that contains the BF
    ''' </summary>
    ''' <param name="objPoly">The BF polyline</param>
    ''' <param name="dblFacteurRecherche"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function GetCSthatContainsBF2(ByVal objPoly As Polyline, ByVal dblFacteurRecherche As Integer) As ObjectId
        Try
            Dim retVal As ObjectId = Nothing 'Return value
            'Get the vertices of the BF polyline after scaling it up 
            Dim Verts As Point3dCollection = GetScaledPolyVertices(objPoly.ObjectId, dblFacteurRecherche, True)
            'Set up typedvalue to select polylines
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_CS*,MUT_CS_*")}
            'Set up typedvalues to select polylines on the BF polyline layer
            Dim Values2() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, objPoly.Layer)}
            'Select all polylines within the scaled up FB polyline
            Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)
            'Check that we found something
            If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                'Get the database and start a transaction 
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Loop through the objectids selecting polylines within each one on the BF layer
                    For Each objPLCSID As ObjectId In ids
                        'Get the polyline
                        Dim objPLCS As Polyline = CType(trans.GetObject(objPLCSID, OpenMode.ForRead), Polyline)
                        'Get the layer that this polyline is on
                        Dim Lay As String = objPLCS.Layer
                        'Check that it's on a layer that we're interested in 'Ajout de TH 09.05.2016
                        If Lay <> "MO_CS_BAT" And _
                            (Lay.StartsWith("MO_CS_") And Not Lay.StartsWith("MO_CS_PTS") And Not Lay.Contains("HACH")) Or _
                            (Lay.StartsWith("MUT_CS") And Not Lay.StartsWith("MUT_CS_PTS") And Not Lay.Contains("HACH")) Then
                            'Get the vertices of this polyline
                            Verts = GetPolyVertices(objPLCSID)
                            'Select all polylines on the BF layer within this polyline
                            Dim ids2() As ObjectId = SelectCrossingPoly(Values2, Verts)
                            'Check that we found something
                            If Not ids2 Is Nothing AndAlso Not ids2.Count = 0 Then
                                'Loop through the found objects
                                For Each objPLbf As ObjectId In ids2
                                    'If this polyline is the BF polyline
                                    If objPLbf.Equals(objPoly.ObjectId) Then
                                        'Set the return value to the containing polyline
                                        retVal = objPLCSID
                                        'Dispose of this CS polyline as we're done with it
                                        objPLCS.Dispose()
                                        'Return the objectid of the CS polyline
                                        Return retVal
                                    End If
                                Next
                            End If
                        End If
                    Next
                End Using
            End If
        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("GetCSthatContainsBF2", "Erreur de recherche d'ilôts BF dans CS" & vbCrLf & ex.Message, False, 100, 100, "critical")

          End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Finds holes inside the CS objects and subtracts the area of these holes from the area of the CS object
    ''' </summary>
    ''' <param name="intProgressDivider"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function SearchForHolesInCS2(ByVal intProgressDivider As Integer, NumParc As String) As Boolean
        Try
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'Loop through the CS objects in the list
            For Each oCS In lstCS
                If oCS.BF_Ilot = False Then
                    'Get the CS Region
                    Dim objCSreg As Region = CType(oCS.RegCS.GetObject(OpenMode.ForRead), Region)
                    If objCSreg.Area > 0 Then
                        'Convert the region to a polyline retaining the original region for the following intersections
                        'Dim objCSpoly As ObjectId = ConvertRegionToPolyline(objCSreg, True)
                        Dim objCSpoly As ObjectId = ConvertRegionToPolyline2(objCSreg, True)

                        'Check that the region converted to a polyline OK
                        If objCSpoly.IsNull Then
                            'The object could not be converted
                        Else
                            'Get the vertices of the polyline
                            Dim Verts As Point3dCollection = GetPolyVertices(objCSpoly)
                            If Verts.Count > 2 Then 'Ignore lines
                                'Set up the values to filter green regions
                                Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "REGION"), _
                                                                                      New TypedValue(DxfCode.LayerName, "MO_BF_INCOMPLET"), _
                                                                                      New TypedValue(DxfCode.Color, 3)}
                                'Select the green regions inside the CS polyline
                                Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)
                                'Cehck that we found some
                                If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                                    'Copy the CS region
                                    Dim objRefReg As Region = CType(objCSreg.Clone, Region)
                                    'Loop through the found items sunbtracting each one from the copy of the CS region
                                    For Each id As ObjectId In ids
                                        Dim objSubreg As Region = CType(id.GetObject(OpenMode.ForRead), Region)
                                        'Check that this sub region is not the CS region
                                        If objCSreg.ObjectId <> objSubreg.ObjectId And objSubreg.Area > 0 Then
                                            objRefReg.BooleanOperation(BooleanOperationType.BoolSubtract, CType(objSubreg.Clone, Region))
                                        End If
                                        objSubreg.Dispose()
                                    Next
                                    If objRefReg.Area > 0 Then
                                        oCS.Surface = objRefReg.Area
                                    End If

                                    objRefReg.Dispose()
                                End If
                            End If
                        End If
                        objCSreg.Color = Color.FromColorIndex(ColorMethod.None, 3)
                    End If
                    Connect.Message("Analyse : BF " & NumParc, "Recherche d'ilôts", False, intProgressDivider, lstCS.Count)

                End If
            Next
            Return True
        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("SearchForHolesInCS2 : BF " & NumParc, "Erreur de recherche d'ilôts CS" & vbCrLf & ex.Message, False, 100, 100, "critical")

        End Try

    End Function

    Private Function GetDDPInsideBF2(ByVal polyID As ObjectId) As Boolean
        lstDDP = New List(Of DDP)
        Dim Verts As Point3dCollection = GetPolyVertices(polyID)
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_BF_DDP,MUT_BF_DDP")}
        Dim Ids() As ObjectId = SelectCrossingPoly(Values, Verts)
        If Not Ids Is Nothing AndAlso Ids.Length > 0 Then
            'Get the database and start a transaction
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Loop through the found objects
                For Each id As ObjectId In Ids
                    Dim pLine As Polyline = CType(trans.GetObject(id, OpenMode.ForRead), Polyline)
                    'Convert this entity to a region
                    Dim objDDPreg As Region = ConvertPolylineToRegion(id)
                    'Convert the supplied polyline to a region
                    Dim objBFReg As Region = ConvertPolylineToRegion(polyID)
                    'Subtract this region from the BF region 
                    objDDPreg.BooleanOperation(BooleanOperationType.BoolIntersect, objBFReg)
                    'If the area of the resulting region is greater than 0
                    If objDDPreg.Area > 0.01 Then
                        'Create a new DDP object
                        Dim objDDP As New DDP
                        'Set the PolyID property 
                        objDDP.objPolyID = id
                        'And the area property 
                        objDDP.Surface = objDDPreg.Area
                        'Set the Partiel flag
                        objDDP.Partiel = (Math.Abs(objDDPreg.Area - pLine.Area) > 0.01)
                        'try to get the block from the polyline
                        Dim ddpBlockID As ObjectId = GetBlockfromLWP2(id, pLine.Layer & "_IMMEUBLE", True)
                        'If we found a block 
                        If Not ddpBlockID.IsNull Then
                            'then get the value of the NUMERO attribute and store it
                            objDDP.Numero = GetBlockAttribute(ddpBlockID, "NUMERO")
                        Else
                            'Otherwise set set the Numero property to a non breaking space
                            objDDP.Numero = "&nbsp;"
                        End If
                    End If
                Next
            End Using
        End If
    End Function
    ''' <summary>
    ''' Delete the temporary green regions
    ''' </summary>
    ''' <returns>True if there were any regions to delete and therefore the outer transaction needs to be committed too</returns>
    ''' <remarks></remarks>
    Private Function DeleteGreenTempObjects2() As Boolean
        Try
            Dim retVal As Boolean = False
            'Set up the typedvalues to select green regions on the desired layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "REGION"), _
                                        New TypedValue(DxfCode.LayerName, "MO_BF_INCOMPLET"), _
                                        New TypedValue(DxfCode.Color, 3)}
            'Selct all items that pass the filter
            Dim pres As PromptSelectionResult = SelectAllItems(Values)
            'Check whether we found anything
            If pres.Status = PromptStatus.OK Then
                'Set the flag to say that the outer transaction needs to be committed
                retVal = True
                'Get the databse abnd start a transaction
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Loop through the found ids
                    For Each id As ObjectId In pres.Value.GetObjectIds
                        'Get the entity from the database
                        Dim ent As Entity = CType(trans.GetObject(id, OpenMode.ForWrite), Entity)
                        'and delete it
                        ent.Erase()
                    Next
                    'save the changes (if this function was called from within another transaction
                    'then the outer transaction needs to be committed too)
                    trans.Commit()
                End Using
            End If
            Return retVal
        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("DeleteGreenTempObjects2", "Erreur de suppression des objets temporaires (vert)" & vbCrLf & ex.Message, False, 100, 100, "critical")

        End Try
    End Function

    ''' <summary>
    ''' Gets the RPL from the BF
    ''' </summary>
    ''' <param name="polyID">The objectid of the BF polyline</param>
    ''' <returns>The objectid of the RPL block</returns>
    ''' <remarks></remarks>
    Private Function GetRPLFromBF2(ByVal polyID As ObjectId) As ObjectId
        Dim objBlockPlan As ObjectId = Nothing
        Try
            'get the vertices of the polyline
            Dim Verts As Point3dCollection = GetPolyVertices(polyID)
            'Set up typedvalues to select polylines on the desired layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_RPL")}
            Dim IDs() As ObjectId = SelectCrossingPoly(Values, Verts)
            'Convert the BF polyline to a region
            Dim objRegBF As Region = ConvertPolylineToRegion(polyID)
            Dim objRegPL As Region = Nothing
            If Not IDs Is Nothing AndAlso Not IDs.Length = 0 Then
                'Loop through the found polylines converting them to regions and subtracting the BF polyline from them
                For Each id As ObjectId In IDs
                    objRegPL = ConvertPolylineToRegion(id)
                    'Store the area of the BF poly
                    Dim dblSurfInit As Double = objRegPL.Area
                    'Subtract the BF from this region
                    objRegPL.BooleanOperation(BooleanOperationType.BoolSubtract, CType(objRegBF.Clone, Region))
                    'Check the size of the resulting region
                    If Math.Abs(dblSurfInit - objRegPL.Area) > 0.01 Then
                        'Get the block from this polyline
                        objBlockPlan = GetBlockfromLWP2(id, "MO_RPL_PLAN", True)
                    End If
                Next
                If Not objRegPL Is Nothing Then objRegPL.Dispose()
                objRegBF.Dispose()
            Else
                'Select everything that passes the filter
                Dim pres As PromptSelectionResult = SelectAllItems(Values)
                'Check that we found something
                If pres.Status = PromptStatus.OK Then
                    'Get the database and start a transaction
                    Dim db As Database = HostApplicationServices.WorkingDatabase
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        'Loop through the found objects
                        For Each id As ObjectId In pres.Value.GetObjectIds
                            'Get the object from the database
                            Dim pline As Polyline = CType(trans.GetObject(polyID, OpenMode.ForRead), Polyline) 'Modif TH 27.02.2016
                            'Dim pline As Polyline = CType(trans.GetObject(id, OpenMode.ForRead), Polyline) 

                            'Check whether this object is within the BF polyline
                            If IsObjectInsidePolyline(pline, id) Then 'Modif TH 27.02.2016
                                'If IsObjectInsidePolyline(pline, polyID) Then
                                'Get the block from the polyline
                                objBlockPlan = GetBlockfromLWP2(id, "MO_RPL_PLAN", True)
                                pline.Dispose()
                                Exit For
                            End If
                        Next
                    End Using
                Else
                    objBlockPlan = Nothing
                End If
            End If
        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("GetRPLFromBF2", "Erreur de recherche RPL-BF" & vbCrLf & ex.Message, False, 100, 100, "critical")

        End Try
        Return objBlockPlan
    End Function

    Private Function GetBlockfromLWP2(ByVal polyID As ObjectId, ByVal strLayer As String, Optional ByVal booWaitFormLoaded As Boolean = False) As ObjectId
        Dim retVal As ObjectId = Nothing 'Return value
        Try
            'Get the vertices of the polyline
            Dim Verts As Point3dCollection = GetPolyVertices(polyID)
            'Set up typedvalues to select blocks on the correct layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayerName, strLayer)}
            Dim IDs() As ObjectId = SelectCrossingPoly(Values, Verts)
            If Not IDs Is Nothing AndAlso IDs.Length = 1 Then
                'Return the objectid
                Return IDs(0)
            Else
                'Found no objects or more than one
                'So get the database
                Dim db As Database = HostApplicationServices.WorkingDatabase
                'Start a transaction
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Get the polyline from the database
                    Dim pline As Polyline = CType(trans.GetObject(polyID, OpenMode.ForRead, True), Polyline)
                    'Zoom to the extents of the polyline
                    Zooming.ZoomToObject(polyID)
                    'And highlight the polyline
                    pline.Highlight()
                    'If the wait form (progress form) is loaded then hide it
                    'If booWaitFormLoaded Then Wait.Hide() : frms.Application.DoEvents()
                    'Call the function to ask the user to select the block manually
                    retVal = GetBlockBySelection(strLayer)
                    'If the wait form is loaded then show it again
                    ' If booWaitFormLoaded Then Wait.Show()
                    'unhighlight the polyline
                    pline.Unhighlight()
                    'If a block has been selected manually then we assume that the zoom has been modified and therefore need to zoom again
                    If Not retVal.IsNull Then
                        Zooming.ZoomExtents()
                    Else
                        Zooming.ZoomPrevious() 'otherwise only one previous zoom
                    End If
                End Using
            End If
            frms.Application.DoEvents()
        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("GetBlockfromLWP2", "Erreur de recherche LWP" & vbCrLf & ex.Message, False, 100, 100, "critical")

        End Try
        Return retVal 'Return the selected id
    End Function





    ''' <summary>
    ''' Creates the HTML document to display the results
    ''' </summary>
    ''' <param name="strBFID">The ID of the BF polyline</param>
    ''' <param name="booOpenHTML">Optional boolean to open the document once created or not</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Private Function CreateEDTHtml2(ByVal strBFID As String, NumBF As String, Optional ByVal booOpenHTML As Boolean = True) As Boolean



        Try

            'MsgBox(My.Application.Culture.ToString) 'affiche 'fr-FR'
            My.Application.ChangeCulture("fr-CH")
            'MsgBox(My.Application.Culture.ToString) 'affiche 'fr-FR'

            Dim dblSurfTotale As Double = 0 ' Int64 MODIF TH / 10.01.2012
            'Dim dt As System.Data.DataTable = GetRecordset("EDT_CS", "SELECT SUM(Surface) FROM EDT_CS WHERE Immeuble='" & strBFID & "'", "System\System.db3")
            Dim dt As System.Data.DataTable = GetRecordsetUsingWrapper("EDT_CS", "SELECT SUM(Surface) FROM EDT_CS WHERE Immeuble='" & strBFID & "'", RVinfo.DB3system)
            If Not dt Is Nothing AndAlso Not dt.Rows.Count = 0 Then

                If Convert.ToString(dt.Rows(0)(0)) <> "" Then
                    dblSurfTotale = Convert.ToDouble(dt.Rows(0)(0)) 'MODIF TH / 10.01.2012
                End If

            Else
                Return False
            End If
            Dim strFmtSurf As String = "#,##0.0000"
            Dim strReport As String = RVinfo.EDTFolder ' LireBaseRegistre(2, "Software\SWWS\JourCAD\Settings", "EDTFolder")
            If strReport = String.Empty Then
                strReport = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)
            End If
            If Not strReport.EndsWith("\") Then strReport = strReport & "\"
            strReport = strReport & strBFID.Replace(" / ", "_") & "_1.htm"

            'Chargement du Templates
            Dim Connect As New Revo.connect
            Dim HTMLtemplate As String = ""
            If System.IO.File.Exists(RVinfo.EDTtemplate) Then
                Try
                    Dim sr As System.IO.StreamReader = New System.IO.StreamReader(RVinfo.EDTtemplate, System.Text.Encoding.Default)
                    While Not sr.EndOfStream()
                        HTMLtemplate += sr.ReadLine() & vbCrLf   '--- Traitement du fichier ligne par ligne
                    End While
                    sr.Close()
                Catch ex As Exception
                    Throw ex
                    Connect.RevoLog(Connect.DateLog & "Write EDT" & vbTab & False & vbTab & "Erreur lecture: " & ex.Message & vbTab & RVinfo.EDTtemplate)
                End Try
            Else
                Connect.RevoLog(Connect.DateLog & "Write EDT" & vbTab & False & vbTab & "Erreur lecture: " & "fichier inexistant" & vbTab & RVinfo.EDTtemplate)
            End If

            'Replace Variable HTML
            ' [[REVO]]
            HTMLtemplate = Replace(HTMLtemplate, "[[REVO]]", "<a href=""http://platform5rd.com"" target=""_blank"" style=""color:#000; font-style:normal ; font:Arial, Helvetica, sans-serif; font-size:10px"">" & "REVO " & RVinfo.xVersion & "</a> ")
            HTMLtemplate = Replace(HTMLtemplate, "[[User]]", GetCurrentWinLogin())              ' [[User]]
            HTMLtemplate = Replace(HTMLtemplate, "[[Date]]", Date.Today.ToLongDateString())     ' [[Date]] 
            HTMLtemplate = Replace(HTMLtemplate, "[[Hour]]", TimeOfDay.ToShortTimeString())     ' [[Hour]]

            HTMLcontent = ""         'Reset content 

            'Write the header
            'Voir Templates <<  Header

            'dt = GetRecordset("EDT_plans", "SELECT * FROM EDT_plans WHERE Immeuble='" & strBFID & "'", "System\System.db3")
            dt = GetRecordsetUsingWrapper("EDT_plans", "SELECT * FROM EDT_plans WHERE Immeuble='" & strBFID & "'", RVinfo.DB3system)
            If Not dt Is Nothing AndAlso Not dt.Rows.Count = 0 Then
                Dim strBF As String = dt.Rows(0)("NoBF").ToString
                If Convert.ToBoolean(dt.Rows(0)("DP")) = True Then strBF = "DP" & strBF

                'District, parcelle, commune, surface
                HTMLwriter2("    <tr>")
                HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                HTMLwriter2("        <tr>")
                HTMLwriter2("          <td width=""50%"">District : <b><span class=""Titre2"">" & dt.Rows(0)("NomDistrict").ToString & "</span></b></td>")
                HTMLwriter2("          <td width=""25%"" align=""center""><b>Parcelle:</b></td>")
                HTMLwriter2("          <td width=""25%""><b><span class=""Titre2"">" & strBF & "</span></b></td>")
                HTMLwriter2("        </tr>")
                HTMLwriter2("        <tr>")
                HTMLwriter2("          <td colsplan=""3"">&nbsp;</td>")
                HTMLwriter2("        </tr>")
                HTMLwriter2("        <tr>")
                HTMLwriter2("          <td width=""50%"">Commune : <b><span class=""Titre2"">" & dt.Rows(0)("NoCom").ToString & " (" & dt.Rows(0)("NoComCH").ToString & ") " & dt.Rows(0)("NomCommune").ToString & "</span></b></td>")
                HTMLwriter2("          <td width=""25%"">Surface technique [m&sup2;]:</td>")
                HTMLwriter2("          <td width=""25%"">" & Format(dblSurfTotale, strFmtSurf) & "</td>")
                HTMLwriter2("        </tr>")
                HTMLwriter2("        <tr>")
                HTMLwriter2("          <td width=""50%"">&nbsp;</td>")
                HTMLwriter2("          <td width=""25%"">Surface arrondie [m&sup2;]:</td>")
                HTMLwriter2("          <td width=""25%"">" & Format(dblSurfTotale, "#,##0") & "</td>")
                HTMLwriter2("        </tr>")
                HTMLwriter2("        <tr>")
                HTMLwriter2("          <td colsplan=""3"">&nbsp;</td>")
                HTMLwriter2("        </tr>")
                'Local Names
                HTMLwriter2("      </table></td>")
                HTMLwriter2("    </tr>")
                'Surfaces by plan
                HTMLwriter2("    <tr>")
                HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                HTMLwriter2("        <tr>")
                HTMLwriter2("          <td width=""5%"" align=""right""><b>Plan</b></td>")
                HTMLwriter2("          <td width=""7%"">&nbsp;</td>")
                HTMLwriter2("          <td width=""28%""><b>Type de mensuration</b></td>")
                HTMLwriter2("          <td width=""12%"" align=""right""><b>Code plan</b></td>")
                HTMLwriter2("          <td width=""28%"">&nbsp;</td>")
                HTMLwriter2("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
                HTMLwriter2("        </tr>")
                HTMLwriter2("        <tr>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td align=""center""><span class=""Legende"">[m&sup2;]</span></td>")
                HTMLwriter2("        </tr>")
                'Loop for each plan
                Dim intPlan As Integer, strCodePlan As String, strTypeMens As String
                Dim dblSomme As Double = 0
                For i As Integer = 0 To dt.Rows.Count - 1
                    intPlan = Convert.ToInt32(dt.Rows(i)("NoPlan"))
                    strCodePlan = dt.Rows(i)("CodePlan").ToString
                    strTypeMens = Replace(CodePlanToTypeMens(intPlan, strCodePlan), "é", "&eacute;")
                    HTMLwriter2("        <tr>")
                    If intPlan > 9000 Then
                        HTMLwriter2("          <td align=""right"">&nbsp;</td>")
                        HTMLwriter2("          <td>&nbsp;</td>")
                        HTMLwriter2("          <td>&nbsp;</td>")
                        HTMLwriter2("          <td align=""right"">&nbsp;</td>")
                    Else
                        HTMLwriter2("          <td align=""right"">" & intPlan & "</td>")
                        HTMLwriter2("          <td>&nbsp;</td>")
                        HTMLwriter2("          <td>" & strTypeMens & "</td>")
                        HTMLwriter2("          <td align=""right"">" & strCodePlan & "</td>")
                    End If
                    HTMLwriter2("          <td>&nbsp;</td>")
                    HTMLwriter2("          <td align=""right"">" & Format(Convert.ToDouble(dt.Rows(i)("Surface")), strFmtSurf) & "</td>")
                    HTMLwriter2("        </tr>")
                    dblSomme = dblSomme + Convert.ToDouble(dt.Rows(i)("Surface"))
                Next
                'Total
                HTMLwriter2("        <tr>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td><b>Total</b></td>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td>&nbsp;</td>")
                HTMLwriter2("          <td align=""right"" style=""border-top-color:#000000; border-top-style:solid; border-top-width:3px;""><b>" & Format(dblSomme, strFmtSurf) & "</b></td>")
                HTMLwriter2("        </tr>")
                HTMLwriter2("      </table></td>")
                HTMLwriter2("    </tr>")

                HTMLwriter2("    <tr>")
                HTMLwriter2("      <td height=""30"">&nbsp;</td>")
                HTMLwriter2("    </tr>")
                'DDP
                If lstDDP.Count > 0 Then
                    'Separator
                    HTMLwriter2("    <tr>")
                    HTMLwriter2("      <td height=""4px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000; border-top-style:solid; border-top-width:2px; border-top-color:#000000;""></td>")
                    HTMLwriter2("    </tr>")
                    HTMLwriter2("    <tr>")
                    HTMLwriter2("      <td>&nbsp;</td>")
                    HTMLwriter2("    </tr>")
                    If lstDDP.Count = 1 Then
                        HTMLwriter2("    <tr>")
                        HTMLwriter2("      <td>Parcelle contenant le DDP " & lstDDP(0).Numero & "</td>")
                        HTMLwriter2("    </tr>")
                    Else
                        HTMLwriter2("    <tr>")
                        HTMLwriter2("      <td>Parcelle contenant les DDP " & lstDDP(0).Numero)
                        For i = 1 To lstDDP.Count - 1
                            HTMLwriter2(", " & lstDDP(i).Numero)
                        Next
                        HTMLwriter2("</td>")
                        HTMLwriter2("    </tr>")
                    End If
                    HTMLwriter2("    <tr>")
                    HTMLwriter2("      <td>&nbsp;</td>")
                    HTMLwriter2("    </tr >")
                End If
                WriteEDTCSTitle2()
                'dt = GetRecordset("EDT_CS", "SELECT * FROM EDT_CS WHERE Immeuble='" & strBFID & "'", "System\system.db3")
                dt = GetRecordsetUsingWrapper("EDT_CS", "SELECT * FROM EDT_CS WHERE Immeuble='" & strBFID & "'", RVinfo.DB3system)
                If Not dt Is Nothing AndAlso Not dt.Rows.Count = 0 Then
                    Dim strNat As String = ""

                    Dim strGenre As String = CStr(IIf(dt.Rows(0)("Categorie") Is Nothing, "", dt.Rows(0)("Categorie").ToString))
                    intPlan = Convert.ToInt32(dt.Rows(0)("NoPlan"))
                    Dim dblSommeGenre As Double = 0
                    Dim dblSommeNat As Double = 0.0
                    dblSomme = 0
                    WriteEDTCSHeader2(strGenre = "Bâtiment")
                    'Start of Table
                    HTMLwriter2("    <tr>")
                    HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                    Dim rowCounter As Integer = 0
                    Do
                        If rowCounter = dt.Rows.Count And strGenre = "Bâtiment" Then
                            Exit Do
                        ElseIf rowCounter = dt.Rows.Count And strGenre <> "Bâtiment" Then
                            GoTo TotalGenre
                        End If
                        If rowCounter = dt.Rows.Count OrElse strGenre <> dt.Rows(rowCounter)("Categorie").ToString Then
TotalGenre:

                            If strGenre <> "Bâtiment" Then
                                'Total of previous type
                                HTMLwriter2("        <tr>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td><b>Total du genre " & strGenre & "</b></td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td align=""right"" style=""border-top-color:#000000; border-top-style:solid; border-top-width:3px;""><b>" & Format(dblSommeGenre, strFmtSurf) & "</b></td>")
                                HTMLwriter2("        </tr>")
                                HTMLwriter2("        <tr>")
                                HTMLwriter2("          <td colspan=""7"">&nbsp;</td>")
                                HTMLwriter2("        </tr>")
                            Else
                                'Total and separator at the end of the building
                                HTMLwriter2("      </table></td>")
                                HTMLwriter2("    </tr>")
                                Call WriteEDTCSHeader2(False) 'Untitled "CS", separator and header only
                                HTMLwriter2("    <tr>")
                                HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                            End If
                            If rowCounter = dt.Rows.Count Then Exit Do
                            strGenre = dt.Rows(rowCounter)("Categorie").ToString
                            'Title (except for buildings)
                            If strGenre <> "Bâtiment" Then
                                HTMLwriter2("        <tr>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td><b>" & strGenre & "</b></td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("          <td>&nbsp;</td>")
                                HTMLwriter2("        </tr>")
                            End If
                            'Reset the totals
                            dblSommeGenre = 0
                            dblSommeNat = 0
                        End If
                        dblSomme += Convert.ToDouble(dt.Rows(rowCounter)("Surface"))
                        dblSommeGenre += Convert.ToDouble(dt.Rows(rowCounter)("Surface"))
                        dblSommeNat += Convert.ToDouble(dt.Rows(rowCounter)("Surface"))
                        If strGenre = "Bâtiment" Then
                            HTMLwriter2("        <tr>")
                            If Convert.ToInt32(dt.Rows(rowCounter)("NoPlan")) < 9000 Then
                                HTMLwriter2("          <td width=""5%"" align=""right"">" & dt.Rows(rowCounter)("NoPlan").ToString & "</td>")
                            Else
                                HTMLwriter2("          <td width=""5%"" align=""right"">&nbsp;</td>")
                            End If
                            HTMLwriter2("          <td width=""3%"">&nbsp;</td>")
                            HTMLwriter2("          <td width=""32%""><b>Bâtiment</b></td>")
                            HTMLwriter2("          <td width=""14%"" align=""right"">" & dt.Rows(rowCounter)("ID_objet").ToString & "</td>")
                            HTMLwriter2("          <td width=""6%"">&nbsp;</td>")
                            If (dt.Rows(rowCounter)("Info_objet").ToString & "") <> "" Then
                                HTMLwriter2("          <td width=""20%"" align=""right"">" & dt.Rows(rowCounter)("Info_objet").ToString & "</td>")
                            Else
                                HTMLwriter2("          <td width=""20%"">&nbsp;</td>")
                            End If
                            HTMLwriter2("          <td width=""20%"" align=""right"">" & Format(Convert.ToDouble(dt.Rows(rowCounter)("Surface")), strFmtSurf) & "</td>")
                            HTMLwriter2("        </tr>")
                            HTMLwriter2("        <tr>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>" & dt.Rows(rowCounter)("Designation_objet").ToString & "</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("        </tr>")
                        Else
                            HTMLwriter2("        <tr>")
                            If Convert.ToInt32(dt.Rows(rowCounter)("NoPlan")) < 9000 Then
                                HTMLwriter2("          <td align=""right"">" & dt.Rows(rowCounter)("NoPlan").ToString & "</td>")
                            Else
                                HTMLwriter2("          <td align=""right"">&nbsp;</td>")
                            End If
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>" & dt.Rows(rowCounter)("EDT").ToString & "</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td align=""right"">" & Format(Convert.ToDouble(dt.Rows(rowCounter)("Surface")), strFmtSurf) & "</td>")
                            strNat = dt.Rows(rowCounter)("EDT").ToString
                            'Check if this is the last record in the table
                            If Not rowCounter + 1 >= dt.Rows.Count Then
                                'This is not the last record then we can check the next record
                                If strNat <> dt.Rows(rowCounter + 1)("EDT").ToString Then
                                    HTMLwriter2("          <td align=""right"">" & Format(dblSommeNat, strFmtSurf) & "</td>")
                                    dblSommeNat = 0
                                Else
                                    HTMLwriter2("          <td>&nbsp;</td>")
                                End If
                            Else
                                'This is the last record
                                HTMLwriter2("          <td align=""right"">" & Format(dblSommeNat, strFmtSurf) & "</td>")
                                dblSommeNat = 0
                            End If
                            HTMLwriter2("        </tr>")
                        End If
                        rowCounter += 1
                    Loop
                    'Grand total of the plot
                    HTMLwriter2("        <tr>")
                    HTMLwriter2("          <td colspan=""7"">&nbsp;</td>")
                    HTMLwriter2("        </tr>")

                    HTMLwriter2("        <tr>")
                    HTMLwriter2("          <td width=""5%"">&nbsp;</td>")
                    HTMLwriter2("          <td width=""3%"">&nbsp;</td>")
                    HTMLwriter2("          <td width=""32%""><b>Total de la parcelle</b></td>")
                    HTMLwriter2("          <td width=""14%"">&nbsp;</td>")
                    HTMLwriter2("          <td width=""6%"">&nbsp;</td>")
                    HTMLwriter2("          <td width=""20%"">&nbsp;</td>")
                    HTMLwriter2("          <td width=""20%"" align=""right""><b>" & Format(dblSomme, strFmtSurf) & "</b></td>")
                    HTMLwriter2("        </tr>")
                    'Foot section (end of table)
                    WriteEDTCSFooter2()
                    'OD
                    'dt = GetRecordset("EDT_CS", "SELECT * FROM EDT_OD WHERE Immeuble='" & strBFID & "'", "System\system.db3")
                    dt = GetRecordsetUsingWrapper("EDT_CS", "SELECT * FROM EDT_OD WHERE Immeuble='" & strBFID & "'", RVinfo.DB3system)
                    If Not dt Is Nothing AndAlso Not dt.Rows.Count = 0 Then
                        WriteEDTODSHeader2(False)
                        'Start of table
                        HTMLwriter2("    <tr>")
                        HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
                        rowCounter = 0
                        Do Until rowCounter = dt.Rows.Count
                            'Writing the ODS object
                            HTMLwriter2("        <tr>")
                            If Convert.ToInt32(dt.Rows(rowCounter)("NoPlan")) < 9000 Then
                                HTMLwriter2("          <td width=""5%"" align=""right"">" & dt.Rows(rowCounter)("NoPlan").ToString & "</td>")
                            Else
                                HTMLwriter2("          <td width=""5%"" align=""right"">&nbsp;</td>")
                            End If
                            HTMLwriter2("          <td width=""3%"">&nbsp;</td>")
                            HTMLwriter2("          <td width=""32%""><b>" & dt.Rows(rowCounter)("Categorie").ToString & "</b></td>")
                            HTMLwriter2("          <td width=""14%"" align=""right"">" & dt.Rows(rowCounter)("ID_objet").ToString & "</td>")
                            HTMLwriter2("          <td width=""6%"">&nbsp;</td>")
                            If (dt.Rows(rowCounter)("Info_objet").ToString & "") <> "" Then
                                HTMLwriter2("          <td width=""20%"" align=""right"">" & dt.Rows(rowCounter)("Info_objet").ToString & "</td>") 'Already formatted
                            Else
                                HTMLwriter2("          <td width=""20%"">&nbsp;</td>")
                            End If
                            HTMLwriter2("          <td width=""20%"" align=""right"">" & Format(Convert.ToDouble(dt.Rows(rowCounter)("Surface")), strFmtSurf) & "</td>")
                            HTMLwriter2("        </tr>")

                            '2nd line for the designation
                            HTMLwriter2("        <tr>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>" & dt.Rows(rowCounter)("Designation_objet").ToString & "</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("          <td>&nbsp;</td>")
                            HTMLwriter2("        </tr>")
                            rowCounter += 1
                        Loop
                        'End of section
                        HTMLwriter2("      </table></td>")
                        HTMLwriter2("    <tr>")
                        HTMLwriter2("    <tr>")
                        HTMLwriter2("      <td>&nbsp;</td>")
                        HTMLwriter2("    </tr>")
                    End If
                    'Foot of page
                    'Voir Templates <<  Footer : WriteEDTFooter(sWriter)


                    Dim htmlWrite As New System.IO.StreamWriter(strReport, False, System.Text.Encoding.Unicode)
                    HTMLtemplate = Replace(HTMLtemplate, "[[EDT]]", HTMLcontent)
                    htmlWrite.WriteLine(HTMLtemplate)
                    htmlWrite.Close()
                    htmlWrite.Dispose()
                End If
                If booOpenHTML Then System.Diagnostics.Process.Start(strReport)
            End If
        Catch ex As Exception

            Dim ConnErr As New Revo.connect
            ConnErr.Message("CreateEDTHtml2 : BF " & NumBF, "Erreur lors de l'écriture du rapport" & vbCrLf & ex.Message, False, 100, 100, "critical")

            Return False
        End Try

    End Function
    Private Function HTMLwriter2(Contents As String)

        HTMLcontent += Contents & vbCrLf
        Return True

    End Function
    Private Function WriteEDTHeader2(ByRef sWriter As System.IO.StreamWriter) As Boolean
        Try

            Dim Ass As New Revo.RevoInfo

            sWriter.WriteLine("<!-- START EDTHeader -->")
            sWriter.WriteLine("<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">")
            sWriter.WriteLine("<html>")
            sWriter.WriteLine("")
            sWriter.WriteLine("<head>")
            sWriter.WriteLine("  <title>EDT</title>")
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
            sWriter.WriteLine("    .Titre2 { font-size: 10pt; }")
            sWriter.WriteLine("    .Legende { font-size: 8pt; }")
            sWriter.WriteLine("  </style>")
            sWriter.WriteLine("</head>")
            sWriter.WriteLine("")
            sWriter.WriteLine("<body>")
            sWriter.WriteLine("  <table border=""0"" cellspacing=""0"" cellpadding=""0"" style=""width: 170mm;"">")
            sWriter.WriteLine("    <tr>")
            sWriter.WriteLine("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            sWriter.WriteLine("        <tr>")
            'Ignore call to CheckAppState as binary values for License are not in registry
            'CheckAppState
            'If Not booDemo Then
            ' Dim strLicence As String = LireBaseRegistre(0, "SoftWare\SWWS\JourCAD", "LicenceName") & "<br>" & LireBaseRegistre(0, "SoftWare\SWWS\JourCAD", "LicenceSite")
            ' sWriter.WriteLine("          <td width=""40%""><div align=""left"">" & strLicence & "</div></td>")
            sWriter.WriteLine("<td width=""60%""><div align=""left"">REVO " & Ass.xVersion & "<br>&nbsp;</div></td>")
            'Else
            '    sWriter.WriteLine("          <td width=""60%""><div align=""left"">REVO " & Ass.xVersion & "<br>Version d'&eacute;valuation</div></td>")
            'End If
            Dim strInfos As String = ""
            Dim logon As String = GetCurrentWinLogin()
            If logon <> String.Empty Then
                strInfos = "Etabli par " & logon & "<br>le " & Date.Today.ToLongDateString & "&nbsp;&agrave;&nbsp;" & TimeOfDay.ToShortTimeString
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
            sWriter.WriteLine("      <td><div align=""center""><span class=""Titre"">ETAT DESCRIPTIF TECHNIQUE</span></div></td>")
            sWriter.WriteLine("    </tr>")
            sWriter.WriteLine("    <tr>")
            sWriter.WriteLine("      <td>&nbsp;</td>")
            sWriter.WriteLine("    </tr>")
            sWriter.WriteLine("    <tr>")
            sWriter.WriteLine("      <td height=""4px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000; border-top-style:solid; border-top-width:2px; border-top-color:#000000;""></td>")
            sWriter.WriteLine("    </tr>")
            sWriter.WriteLine("    <tr>")
            sWriter.WriteLine("      <td height=""30"">&nbsp;</td>")
            sWriter.WriteLine("    </tr>")
            sWriter.WriteLine("<!-- END EDTHeader -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTHeader", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try

    End Function

    Public Function WriteEDTCSTitle2() As Boolean
        Try
            HTMLwriter2("<!-- START EDTCSTitle -->")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td height=""4px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000; border-top-style:solid; border-top-width:2px; border-top-color:#000000;""></td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td>&nbsp;</td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td bgcolor=""#CCCCCC""><div align=""center""><span class=""Titre"">COUVERTURE DU SOL</span></div></td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("<!-- END EDTCSTitle -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTCSTitle", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try

    End Function

    Public Function WriteEDTCSHeader2(ByVal booBat As Boolean) As Boolean
        Try
            'New Table
            HTMLwriter2("<!-- START EDTCSHeader -->")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td>&nbsp;</td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td height=""4px"" style=""border-bottom-style:solid; border-bottom-width:2px; border-bottom-color:#000000; border-top-style:solid; border-top-width:2px; border-top-color:#000000;""></td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td>&nbsp;</td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            HTMLwriter2("        <tr>")
            HTMLwriter2("          <td width=""5%"" align=""right""><b>Plan</b></td>")
            HTMLwriter2("          <td width=""3%"">&nbsp;</td>")
            HTMLwriter2("          <td width=""32%""><b>Genre</b></td>")
            If booBat = True Then 'Buildings
                HTMLwriter2("          <td width=""14%"" align=""right""><b>No b&acirc;timent</b></td>")
            Else
                HTMLwriter2("          <td width=""14%"">&nbsp;</td>")
            End If
            HTMLwriter2("          <td width=""6%"">&nbsp;</td>")
            HTMLwriter2("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
            HTMLwriter2("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
            HTMLwriter2("        </tr>")
            HTMLwriter2("        <tr>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>D&eacute;signation</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td align=""center""><span class=""Legende"">[m&sup2;]</span></td>")
            HTMLwriter2("          <td align=""center""><span class=""Legende"">[m&sup2;]</span></td>")
            HTMLwriter2("        </tr>")
            HTMLwriter2("        <tr>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            If booBat = True Then
                HTMLwriter2("          <td align=""center""><b>B&acirc;timent complet</b></td>")
                HTMLwriter2("          <td>&nbsp;</td>")
            Else
                HTMLwriter2("          <td align=""center""><b><span class=""Legende"">Objet par plan</span></b></td>")
                HTMLwriter2("          <td align=""center""><b><span class=""Legende"">Total par objet / genre</span></b></td>")
            End If
            HTMLwriter2("        </tr>")
            HTMLwriter2("        <tr>")
            HTMLwriter2("          <td colspan=""7"" height=""30"">&nbsp;</td>")
            HTMLwriter2("        </tr>")
            HTMLwriter2("      </table></td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("<!-- END EDTCSHeader -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTCSHeader", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    Public Function WriteEDTCSFooter2() As Boolean
        Try
            HTMLwriter2("<!-- START EDTCSFooter -->")
            HTMLwriter2("      </table></td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td>&nbsp;</td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("<!-- END EDTCSFooter -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTCSFooter", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    Public Function WriteEDTODSHeader2(ByVal booBat As Boolean) As Boolean
        Try
            HTMLwriter2("<!-- START EDTODSHeader -->")
            'Title and separator
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td bgcolor=""#CCCCCC""><div align=""center""><span class=""Titre"">OBJETS DIVERS (hors statistique)</span></div></td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td>&nbsp;</td>")
            HTMLwriter2("    </tr>")

            'new table
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")

            'Separator and title of columns
            HTMLwriter2("    <tr>")
            HTMLwriter2("      <td><table width=""100%"" border=""0"" cellspacing=""0"" cellpadding=""0"">")
            HTMLwriter2("        <tr>")
            HTMLwriter2("          <td width=""5%"" align=""right""><b>Plan</b></td>")
            HTMLwriter2("          <td width=""3%"">&nbsp;</td>")
            HTMLwriter2("          <td width=""32%""><b>Genre</b></td>")
            HTMLwriter2("          <td width=""14%"" align=""right""><b>No b&acirc;timent</b></td>")
            HTMLwriter2("          <td width=""6%"">&nbsp;</td>")
            HTMLwriter2("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
            HTMLwriter2("          <td width=""20%"" align=""center""><b>Surface technique</b></td>")
            HTMLwriter2("        </tr>")
            HTMLwriter2("        <tr>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>D&eacute;signation</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td align=""cente)""><span class=""Legende"">[m&sup2;]</span></td>")
            HTMLwriter2("          <td align=""center""><span class=""Legende"">[m&sup2;]</span></td>")
            HTMLwriter2("        </tr>")
            HTMLwriter2("        <tr>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("          <td align=""center""><b>B&acirc;timent complet</b></td>")
            HTMLwriter2("          <td>&nbsp;</td>")
            HTMLwriter2("        </tr>")
            HTMLwriter2("        <tr>")
            HTMLwriter2("          <td colspan=""7"" height=""30"">&nbsp;</td>")
            HTMLwriter2("        </tr>")
            HTMLwriter2("      </table></td>")
            HTMLwriter2("    </tr>")
            HTMLwriter2("<!-- END EDTODSHeader -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTODSHeader", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    Public Function WriteEDTFooter2(ByRef sWriter As System.IO.StreamWriter) As Boolean
        Try
            sWriter.WriteLine("<!-- START EDTFooter -->")
            sWriter.WriteLine("  </table>")
            sWriter.WriteLine("</body>")
            sWriter.WriteLine("")
            sWriter.WriteLine("</html>")
            sWriter.WriteLine("<!-- END EDTFooter -->")
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "WriteEDTFooter", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    Private Function WriteODSObjectsToTable2(ByRef dt As System.Data.DataTable, ByVal strCom As String, ByVal intPlan As Integer, _
                                       ByVal strNum As String, ByVal strId As String, ByVal strCodePlan As String) As Boolean
        For Each oODS As ODS In lstODS
            Dim dr As DataRow = dt.NewRow
            dr("Theme") = 3
            dr("NoCom") = strCom
            dr("NoPlan") = intPlan
            dr("NoBF") = strNum.Replace("DP", "").Trim
            dr("DP") = strNum.Contains("DP")
            dr("Immeuble") = strId
            dr("CodePlan") = strCodePlan
            dr("JC_Layer") = oODS.Type
            dr("ID_objet") = oODS.ID
            dr("Designation_objet") = oODS.Description.Replace(" ", "_")
            dr("Info_objet") = oODS.Divers
            dr("Surface") = oODS.Surface
            dt.Rows.Add(dr)
        Next
        UpdateTable(dt, RVinfo.DB3system)
    End Function





End Module
