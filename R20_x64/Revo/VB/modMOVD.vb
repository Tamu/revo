Imports System.IO
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports frms = System.Windows.Forms
Module modMOVD

    'modRecentSettings > Insertion d'un point
    Public intInsPtCat As Integer, intInsPtTheme As Integer, intInsPtType As Integer
    Public intInsPtSigne As Integer, intInsPtDefini As Integer, intInsPtVar As Integer
    Public strInsPtNum As String
    Public intInsPtFiabPlan As Integer, intInsPtFiabAlt As Integer
    Public strInsPtPrecPlan As String, strInsPtPrecAlt As String

    'modRecentSettings > Insertion d'un texte / symbole
    'Public m_ps1 As Autodesk.AutoCAD.Windows.PaletteSet = Nothing


    Public Sub ManageTerritoryBoundaries()

        Dim Ass As New Revo.RevoInfo
        DeleteAllInLayer("MO_NAT_CA")
        DeleteAllInLayer("MO_CAN_CA")
        DeleteAllInLayer("MO_DIS_CA")
        DeleteAllInLayer("MO_COM_CA")
        DeleteAllInLayer("MO_RPL_CA")

        DuplicateAllObjectsInLayer("MO_NAT", "MO_NAT_CA")
        DuplicateAllObjectsInLayer("MO_CAN", "MO_CAN_CA")
        DuplicateAllObjectsInLayer("MO_DIS", "MO_DIS_CA")
        DuplicateAllObjectsInLayer("MO_COM", "MO_COM_CA")
        DuplicateAllObjectsInLayer("MO_RPL", "MO_RPL_CA")

        Try
            'Open the impression.dat file
            Dim DrawOrderXML As String = Ass.ConfigMOVD '(Config-MOVD.xml) anciennement : impression.dat
           
            'Check that it exists first
            If File.Exists(DrawOrderXML) Then
                'Open the file
                Dim configXML As New System.Xml.XmlDocument
                configXML.Load(DrawOrderXML)
                Dim Xroot As String = "revo/draworder"
                Dim NodeDrawOrders As System.Xml.XmlNodeList
                NodeDrawOrders = configXML.SelectNodes(Xroot)
                For Each NodeDrawOrder As System.Xml.XmlNode In NodeDrawOrders
                    For Each NodeOnTB As System.Xml.XmlNode In NodeDrawOrder
                        If NodeOnTB.Name.ToLower = "ontop" Then
                            PlaceObjectsOnTop(NodeOnTB.InnerText.ToUpper, NodeOnTB.Attributes.ItemOf(0).InnerText.ToUpper)
                        ElseIf NodeOnTB.Name.ToLower = "onbottom" Then
                            PlaceObjectsOnBottom(NodeOnTB.InnerText.ToUpper, NodeOnTB.Attributes.ItemOf(0).InnerText.ToUpper)
                        End If
                    Next
                Next
            Else
                frms.MessageBox.Show("File not found: " & DrawOrderXML, "Manage Territory Boundaries", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
            End If
        Catch
            Dim Connect As New Revo.connect
            Connect.RevoLog(Connect.DateLog & "Cmd DrawOrder" & vbTab & False & vbTab & "Reading error: " & vbTab)
        End Try
    End Sub


    ''' <summary>
    ''' Gets a polyline from the user
    ''' </summary>
    ''' <param name="strLayers">A list of layers to restrict the selection to</param>
    ''' <param name="blkID">Byref blockreference to return the block inside the polyline</param>
    ''' <returns>The selected polyline</returns>
    ''' <remarks></remarks>
    Public Function GetPolylineOnScreen(ByVal strLayers As String, ByRef blkID As ObjectId) As Polyline

        Try

            'Split the layers into an array
            Dim arrLayers() As String = strLayers.Split(Convert.ToChar(","))
            'Set up the list to hold the allowed object types to select
            '(Allowed object types are Polyline and BlockReference)
            Dim lstTypes As New List(Of Type)
            lstTypes.Add(GetType(Polyline))
            lstTypes.Add(GetType(BlockReference))
            'Call the function to get the message to show to the user from the layers
            Dim msg As String = GetMsgForEntitySelection(strLayers)
            Dim altMsg As String = GetMsgFromLayers(strLayers)
            Dim LoopAgain As Boolean = True
            Do
                'Get the entity from the user
                Dim SelectedID As ObjectId = GetEntityFromUser(lstTypes, msg)
                'Check that something was selected
                If Not SelectedID.IsNull Then
                    'Get the database and start a transaction
                    Dim db As Database = HostApplicationServices.WorkingDatabase
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        'Get the selected entity from the database
                        Dim ent As Entity = CType(trans.GetObject(SelectedID, OpenMode.ForRead), Entity)
                        Dim ReturnObj As Polyline = Nothing
                        If TypeOf ent Is Polyline Then
                            'Check that the polyline is on one of the allowed layers
                            If EntityIsOnAllowedLayer(ent.ObjectId, arrLayers.ToList) Then
                                'Set the return object
                                ReturnObj = CType(ent, Polyline)
                                'Find the block inside the polyline
                                If blkID.IsNull Then 'It may already have been selected but the polyline could not be found
                                    If ent.Layer.Contains("_BF") Then
                                        'Call the function to find the block within the polyline
                                        Dim BlockID As ObjectId = GetBFBlockByPolyline(ReturnObj, ReturnObj.Layer & "_IMMEUBLE")
                                        If Not BlockID.IsNull Then 'If it's not nothing
                                            'Set the byref block id 
                                            blkID = BlockID
                                            Return ReturnObj
                                        End If
                                    Else
                                        blkID = Nothing
                                    End If
                                End If
                            Else
                                'Polyline is not on one of the allowed layers so display a warning to the user and allow them 
                                'to re-select
                                If frms.MessageBox.Show(altMsg, "Sélection incorrecte", Windows.Forms.MessageBoxButtons.OKCancel, _
                                                        Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Cancel Then
                                    Return Nothing
                                End If
                            End If
                        ElseIf TypeOf ent Is BlockReference Then
                            'Convert the entity to a block reference
                            Dim blkRef As BlockReference = CType(ent, BlockReference)
                            'Loop through the layers checking whether the block is on one of the allowed layers
                            For Each l In arrLayers
                                Dim tmpMsg As String = ""
                                Dim tmpMsg2 As String = ""
                                If blkRef.Layer = l & "_IMMEUBLE" Then
                                    blkID = blkRef.ObjectId
                                    tmpMsg = "Le contour du bien-fonds correspondant au numéro sélectionné n'a pas pu être déterminé automatiquement !" & vbCrLf _
                                            & "Sélectionner la polyligne de contour du bien-fonds désiré."
                                    tmpMsg2 = "Le contour du bâtiment correspondant au numéro sélectionné n'a pas pu être déterminé  !" & vbCrLf _
                                                                    & "Sélectionner la polyligne de contour du bâtiment désiré."
                                ElseIf blkRef.Layer = l & "_NUM" Then
                                    blkID = blkRef.ObjectId
                                    tmpMsg = "Le contour du bâtiment correspondant au numéro sélectionné n'a pas pu être déterminé automatiquement !" & vbCrLf _
                                    & "Sélectionner la polyligne de contour du bâtiment désiré."
                                    tmpMsg2 = "Le contour du bâtiment correspondant au numéro sélectionné n'a pas pu être déterminé  !" & vbCrLf _
                                        & "Sélectionner la polyligne de contour du bâtiment désiré."
                                End If
                                If tmpMsg <> "" Then 'The block is on one of the allowed layers
                                    'Get the polyline from the block
                                    Dim plID As ObjectId = GetLWPByCentroidAndRect(blkRef.Position, 1000, l)
                                    'Check whether we found the polyline
                                    If plID.IsNull Then
                                        'No polyline found so show message giving the user the option to re-select
                                        If frms.MessageBox.Show(tmpMsg, "Sélection incorrecte", Windows.Forms.MessageBoxButtons.OKCancel, _
                                            Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Cancel Then
                                            Return Nothing
                                        Else
                                            Exit For
                                        End If
                                    Else
                                        ReturnObj = CType(trans.GetObject(plID, OpenMode.ForRead), Polyline)
                                        Dim ObjPol As Polyline = CType(trans.GetObject(plID, OpenMode.ForRead), Polyline)
                                        Dim PtsColl As New Point2dCollection
                                        For i = 0 To ObjPol.NumberOfVertices - 1
                                            PtsColl.Add(ObjPol.GetPoint2dAt(i))
                                        Next
                                        Dim ObjOK As New ObjectIdCollection

                                        'Check whether there is more than one block inside the polyline
                                        If CountObjectsInsidePolyline(plID, "INSERT", blkRef.Layer) > 1 Then
                                            Dim ObjColl As ObjectIdCollection = SelectObjectsInsidePolyline(plID, "INSERT", blkRef.Layer)
                                            Dim CoordXY As String = ""
                                            For Each ObjID In ObjColl
                                                Dim Obj As Entity = CType(trans.GetObject(ObjID, OpenMode.ForRead), Entity)
                                                If TypeOf Obj Is BlockReference Then
                                                    Dim BL As BlockReference = Obj
                                                    CoordXY += BL.Position.X & "," & BL.Position.Y & vbLf
                                                    If PointInsidePolygon(PtsColl, New Point2d(BL.Position.X, BL.Position.Y), New Integer) Then
                                                        ObjOK.Add(ObjID)
                                                    End If
                                                End If
                                            Next

                                            If ObjOK.Count = 1 Then
                                                LoopAgain = False
                                            Else

                                                Dim SelVal As Double = -1
                                                Dim frmStat As New frmState
                                                frmStat.BoxList.Visible = True : frmStat.BoxList.DropDownStyle = Windows.Forms.ComboBoxStyle.DropDownList : frmStat.BtnValid.Visible = True : frmStat.ProgBar.Visible = False
                                                frmStat.Text = "Plusieurs numéros de parcelles ont été trouvés"
                                                frmStat.lbl_infos.Text = "Sélectionner un numéro de parcelle :"
                                                Dim ObjParc As New ObjectIdCollection
                                                For Each ObjID In ObjOK ' Demande de faire un choix
                                                    Dim Obj As Entity = CType(trans.GetObject(ObjID, OpenMode.ForRead), Entity)
                                                    If TypeOf Obj Is BlockReference Then

                                                        'Lecture des attributs
                                                        Dim BL As BlockReference = CType(trans.GetObject(ObjID, OpenMode.ForRead), BlockReference)
                                                        Dim attCol As AttributeCollection = BL.AttributeCollection
                                                        For Each attid As ObjectId In attCol
                                                            Dim att As AttributeReference = CType(trans.GetObject(attid, OpenMode.ForRead), AttributeReference)
                                                            Select Case att.Tag.ToUpper
                                                                Case "NUMERO"
                                                                    frmStat.BoxList.Items.Add(att.TextString)
                                                                    ObjParc.Add(ObjID)
                                                                    Exit For
                                                            End Select
                                                        Next
                                                    End If
                                                Next
                                                If frmStat.BoxList.Items.Count > 0 Then
                                                    frmStat.BoxList.Text = frmStat.BoxList.Items(0)
                                                    frmStat.ShowDialog()
                                                    SelVal = frmStat.BoxList.SelectedIndex
                                                End If

                                                If SelVal <> -1 Then 'Si fait un bon choix
                                                    '  MsgBox(SelVal)
                                                    blkID = ObjParc(SelVal)
                                                    LoopAgain = False
                                                Else
                                                    'More than one block so warn the user and give them the option to re-select
                                                    If frms.MessageBox.Show(tmpMsg2, "Sélection incorrecte", _
                                                                            Windows.Forms.MessageBoxButtons.OKCancel, Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Cancel Then
                                                        Return Nothing
                                                    Else
                                                        Exit For
                                                    End If
                                                End If

                                            End If
                                        Else
                                            LoopAgain = False
                                        End If

                                    End If
                                    Exit For
                                End If
                            Next
                            If Not ReturnObj Is Nothing And Not LoopAgain Then Return ReturnObj
                        Else
                            If frms.MessageBox.Show(altMsg, "Sélection incorrecte", Windows.Forms.MessageBoxButtons.OKCancel, _
                                                    Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Cancel Then
                                Return Nothing
                            End If
                        End If
                    End Using
                Else
                    Return Nothing
                End If
            Loop


        Catch 'ex As Exception
            Return Nothing
        End Try

    End Function

    ''' <summary>
    ''' Checks whether the entity supplied is on a layer in the supplied list of layers
    ''' </summary>
    ''' <param name="entityID">The objectid of the entity to check</param>
    ''' <param name="Layers">The list of allowed layers</param>
    ''' <returns>True if the entity is on a layer in the supplied list</returns>
    ''' <remarks></remarks>
    Public Function EntityIsOnAllowedLayer(ByVal entityID As ObjectId, ByVal Layers As List(Of String)) As Boolean
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Try
            'Start a transaction
            Using trans As Transaction = db.TransactionManager.StartTransaction
                Dim ent As Entity = CType(trans.GetObject(entityID, OpenMode.ForRead), Entity)
                'Get the layer table record that the entity is on
                Dim lt As LayerTableRecord = CType(trans.GetObject(ent.LayerId, OpenMode.ForRead), LayerTableRecord)
                'Check if the layer is in the list
                If Layers.Contains(lt.Name, StringComparer.OrdinalIgnoreCase) Then
                    lt.Dispose()
                    Return True
                End If
            End Using
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "EntityIsOnAllowedLayer", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Checks whether the entity supplied is on the layer supplied
    ''' </summary>
    ''' <param name="entityID">The objectid of the entity to check</param>
    ''' <param name="Layer">The layer to check for</param>
    ''' <returns>True if the entity is on the layer supplied</returns>
    ''' <remarks></remarks>
    Public Function EntityIsOnAllowedLayer(ByVal entityID As ObjectId, ByVal Layer As String) As Boolean
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Try
            'Start a transaction
            Using trans As Transaction = db.TransactionManager.StartTransaction
                Dim ent As Entity = CType(trans.GetObject(entityID, OpenMode.ForRead), Entity)
                'Get the layer table record that the entity is on
                Dim lt As LayerTableRecord = CType(trans.GetObject(ent.LayerId, OpenMode.ForRead), LayerTableRecord)
                'Check if the layer is in the list
                If Layer.ToUpper = lt.Name.ToUpper Then
                    lt.Dispose()
                    Return True
                End If
            End Using
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "EntityIsOnAllowedLayer", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Return False
    End Function

    ''' <summary>
    ''' Get the message to display to the user while selecting an entity
    ''' </summary>
    ''' <param name="strLayers">The list of layers the selection is restricted to</param>
    ''' <returns></returns>
    ''' <remarks>The message differs depending on the layers that are passed in</remarks>
    Private Function GetMsgForEntitySelection(ByVal strLayers As String) As String

        If strLayers.Contains("_CS_BAT") And strLayers.Contains("_BF") And strLayers.Contains("_ODS_BATSOUT") Then
            Return "Sélectionner un objet (BF, CS_BAT ou ODS_BATSOUT): "
        ElseIf strLayers.Contains("_CS_BAT") And strLayers.Contains("_BF") Then
            Return "Sélectionner un objet (BF ou CS_BAT): "
        ElseIf strLayers.Contains("_CS_BAT") Then
            Return "Sélectionner un objet (CS_BAT): "
        ElseIf strLayers.Contains("_ODS_BATSOUT") Then
            Return "Sélectionner un objet (ODS_BATSOUT): "
        ElseIf strLayers.Contains("_ODS_COUVINDEP") Then
            Return "Sélectionner un objet (ODS_COUVINDEP): "
        Else
            Return "Sélectionner un objet (BF): "
        End If
    End Function

    Private Function GetMsgFromLayers(ByVal strLayers As String) As String
        If strLayers.Contains("_BAT") And strLayers.Contains("_BF") Then
            Return "Sélectionner un bien-fonds ou un bâtiment (polyligne de contour ou bloc de définition des attributs) !"
        ElseIf strLayers.Contains("_BF") Then
            Return "Sélectionner un bien-fonds (polyligne de contour ou bloc de définition des attributs) !"
        ElseIf strLayers.Contains("_BAT") Then
            Return "Sélectionner un bâtiment (polyligne de contour ou bloc de définition des attributs) !"
        Else
            Return "Sélectionner un objet surfacique (polyligne de contour ou bloc de définition des attributs) !"
        End If
    End Function

    Public Function GetBFBlockByPolyline(ByVal objBF As Polyline, ByVal strLayer As String) As ObjectId
        Try
            'Get the vertices of the polyline
            Dim Verts As Point3dCollection = GetPolyVertices(objBF.ObjectId)
            'Zoom to the extents of the object otherwise the selection may not work
            Zooming.ZoomToObject(objBF.ObjectId)
            'Get the editor
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            'Create typedvalues to select Blocks (INSERT) on the supplied layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayerName, strLayer)}
            'Create the selectionfilter from the typedvalues
            Dim sFilter As New SelectionFilter(Values)
            'Make the selection
            Dim pres As PromptSelectionResult = ed.SelectCrossingPolygon(Verts, sFilter)
            'Check that we found something
            If pres.Status = PromptStatus.OK Then
                Dim ids() As ObjectId = pres.Value.GetObjectIds
                'No blocks found or more than one so the block must be selected manually
                If ids.Length <> 1 Then
                    'Warn the user 
                    If frms.MessageBox.Show("Le bloc contenant la définition du bien-fonds à analyser n'a pas pu être localisé automatiquement" _
                                            & vbCrLf & "(numéro en dehors du périmètre du bien-fonds ou plusieurs numéros à proximité) !" _
                                            & vbCrLf & vbCrLf & _
                                            "Pour pouvoir générer le rapport, vous devez sélectionner ce bloc manuellement.", _
                                            "Impossible de récupérer les attributs du bien-fonds à traiter", Windows.Forms.MessageBoxButtons.OKCancel, _
                                             Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.OK Then
                        'Call the function that asks the user to select the block manually
                        Return GetMO_BFBlock(strLayer)
                    Else
                        Return Nothing
                    End If
                Else
                    'Get the block
                    Return ids(0)
                End If
            End If
            Zooming.ZoomPrevious()
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetBFBlockByPolyline", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    ''' <summary>
    ''' Ask the user to select the block
    ''' </summary>
    ''' <param name="strLayer">The layer that the selected block should be on</param>
    ''' <returns>The objectid of the selected block</returns>
    ''' <remarks></remarks>
    Private Function GetMO_BFBlock(ByVal strLayer As String) As ObjectId
        'Set the message to display to the user while selecting
        Dim msg As String = "Sélectionner le numéro du bien-fonds (bloc): "
        'Set up the types of object to be selected
        Dim lstTypes As New List(Of Type)
        lstTypes.Add(GetType(BlockReference)) 'Only allowing BlockReferences
        'Loop around until either the user selects a valid block or presses Cancel on the message box
        Do
            'Ask the user to select an entity
            Dim SelectedID As ObjectId = GetEntityFromUser(lstTypes, msg)
            'Check that they selected something
            If Not SelectedID.IsNull Then
                'Check that it's on the allowed layer
                If EntityIsOnAllowedLayer(SelectedID, strLayer) Then
                    'Return the objectid
                    Return SelectedID
                End If
            End If
            'Either the selection was null or it was not on the allowed layer
            'Display a message to the user allowing them to select again if desired
        Loop While frms.MessageBox.Show("Sélectionner un bien-fonds (bloc de définition des attributs) !", "Sélection incorrecte", _
                                         Windows.Forms.MessageBoxButtons.OKCancel, Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.OK
        'User cancelled messagebox so return nothing
        Return Nothing
    End Function

    ''' <summary>
    ''' Counts the objects in the polyline
    ''' </summary>
    ''' <param name="plID">The objectid of the polyline</param>
    ''' <param name="strEntity">The DXFCode entity type to search for (e.g. INSERT for blockreferences) </param>
    ''' <param name="strLayer">The layer to restrict the search to</param>
    ''' <returns>The number of items found</returns>
    ''' <remarks></remarks>
    Public Function CountObjectsInsidePolyline(ByVal plID As ObjectId, ByVal strEntity As String, ByVal strLayer As String) As Integer
        'Set up the typedvalues to select the desired entity types on the desired layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, strEntity), New TypedValue(DxfCode.LayerName, strLayer)}
        'Get its vertices
        Dim Verts As Point3dCollection = GetPolyVertices(plID)
        'Zoom to the object (so we don't get problems with the selection)
        Zooming.ZoomToObject(plID)
        Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)
        'Check that we found something
        If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
            'Return the number of items found
            Return ids.Length
        End If
        Return 0
    End Function

    ''' <summary>
    ''' Select the objects in the polyline
    ''' </summary>
    ''' <param name="plID">The objectid of the polyline</param>
    ''' <param name="strEntity">The DXFCode entity type to search for (e.g. INSERT for blockreferences) </param>
    ''' <param name="strLayer">The layer to restrict the search to</param>
    ''' <returns>Object ID collection</returns>
    ''' <remarks></remarks>
    Public Function SelectObjectsInsidePolyline(ByVal plID As ObjectId, ByVal strEntity As String, ByVal strLayer As String) As ObjectIdCollection
        'Set up the typedvalues to select the desired entity types on the desired layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, strEntity), New TypedValue(DxfCode.LayerName, strLayer)}
        'Get its vertices
        Dim Verts As Point3dCollection = GetPolyVertices(plID)
        'Zoom to the object (so we don't get problems with the selection)
        Zooming.ZoomToObject(plID)
        Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)

        'Check that we found something
        If Not ids Is Nothing AndAlso Not ids.Count = 0 Then

            Dim SelectID As ObjectId
            'Charge les variables
            Dim EDTList As New Collection
            For Each id As ObjectId In ids
                Dim ListParc As EDTListItem = GetInfoFromBlock(id)
                EDTList.Add(ListParc)
                SelectID = ListParc.BlockID
            Next

            Dim ObjColl As New ObjectIdCollection
            For Each ObjId In ids
                ObjColl.Add(ObjId)
            Next

            'Return the number of items found
            Return ObjColl 'SelectID

        End If
        Return Nothing
    End Function

    ''' <summary>
    ''' Asks the user to select a block on screen
    ''' </summary>
    ''' <param name="strLayer">The leayer that the selected block must be on to be valid</param>
    ''' <returns>The objectid of the selected block</returns>
    ''' <remarks></remarks>
    Public Function GetBlockBySelection(ByVal strLayer As String) As ObjectId
        Try
            'Variable for message
            Dim strObj As String = "objet"
            Select Case strLayer
                Case "MO_RPL_PLAN" : strObj = "plan"
                Case "MO_BF_IMMEUBLE" : strObj = "numéro du bien-fonds"
                Case "MO_BF_DDP_IMMEUBLE" : strObj = "numéro d'un DDP"
            End Select
            'Display the message with the instruction for selection to the user
            If frms.MessageBox.Show("Le " & strObj & " concerné, nécessaire pour la lecture des attributs administratifs," _
                                    & vbCrLf & "n'a pas pu être trouvé automatiquement !" & vbCrLf & vbCrLf & _
                                    "Sélectionner manuellement le bloc de définition du " & strObj & " sur le dessin.", _
                                    "Sélection du " & strObj, Windows.Forms.MessageBoxButtons.OKCancel, _
                                    Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.OK Then
                'Check that they clicked OK to continue
                'Set up the allowed types
                Dim lstTypes As New List(Of Type)
                lstTypes.Add(GetType(BlockReference))
                'Loop until the user selects a block on the correct layer or they cancel
                Do
                    'Call the function to get the selection from the user
                    Dim selID As ObjectId = GetEntityFromUser(lstTypes, "Sélectionner une définition de " & strObj & ": ")
                    'Check that they selected something
                    If Not selID.IsNull Then
                        'Check that it's on an allowed layer
                        If EntityIsOnAllowedLayer(selID, strLayer) Then
                            'It is so return the objectid
                            Return selID
                        Else
                            'It's not on the correct layer so warn the user
                            If Not frms.MessageBox.Show("Sélectionner un " & strObj & " (bloc de définition des attributs) !", _
                                "Sélection incorrecte", Windows.Forms.MessageBoxButtons.OKCancel, _
                                Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.OK Then
                                'they don't want to carry on so return nothing
                                Return Nothing
                            End If
                        End If
                    Else
                        'Selected id is nothing
                        Return Nothing
                    End If
                Loop
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetBlockBySelection", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Return Nothing
    End Function

    ''' <summary>
    ''' Determines whether the supplied entity is within the supplied polyline
    ''' </summary>
    ''' <param name="ent">The entity to test</param>
    ''' <param name="Polyid">The polyline we're checking within</param>
    ''' <returns>True if the object is whithin the polyline</returns>
    ''' <remarks></remarks>
    Public Function IsObjectInsidePolyline(ByVal ent As Entity, ByVal Polyid As ObjectId) As Boolean
        Try
            'get the vertices of the polyline
            Dim Verts As Point3dCollection = GetPolyVertices(Polyid)
            Dim strEntity As String = ""
            'Determine which type of entity to look for based on the supplied entity type
            Dim entType As String = ent.GetType.ToString
            entType = entType.Substring(entType.LastIndexOf(".") + 1)
            Select Case entType.ToUpper
                Case "LINE" : strEntity = "LINE"
                Case "POLYLINE" : strEntity = "LWPOLYLINE"
                Case "REGION" : strEntity = "REGION"
                Case "BLOCKREFERENCE" : strEntity = "INSERT"
            End Select
            'Set up the typedvalues to look for that type of entity on the 
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, strEntity), New TypedValue(DxfCode.LayerName, ent.Layer)}
            Dim IDs() As ObjectId = SelectCrossingPoly(Values, Verts)
            If Not IDs Is Nothing AndAlso Not IDs.Count = 0 Then
                'Loop through the found objects checking if any of them have the objectid of the entity that we're looking for
                For Each id As ObjectId In IDs
                    'If the objectids match then it's the same entity and it is within the polyline so return true
                    If id.Equals(ent.ObjectId) Then Return True
                Next
                Return False
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "IsObjectInsidePolyline", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Gets the centre of the polyline
    ''' </summary>
    ''' <param name="polyID">The objectid of the polyline</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetGravityCentre(ByVal polyID As ObjectId) As Point3d
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim pline As Polyline = CType(trans.GetObject(polyID, OpenMode.ForRead), Polyline)
            Dim NumVerts As Integer = pline.NumberOfVertices
            Dim X As Double = 0
            Dim Y As Double = 0
            For i As Integer = 0 To NumVerts - 1
                Dim pt As Point3d = pline.GetPoint3dAt(i)
                X += pt.X
                Y += pt.Y
            Next
            Dim retPt As New Point3d(X / NumVerts, Y / NumVerts, 0)
            pline.Dispose()
            Return retPt
        End Using
    End Function

    Public Function CodePlanToTypeMens(ByVal intplan As Integer, ByVal strCodePlan As String) As String
        Dim retVal As String = ""
        'Type mens.
        Select Case Left(strCodePlan, 1)
            Case "0" : retVal = "?"
            Case "1" : retVal = "Transitoire"
            Case "2" : retVal = "Graphique"
            Case "3"
                If intplan >= 1000 Then retVal = "Semi-num." Else retVal = "Semi-numérique"
            Case "4" : retVal = "Numérique"
        End Select
        'Numérisé
        If intplan >= 1000 Then
            retVal = retVal & " numérisé"
        End If
        Return retVal
    End Function

    Public Function GetBATBlockByPolyline(ByVal objBAT As ObjectId, ByVal strLayer As String) As ObjectId
        'Get the polyline vertices
        Dim Verts As Point3dCollection = GetPolyVertices(objBAT)
        'Zoom to the extents of the drawing
        Zooming.ZoomExtents()
        Try

            'Set up the values to select blocks on the correct layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayerName, strLayer)}
            'Make the selection
            Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)
            If Not ids Is Nothing AndAlso Not ids.Count = 1 Then

                If frms.MessageBox.Show("Le bloc contenant la définition du bâtiment à analyser n'a pas pu être localisé automatiquement" & _
                                        vbCrLf & "(numéro en dehors du périmètre du bâtiment ou plusieurs numéros à proximité) !" & _
                                        vbCrLf & vbCrLf & "Pour pouvoir générer le rapport, vous devez sélectionner ce bloc manuellement.", _
                                        "Impossible de récupérer les attributs du bâtiment à traiter", Windows.Forms.MessageBoxButtons.OKCancel, _
                                        Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.OK Then
                
                    'Select on screen
                    Zooming.ZoomToObject(objBAT)
                    GetBATBlockByPolyline = GetMO_BATBlock(strLayer)
                    Zooming.ZoomPrevious()
                Else
                    'Cancel
                    GetBATBlockByPolyline = Nothing
                End If
            Else
                'Return the block
                GetBATBlockByPolyline = ids(0)
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetBATBlockByPolyline", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Zooming.ZoomPrevious()
    End Function

    Private Function GetMO_BATBlock(ByVal strLayer As String) As ObjectId
        Try
            'Add the blockreference type to a list so it can be passed to the GetEntityFromUser function 
            Dim Types As New List(Of Type)
            Types.Add(GetType(BlockReference))
            Do
                'Call the function to ask the user to select an entity (limited to the types in the Types list)
                Dim objId As ObjectId = GetEntityFromUser(Types, "Sélectionner le numéro du bâtiment (bloc): ")
                'Check that they selected something
                If Not objId.IsNull Then
                    'Get the database and start a transaction
                    Dim db As Database = HostApplicationServices.WorkingDatabase
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        'Get the block fromthe database
                        Dim blkRef As BlockReference = CType(trans.GetObject(objId, OpenMode.ForRead), BlockReference)
                        'Check that the selected block is on the layer we want
                        If blkRef.Layer.Equals(strLayer, StringComparison.OrdinalIgnoreCase) Then
                            'If so, then return the objectid of the block
                            Return objId
                        Else
                            'The block is not on the correct layer so ask the user to if they want to select again
                            If frms.MessageBox.Show("Sélectionner un bâtiment (bloc de définition des attributs) !", "Sélection incorrecte", _
                                                     Windows.Forms.MessageBoxButtons.OKCancel, Windows.Forms.MessageBoxIcon.Exclamation) = Windows.Forms.DialogResult.Cancel Then
                                'They cancelled so return nothing
                                Return Nothing
                            End If
                        End If
                    End Using
                Else
                    'Nothing selected so return nothing
                    Return Nothing
                End If
            Loop
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetMO_BATBlock", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
    End Function

    Public Function GetBFNumFromBAT(ByVal objCSID As ObjectId, Optional ByRef strCom As String = "?") As String
        GetBFNumFromBAT = ""
        Try
            'Get the vertices of the polyline
            Dim Verts As Point3dCollection = GetPolyVertices(objCSID)
            'Zoom extents (otherwise problems with selection)
            Zooming.ZoomExtents()
            'Set up the typedvalues to select polylines on the MO_BF layer
            Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, "MO_BF")}
            'Make the selection using the vertices and the values
            Dim ids() As ObjectId = SelectCrossingPoly(Values, Verts)
            Dim blkBF As ObjectId = Nothing
            Dim objPLbfID As ObjectId
            'Check that we found something inside the CS poly
            If Not ids Is Nothing AndAlso Not ids.Count = 0 Then
                Dim objRegBF As Region
                'Convert the CS poly to a region
                Dim objRegCS As Region = ConvertPolylineToRegion(objCSID)
                Dim dblSurfInit As Double
                'Loop through the polylines inside the CS polyline
                For Each objPLbfID In ids
                    'Convert it to a region
                    objRegBF = ConvertPolylineToRegion(objPLbfID)
                    'Check that it converted OK
                    If Not objRegBF.IsNull Then
                        'Get the area of the region
                        dblSurfInit = objRegBF.Area
                        'Intersect it with the CS region
                        objRegBF.BooleanOperation(BooleanOperationType.BoolIntersect, CType(objRegCS.Clone, Region))
                        'Compare the areas
                        If Math.Abs(dblSurfInit - objRegBF.Area) > 0.01 Then
                            blkBF = GetBlockfromLWP1(objPLbfID, "MO_BF_IMMEUBLE")
                            GetBFNumFromBAT = GetBFNumFromBAT & GetBlockAttribute(blkBF, "NUMERO") & ","
                        End If
                    Else
                        'Failed to convert the polyline to a region
                    End If
                    objRegBF.Dispose()
                Next
                objRegCS.Dispose()
            Else
                Dim pres As PromptSelectionResult = SelectAllItems(Values)
                If pres.Status = PromptStatus.OK Then
                    Dim db As Database = HostApplicationServices.WorkingDatabase
                    Using trans As Transaction = db.TransactionManager.StartTransaction
                        Dim objCS As Entity = CType(trans.GetObject(objCSID, OpenMode.ForRead), Entity)
                        For Each objPLbfID In pres.Value.GetObjectIds
                            If IsObjectInsidePolyline(objCS, objPLbfID) Then
                                blkBF = GetBlockfromLWP1(objPLbfID, "MO_BF_IMMEUBLE")
                                If Not blkBF.IsNull Then GetBFNumFromBAT = GetBlockAttribute(blkBF, "NUMERO") & ","
                                Exit For
                            End If
                        Next
                        objCS.Dispose()
                    End Using
                Else
                    blkBF = GetBlockBySelection("MO_BF_IMMEUBLE")
                    If Not blkBF.IsNull Then GetBFNumFromBAT = GetBlockAttribute(blkBF, "NUMERO") & ","
                End If
            End If
            If Not blkBF.IsNull Then
                Dim strIdentDN As String = GetBlockAttribute(blkBF, "IDENTDN")
                strCom = strIdentDN.Substring(3, 3)
            End If
            'Remove the last ","
            If GetBFNumFromBAT <> "" Then
                GetBFNumFromBAT = GetBFNumFromBAT.Substring(0, GetBFNumFromBAT.Length - 1)
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetBFNumFromBAT", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Zooming.ZoomPrevious()
    End Function


    Public Sub PlanToList(ByRef strPlanInfos As String, ByRef objLV As System.Windows.Forms.ListView)

        On Error Resume Next 'POur géréer les infos de plan hors VD

        'Déclarations
        Dim arrInfos() As String
        Dim intIdx As Short = objLV.Items.Count + 1

        'Attributs de plan
        arrInfos = Split(strPlanInfos, " ")


        With objLV.Items

            .Add("IDX" & intIdx, CStr(Val(Mid(arrInfos(3), 4, 3))), "") 'Commune
            'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            If .Item("IDX" & intIdx).SubItems.Count > 1 Then
                .Item("IDX" & intIdx).SubItems(1).Text = arrInfos(4)
            Else
                .Item("IDX" & intIdx).SubItems.Insert(1, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, arrInfos(4)))
            End If 'Plan

            'Echelle
            Select Case Mid(arrInfos(8), 4, 1)
                Case "0"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 2 Then
                        .Item("IDX" & intIdx).SubItems(2).Text = "?"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "?"))
                    End If
                Case "1"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 2 Then
                        .Item("IDX" & intIdx).SubItems(2).Text = "1 :10'000"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "1 :10'000"))
                    End If
                Case "2"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 2 Then
                        .Item("IDX" & intIdx).SubItems(2).Text = "1 : 5'000"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "1 : 5'000"))
                    End If
                Case "3"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 2 Then
                        .Item("IDX" & intIdx).SubItems(2).Text = "1 : 4'000"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "1 : 4'000"))
                    End If
                Case "4"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 2 Then
                        .Item("IDX" & intIdx).SubItems(2).Text = "1 : 2'000"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "1 : 2'000"))
                    End If
                Case "5"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 2 Then
                        .Item("IDX" & intIdx).SubItems(2).Text = "1 : 1'000"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "1 : 1'000"))
                    End If
                Case "6"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 2 Then
                        .Item("IDX" & intIdx).SubItems(2).Text = "1 :   500"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "1 :   500"))
                    End If
                Case "7"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 2 Then
                        .Item("IDX" & intIdx).SubItems(2).Text = "1 :   250"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(2, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "1 :   250"))
                    End If
            End Select

            'NT
            Select Case Mid(arrInfos(8), 3, 1)
                Case "0"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "?"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "?"))
                    End If
                Case "1"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "Digit. non qual."
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Digit. non qual."))
                    End If
                Case "2"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "Digit. graph."
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Digit. graph."))
                    End If
                Case "3"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "Digit. semi-num."
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Digit. semi-num."))
                    End If
                Case "4"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "NT5"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "NT5"))
                    End If
                Case "5"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "NT4"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "NT4"))
                    End If
                Case "6"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "NT3"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "NT3"))
                    End If
                Case "7"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "NT2"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "NT2"))
                    End If
                Case "8"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 3 Then
                        .Item("IDX" & intIdx).SubItems(3).Text = "NT1"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(3, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "NT1"))
                    End If
            End Select

            'Type mens.
            Select Case Mid(arrInfos(8), 1, 1)
                Case "0"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 4 Then
                        .Item("IDX" & intIdx).SubItems(4).Text = "?"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "?"))
                    End If
                Case "1"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 4 Then
                        .Item("IDX" & intIdx).SubItems(4).Text = "Transitoire"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Transitoire"))
                    End If
                Case "2"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 4 Then
                        .Item("IDX" & intIdx).SubItems(4).Text = "Graphique"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Graphique"))
                    End If
                Case "3"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 4 Then
                        .Item("IDX" & intIdx).SubItems(4).Text = "Semi-numérique"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Semi-numérique"))
                    End If
                Case "4"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 4 Then
                        .Item("IDX" & intIdx).SubItems(4).Text = "Numérique"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(4, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Numérique"))
                    End If
            End Select

            'Référentiel
            Select Case Mid(arrInfos(8), 7, 1)
                Case "0"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 5 Then
                        .Item("IDX" & intIdx).SubItems(5).Text = "?"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "?"))
                    End If
                Case "1"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 5 Then
                        .Item("IDX" & intIdx).SubItems(5).Text = "Aucun"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Aucun"))
                    End If
                Case "2"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 5 Then
                        .Item("IDX" & intIdx).SubItems(5).Text = "Proj. Bonne"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Proj. Bonne"))
                    End If
                Case "3"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 5 Then
                        .Item("IDX" & intIdx).SubItems(5).Text = "Proj. MN03"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Proj. MN03"))
                    End If
                Case "4"
                    'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item() est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
                    If .Item("IDX" & intIdx).SubItems.Count > 5 Then
                        .Item("IDX" & intIdx).SubItems(5).Text = "Proj. MN95"
                    Else
                        .Item("IDX" & intIdx).SubItems.Insert(5, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, "Proj. MN95"))
                    End If
            End Select

            'Commune et plan, pour tri croissant
            'UPGRADE_WARNING: La limite inférieure de la collection objLV.ListItems.Item(IDX & intIdx) est passée de 1 à 0. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="A3B628A0-A810-4AE2-BFA2-9E7A29EB9AD0"'
            If .Item("IDX" & intIdx).SubItems.Count > 6 Then
                .Item("IDX" & intIdx).SubItems(6).Text = Format(Val(Mid(arrInfos(3), 4, 3)), "000") & Format(Val(arrInfos(4)), "0000")
            Else
                .Item("IDX" & intIdx).SubItems.Insert(6, New System.Windows.Forms.ListViewItem.ListViewSubItem(Nothing, Format(Val(Mid(arrInfos(3), 4, 3)), "000") & Format(Val(arrInfos(4)), "0000")))
            End If


            'Code en entier => Tag
            .Item("IDX" & intIdx).Tag = arrInfos(8)

        End With

        On Error GoTo 0

        'Exemple
        'TABL Plan
        ''0      1         2          3        4   5     6        7         8        9
        'OBJE 13372507 13369294 VD0152000000 1030 NPC 19470101 19940101 323542322 8115743
        'OBJE 13372508 13369294 VD0152000000 1029 NPC 19530101 19950101 323542322 8115744
        'ETAB

    End Sub

    Public Function GetMOplan(Pts3D As Point3d)

        Dim strEntity As String = "INSERT"
        Dim BlocName As String = "PLAN,MUT_PLAN"
        Dim strLayer As String = "MO_RPL_PLAN,MUT_RPL_PLAN"
        Dim ComList As New List(Of String)
        Dim PlanList As New List(Of String)
        Dim CodeList As New List(Of String)

        'Set up the typedvalues to filter out the objects we're looking for
        Dim Values() As TypedValue
        If strLayer = "" Then
            Values = New TypedValue() {New TypedValue(DxfCode.Start, strEntity), New TypedValue(DxfCode.BlockName, BlocName)}
        Else
            Values = New TypedValue() {New TypedValue(DxfCode.Start, strEntity), New TypedValue(DxfCode.LayerName, strLayer), New TypedValue(DxfCode.BlockName, BlocName)}
        End If
        'Select all items using the filter values
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then ' If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'Declare a objectidcollection (required for the MoveToTop method of the draw order table)
            ' Dim entCol As New ObjectIdCollection
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Get the blocktablerecord for modelspace
                '  Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                ' Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                'Loop through the found objects adding them to the objectidcollection

                For Each oID As ObjectId In pres.Value.GetObjectIds
                    Try

                        'Lecture des attributs
                        Dim blkRef As BlockReference = CType(trans.GetObject(oID, OpenMode.ForRead), BlockReference)
                        'Get the attributecollection
                        Dim attCol As AttributeCollection = blkRef.AttributeCollection
                        'Loop through the attributes
                        For Each attid As ObjectId In attCol
                            'Get the attribute reference
                            Dim attRef As AttributeReference = CType(trans.GetObject(attid, OpenMode.ForRead), AttributeReference)
                            'Check which attribute this is
                            Select Case attRef.Tag.ToUpper
                                Case "IDENTDN"
                                    ComList.Add(Mid(attRef.TextString, 4, 3))
                                Case "NUMERO"
                                    PlanList.Add(attRef.TextString)
                                Case "CODEPLAN"
                                    CodeList.Add(attRef.TextString)
                            End Select
                        Next

                    Catch ex As Exception
                        ' MsgBox(ex.Message)
                    End Try
                Next

                'Save the changes
                'trans.Commit()

            End Using
        End If


        Return ComList

    End Function

    Public Function PropRF(Canton As String)

       
        'Sélection des parcelles / bloc :
        'IMMEUBLE_NUM, IMMEUBLE_NUM_DDP, IMMEUBLE_NUM_DDP_PROJ, IMMEUBLE_NUM_PROJ, MUT_IMMEUBLE_NUM, MUT_IMMEUBLE_NUM_DDP, RAD_IMMEUBLE_DDP, RAD_IMMEUBLE_NUM

        Dim RVinfo As New Revo.RevoInfo
        Dim Connect As New Revo.connect
        Dim NbreParc As Double = 0

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim CollBlock As New Collection
        Dim CollTxt As New Collection

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acOPrompt As PromptSelectionOptions = New PromptSelectionOptions()
                acOPrompt.MessageForAdding = "Sélectionner les parcelles à rechercher (Bloc MO)"
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
                            Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)
                            If Not IsDBNull(acEnt) Then
                                If TypeName(acEnt) Like "*BlockReference*" Then
                                    Dim BL As BlockReference = acEnt
                                    If InStr("-IMMEUBLE_NUM-IMMEUBLE_NUM_DDP-IMMEUBLE_NUM_DDP_PROJ-IMMEUBLE_NUM_PROJ-MUT_IMMEUBLE_NUM-MUT_IMMEUBLE_NUM_DDP-RAD_IMMEUBLE_DDP-RAD_IMMEUBLE_NUM-", "-" & BL.Name & "-") <> 0 Then
                                        CollBlock.Add(acEnt)
                                    End If
                                End If

                            End If
                        End If
                    Next
                End If

                '  If CollBlock.Count = 0 Then Return False ' Aucune Sélection

                NbreParc = CollBlock.Count
                Dim DateUpdate As String = ""
                DateUpdate = Format(Day(Today), "00") & "." & Format(Month(Today), "00") & "." & Format(Year(Today), "0000")

                Connect.Message("Connection et traitement des données en ligne RF (gratuit)", "Recherche des informations ... ", False, 10, 100)

                If Canton = "VD" Then

                    Dim RFurl As String = "http://www.rfinfo.vd.ch/interop/rfinfo.php?no_commune=#RFcom#&no_immeuble=#RFparc#"

                    Dim i As Double = 0
                    'boucle dans les parcelles
                    For Each BL As BlockReference In CollBlock

                        'Recherche de Communes
                        'Lecture des attributs
                        i += 1
                        Connect.Message("Connection et traitement des données en ligne RF (gratuit)", "Traitement de la parcelle " & i & "/" & CollBlock.Count, False, i, CollBlock.Count)

                        'Get the attributecollection
                        Dim attCol As AttributeCollection = BL.AttributeCollection
                        Dim Commune As String = ""
                        Dim NumParc As String = ""

                        'Loop through the attributes
                        For Each attid As ObjectId In attCol
                            'Get the attribute reference
                            Dim attRef As AttributeReference = DirectCast(acTrans.GetObject(attid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), AttributeReference)
                            'Check which attribute this is
                            Select Case attRef.Tag.ToUpper
                                Case "IDENTDN" 'VD0132000000
                                    Commune = Trim(Mid(attRef.TextString, 4, 3))
                                Case "NUMERO" 'VD0132000000
                                    NumParc = Trim(attRef.TextString)
                            End Select
                        Next


                        'Rechercher CADASTRE VD

                        'ignorer les DP
                    If Left(Trim(NumParc.ToUpper), 2) <> "DP" Then


                        'www.rfinfo.vd.ch/interop/rfinfo.php?no_commune=#RFcom#&no_immeuble=#RFparc#
                        Dim RFurlX As String = Replace(Replace(RFurl, "#RFcom#", Commune), "#RFparc#", NumParc)
                        Dim RVscript As New Revo.RevoScript
                        Dim VarData As New List(Of String)
                        Dim Collprop As New Collection

                        'Requête RF : HTTP

                        'Var;/CADVD1/;http;[[Wurl]]
                        Collprop.Add(Split("Wurl]]" & RFurlX, "]]"))
                        VarData = RVscript.VarHttp(Collprop, VarData, Nothing, "")
                        Collprop.Clear()

                        'Var;/CADVD2/;http;[[Wtable]]0||1||/RFnum/-/RFparc/
                        Collprop.Add(Split("Wtable]]0||1||/RFnum/-/RFparc/", "]]"))
                        VarData = RVscript.VarHttp(Collprop, VarData, Nothing, "")
                        Collprop.Clear()

                        'Var;/CADVD3/;http;[[WfindT]]/RFnum/-/RFparc/||1||Propriétaire(s)*>*
                        Collprop.Add(Split("WfindT]]/RFnum/-/RFparc/||1||Propriétaire(s)*>*", "]]"))
                        VarData = RVscript.VarHttp(Collprop, VarData, Nothing, "")
                        Collprop.Clear()


                        'Var;/Proprietaires/;Acad;[[Val]];>/CADVD3/;[[Replace]]\P||, ;
                        Dim NomProp As String = ""
                        For Each Var As String In VarData
                            If NomProp <> "" Then NomProp += "\P"
                            NomProp += Replace(Var, ",", "\P")
                        Next


                        'BL;ID:(2127672488);[[Space]];Model;[[!xyz]]Val;
                        'BL;NUMERO.XNUMERO.Y;[[Space]];Model;[[!NOM_PROP]]/Proprietaires/;

                        'Loop through the attributes
                        For Each attid As ObjectId In attCol
                            'Get the attribute reference
                            Dim attRef As AttributeReference = DirectCast(acTrans.GetObject(attid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), AttributeReference)
                            'Check which attribute this is
                            Select Case attRef.Tag.ToUpper
                                Case "NOM_PROP" 'NOM Prénom
                                    attRef.TextString = NomProp
                                Case "MAJ_PROP" '03.12.2012
                                    If Trim(NomProp) = "" Then
                                        attRef.TextString = "-"
                                    Else
                                        attRef.TextString = DateUpdate
                                    End If
                            End Select
                        Next

                    End If

                    Next
                Else
                    Connect.Message(RVinfo.xTitle, "Le canton actif n'es pas disponible pour cette fonction", False, 100, 100, "info")
                    Connect.Message(RVinfo.xTitle, "", True, 100, 100)
                    MsgBox("", vbInformation)
                End If


                Connect.Message(RVinfo.xTitle, "Opération terminée avec succès. Nombre de parcelles traitées : " & NbreParc, False, 100, 100, "info")
                Connect.Message(RVinfo.xTitle, "", True, 100, 100)

                '' Save the new object to the database
                acTrans.Commit()
                '' Dispose of the transaction
            End Using
            ' End selection of objects

       
        Return True

    End Function
    Public Sub EditTxtBlock(ByRef objBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference)

        'Chargement de la palettes
        Dim MyCmd As New Revo.MyCommands
        MyCmd.RevoPalette()


        'Insertion du bloc
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                    Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

        Dim dblInsertPt() As Double = objBlock.InsertionPoint
        Revo.MyCommands.PaletteTxT.txtX.Text = dblInsertPt(0)
        Revo.MyCommands.PaletteTxT.txtY.Text = dblInsertPt(1)
        Revo.MyCommands.PaletteTxT.txtZ.Text = dblInsertPt(2)
        Revo.MyCommands.PaletteTxT.txtOri.Text = CDbl(Format(ConvAngTrigoTopoG((objBlock.Rotation * 200) / Math.PI), "0.000")) ' (objBlock.Rotation / Math.PI) * 200
        Revo.MyCommands.PaletteTxT.cmdNext.Visible = True
        Revo.MyCommands.PaletteTxT.txtBlockID.Text = objBlock.ObjectID
        Revo.MyCommands.PaletteTxT.cmdNext.Text = "Modifier"
        Revo.MyCommands.PaletteTxT.lblTitle.Text = "Modification de texte ou symbole"


        Revo.MyCommands.PaletteTxT.objBlockModif = objBlock


        'Lecture de Domaine / Thème / Type

        Dim docXML As New System.Xml.XmlDocument
        Dim RVinfo As New Revo.RevoInfo
        If IO.File.Exists(RVinfo.PalletteXML) Then
            Try
                docXML.Load(RVinfo.PalletteXML)

                Dim Palettes As System.Xml.XmlNodeList
                Palettes = docXML.SelectNodes("/revo/palettes/palette/field/topic/object/symbol")

                For Each Palette As System.Xml.XmlNode In Palettes
                    If Palette.InnerText.ToUpper = objBlock.EffectiveName Then
                        Revo.MyCommands.PaletteTxT.cboCat.Text = Palette.ParentNode.ParentNode.ParentNode.Attributes.ItemOf(0).InnerText 'field / Domaine
                        Revo.MyCommands.PaletteTxT.cboTheme.Text = Palette.ParentNode.ParentNode.Attributes.ItemOf(0).InnerText 'topic / Thème
                        Revo.MyCommands.PaletteTxT.cboType.Text = Palette.ParentNode.Item("name").InnerText ' Name / Type
                        Exit For 'Quitte après la 1er trouvaille
                    End If
                Next

            Catch ex As Xml.XmlException
                MsgBox("Erreur de lecture XML : " & ex.Message & vbCrLf & RVinfo.PalletteXML)
            Catch
                MsgBox("Erreur de lecture XML")
            End Try
        End If


        'Lecture des attributs
        Dim varAttributes As Object
        Dim strNomAttr As String = ""
        varAttributes = objBlock.GetAttributes


        For j = LBound(varAttributes) To UBound(varAttributes)

            For i = 0 To Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls.Count - 1
                If Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls(i).Name Like "Div*" And Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls(i).Visible = True Then
                    For u = 0 To Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls(i).Controls.Count - 1
                        If Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls(i).Controls(u).Name Like "lbl*" And _
                             Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls(i).Controls(u).Text.ToUpper = varAttributes(j).TagString.ToString.ToUpper Then

                            For y = 0 To Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls(i).Controls.Count - 1
                                If Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls(i).Controls(y).Name Like "cbo*" Then
                                    Revo.MyCommands.PaletteTxT.TLayoutAtt.Controls(i).Controls(y).Text = varAttributes(j).TextString.ToString
                                End If
                            Next
                        End If
                    Next
                End If
            Next

        Next

        'Désactivation de Cat / Theme / Type
        Revo.MyCommands.PaletteTxT.cboCat.Enabled = False
        Revo.MyCommands.PaletteTxT.cboTheme.Enabled = False
        Revo.MyCommands.PaletteTxT.cboType.Enabled = False
        
       
    End Sub

    Public Sub EditPointBlock(ByRef objBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference)

        Dim Editable As Boolean = True
        'Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
        'Dim objBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference = acDoc.HandleToObject(objElement.Handle.ToString)

        'If objBlock.HasAttributes Then
        '    Dim StrBlockValide() As String = Split("IMPL_B_PTS_CONTROL%IMPL_B_PTS_PROJET%IMPL_PTS_CONTROL%IMPL_PTS_PROJET%MUT_PFP3_BORNE%MUT_PFP3_CHEVILLE%MUT_PFP3_CROIX%MUT_PFP3_PIEU%MUT_PTS_BF_BORNE%MUT_PTS_BF_CHEVILLE%MUT_PTS_BF_CROIX%MUT_PTS_BF_NON_MATE_DEFINI%MUT_PTS_BF_NON_MATE_NONDEFINI%MUT_PTS_BF_PIEU%MUT_PTS_COM_BORNE%MUT_PTS_COM_CHEVILLE%MUT_PTS_COM_CROIX%MUT_PTS_COM_NON_MATE_DEFINI%MUT_PTS_COM_NON_MATE_NONDEFINI%MUT_PTS_COM_PIEU%MUT_PTS_CS%MUT_PTS_OD%PFA1%PFA2%PFA3%PFP1_INACCESSIBLE%PFP1_STATIONNABLE%PFP2_INACCESSIBLE%PFP2_STATIONNABLE%PFP3_BORNE%PFP3_CHEVILLE%PFP3_CROIX%PFP3_PIEU%PTS_BF_BORNE%PTS_BF_CHEVILLE%PTS_BF_CROIX%PTS_BF_NON_MATE_DEFINI%PTS_BF_NON_MATE_NONDEFINI%PTS_BF_PIEU%PTS_COM_BORNE%PTS_COM_CHEVILLE%PTS_COM_CROIX%PTS_COM_NON_MATE_DEFINI%PTS_COM_NON_MATE_NONDEFINI%PTS_COM_PIEU%PTS_CS%PTS_OD%SOUT_B_EC_CE%SOUT_B_EC_CH%SOUT_B_EC_CU%SOUT_B_EC_DE%SOUT_B_EC_EX%SOUT_B_EC_GR%SOUT_B_EC_GU%SOUT_B_EC_OS%SOUT_B_EC_PR%SOUT_B_EC_RE%SOUT_B_EC_SE%SOUT_B_EG_CB%SOUT_B_EG_CD%SOUT_B_EG_CE%SOUT_B_EG_PR%SOUT_B_EL_CD%SOUT_B_EL_POT%SOUT_B_EL_PR%SOUT_B_EM_CE%SOUT_B_EM_CH%SOUT_B_EM_OS%SOUT_B_EM_PR%SOUT_B_EP_BH%SOUT_B_EP_CC%SOUT_B_EP_CO%SOUT_B_EP_CVA%SOUT_B_EP_FO%SOUT_B_EP_MA%SOUT_B_EP_PR%SOUT_B_EP_PU%SOUT_B_EP_RE%SOUT_B_EP_VA%SOUT_B_EU_CE%SOUT_B_EU_CH%SOUT_B_EU_FS%SOUT_B_EU_OS%SOUT_B_EU_PR%SOUT_B_EU_RE%SOUT_B_EU_SE%SOUT_B_GA_MA%SOUT_B_GA_PR%SOUT_B_GA_RE%SOUT_B_GA_VA%SOUT_B_PROJ_EC_CE%SOUT_B_PROJ_EC_CH%SOUT_B_PROJ_EC_CU%SOUT_B_PROJ_EC_DE%SOUT_B_PROJ_EC_EX%SOUT_B_PROJ_EC_GR%SOUT_B_PROJ_EC_GU%SOUT_B_PROJ_EC_OS%SOUT_B_PROJ_EC_PR%SOUT_B_PROJ_EC_RE%SOUT_B_PROJ_EC_SE%SOUT_B_PROJ_EP_BH%SOUT_B_PROJ_EP_CC%SOUT_B_PROJ_EP_CO%SOUT_B_PROJ_EP_CVA%SOUT_B_PROJ_EP_FO%SOUT_B_PROJ_EP_PR%SOUT_B_PROJ_EP_PU%SOUT_B_PROJ_EP_RE%SOUT_B_PROJ_EP_VA%SOUT_B_PROJ_EU_CE%SOUT_B_PROJ_EU_CH%SOUT_B_PROJ_EU_FS%SOUT_B_PROJ_EU_OS%SOUT_B_PROJ_EU_PR%SOUT_B_PROJ_EU_RE%SOUT_B_PROJ_EU_SE%SOUT_B_TT_DTT%SOUT_B_TT_POT%SOUT_B_TT_PR%SOUT_B_TT_RTT%SOUT_B_TV_ART%SOUT_B_TV_ATT%SOUT_B_TV_CT%SOUT_B_TV_PR%TOPO_B_ARBRE%TOPO_B_PTS%TOPO_PTS", "%")
        '    For i = 0 To StrBlockValide.Length - 1
        '        If StrBlockValide(i) = objBlock.EffectiveName.ToUpper Then
        '            Editable = True
        '            Exit For
        '        End If
        '    Next
        'End If

        If Editable Then

            Dim j As Object

            'Charge la boite de dialogue d'édition de point
            Dim FrmPts As New frmInsertionPts

            Dim strCat As String = ""
            Dim varAttributes As Object
            Dim strNomAttr As String = ""
            With FrmPts

                .Show()
                '.Hide()
                '.ShowDialog()
                '.Visible = False

                'Paramètres et textes généraux
                .Tag = "EDIT"
                .Text = "Edition de point"
                '.lblTitle.Text = "Edition de point"

                'Supprimer la saisie en séries
                .chkDigitSerie.Visible = False


                'Catégorie
                If InStr(objBlock.Layer, "_") <> 0 Then strCat = Left(objBlock.Layer, InStr(objBlock.Layer, "_") - 1)
                .cboCat.Text = strCat

                Select Case strCat

                    Case "MO", "MUT"

                        'Entité et type => affichage des attributs grâce à l'événement "_click"
                        Select Case objBlock.Layer

                            Case "MO_PFA1"
                                .cboTheme.SelectedIndex = 0
                                .cboType.SelectedIndex = 0
                            Case "MO_PFA2"
                                .cboTheme.SelectedIndex = 0
                                .cboType.SelectedIndex = 1
                            Case "MO_PFA3"
                                .cboTheme.SelectedIndex = 0
                                .cboType.SelectedIndex = 2
                            Case "MO_PFP1"
                                .cboTheme.SelectedIndex = 0
                                .cboType.SelectedIndex = 3
                            Case "MO_PFP2"
                                .cboTheme.SelectedIndex = 0
                                .cboType.SelectedIndex = 4
                            Case "MO_PFP3"
                                .cboTheme.SelectedIndex = 0
                                .cboType.SelectedIndex = 5
                            Case "MUT_PFP3"
                                .cboTheme.SelectedIndex = 0
                                .cboType.SelectedIndex = 0
                            Case "MO_CS_PTS", "MUT_CS_PTS"
                                .cboTheme.SelectedIndex = 1
                                .cboType.SelectedIndex = 0
                            Case "MO_OD_PTS", "MUT_OD_PTS"
                                .cboTheme.SelectedIndex = 2
                                .cboType.SelectedIndex = 0
                            Case "MO_BF_PTS", "MUT_BF_PTS"
                                .cboTheme.SelectedIndex = 3
                                .cboType.SelectedIndex = 0
                            Case "MO_COM_PTS", "MUT_COM_PTS"
                                .cboTheme.SelectedIndex = 4
                                .cboType.SelectedIndex = 0

                        End Select


                    Case "TOPO"
                        .cboTheme.SelectedIndex = 0
                        .cboType.SelectedIndex = 0

                        Select Case objBlock.Layer
                            Case "TOPO_PTS"
                                .cboSigne.SelectedIndex = 0
                            Case "TOPO_ARBRE"
                                .cboSigne.SelectedIndex = 1
                        End Select

                    Case "IMPL"
                        .cboTheme.SelectedIndex = 0
                        .cboType.SelectedIndex = 0

                        Select Case objBlock.Layer
                            Case "IMPL_CONT_PTS"
                                .cboSigne.SelectedIndex = 0
                            Case "IMPL_PROJ_PTS"
                                .cboSigne.SelectedIndex = 1
                        End Select

                    Case "SOUT"
                        'Entité, type et signe => affichage des attributs grâce à l'événement "_click"

                        If Left(objBlock.Layer, 7) = "SOUT_EL" Then 'Electricité
                            .cboTheme.SelectedIndex = 8
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_EL_PR"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_EL_CD"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_EL_POT"
                                    .cboSigne.SelectedIndex = 2
                            End Select

                        ElseIf Left(objBlock.Layer, 7) = "SOUT_TT" Then  'Téléphone
                            .cboTheme.SelectedIndex = 10
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_TT_PR"
                                    .cboSigne.SelectedIndex = 2
                                Case "SOUT_TT_DTT"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_TT_RTT"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_TT_POT"
                                    .cboSigne.SelectedIndex = 3
                            End Select

                        ElseIf Left(objBlock.Layer, 7) = "SOUT_EG" Then  'Eclairage public
                            .cboTheme.SelectedIndex = 7
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_EG_PR"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_EG_CD"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_EG_CB"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_EG_CE"
                                    .cboSigne.SelectedIndex = 2
                            End Select

                        ElseIf Left(objBlock.Layer, 7) = "SOUT_GA" Then  'Gaz
                            .cboTheme.SelectedIndex = 9
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_GA_PR"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_GA_VA"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_GA_MA"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_GA_RE"
                                    .cboSigne.SelectedIndex = 2
                            End Select

                        ElseIf Left(objBlock.Layer, 7) = "SOUT_EP" Then  'Eau potable
                            .cboTheme.SelectedIndex = 0
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_EP_PR"
                                    .cboSigne.SelectedIndex = 5
                                Case "SOUT_EP_VA"
                                    .cboSigne.SelectedIndex = 9
                                Case "SOUT_EP_CVA"
                                    .cboSigne.SelectedIndex = 2
                                Case "SOUT_EP_BH"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_EP_MA"
                                    .cboSigne.SelectedIndex = 4
                                Case "SOUT_EP_PU"
                                    .cboSigne.SelectedIndex = 6
                                Case "SOUT_EP_CO"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_EP_RE"
                                    .cboSigne.SelectedIndex = 7
                                Case "SOUT_EP_CC"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_EP_FO"
                                    .cboSigne.SelectedIndex = 8
                            End Select

                        ElseIf Left(objBlock.Layer, 12) = "SOUT_PROJ_EP" Then  'Eau potable projet
                            .cboTheme.SelectedIndex = 1
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_PROJ_EP_PR"
                                    .cboSigne.SelectedIndex = 5
                                Case "SOUT_PROJ_EP_VA"
                                    .cboSigne.SelectedIndex = 9
                                Case "SOUT_PROJ_EP_CVA"
                                    .cboSigne.SelectedIndex = 2
                                Case "SOUT_PROJ_EP_BH"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_PROJ_EP_MA"
                                    .cboSigne.SelectedIndex = 4
                                Case "SOUT_PROJ_EP_PU"
                                    .cboSigne.SelectedIndex = 6
                                Case "SOUT_PROJ_EP_CO"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_PROJ_EP_RE"
                                    .cboSigne.SelectedIndex = 7
                                Case "SOUT_PROJ_EP_CC"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_PROJ_EP_FO"
                                    .cboSigne.SelectedIndex = 8
                            End Select

                        ElseIf Left(objBlock.Layer, 7) = "SOUT_TV" Then  'TV
                            .cboTheme.SelectedIndex = 11
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_TV_PR"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_TV_ART"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_TV_ATT"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_TV_CT"
                                    .cboSigne.SelectedIndex = 2
                            End Select

                        ElseIf Left(objBlock.Layer, 7) = "SOUT_EC" Then  'Eaux claires
                            .cboTheme.SelectedIndex = 2
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_EC_PR"
                                    .cboSigne.SelectedIndex = 8
                                Case "SOUT_EC_CH"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_EC_CE"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_EC_GR"
                                    .cboSigne.SelectedIndex = 5
                                Case "SOUT_EC_SE"
                                    .cboSigne.SelectedIndex = 10
                                Case "SOUT_EC_OS"
                                    .cboSigne.SelectedIndex = 7
                                Case "SOUT_EC_RE"
                                    .cboSigne.SelectedIndex = 9
                                Case "SOUT_EC_EX"
                                    .cboSigne.SelectedIndex = 4
                                Case "SOUT_EC_GU"
                                    .cboSigne.SelectedIndex = 6
                                Case "SOUT_EC_DE"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_EC_CU"
                                    .cboSigne.SelectedIndex = 2
                            End Select

                        ElseIf Left(objBlock.Layer, 12) = "SOUT_PROJ_EC" Then  'Eaux claires projet
                            .cboTheme.SelectedIndex = 3
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_PROJ_EC_PR"
                                    .cboSigne.SelectedIndex = 8
                                Case "SOUT_PROJ_EC_CH"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_PROJ_EC_CE"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_PROJ_EC_GR"
                                    .cboSigne.SelectedIndex = 5
                                Case "SOUT_PROJ_EC_SE"
                                    .cboSigne.SelectedIndex = 10
                                Case "SOUT_PROJ_EC_OS"
                                    .cboSigne.SelectedIndex = 7
                                Case "SOUT_PROJ_EC_RE"
                                    .cboSigne.SelectedIndex = 9
                                Case "SOUT_PROJ_EC_EX"
                                    .cboSigne.SelectedIndex = 4
                                Case "SOUT_PROJ_EC_GU"
                                    .cboSigne.SelectedIndex = 6
                                Case "SOUT_PROJ_EC_DE"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_PROJ_EC_CU"
                                    .cboSigne.SelectedIndex = 2
                            End Select

                        ElseIf Left(objBlock.Layer, 7) = "SOUT_EU" Then  'Eaux usées
                            .cboTheme.SelectedIndex = 5
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_EU_PR"
                                    .cboSigne.SelectedIndex = 4
                                Case "SOUT_EU_CH"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_EU_CE"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_EU_FS"
                                    .cboSigne.SelectedIndex = 2
                                Case "SOUT_EU_SE"
                                    .cboSigne.SelectedIndex = 6
                                Case "SOUT_EU_OS"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_EU_RE"
                                    .cboSigne.SelectedIndex = 5
                            End Select

                        ElseIf Left(objBlock.Layer, 12) = "SOUT_PROJ_EU" Then  'Eaux usées projet
                            .cboTheme.SelectedIndex = 6
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_PROJ_EU_PR"
                                    .cboSigne.SelectedIndex = 4
                                Case "SOUT_PROJ_EU_CH"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_PROJ_EU_CE"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_PROJ_EU_FS"
                                    .cboSigne.SelectedIndex = 2
                                Case "SOUT_PROJ_EU_SE"
                                    .cboSigne.SelectedIndex = 6
                                Case "SOUT_PROJ_EU_OS"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_PROJ_EU_RE"
                                    .cboSigne.SelectedIndex = 5
                            End Select

                        ElseIf Left(objBlock.Layer, 7) = "SOUT_EM" Then  'Eaux unitaires
                            .cboTheme.SelectedIndex = 4
                            .cboType.SelectedIndex = 0

                            Select Case objBlock.Layer
                                Case "SOUT_EM_PR"
                                    .cboSigne.SelectedIndex = 3
                                Case "SOUT_EM_CH"
                                    .cboSigne.SelectedIndex = 0
                                Case "SOUT_EM_CE"
                                    .cboSigne.SelectedIndex = 1
                                Case "SOUT_EM_OS"
                                    .cboSigne.SelectedIndex = 2
                            End Select

                        End If

                End Select


                'Lecture des attributs
                varAttributes = objBlock.GetAttributes
                Dim Altinfo As Double = 0

                On Error Resume Next

                For j = LBound(varAttributes) To UBound(varAttributes)

                    'MsgBox varAttributes(j).TextString

                    'En fonction du nom de l'attribut
                    Select Case UCase(varAttributes(j).TagString)

                        Case "IDENTDN"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .txtIdentDN.Text = varAttributes(j).TextString
                        Case "NUMERO"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .txtNo.Text = varAttributes(j).TextString
                        Case "SIGNE"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .cboSigne.Text = varAttributes(j).TextString
                        Case "ACCESSIBILITE"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .cboAttrVariable.Text = varAttributes(j).TextString
                        Case "FICHE"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .cboAttrVariable.Text = varAttributes(j).TextString
                        Case "ANC_BORNE_SPECIAL"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .cboAttrVariable.Text = varAttributes(j).TextString
                        Case "DEFINI_EXACTEMENT"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .cboExact.Text = varAttributes(j).TextString
                        Case "BORNE_TERRITORIALE"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .cboAttrVariable.Text = varAttributes(j).TextString
                        Case "PRECPLAN"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .txtPrecPlan.Text = Replace(varAttributes(j).TextString, "?", "")
                        Case "PRECALT"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .txtPrecAlt.Text = Replace(varAttributes(j).TextString, "?", "")
                        Case "FIABPLAN"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .cboFiabPlan.Text = varAttributes(j).TextString
                        Case "FIABALT"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .cboFiabAlt.Text = varAttributes(j).TextString
                        Case "COMMENTAIRE"
                            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet varAttributes().TextString. Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            .txtComment.Text = Replace(varAttributes(j).TextString, "?", "")
                        Case "ALTITUDE"
                            If IsNumeric(CDbl(varAttributes(j).TextString)) Then Altinfo = CDbl(varAttributes(j).TextString)
                    End Select

                Next

                On Error GoTo 0


                'Cas spéciaux: commune et plan
                If .txtCom.Enabled = True Then .txtCom.Text = CStr(Val(Mid(.txtIdentDN.Text, 4, 3)))
                If .txtPlan.Enabled = True Then .txtPlan.Text = CStr(Val(Right(.txtIdentDN.Text, 4)))


                'Coordonnées
                .txtX.Text = Format(objBlock.InsertionPoint(0), "0.000")
                .txtY.Text = Format(objBlock.InsertionPoint(1), "0.000")
                '.txtZ.Text = Format(objBlock.InsertionPoint(2), "0.000")

                'Traitement de l'altitude
                If objBlock.InsertionPoint(2) <> 0 Then '  Si Z<>0 -> 3D 
                    .Alt3D.Checked = True : .Alt2D.Checked = False
                    .txtZ.Text = Format(objBlock.InsertionPoint(2), "0.000")
                Else ' Si(Z = 0 -> 2D)
                    If Altinfo <> 0 Then .txtZ.Text = Format(Altinfo, "0.000") '(Altinfo())
                End If

                'ID du block (handle) pour édition des attributs
                .txtBlockID.Text = objBlock.Handle



                'Affiche le dialogue
                '.ShowDialog()
                '.Visible = True

            End With
        End If
    End Sub



    Public Function BlockNameToNature(ByRef strBlockName As String) As Short

        Dim intNat As Short

        Select Case strBlockName

            Case "SOUT_B_EL_PR" : intNat = 1
            Case "SOUT_B_EL_CD" : intNat = 2
            Case "SOUT_B_EL_POT" : intNat = 3

            Case "SOUT_B_TT_PR" : intNat = 4
            Case "SOUT_B_TT_DTT" : intNat = 5
            Case "SOUT_B_TT_RTT" : intNat = 6
            Case "SOUT_B_TT_POT" : intNat = 7

            Case "TOPO_B_PTS", "TOPO_PTS" : intNat = 8
            Case "TOPO_B_ARBRE" : intNat = 9

            Case "PFP3_Borne" : intNat = 11
            Case "PFP3_Cheville" : intNat = 12
            Case "PFP3_Croix" : intNat = 13
            Case "PFP3_Pieu" : intNat = 14
            Case "PFP2_stationnable" : intNat = 16
            Case "PFP2_stationnable" : intNat = 17
            Case "PFP2_stationnable" : intNat = 18
            Case "PFP2_inaccessible" : intNat = 19

            Case "SOUT_B_EG_PR" : intNat = 20
            Case "SOUT_B_EG_CD" : intNat = 21
            Case "SOUT_B_EG_CB" : intNat = 22
            Case "SOUT_B_EG_CE" : intNat = 23

            Case "PTS_CS" : intNat = 25
            Case "PTS_CS" : intNat = 26

            Case "SOUT_B_GA_PR" : intNat = 30
            Case "SOUT_B_GA_VA" : intNat = 31
            Case "SOUT_B_GA_MA" : intNat = 32
            Case "SOUT_B_GA_RE" : intNat = 33

            Case "PTS_OD" : intNat = 35
            Case "PTS_OD" : intNat = 36

            Case "IMPL_B_PTS_PROJET", "IMPL_PTS_PROJET" : intNat = 38
            Case "IMPL_B_PTS_CONTROL", "IMPL_PTS_CONTROL" : intNat = 39

            Case "SOUT_B_EP_PR" : intNat = 40
            Case "SOUT_B_EP_VA" : intNat = 41
            Case "SOUT_B_EP_CVA" : intNat = 42
            Case "SOUT_B_EP_BH" : intNat = 43
            Case "SOUT_B_EP_MA" : intNat = 44
            Case "SOUT_B_EP_PU" : intNat = 45
            Case "SOUT_B_EP_CO" : intNat = 46
            Case "SOUT_B_EP_RE" : intNat = 47
            Case "SOUT_B_EP_CC" : intNat = 48
            Case "SOUT_B_EP_FO" : intNat = 49

            Case "SOUT_B_TV_PR" : intNat = 50
            Case "SOUT_B_TV_ART" : intNat = 51
            Case "SOUT_B_TV_ATT" : intNat = 52
            Case "SOUT_B_TV_CT" : intNat = 53

            Case "Pts_BF_Borne" : intNat = 61
            Case "Pts_BF_Cheville" : intNat = 62
            Case "Pts_BF_Croix" : intNat = 63
            Case "Pts_BF_Pieu" : intNat = 64
            Case "Pts_BF_Non_mate_defini" : intNat = 65
            Case "Pts_BF_Non_mate_nondefini" : intNat = 66

            Case "SOUT_B_EC_PR" : intNat = 70
            Case "SOUT_B_EC_CH" : intNat = 71
            Case "SOUT_B_EC_CE" : intNat = 72
            Case "SOUT_B_EC_GR" : intNat = 73
            Case "SOUT_B_EC_SE" : intNat = 74
            Case "SOUT_B_EC_OS" : intNat = 75
            Case "SOUT_B_EC_RE" : intNat = 76
            Case "SOUT_B_EC_EX" : intNat = 77
            Case "SOUT_B_EC_GU" : intNat = 78
            Case "SOUT_B_EC_DE" : intNat = 79

            Case "SOUT_B_EU_PR" : intNat = 80
            Case "SOUT_B_EU_CH" : intNat = 81
            Case "SOUT_B_EU_CE" : intNat = 82
            Case "SOUT_B_EU_FS" : intNat = 83
            Case "SOUT_B_EU_SE" : intNat = 84
            Case "SOUT_B_EU_OS" : intNat = 85
            Case "SOUT_B_EU_RE" : intNat = 86

            Case "SOUT_B_EM_PR" : intNat = 90
            Case "SOUT_B_EM_CH" : intNat = 91
            Case "SOUT_B_EM_CE" : intNat = 92
            Case "SOUT_B_EM_OS" : intNat = 95


            Case Else : intNat = 0

        End Select

        BlockNameToNature = intNat

    End Function


    Public Function RadierElements() As Object


        Dim CollBlock As New Collection
        Dim CollPoly As New Collection

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            Dim acOPrompt As PromptSelectionOptions = New PromptSelectionOptions()
            acOPrompt.MessageForAdding = "Sélectionnez les objets à radier, puis appuyez sur [ENTER]."
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
                        Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                        If Not IsDBNull(acEnt) Then
                            If TypeName(acEnt) = "BlockReference" Then
                                CollBlock.Add(acEnt)
                            ElseIf TypeName(acEnt) = "Polyline" Then
                                CollPoly.Add(acEnt)
                            End If
                        End If
                    End If
                Next
            End If
            ' End selection of objects


            Dim strLayer As String
            Dim Numi As Double = 0
            Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, True, True)
            Dim btrSpace As BlockTableRecord = DirectCast(acTrans.GetObject(bt(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), BlockTableRecord)
            Dim NombreObjOK As Double = 0

            If CollBlock.Count > 0 Then 'objSelSetRad

                'Initialisation des infos d'état d'avancement
                conn.Message("Radiation d'objets", "Patienter pendant le déplacement des objets : 0 / " & CollBlock.Count & " radié (0%)" & vbCrLf & "Transfert de la sélection dans la catégorie 'RAD'", False, 0, CollBlock.Count)

                'Traitement des blocs
                For Each AcBlock As BlockReference In CollBlock

                    'MO BF/CS/OD/PFP3 uniquement
                    If Left(AcBlock.Layer, 5) = "MO_BF" Or Left(AcBlock.Layer, 5) = "MO_CS" Or Left(AcBlock.Layer, 5) = "MO_OD" Or Left(AcBlock.Layer, 7) = "MO_PFP3" Then

                        'Etat d'avancement
                        conn.Message("Radiation d'objets", "Traiment de " & AcBlock.Layer & "...", False, 0, CollBlock.Count)

                        'Déplacement dans RAD
                        strLayer = Replace(AcBlock.Layer, "MO_", "RAD_")
                        Dim NewNameBL As String = "RAD_" & AcBlock.Name

                        'Le bloc doit être recréé dans la nouvelle couche pour que les attributs suivent également
                        Try
                            If modAcad.IsBlockIN(NewNameBL) Then 'Si block existant dans le dessin

                                Dim NewBL As BlockReference '= CType(acTrans.GetObject(IDobjBL, OpenMode.ForWrite), BlockReference)
                                Dim NewBLid As ObjectId
                                Dim btrSource As BlockTableRecord = GetBlock(NewNameBL, "", acCurDb) ', strBlockNewName)
                                Dim btr As BlockTableRecord = acTrans.GetObject(btrSource.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, True, True)

                                'if getblock returns nothing then the rest of this code is worthless, skip it and dispose the transaction
                                If btrSource IsNot Nothing Then

                                    Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                    Dim attColl As AttributeCollection
                                    Dim ent As Entity
                                    NewBL = New BlockReference(AcBlock.Position, btr.ObjectId)
                                    acBlkTblRec.AppendEntity(NewBL)
                                    acTrans.AddNewlyCreatedDBObject(NewBL, True)
                                    attColl = NewBL.AttributeCollection
                                    For Each oid As ObjectId In btr 'accepts only key string
                                        ent = acTrans.GetObject(oid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, True, True)
                                        If TypeOf ent Is AttributeDefinition Then
                                            Dim attdef As AttributeDefinition = ent
                                            Dim attref As New AttributeReference
                                            attref.SetAttributeFromBlock(attdef, NewBL.BlockTransform)
                                            attref.TextString = attdef.TextString
                                            attColl.AppendAttribute(attref)
                                            acTrans.AddNewlyCreatedDBObject(attref, True)
                                        End If
                                    Next oid
                                    NewBL.Layer = strLayer 'Couche cible
                                    NewBL.ScaleFactors = New Scale3d(Val(GetCurrentScale()) / 1000)
                                    NewBLid = NewBL.ObjectId

                                    'Copie des attributs
                                    Dim NewBLLL As BlockReference = CType(acTrans.GetObject(NewBLid, OpenMode.ForRead), BlockReference)

                                    'Get the block attributue collection
                                    Dim attCol As AttributeCollection = AcBlock.AttributeCollection
                                    Dim NewattCol As AttributeCollection = NewBLLL.AttributeCollection
                                    'Loop through the attribute collection
                                    For Each attId As ObjectId In attCol
                                        'Get this attribute reference
                                        Dim attRef As AttributeReference = DirectCast(acTrans.GetObject(attId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), AttributeReference)
                                        Try 'Store its name
                                            For Each NewAttID As ObjectId In NewattCol
                                                Dim NewattRef As AttributeReference = DirectCast(acTrans.GetObject(NewAttID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), AttributeReference)
                                                If attRef.Tag.ToUpper = NewattRef.Tag.ToUpper Then NewattRef.TextString = attRef.TextString
                                            Next
                                        Catch ex As Exception
                                        End Try
                                    Next

                                    ' Efface(l) 'ancien bloc
                                    Dim AcBlockDel As BlockReference = CType(acTrans.GetObject(AcBlock.ObjectId, OpenMode.ForWrite), BlockReference)
                                    AcBlockDel.Erase()
                                    NombreObjOK += 1 ' Objet OK

                                End If
                            Else
                                MsgBox("Erreur de gabarit: calque " & strLayer & " inexistant !")
                            End If

                        Catch ex As Exception
                            MsgBox("Erreur de la fonction Radier MO." & vbCrLf & ex.Message)
                        End Try
                    End If

                    'Etat d'avancement
                    Numi += 1
                    conn.Message("Radiation d'objets (block)", "Block en traitement : " & Numi & " / " & CollBlock.Count, False, Numi, CollBlock.Count)

                Next

                Dim RVscript As New Revo.RevoScript 'Mise à jour des échelles
                RVscript.cmdCmd("#Cmd;ANNOAUTOSCALE|4|_CANNOSCALE|1:500|_CANNOSCALE|1:1000|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1")

            Else

                '    MsgBox("Opération annulée: aucun objet radié !", MsgBoxStyle.Exclamation, "Aucune sélection")
            End If



            'Traitement des polylignes

            Numi = 0
            'Copie dans MUT (option)'Demande la première fois
            Dim intCopieMut As Boolean = False
            If CollPoly.Count <> 0 Then
                If MsgBox("Copier les polylignes dans la catégorie 'MUT' avant de le radier (catégoire 'RAD')", MsgBoxStyle.YesNo + vbInformation, "Radiation") = MsgBoxResult.Yes Then intCopieMut = True
                Dim CollPolID As New ObjectIdCollection

                For Each AcPoly As Polyline In CollPoly
                    Dim NewPoly As Polyline = CType(acTrans.GetObject(AcPoly.ObjectId, OpenMode.ForWrite), Polyline)

                    If intCopieMut Then 'OUI
                        strLayer = Replace(AcPoly.Layer, "MO_", "MUT_")
                        Dim acBlkTblRec As BlockTableRecord = acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                        'Add the cloned circle
                        Dim CopyPoly As Polyline = AcPoly.Clone
                        CopyPoly.Layer = strLayer
                        acBlkTblRec.AppendEntity(CopyPoly)
                        acTrans.AddNewlyCreatedDBObject(CopyPoly, True)
                        CollPolID.Add(NewPoly.ObjectId)
                    End If

                    'Déplacement dans RAD
                    strLayer = Replace(AcPoly.Layer, "MO_", "RAD_")
                    NewPoly.Layer = strLayer
                    NombreObjOK += 1 ' Objet OK

                    'Etat d'avancement
                    Numi += 1
                    conn.Message("Radiation d'objets (polylignes)", "Polylignes en traitement  : " & Numi & " / " & CollPoly.Count, False, Numi, CollPoly.Count)
                Next


                'Placer les poylignes en arrière fond
                Dim btDirect As BlockTable = DirectCast(acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = DirectCast(acTrans.GetObject(btDirect.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                Dim dat As DrawOrderTable = DirectCast(acTrans.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite), DrawOrderTable)
                'Move the objects in the collection to the top
                Try
                    dat.MoveToBottom(CollPolID)
                Catch
                End Try
            End If

            'Fin
            conn.Message("Radiation d'objets", "Objets radiés : " & NombreObjOK & "          Objets ignorés : " & CollBlock.Count + CollPoly.Count - NombreObjOK, False, 99, 100, "info")
            conn.Message("Radiation d'objets", "", True, 99, 100)

            acTrans.Commit() ' Commit the transaction

        End Using

        Return True

    End Function

    Public Function ModelSpaceInsertBlock(ByVal InsPt As Point3d, ByVal BlockName As String, ByVal XScale As Double, _
 ByVal YScale As Double, ByVal ZScale As Double) As ObjectId
        Dim ObjID As ObjectId = Nothing
        Dim myBlkID As ObjectId = ObjectId.Null
        Try

            Dim myDwg As Document = Application.DocumentManager.MdiActiveDocument
            Using myTrans As Transaction = myDwg.TransactionManager.StartTransaction

                'Open the database for Write
                Dim myBT As BlockTable = myDwg.Database.BlockTableId.GetObject(OpenMode.ForRead)
                Dim myModelSpace As BlockTableRecord = myBT(BlockTableRecord.ModelSpace).GetObject(OpenMode.ForWrite)

                'Insert the Block
                Dim myBlockDef As BlockTableRecord = myBT(BlockName).GetObject(OpenMode.ForRead)
                Dim myBlockRef As New Autodesk.AutoCAD.DatabaseServices.BlockReference(InsPt, myBT(BlockName))
                myBlockRef.ScaleFactors = New Scale3d(XScale, YScale, ZScale)

                myModelSpace.AppendEntity(myBlockRef)
                myTrans.AddNewlyCreatedDBObject(myBlockRef, True)
                myBlkID = myBlockRef.ObjectId

                'Append Attribute References to the BlockReference
                Dim myAttColl As AttributeCollection
                Dim myEntID As ObjectId
                Dim myEnt As Autodesk.AutoCAD.DatabaseServices.Entity
                myAttColl = myBlockRef.AttributeCollection
                For Each myEntID In myBlockDef
                    myEnt = myEntID.GetObject(OpenMode.ForWrite)
                    If TypeOf myEnt Is AttributeDefinition Then
                        Dim myAttDef As AttributeDefinition = CType(myEnt, AttributeDefinition)
                        Dim myAttRef As New AttributeReference
                        myAttRef.SetAttributeFromBlock(myAttDef, myBlockRef.BlockTransform)
                        myAttColl.AppendAttribute(myAttRef)
                        myTrans.AddNewlyCreatedDBObject(myAttRef, True)
                    End If
                Next

                ObjID = myBlkID

                'Commit the Transaction
                myTrans.Commit()
            End Using

        Catch ex As System.Exception
            myBlkID = ObjectId.Null
        End Try

        Return ObjID
    End Function 'Receives block id then inserts at point


    Public Function ValiderMutations()

        Dim RVinfo As New Revo.RevoInfo
        Dim Connect As New Revo.connect
        Dim MutationXML As String = RVinfo.ConfigMOVD '(Config-MOVD.xml) anciennement : mutation.dat

        'Initialisation des infos d'état d'avancement
        Connect.Message("Validation de la mutation", "Transfert des éléments ""MUT"" dans la catégorie ""MO""", False, 10, 100)

        'Analyse des exceptions: données à supprimer, ignorer ou déplacer ?
        Dim ExeptionLay As New List(Of String)
        Dim ExeptionLayOp As New List(Of String)
        Try
            If File.Exists(MutationXML) Then  ' <validate-mutation>  <exeption layer="MUT_ADHOC">DELETE
                Dim configXML As New System.Xml.XmlDocument
                configXML.Load(MutationXML)
                Dim Xroot As String = "revo/validate-mutation"
                Dim NodeExeptions As System.Xml.XmlNodeList
                NodeExeptions = configXML.SelectNodes(Xroot & "/exeption")
                For Each NodeExeption As System.Xml.XmlNode In NodeExeptions
                    ExeptionLay.Add(NodeExeption.Attributes.ItemOf(0).InnerText.ToUpper)
                    ExeptionLayOp.Add(NodeExeption.InnerText.ToUpper)
                Next
            Else
                Connect.RevoLog(Connect.DateLog & "Validation Mutation" & vbTab & False & vbTab & "Erreur lecture: " & vbTab & MutationXML)
            End If
        Catch
            Connect.RevoLog(Connect.DateLog & "Validation Mutation" & vbTab & False & vbTab & "Erreur lecture: " & vbTab & MutationXML)
        End Try


        'Sélection des calques "MUT" (et comptage des objets)
        Dim Values() As TypedValue
        Dim Collent As New ObjectIdCollection
        Dim CollBL As New Collection
        Dim ObjValide As Double = 0
        Dim ObjDeleted As Double = 0

        Values = New TypedValue() {New TypedValue(DxfCode.Start, "INSERT,LWPOLYLINE,LINE"), New TypedValue(DxfCode.LayerName, "MUT_*,RAD_*")}
        'Select all items using the filter values
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then ' If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'Declare a objectidcollection (required for the MoveToTop method of the draw order table)

            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Get the blocktablerecord for modelspace
                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                'Loop through the found objects adding them to the objectidcollection
                For Each oid As ObjectId In pres.Value.GetObjectIds
                    Collent.Add(oid)

                    'POUR CHAQUE OBJET MUT
                    '**********************

                    Try 'Get the entity from the database
                        Dim AcObj As Entity = DirectCast(trans.GetObject(oid, OpenMode.ForWrite), Entity)
                        Dim AcObjType As String = AcObj.GetType.ToString
                        AcObjType = AcObjType.Substring(AcObjType.LastIndexOf(".") + 1).ToUpper
                        Dim StrOperation As String = "MOVE"

                        'Si l'objet est dans un calque avec une expeption (dans mutation)
                        'Test Layer exeption si : 'DELETE ou IGNORE  (ou MOVE)
                        '  If ExeptionLay.Contains(AcObj.Layer.ToUpper)  Then
                        For u = 0 To ExeptionLay.Count - 1 ' DELETE ou IGNORE
                            If AcObj.Layer.ToUpper Like ExeptionLay(u) Then
                                StrOperation = ExeptionLayOp(u)
                                Exit For
                            End If
                        Next
                        'End If

                        If StrOperation = "MOVE" Then ' Operation par défault : MOVE
                            'Déplacement du calque MUT dans le calque MO (POLYLINE + LINE + BLOCKREFERENCE)

                            AcObj.Layer = Replace(AcObj.Layer, "MUT_", "MO_")      'Nouveau calque pour les objets déplacés
                            ObjValide += 1 ' OBjet Déplacé + supprimé

                            If AcObjType = "BLOCKREFERENCE" Then   'BLOCKREFERENCE
                                Dim BLObj As BlockReference = DirectCast(trans.GetObject(oid, OpenMode.ForWrite), BlockReference)
                                'Attention aux cas particuliers ou l'objet mutation est supprimé et non déplacé !
                                CollBL.Add(BLObj)
                            End If

                        ElseIf StrOperation = "DELETE" Then ' Operation par défault : MOVE
                            AcObj.Erase() 'supprimé
                            ObjDeleted += 1 ' OBjet Déplacé + supprimé
                        Else
                            'Elément ignoré / non "touché"
                        End If

                        'Etat d'avancement
                        Connect.Message("Validation de la mutation : " & ObjValide & "/" & Collent.Count, "Transfert des éléments 'MUT' dans la catégorie 'MO'", False, 60, 100)

                    Catch
                        MsgBox("Erreur Validation de la mutations")
                    End Try
                Next
                trans.Commit()
            End Using
        End If

        'Remplace les Blocks
        For Each BLObj As BlockReference In CollBL
            ReplaceBlock(BLObj, Replace(BLObj.Name, "MUT_", "")) 'Le bloc doit être recréé dans la nouvelle couche pour que les attributs suivent également
        Next

        If Collent.Count = 0 Then 'Si trouve objet
            'Aucune mutation à valider
            Connect.Message("Validation des mutations actives", "Aucune mutation à valider dans le projet courant ", False, 100, 100, "critical")
            Return -1
        Else
            Connect.Message("Validation des mutations actives", ObjValide & " objets validés avec succès (passé de la catégorie MUT à MO)" & vbCrLf & _
                             ObjDeleted & " objets supprimés avec succès", False, 100, 100, "info")
            Connect.Message("Fin", "... ", True, 0, 0, "hide") 'Fin
            Return Collent.Count
        End If

    End Function

    Function InterlisProjectTest() As Boolean

        'Si nécessite modèles interlis :     
        Dim StrVersion As String = "0"
        ' Dim StrVersionJourCAD As String = "0"
        Dim Ass As New Revo.RevoInfo
        StrVersion = GetFileProperty("VersionRevo")
        If StrVersion = "" Then StrVersion = 0
        'Dim FileDraw As String = curDwg.Name
        'Dim SavedFile As Boolean = CDbl(Application.GetSystemVariable("DWGTITLED"))

        'Create a new drawing using the template
        If CDbl(StrVersion) >= Ass.TemplateVersion Then
            Return True  '  C'est tout ok : rien à faire
        Else
            Return False
        End If
    End Function


    ''' <summary>
    ''' Update Polyline for Batiment
    ''' </summary>
    ''' <param name="ObjID">Perimètre du bâtiment</param>
    ''' <returns>The objectid of the selected block</returns>
    ''' <remarks></remarks>
    Public Function UpdateXdataBatiment(ByVal ObjPoly As Polyline)

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        doc.LockDocument()

        'Mise à jour des Xdata des bâtiments 
        Dim CollPol As New Collection
        If ObjPoly Is Nothing Then
            CollPol = SelectPolyline(False, False, "Sélectionner les poylignes des bâtiments")
        Else
            CollPol.Add(ObjPoly)
        End If


        Try
            Dim NbreOK As Double = 0
            For Each objPolyline As Polyline In CollPol

                Dim strNo As String = 0
                Dim strDes As String = ""
                Dim Verts As Point3dCollection = GetPolyVertices(objPolyline.ObjectId)
                Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayerName, "MO_CS_BAT_NUM")}
                Dim IDs() As ObjectId = SelectWindowPoly(Values, Verts)


                If IDs Is Nothing Then
                    Zooming.ZoomToObject(objPolyline.ObjectId, 10)
                    SetBlockColour(objPolyline.ObjectId, Autodesk.AutoCAD.Colors.Color.FromColor(Drawing.Color.Blue))

                    Dim NumCOll As Collection
                    NumCOll = SelectObj(True, "Sélectionner le numéro du bâtiment", False, False, True)
                    For Each BL As BlockReference In NumCOll
                        Dim IDsTxt(0 To 0) As ObjectId
                        IDsTxt(0) = BL.ObjectId
                        IDs = IDsTxt
                        Exit For
                    Next

                    SetBlockColour(objPolyline.ObjectId, Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 256))

                End If


                If IDs IsNot Nothing Then

                    strNo = GetBlockAttribute(IDs(0), "NUMERO")
                    strDes = GetBlockAttribute(IDs(0), "DESIGNATION")
                    ' Dim doc As Document = Application.DocumentManager.MdiActiveDocument
                    Dim db As Database = HostApplicationServices.WorkingDatabase

                    Using tr As Transaction = db.TransactionManager.StartTransaction()

                        'Effacer les Xdata
                        ClearAllXData(objPolyline.ObjectId)

                        'Ecriture des Xdata
                        Dim obj As DBObject = tr.GetObject(objPolyline.ObjectId, OpenMode.ForWrite)
                        AddRegAppTableRecord("REVO")
                        Dim rb As New ResultBuffer(New TypedValue(1001, "REVO"), New TypedValue(1000, "ID" & strNo), New TypedValue(1000, strNo), New TypedValue(1000, strDes))
                        obj.XData = rb
                        rb.Dispose()
                        tr.Commit()

                        NbreOK += 1
                    End Using


                End If

            Next

            MsgBox(NbreOK & " bâtiment(s) ont été mis à jour avec succès !", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Mise à jour des numéros (" & NbreOK & "/" & CollPol.Count & ")")

        Catch ex As System.Exception
            MsgBox("Erreur Bâtiment Xdata update : " & ex.Message)
        End Try

        doc.LockDocument.Dispose()


            Return True

    End Function



    ''' <summary>
    ''' Check MO rules
    ''' </summary>
    ''' <remarks></remarks>
    Public Function CheckMO()

        'Mise à jour des Xdata des bâtiments 
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim NbreErreur As Double = 0
        Dim NbrePoly As Double = 0

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument


        Try
            doc.LockDocument()
            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim CheckBF As List(Of Double) = CheckSurfaceClosed(tr, "*_BF")
                NbreErreur += CheckBF(0)
                NbrePoly += CheckBF(1)

                Dim CheckCSBAT As List(Of Double) = CheckSurfaceClosed(tr, "*CS_BAT")
                NbreErreur += CheckCSBAT(0)
                NbrePoly += CheckCSBAT(1)
                Dim CheckCSSV As List(Of Double) = CheckSurfaceClosed(tr, "*CS_SV_*")
                NbreErreur += CheckCSSV(0)
                NbrePoly += CheckCSSV(1)
                Dim CheckCSRD As List(Of Double) = CheckSurfaceClosed(tr, "*CS_RD_")
                NbreErreur += CheckCSRD(0)
                NbrePoly += CheckCSRD(1)

                Dim CheckOD As List(Of Double) = CheckSurfaceClosed(tr, "*ODS_BATSOUT")
                NbreErreur += CheckOD(0)
                NbrePoly += CheckOD(1)

                ' ed.Regen()
                tr.Commit()
            End Using

            doc.LockDocument.Dispose()

        Catch ex As System.Exception
            MsgBox("Erreur d'analyse : " & ex.Message)
        End Try




        If NbreErreur = 0 Then
            MsgBox("Aucune erreur n'a été trouvée" & vbCrLf & "", MsgBoxStyle.Information + vbOKOnly, "Analyse terminée ( " & NbrePoly & " polylignes analysées)")
            Return True
        Else

            MsgBox("Nous avons trouvé " & NbreErreur & " erreur(s)" & vbCrLf & vbCrLf & "Les périmètres contenant une erreur ont été modifiés en bleu (couleur 5)." _
                          , MsgBoxStyle.Information + vbOKOnly, "Analyse terminée ( " & NbrePoly & " polylignes analysées)")
            Return False
        End If


    End Function

    Private Function CheckSurfaceClosed(tr As Transaction, LayerName As String) As List(Of Double)

        Dim Check As New List(Of Double)
        Dim NbrePoly As Double = 0
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, LayerName)}
        Dim sFilter As New SelectionFilter(Values)
        Dim pres As PromptSelectionResult = ed.SelectAll(sFilter)
        Dim PolyColl As New Collection
        Dim NbreErreur As Double = 0


        If pres.Status = PromptStatus.OK Then
            Dim ids() As ObjectId = pres.Value.GetObjectIds
            NbrePoly = ids.Length

            For i = 0 To ids.Length - 1


                'Charge polyligne
                Dim CheckPoly As Polyline = CType(tr.GetObject(ids(i), OpenMode.ForWrite), Polyline) 'DirectCast(tr.GetObject(ids(i), OpenMode.ForRead), Polyline)
                'As Polyline = CType(tr.GetObject(ids(i), OpenMode.ForWrite), Polyline)
                'tr.GetObject(ids(i), OpenMode.ForWrite)


                'Check si la polyligne est fermée
                If CheckPoly.Closed = False Then
                    NbreErreur += 1
                    PolyColl.Add(CheckPoly)
                    CheckPoly.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.None, 5)

                End If

            Next

        End If

        Check.Add(NbreErreur)
        Check.Add(NbrePoly)
        Return Check

    End Function

    Public Function ChangeScale(ByRef intScale As Short)


        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database


        'Chargement du facteur d'Annotation
        Dim EchPS As String = GetScalePS()


        'Test si le dessin est annotatif  avec le bloc   PTS_BF_BORNE
        Dim Annotatif As Boolean = False

        'Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        'Dim db As Database = doc.Database
        Using trx As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = trx.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
            If bt.Has("PTS_BF_BORNE") Then
                Dim btr As BlockTableRecord = trx.GetObject(bt("PTS_BF_BORNE"), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                If btr.Annotative = AnnotativeStates.True Then Annotatif = True
            End If
        End Using


        'Définition de l'échelle de papier
        SetFileProperty("PS-Scale", EchPS)


        'Définition de l'échelle
        SetFileProperty("Echelle", (intScale))



        Dim strEch As String
        Dim dblFacteEch As Double
        strEch = CStr(intScale)
        dblFacteEch = (intScale / 1000)




        '-------------------- Appliction de l'échelle sur les objets -------------------


#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If


        'Application du facteur aux types de lignes             ------------------------------------------------
        If EchPS <> 1 Then 'Si le facteur de l'échelle de l'espace papier =>  modifier le type de ligne
            CommandLine.CommandC("echltp", dblFacteEch) ' Mise à jour 2015 (29.09.2014)
        End If

        'Inutile d'aller plus loin (blocks et hachures) si l'espace objet est vide (aucun objet)
        If acDoc.ModelSpace.Count <> 0 Then



            'Message
            conn.Message("Mise à l'échelle", "Application du changement d'échelle ...", False, 40, 100)
            System.Windows.Forms.Application.DoEvents()


            'Modifie l'échelle de tous les blocs saud EXCEPTIONS ------------------------------------------------

            'Sélection
            Dim objSelSet As Autodesk.AutoCAD.Interop.AcadSelectionSet
            Dim objAcadBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference


            'Restrictions pour la sélection
            Dim adDXFCode(1) As Short
            Dim adDXFGroup(1) As Object


            objSelSet = acDoc.SelectionSets.Add("JourCAD" & acDoc.SelectionSets.Count) 'Nom de la sélection
            adDXFCode(0) = 0
            adDXFGroup(0) = "INSERT" 'Type d'objet
            adDXFCode(1) = 67
            adDXFGroup(1) = 0 'Espace objet ou papier: ModelSpace

            'Effectue la sélection
            objSelSet.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , adDXFCode, adDXFGroup)

            'Boucle pour tous les éléments sélectionnés (blocs) 
            If objSelSet.Count <> 0 Then

                For Each objAcadBlock In objSelSet

                    'Ne traite pas les références externes
                    If acDoc.Blocks.Item(objAcadBlock.Name).IsXRef = False Then

                        'Ne modifie pas l'échelle des arbres ' Retrait TH 10.05.2016 : objAcadBlock.Name <> "TOPO_B_ARBRE"

                        ' Ajout de l'expetion RESINEUX + FEUILLU TH 10.05.2016
                        If objAcadBlock.Name <> "TOPO_ARBRE_RESINEUX" And objAcadBlock.Name <> "TOPO_ARBRE_FEUILLU" Then

                            If objAcadBlock.Name = "TOPO_B_ARBRE" And objAcadBlock.IsDynamicBlock = True Then
                                'Si le bloc TOPO_B_ARBRE est dynamique : ignorer le bloc
                            Else

                                Dim NameBL As String = Replace(Replace(objAcadBlock.Name.ToUpper, "MUT_", ""), "RAD_", "")
                                Dim ListBL As String = "[AD][AD_LOCALISATION][ADHOC][B_CS_BAT][B_CS_NOM][B_CS_PROJ_BAT][B_CS_PROJ_NOM][B_CS_PROJ_SENS][B_CS_SENS][B_NO_LIEU][B_NO_LIEUDIT][B_NO_LOCAL][B_ODL_BAC][B_ODL_BAT][B_ODL_COUV][B_ODL_NOM][B_ODL_SENS][BIFFER_LIGNES][BIFFER_PTS][CHECK_ITF_AD][CHECK_ITF_BF][CHECK_ITF_CS][CHECK_ITF_DN][CHECK_ITF_LC][CHECK_ITF_NOM][CHECK_ITF_OD][CHECK_ITF_PFP][CHECK_ITF_RP][ENQ_ARBRE_ABAT][ENQ_ARBRES_EXIST][ENQ_SERV_REG_EC][ENQ_SERV_REG_EU][ETIQUETTE_ELEMENTS][FORMAT_TEMPLATE][IMMEUBLE_DDP][IMMEUBLE_NUM][IMMEUBLE_NUM_DDP][IMMEUBLE_NUM_DDP_PROJ][IMMEUBLE_NUM_PROJ][IMPL_B_PTS_CONTROL][IMPL_B_PTS_PROJET][NO_LIEU][NOM_CANTON][NOM_COMMUNE][NOM_COMMUNE_PROJ][NOM_DISTRICT][NOM_PAYS][NORD][NT][PFA1][PFA2][PFA3][PFP1_INACCESSIBLE][PFP1_STATIONNABLE][PFP2_INACCESSIBLE][PFP2_STATIONNABLE][PFP3_BORNE][PFP3_CHEVILLE][PFP3_CROIX][PFP3_PIEU][PLAN][PTS_BF_BORNE][PTS_BF_CHEVILLE][PTS_BF_CROIX][PTS_BF_NON_MATE_DEFINI][PTS_BF_NON_MATE_NONDEFINI][PTS_BF_PIEU][PTS_COM_BORNE][PTS_COM_CHEVILLE][PTS_COM_CROIX][PTS_COM_NON_MATE_DEFINI][PTS_COM_NON_MATE_NONDEFINI][PTS_COM_PIEU][PTS_CS][PTS_OD][QUADRI_DYN][SCHEMA_ORIEL_CLOCHE][SCHEMA_ORIEL_DIVERS][SCHEMA_STATIONS][SERV_EXIST_OBJ_BH][SERV_EXIST_OBJ_VANNE_EAU_P][SERV_EXIST_REG_EC][SERV_EXIST_REG_EU][SERV_EXIST_TXT_AUTRE][SERV_EXIST_TXT_EC][SERV_EXIST_TXT_EU][SERV_NOUV_OBJ_BH][SERV_NOUV_OBJ_VANNE_EAU_PO][SERV_NOUV_REG_EC][SERV_NOUV_REG_EU][SERV_NOUV_TXT][SOUT_B_EC_CE][SOUT_B_EC_CH][SOUT_B_EC_CU][SOUT_B_EC_DE][SOUT_B_EC_EX][SOUT_B_EC_GR][SOUT_B_EC_GU][SOUT_B_EC_OS][SOUT_B_EC_PR][SOUT_B_EC_RE][SOUT_B_EC_SE][SOUT_B_EG_CB][SOUT_B_EG_CD][SOUT_B_EG_CE][SOUT_B_EG_PR][SOUT_B_EL_CD][SOUT_B_EL_POT][SOUT_B_EL_PR][SOUT_B_EM_CE][SOUT_B_EM_CH][SOUT_B_EM_OS][SOUT_B_EM_PR][SOUT_B_EP_BH][SOUT_B_EP_CC][SOUT_B_EP_CO][SOUT_B_EP_CVA][SOUT_B_EP_FO][SOUT_B_EP_MA][SOUT_B_EP_PR][SOUT_B_EP_PU][SOUT_B_EP_RE][SOUT_B_EP_VA][SOUT_B_EU_CE][SOUT_B_EU_CH][SOUT_B_EU_FS][SOUT_B_EU_OS][SOUT_B_EU_PR][SOUT_B_EU_RE][SOUT_B_EU_SE][SOUT_B_GA_MA][SOUT_B_GA_PR][SOUT_B_GA_RE][SOUT_B_GA_VA][SOUT_B_PROJ_EC_CE][SOUT_B_PROJ_EC_CH][SOUT_B_PROJ_EC_CU][SOUT_B_PROJ_EC_DE][SOUT_B_PROJ_EC_EX][SOUT_B_PROJ_EC_GR][SOUT_B_PROJ_EC_GU][SOUT_B_PROJ_EC_OS][SOUT_B_PROJ_EC_PR][SOUT_B_PROJ_EC_RE][SOUT_B_PROJ_EC_SE][SOUT_B_PROJ_EP_BH][SOUT_B_PROJ_EP_CC][SOUT_B_PROJ_EP_CO][SOUT_B_PROJ_EP_CVA][SOUT_B_PROJ_EP_FO][SOUT_B_PROJ_EP_PR][SOUT_B_PROJ_EP_PU][SOUT_B_PROJ_EP_RE][SOUT_B_PROJ_EP_VA][SOUT_B_PROJ_EU_CE][SOUT_B_PROJ_EU_CH][SOUT_B_PROJ_EU_FS][SOUT_B_PROJ_EU_OS][SOUT_B_PROJ_EU_PR][SOUT_B_PROJ_EU_RE][SOUT_B_PROJ_EU_SE][SOUT_B_TT_DTT][SOUT_B_TT_POT][SOUT_B_TT_PR][SOUT_B_TT_RTT][SOUT_B_TV_ART][SOUT_B_TV_ATT][SOUT_B_TV_CT][SOUT_B_TV_PR][TOPO_ARBRE_FEUILLUS][TOPO_ARBRE_RESINEUX][TOPO_B_ARBRE][TOPO_B_BORD_ROUTE][TOPO_B_ESCALIER][TOPO_B_FAITE_CORNICHE][TOPO_B_HAUT_MUR][TOPO_B_PIED_MUR][TOPO_B_PTS][TOPO_B_PTS_COTE][TOPO_B_PTS_LIGNE_RUPTURE][TOPO_B_TROTTOIR][VAL_FIABLE][VAL_PFP12A][VAL_PFP12I][VAL_PFP3]"

                                If InStr(ListBL, "[" & NameBL & "]") <> 0 Then 'Check si bloc MOVD

                                    If Annotatif Then
                                        objAcadBlock.XScaleFactor = 1
                                        objAcadBlock.YScaleFactor = 1
                                        objAcadBlock.ZScaleFactor = 1
                                    Else
                                        objAcadBlock.XScaleFactor = dblFacteEch * CDbl(EchPS)
                                        objAcadBlock.YScaleFactor = dblFacteEch * CDbl(EchPS)
                                        objAcadBlock.ZScaleFactor = dblFacteEch * CDbl(EchPS)
                                    End If

                                End If


                            End If
                        End If

                        ' Expetion BLOC 1000


                    End If

                Next objAcadBlock

            End If


            If Annotatif Then 'Si annotatif ajoute l'echelle (si disponible)
                Dim RVscript As New Revo.RevoScript
                Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
                If db.Cannoscale.Name = "1:1000" Then
                    RVscript.cmdCmd("#Cmd;_CANNOSCALE|1:500|ANNOAUTOSCALE|4|_CANNOSCALE|1:1000|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1")
                    If intScale <> 1000 Then RVscript.cmdCmd("#cmd;_CANNOSCALE|1:" & intScale)
                ElseIf db.Cannoscale.Name = "1:" & intScale Then
                    RVscript.cmdCmd("#Cmd;_CANNOSCALE|1:1000|ANNOAUTOSCALE|4|_CANNOSCALE|1:" & intScale & "|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1")
                Else
                    RVscript.cmdCmd("#Cmd;ANNOAUTOSCALE|4|_CANNOSCALE|1:" & intScale & "|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1")
                End If

                RVscript.cmdCmd("#cmd;_ATTSYNC|_N|*")

            End If

            'Effacement de la sélection
            objSelSet.Delete()

            conn.Message("Mise à l'échelle", "Application du changement d'échelle ...", False, 60, 100)
            System.Windows.Forms.Application.DoEvents()



            'Modifie l'échelle des hachures de type "DOTS"    ------------------------------------------------
            conn.Message("Mise à l'échelle", "Application du changement d'échelle ...", False, 80, 100)
            System.Windows.Forms.Application.DoEvents()

            'Sélection
            'Dim objSelSet As AcadSelectionSet
            Dim objAcadHatch As Autodesk.AutoCAD.Interop.Common.AcadHatch

            'Restrictions pour la sélection
            Dim adDXFCode2(11) As Short
            Dim adDXFGroup2(11) As Object
            objSelSet = acDoc.SelectionSets.Add("JourCAD" & acDoc.SelectionSets.Count) 'Nom de la sélection
            adDXFCode2(0) = 0
            adDXFGroup2(0) = "HATCH" 'Type d'objet

            adDXFCode2(1) = -4
            adDXFGroup2(1) = "<or"

            adDXFCode2(2) = -4
            adDXFGroup2(2) = "<and"
            adDXFCode2(3) = 8
            adDXFGroup2(3) = "MO_*" 'Filtre uniquement les Hachures MO + MUT + RAD
            adDXFCode2(4) = -4
            adDXFGroup2(4) = "and>"

            adDXFCode2(5) = -4
            adDXFGroup2(5) = "<and"
            adDXFCode2(6) = 8
            adDXFGroup2(6) = "MUT_*" 'Filtre uniquement les Hachures MO + MUT + RAD
            adDXFCode2(7) = -4
            adDXFGroup2(7) = "and>"

            adDXFCode2(8) = -4
            adDXFGroup2(8) = "<and"
            adDXFCode2(9) = 8
            adDXFGroup2(9) = "RAD_*" 'Filtre uniquement les Hachures MO + MUT + RAD
            adDXFCode2(10) = -4
            adDXFGroup2(10) = "and>"

            adDXFCode2(11) = -4
            adDXFGroup2(11) = "or>"


            'Effectue la sélection
            objSelSet.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , adDXFCode2, adDXFGroup2)

            'Facteur de mise à l'échelle (différence entre prcédente et actuelle)
            Dim dblCorrEch As Double = dblFacteEch

            'Boucle pour tous les éléments sélectionnés
            If objSelSet.Count <> 0 Then

                Try 'Erreur si échelle trop grande (ex:  1:10'000 et qu'aun point (DOTS) n'est dans une surface)

                    For Each objAcadHatch In objSelSet

                        dblCorrEch = dblFacteEch

                        If objAcadHatch.Layer = "MO_CS_BAT_HACH" Then
                            dblCorrEch = dblFacteEch * 0.479999989271164
                        ElseIf objAcadHatch.Layer = "MO_CS_SSV_ROCHER_HACH" Then
                            dblCorrEch = dblFacteEch * 3.45000004768371
                        ElseIf objAcadHatch.Layer = "MO_CS_SSV_EBOULIS_HACH" Then
                            dblCorrEch = dblFacteEch * 3.45000004768371
                        ElseIf objAcadHatch.Layer = "MO_CS_SB_PAT_OUVERT_HACH" Then
                            dblCorrEch = dblFacteEch * 5.03999996185302
                        ElseIf objAcadHatch.Layer = "MO_CS_SB_PAT_DENSE_HACH" Then
                            dblCorrEch = dblFacteEch * 2.51999998092651
                        ElseIf objAcadHatch.Layer = "MO_CS_SB_FORET_HACH" Then
                            dblCorrEch = dblFacteEch * 1.25999999046325
                        End If

                        If Annotatif Then
                            objAcadHatch.PatternScale = dblCorrEch / EchPS
                        Else
                            objAcadHatch.PatternScale = dblCorrEch
                        End If

                        objAcadHatch.Evaluate()

                    Next objAcadHatch

                Catch ex As Exception
                End Try

            End If


            conn.Message("Mise à l'échelle", "Application du changement d'échelle ...", False, 100, 100)
            conn.Message("Mise à l'échelle", "", True, 100, 100)

            'Effacement de la sélection
            objSelSet.Delete()

            'Met à jour le dessin
            Application.DocumentManager.MdiActiveDocument.Editor.Regen()

        Else
            Return False
        End If


        Return True

    End Function


    Public Function MOAnnotatif(Active As Boolean)

        Dim RVscript As New Revo.RevoScript
        Dim ValActive As String

        If Active = True Then 'Activer le mode annotatif
            ValActive = "1"
        Else 'Supprimer le mode annotatif
            ValActive = "0"
        End If

        ActivateModel()

        ' Orientation paper space : Activer / désactiver

        RVscript.cmdBL("#BL;SERV_NOUV_TXT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_NOUV_REG_EU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_CE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_CU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_CH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_NOUV_OBJ_BH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_OD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_CS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_OBJ_BH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_TXT_AUTRE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_REG_EU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_SE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_RE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EG_CB;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EG_CE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EG_CD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_EX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_DE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_GR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_OS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_GU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_COUV;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_BAT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_NOM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_NOM_COMMUNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_SENS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_NO_LOCAL;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_CS_BAT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_AD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_CS_NOM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_NO_LIEUDIT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_CS_SENS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PLAN;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_NON_MATE_NONDEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PFP3_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_NT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_NO_LIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PFP3_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PFP3_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PFP3_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TT_DTT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_BH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TT_POT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TT_RTT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TT_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_GA_VA;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_SE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_RE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_GA_MA;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_GA_RE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_GA_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;VAL_PFP12A;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;VAL_FIABLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;VAL_PFP12I;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;VAL_PFP3;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;TOPO_B_PTS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TV_ATT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TV_ART;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TV_CT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;TOPO_B_ARBRE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TV_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EM_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EM_OS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_BH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_CO;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_CC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EM_CH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EL_CD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EG_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EL_POT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EM_CE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EL_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_CE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_VA;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_CH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_OS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_FS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_RE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_FO;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_CVA;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_MA;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_PU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_OD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_BIFFER_PTS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_BIFFER_LIGNES;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_CS_BAT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_CS_SENS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_CS_NOM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_ADHOC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;ENQ_SERV_REG_EC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_RP;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMPL_B_PTS_CONTROL;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_AD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMPL_B_PTS_PROJET;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NOM_COMMUNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NOM_CANTON;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NOM_DISTRICT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NO_LIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NOM_PAYS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_SENS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_NO_LOCAL;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_NO_LIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_BAC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_NOM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_COUV;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_PFP;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_NO_LIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_SENS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_NO_LIEUDIT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_BAC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_NO_LOCAL;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_PROJ_SENS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_BAT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;AD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_NOM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_PROJ_NOM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_PROJ_BAT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_DN;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_CS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_LC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_OD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_NOM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_BF;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_COUV;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_BAT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_NOM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_AD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_SENS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP3_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP3_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP3_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PLAN;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP3_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP2_STATIONNABLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFA3;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFA2;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP1_INACCESSIBLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP2_INACCESSIBLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP1_STATIONNABLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_NON_MATE_NONDEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_CS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_NON_MATE_DEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFA1;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_NON_MATE_NONDEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PLAN;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PFP3_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PFP3_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PFP3_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PFP3_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_DISTRICT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_COMMUNE_PROJ;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_PAYS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NORD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_COMMUNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_OD;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_CS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_SCHEMA_ORIEL_CLOCHE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_CANTON;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_SCHEMA_STATIONS;[[PaperOrientation]]" & ValActive & ";")

        RVscript.cmdBL("#BL;AD_LOCALISATION;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_AD_LOCALISATION;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_AD_LOCALISATION;[[PaperOrientation]]" & ValActive & ";")

        RVscript.cmdBL("#BL;ENQ_SERV_REG_EU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMMEUBLE_NUM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMMEUBLE_NUM_DDP;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMMEUBLE_NUM_DDP_PROJ;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMMEUBLE_NUM_PROJ;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_NO_LIEUDIT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_BAT;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_IMMEUBLE_NUM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_IMMEUBLE_NUM_DDP;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_NON_MATE_DEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_BORNE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_NON_MATE_DEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_NON_MATE_NONDEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_SCHEMA_ORIEL_DIVERS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_NON_MATE_NONDEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_NON_MATE_DEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_NO_LIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_BAC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_IMMEUBLE_DDP;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_IMMEUBLE_NUM;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_NON_MATE_DEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_CHEVILLE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_CROIX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_NON_MATE_DEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_NON_MATE_NONDEFINI;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_PIEU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_OBJ_VANNE_EAU_P;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_REG_EC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_TXT_EC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_TXT_EU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_NOUV_OBJ_VANNE_EAU_PO;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_NOUV_REG_EC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_CE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_CH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_CU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_DE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_EX;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_GR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_GU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_OS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_RE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_SE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_CC;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_CO;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_CVA;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_FO;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_PU;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_RE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_VA;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_CE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_CH;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_FS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_OS;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_PR;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_RE;[[PaperOrientation]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_SE;[[PaperOrientation]]" & ValActive & ";")




        ' Annotatif : Activer / désactiver

        RVscript.cmdBL("#BL;SERV_NOUV_TXT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_NOUV_REG_EU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_CE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_CU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_CH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_NOUV_OBJ_BH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_OD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_CS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_OBJ_BH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_TXT_AUTRE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_REG_EU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_SE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_RE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EG_CB;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EG_CE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EG_CD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_EX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_DE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_GR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_OS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EC_GU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_COUV;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_BAT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_NOM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_NOM_COMMUNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_SENS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_NO_LOCAL;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_CS_BAT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_AD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_CS_NOM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_NO_LIEUDIT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_CS_SENS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PLAN;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_NON_MATE_NONDEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PFP3_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_NT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_NO_LIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PFP3_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PFP3_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PFP3_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TT_DTT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_BH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TT_POT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TT_RTT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TT_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_GA_VA;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_SE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_RE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_GA_MA;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_GA_RE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_GA_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;VAL_PFP12A;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;VAL_FIABLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;VAL_PFP12I;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;VAL_PFP3;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;TOPO_B_PTS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TV_ATT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TV_ART;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TV_CT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;TOPO_B_ARBRE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_TV_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EM_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EM_OS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_BH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_CO;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_CC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EM_CH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EL_CD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EG_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EL_POT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EM_CE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EL_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_CE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_VA;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_CH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_OS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EU_FS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_RE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_FO;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_CVA;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_MA;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_PU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_EP_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_OD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_BIFFER_PTS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_BIFFER_LIGNES;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_CS_BAT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_CS_SENS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_CS_NOM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_ADHOC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;ENQ_SERV_REG_EC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_RP;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMPL_B_PTS_CONTROL;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_AD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMPL_B_PTS_PROJET;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NOM_COMMUNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NOM_CANTON;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NOM_DISTRICT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NO_LIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NOM_PAYS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_SENS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_NO_LOCAL;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_NO_LIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_BAC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_NOM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_COUV;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_PFP;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_NO_LIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_SENS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_NO_LIEUDIT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_BAC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_NO_LOCAL;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_PROJ_SENS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_BAT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;AD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_NOM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_PROJ_NOM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_CS_PROJ_BAT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_DN;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_CS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_LC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_OD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_NOM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_BF;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_COUV;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_BAT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_NOM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;CHECK_ITF_AD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;B_ODL_SENS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP3_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP3_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP3_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PLAN;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP3_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP2_STATIONNABLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFA3;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFA2;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP1_INACCESSIBLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP2_INACCESSIBLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFP1_STATIONNABLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_NON_MATE_NONDEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_CS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_NON_MATE_DEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PFA1;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_NON_MATE_NONDEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PLAN;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PFP3_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_NT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PFP3_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PFP3_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PFP3_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_DISTRICT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_COMMUNE_PROJ;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_PAYS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NORD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_COMMUNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_OD;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_CS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_SCHEMA_ORIEL_CLOCHE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;NOM_CANTON;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_SCHEMA_STATIONS;[[Annotative]]" & ValActive & ";")

        RVscript.cmdBL("#BL;AD_LOCALISATION;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_AD_LOCALISATION;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_AD_LOCALISATION;[[Annotative]]" & ValActive & ";")

        RVscript.cmdBL("#BL;ENQ_SERV_REG_EU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMMEUBLE_NUM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMMEUBLE_NUM_DDP;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMMEUBLE_NUM_DDP_PROJ;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;IMMEUBLE_NUM_PROJ;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_NO_LIEUDIT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_B_ODL_BAT;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_IMMEUBLE_NUM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_IMMEUBLE_NUM_DDP;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_BF_NON_MATE_DEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_BORNE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_NON_MATE_DEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_NON_MATE_NONDEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_PTS_COM_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;MUT_SCHEMA_ORIEL_DIVERS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_BF_NON_MATE_NONDEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;PTS_COM_NON_MATE_DEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_NO_LIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_B_ODL_BAC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_IMMEUBLE_DDP;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_IMMEUBLE_NUM;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_BF_NON_MATE_DEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_CHEVILLE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_CROIX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_NON_MATE_DEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_NON_MATE_NONDEFINI;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;RAD_PTS_COM_PIEU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_OBJ_VANNE_EAU_P;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_REG_EC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_TXT_EC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_EXIST_TXT_EU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_NOUV_OBJ_VANNE_EAU_PO;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SERV_NOUV_REG_EC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_CE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_CH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_CU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_DE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_EX;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_GR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_GU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_OS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_RE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EC_SE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_CC;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_CO;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_CVA;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_FO;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_PU;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_RE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EP_VA;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_CE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_CH;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_FS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_OS;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_PR;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_RE;[[Annotative]]" & ValActive & ";")
        RVscript.cmdBL("#BL;SOUT_B_PROJ_EU_SE;[[Annotative]]" & ValActive & ";")

        ' Sans annotatif dans le modèle par default
        ' TOPO_ARBRE_FEUILLUS
        ' TOPO_ARBRE_RESINEUX
        ' MUT_ETIQUETTE_ELEMENTS
        ' ENQ_ARBRES_EXIST
        ' ENQ_ARBRE_ABAT
        ' FORMAT_TEMPLATE



        Return True
    End Function

    Public Function PrecisionToValue(ByRef sngPrecPlan As Single) As Short
        Dim intVal As Object

        'Conversion en valeur
        Select Case sngPrecPlan
            'PFP2 --------------------------
            Case 2.1
                intVal = 11
            Case 3.1
                intVal = 21
            Case 4.1
                intVal = 31
            Case 5.1
                intVal = 41
            Case 6.1
                intVal = 51

                'PFP3 --------------------------
            Case 3.2
                intVal = 12
            Case 4.2
                intVal = 22
            Case 5.2
                intVal = 32
            Case 10.2
                intVal = 42
            Case 12.2
                intVal = 52

                'PL / détail --------------------------
            Case 3.3
                intVal = 13
            Case 3.4
                intVal = 14
            Case 3.5
                intVal = 15
            Case 4.3
                intVal = 23
            Case 4.4
                intVal = 24
            Case 4.5
                intVal = 25
            Case 7.3
                intVal = 33
            Case 7.4
                intVal = 34
            Case 7.5
                intVal = 35
            Case 16.3
                intVal = 43
            Case 16.4
                intVal = 44
            Case 16.5
                intVal = 45
            Case 35.3
                intVal = 53
            Case 35.4
                intVal = 54
            Case 35.5
                intVal = 55
                'PL / détail (tableau ne tenant pas compte de l'échelle, selon outil ConvertirPtVd)
            Case 13.6
                intVal = 16
            Case 13.7
                intVal = 17
            Case 24.6
                intVal = 26
            Case 24.7
                intVal = 27
            Case 47.6
                intVal = 36
            Case 47.7
                intVal = 37
            Case 116.6
                intVal = 46
            Case 116.7
                intVal = 47
            Case 135.6
                intVal = 56
            Case 135.7
                intVal = 57
                'PL / détail  --------------------------------------------------
            Case 8.2, 11.2, 20.2, 40.2, 80.2, 100.2
                intVal = 62
            Case 8.3, 11.3, 20.3, 40.3, 80.3, 100.3
                intVal = 63
            Case 8.4, 11.4, 20.4, 40.4, 80.4, 100.4
                intVal = 64
            Case 8.5, 11.5, 20.5, 40.5, 80.5, 100.5
                intVal = 65
            Case 8.6, 11.6, 20.6, 40.6, 80.6, 100.6
                intVal = 66
            Case 8.7, 11.7, 20.7, 40.7, 80.7, 100.7
                intVal = 67
            Case 8.8, 11.8, 20.8, 40.8, 80.8, 100.8
                intVal = 68
            Case 8.9, 11.9, 20.9, 40.9, 80.9, 100.9
                intVal = 69
                'PL / détail ---------------------------------------------------
            Case 9.2, 14.2, 30.2, 60.2, 120.2, 150.2
                intVal = 72
            Case 9.3, 14.3, 30.3, 60.3, 120.3, 150.3
                intVal = 73
            Case 9.4, 14.4, 30.4, 60.4, 120.4, 150.4
                intVal = 74
            Case 9.5, 14.5, 30.5, 60.5, 120.5, 150.5
                intVal = 75
            Case 9.6, 14.6, 30.6, 60.6, 120.6, 150.6
                intVal = 76
            Case 9.7, 14.7, 30.7, 60.7, 120.7, 150.7
                intVal = 77
            Case 9.8, 14.8, 30.8, 60.8, 120.8, 150.8
                intVal = 78
            Case 9.9, 14.9, 30.9, 60.9, 120.9, 150.9
                intVal = 79
                'PL / détail  ---------------------------------------------------
            Case 18.2, 25.2, 50.2, 90.2, 180.2, 230.2
                intVal = 82
            Case 18.3, 25.3, 50.3, 90.3, 180.3, 230.3
                intVal = 83
            Case 18.4, 25.4, 50.4, 90.4, 180.4, 230.4
                intVal = 84
            Case 18.5, 25.5, 50.5, 90.5, 180.5, 230.5
                intVal = 85
            Case 18.6, 25.6, 50.6, 90.6, 180.6, 230.6
                intVal = 86
            Case 18.7, 25.7, 50.7, 90.7, 180.7, 230.7
                intVal = 87
            Case 18.8, 25.8, 50.8, 90.8, 180.8, 230.8
                intVal = 88
            Case 18.9, 25.9, 50.9, 90.9, 180.9, 230.9
                intVal = 89
            Case Else
                intVal = 0
        End Select

        'Retourne la valeur résulat
        Return intVal


    End Function
End Module
