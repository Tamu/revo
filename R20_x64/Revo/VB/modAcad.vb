Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Colors
Imports frms = System.Windows.Forms
Imports Autodesk.AutoCAD.Interop
Imports Autodesk.AutoCAD.Interop.Common
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime

Module modAcad

    Dim Ncolor As Integer = 2
    Dim Fen As New Viewport

    Public UCSVar As Integer
    Public Processing As Boolean = False
    'Coord structure to hold X, Y and ID of coordinate
    Public Structure Coord
        Public X As Double
        Public Y As Double
        Public ID As String
    End Structure



    'Coordinate collection class to hold collection of coord objects
    Class CoordCollection(Of coord)
        Inherits CollectionBase
        'Sub to add a coord object to the collection
        Public Sub Add(ByVal C As coord)
            InnerList.Add(C)
        End Sub
        'Sub to remove a coord object from the collection
        Public Sub Remove(ByVal index As Integer)
            InnerList.RemoveAt(index)
        End Sub
        'Function to retrieve a coord object from the collection
        Public Function Item(ByVal index As Integer) As coord
            Return DirectCast(InnerList.Item(index), coord)
        End Function
    End Class



    ''' <summary>
    ''' Sets the frozen state of all layers to the state specified in the booFreeze variable
    ''' </summary>
    ''' <param name="booFreeze">The frozen state to set the layers to. True = Frozen, False = Thawed</param>
    ''' <remarks></remarks>
    Public Sub FreezeAllLayers(ByVal booFreeze As Boolean)
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Get the ID of current layer in the drawing
        Dim curLayerID As ObjectId = db.Clayer
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartOpenCloseTransaction
            'Get the layertable from the database
            Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            'Loop through the records in the layer table
            For Each oid As ObjectId In lt
                'Check that this layer is not the current one
                If oid <> curLayerID Then
                    'Get the layer table rcord from the objectid
                    Dim ltr As LayerTableRecord = DirectCast(trans.GetObject(oid, OpenMode.ForRead), LayerTableRecord)
                    'Upgrade its open status to write
                    ltr.UpgradeOpen()
                    'set its frozen state
                    ltr.IsFrozen = booFreeze
                End If
            Next
            trans.Commit()
        End Using

    End Sub

    ''' <summary>
    ''' Sets the required layer to the required frozen state
    ''' </summary>
    ''' <param name="strLayer">The name of the layer to freeze/unfreeze</param>
    ''' <param name="booFreeze">The frozen state to set the layer to. True if the layer is to be frozen.</param>
    ''' <remarks></remarks>
    Public Sub FreezeLayer(ByVal strLayer As String, ByVal booFreeze As Boolean)
        'Get the active document
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        'Get the ative document's database
        Dim db As Database = doc.Database
        'Lock the document
        Using docLock As DocumentLock = doc.LockDocument
            'Start a transaction with the database
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                'Get the ID of current layer in the drawing
                Dim curLayerID As ObjectId = db.Clayer
                'Get the layer table for this drawing
                Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
                'if the drawing contains the layer to set
                If lt.Has(strLayer) Then
                    'then get the layer table record for it
                    Dim ltr As LayerTableRecord = DirectCast(trans.GetObject(lt.Item(strLayer), OpenMode.ForRead), LayerTableRecord)
                    'Check that we're not trying to freeze the current layer
                    '(even if booFreeze is False i.e. we're trying to thaw the layer, there's still no need 
                    'to do the freezing/thawing as the current layer must already be thawed. If it wasn't, it couldn't be the current layer)
                    If ltr.ObjectId <> curLayerID Then
                        'Make the layertablerecord record writeable
                        ltr.UpgradeOpen()
                        'then freeze it
                        ltr.IsFrozen = booFreeze
                    End If
                End If
                'Commit the transaction to save any changes to the database
                trans.Commit()
            End Using
        End Using
        'Regen the drawing
        doc.Editor.Regen()
    End Sub

    ''' <summary>
    ''' Sets the visibility of the required layer to the required state
    ''' </summary>
    ''' <param name="strLayer">The name of the layer</param>
    ''' <param name="booVisible">The visibility state to set the layer to. True = On, False = Off</param>
    ''' <remarks></remarks>
    Public Sub ShowLayer(ByVal strLayer As String, ByVal booVisible As Boolean)
        'Get the database for this drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the layertable
            Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            'See if the layer that we're looking for is in the layer table
            If lt.Has(strLayer) Then
                Try
                    'Get the layertablerecord for the layer we're looking for
                    Dim ltr As LayerTableRecord = DirectCast(trans.GetObject(lt.Item(strLayer), OpenMode.ForRead), LayerTableRecord)
                    'Set it's open state to writeable
                    ltr.UpgradeOpen()
                    'Set the IsOff property to the inverse of the booVisible parameter
                    ltr.IsOff = Not (booVisible)
                    'Commit the transaction to save the change
                    trans.Commit()
                Catch ex As Exception
                    'Something went wrong so notify the user
                    frms.MessageBox.Show(ex.Message, "Show Layer", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                End Try
            End If
        End Using

    End Sub

    ''' <summary>
    ''' Sets the elevation of the drawing to the required value
    ''' </summary>
    ''' <param name="dblHeight">The elevation to set</param>
    ''' <remarks></remarks>
    Public Sub DefineElevation(ByVal dblHeight As Double)
        Dim db As Database = HostApplicationServices.WorkingDatabase
        db.Elevation = dblHeight
    End Sub

    ''' <summary>
    ''' Sets the required layer as the current layer
    ''' </summary>
    ''' <param name="strLayer">The layer to make current</param>
    ''' <remarks></remarks>
    Public Sub ActivateLayer(ByVal strLayer As String)
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Open a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the layertable
            Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            'See if the required layer is in the layertable
            If lt.Has(strLayer) Then
                Try
                    'Set that layer to be current
                    db.Clayer = lt.Item(strLayer)
                Catch ex As Exception
                End Try
            End If
            trans.Commit()
        End Using
    End Sub

    ''' <summary>
    ''' Sets the model space actived
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub ActivateModel()

        ' Active Model space
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
            Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
        Dim SpaceCur As Autodesk.AutoCAD.Interop.Common.AcadLayout = acDoc.ActiveLayout 'mémo du current Space
        If SpaceCur.Name <> "Model" Then
            For Each acLay As Autodesk.AutoCAD.Interop.Common.AcadLayout In acDoc.Layouts
                If acLay.Name = "Model" Then
                    acDoc.ActiveLayout = acLay 'Space = acdoc.ModelSpace
                End If
            Next
        End If


    End Sub


    ''' <summary>
    ''' Creates a polyline in the drawing from the supplied vertices
    ''' </summary>
    ''' <param name="Verts">The vertices of the polyline to create</param>
    ''' <returns>The created polyline</returns>
    ''' <remarks></remarks>
    Public Function InsertLWPolyline(ByVal Verts As Point2dCollection, ByVal strLayer As String) As Polyline
        Dim LayID As ObjectId = GetLayerIDFromLayerName(strLayer)
        'Declare a polyline
        Dim retLine As Polyline = Nothing
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Try
                'Create the polyline
                Dim pLine As New Polyline(Verts.Count)
                Dim Counter As Integer = 0
                'Add the vertices to the polyline
                For Each p As Point2d In Verts
                    pLine.AddVertexAt(Counter, p, 0, 0, 0)
                    Counter += 1
                Next
                pLine.LayerId = LayID
                'Get the modelspace block table record
                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                'Make it writeable
                btr.UpgradeOpen()
                'Add the polyline to modelspace
                Dim Id As ObjectId = btr.AppendEntity(pLine)
                'Let the transaction know about the polyline
                trans.AddNewlyCreatedDBObject(pLine, True)
                retLine = DirectCast(trans.GetObject(Id, OpenMode.ForRead), Polyline)
                trans.Commit()
            Catch ex As Exception
                frms.MessageBox.Show("Error creating polyline." & vbCrLf & ex.Message, "InsertLWPolyline", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
        Return retLine
    End Function

    ''' <summary>
    ''' Creates a polyline in the drawing from the supplied vertices
    ''' </summary>
    ''' <param name="Verts">The vertices of the polyline to create</param>
    ''' <returns>The created polyline</returns>
    ''' <remarks>Overloaded method to create polyline from dictionary of vertices</remarks>
    Public Function InsertLWPolyline(ByVal Verts As Dictionary(Of Integer, Point2d), ByVal strLayer As String) As Polyline
        Dim LayID As ObjectId = GetLayerIDFromLayerName(strLayer)
        'Convert the dictionary to a list so we can sort the vertices in order
        Dim VertList As List(Of KeyValuePair(Of Integer, Point2d)) = ConvertDictionary(Of Integer, Point2d)(Verts)
        VertList.Sort(AddressOf CompareIntKeys)
        'Declare a polyline
        Dim retLine As Polyline = Nothing
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Try
                'Create the polyline
                Dim pLine As New Polyline(Verts.Count)
                Dim Counter As Integer = 0
                'Add the vertices to the polyline
                For Each kvp As KeyValuePair(Of Integer, Point2d) In VertList
                    Dim p As Point2d = kvp.Value
                    pLine.AddVertexAt(Counter, p, 0, 0, 0)
                    Counter += 1
                Next
                pLine.LayerId = LayID
                'Get the modelspace block table record
                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                'Make it writeable
                btr.UpgradeOpen()
                'Add the polyline to modelspace
                Dim Id As ObjectId = btr.AppendEntity(pLine)
                'Let the transaction know about the polyline
                trans.AddNewlyCreatedDBObject(pLine, True)
                retLine = DirectCast(trans.GetObject(Id, OpenMode.ForRead), Polyline)
                trans.Commit()
            Catch ex As Exception
                frms.MessageBox.Show("Error creating polyline." & vbCrLf & ex.Message, "InsertLWPolyline", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
        Return retLine
    End Function

    ''' <summary>
    ''' Calculates the bulge value of the arc
    ''' </summary>
    ''' <param name="p1">The start point of the arc</param>
    ''' <param name="p2">The mid point of the arc</param>
    ''' <param name="p3">The end point of the arc</param>
    ''' <returns>The bulge value of the arc</returns>
    ''' <remarks></remarks>
    Public Function GetBulgeValue(ByVal p1 As Point2d, ByVal p2 As Point2d, ByVal p3 As Point2d) As Double
        'Create a circular arc in memory from the three points supplied
        Dim a As New CircularArc2d(p1, p2, p3)
        'Store the radius
        Dim dblR As Double = a.Radius
        'Store whether it's clockwise
        Dim booClockwise As Boolean = a.IsClockWise
        Dim newStart As Double = 0
        If a.StartAngle > a.EndAngle Then
            newStart = a.StartAngle - 8 * Math.Atan(1)
        Else
            newStart = a.StartAngle
        End If
        Dim dblBulge As Double = Math.Tan((a.EndAngle - newStart) / 4)
        If booClockwise Then dblBulge = -dblBulge
        Return dblBulge
    End Function

    ''' <summary>
    ''' Sets the bulge of the polyline at the desired index
    ''' </summary>
    ''' <param name="oId">The objectid of the polyline to set the bulge in</param>
    ''' <param name="Indx">The index of the vertex at which to set the bulge</param>
    ''' <param name="blg">The bulge amount</param>
    ''' <remarks></remarks>
    Public Sub SetPolyBulge(ByVal oId As ObjectId, ByVal Indx As Integer, ByVal blg As Double)
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the polyline form the database
            Dim pline As Polyline = DirectCast(trans.GetObject(oId, OpenMode.ForRead), Polyline)
            'Make it writeable
            pline.UpgradeOpen()
            'Set the bulge
            pline.SetBulgeAt(Indx, blg)
            'Save the change
            trans.Commit()
        End Using
    End Sub



    ''' <summary>
    ''' Inserts a block reference by name at a specified point in a specified space.
    ''' </summary>
    ''' <param name="dblInsert">3D Point of insertion</param>
    ''' <param name="btrSpace">BlockTableRecord of parent modelspace, paperspace, whatever</param>
    ''' <param name="strSourceBlockName">Name of block</param>
    ''' <param name="strSourceBlockPath">Name of block's dwg file. (Path+FileName !!!)</param>
    ''' <param name="db">db as DWG Database </param>
    ''' <param name="strBlockNewName">Name to rename block if desired</param>
    ''' <returns>Block reference returned</returns>
    ''' <remarks>Processes attributes upon insertion. Returns nothing if failed.</remarks>
    Public Function InsertBlockRef(ByVal dblInsert As Point3d, ByVal btrSpace As BlockTableRecord, ByVal strSourceBlockName As String, ByVal strSourceBlockPath As String, ByVal db As Database, Optional ByVal strBlockNewName As String = "") As BlockReference
        Dim bt As BlockTable
        Dim btr As BlockTableRecord = Nothing
        Dim br As BlockReference = Nothing
        'Dim id As ObjectId
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Try

                'insert block and rename it
                bt = trans.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, True, True)
                Dim btrSource As BlockTableRecord = GetBlock(strSourceBlockName, strSourceBlockPath, db) ', strBlockNewName)
                'if getblock returns nothing then the rest of this code is worthless, skip it and dispose the transaction
                If btrSource IsNot Nothing Then
                    btr = trans.GetObject(btrSource.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, True, True)
                    btrSpace = trans.GetObject(btrSpace.ObjectId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, True, True)
                    'Get the Attributes from the block source and add references to them in the blockref's attribute collection
                    Dim attColl As AttributeCollection
                    Dim ent As Entity
                    br = New BlockReference(dblInsert, btr.ObjectId)
                    btrSpace.AppendEntity(br)
                    trans.AddNewlyCreatedDBObject(br, True)
                    attColl = br.AttributeCollection
                    For Each oid As ObjectId In btr 'accepts only key string
                        ent = trans.GetObject(oid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, True, True)
                        If TypeOf ent Is AttributeDefinition Then
                            Dim attdef As AttributeDefinition = ent
                            Dim attref As New AttributeReference
                            attref.SetAttributeFromBlock(attdef, br.BlockTransform)
                            attref.TextString = attdef.TextString
                            attColl.AppendAttribute(attref)
                            trans.AddNewlyCreatedDBObject(attref, True)
                        End If
                    Next oid
                    trans.Commit()
                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        End Using
        Return br
    End Function

    ''' <summary>
    ''' Finds / inserts a block definition from file.
    ''' </summary>
    ''' <param name="strSourceBlockName">Name of inserted block</param>
    ''' <param name="strSourceBlockPath">Name of block's dwg file</param>
    ''' <param name="db">db as DWG Database </param>
    ''' <returns>BlockTableRecord of the found/inserted block</returns>
    ''' <remarks>If block is not found, then an error is thrown, and nothing is returned.</remarks>
    Public Function GetBlock(ByVal strSourceBlockName As String, ByVal strSourceBlockPath As String, ByVal db As Database) As BlockTableRecord
        Dim oid As ObjectId = Nothing
        Dim btr As BlockTableRecord = Nothing
        Dim bt As BlockTable = db.BlockTableId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
        If bt.Has(strSourceBlockName) Then
            btr = bt(strSourceBlockName).GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
        Else
            Using trans As Transaction = db.TransactionManager.StartTransaction
                Using sourcedb As Database = New Database(False, False)
                    Try
                        sourcedb.ReadDwgFile(strSourceBlockPath, IO.FileShare.Read, True, "")
                        oid = db.Insert(strSourceBlockPath, sourcedb, True)
                        btr = trans.GetObject(oid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, True, True)
                        btr.Name = strSourceBlockName
                        trans.Commit()
                    Catch ex As System.Exception
                        Throw New System.Exception("Block file not found " & strSourceBlockPath & ": " & ex.Message)
                    End Try
                End Using
            End Using
        End If
        Return btr
    End Function



    ''' <summary>
    ''' Inserts a DWG into the drawing
    ''' </summary>
    ''' <param name="DataFile">Nom de fichier</param>
    ''' <param name="Space">Space / Layout</param>
    ''' <param name="CoordXYZ">Points d'insertion</param>
    ''' <remarks></remarks>
    Public Function InsertDWG(DataFile As String, ByVal Space As String, CoordXYZ As String)

        Space = Space.ToLower

        'Importer l'espace Objet (insertion bloc)
        If "model" Like Space Then

            'insertion point
            Dim PtOrig As Point3d = New Point3d(0, 0, 0)
            If CoordXYZ IsNot Nothing Or CoordXYZ <> "" Then 'XYZInsertionPoint = dbl,dbl,dbl (@)
                Dim SplitDbl() As String = Split(CoordXYZ, ",")
                If SplitDbl.Length = 3 Then PtOrig = New Point3d(CDbl(SplitDbl(0)), CDbl(SplitDbl(1)), CDbl(SplitDbl(2)))
            End If

            'Nom du fichier
            Dim BlocName As String = IO.Path.GetFileNameWithoutExtension(DataFile)
            Dim BlocPathName As String = IO.Path.GetFullPath(DataFile)

            'Importation du fichier DWG
            Dim doc As Document = Application.DocumentManager.MdiActiveDocument
            Dim db As Database = doc.Database
            Dim ed As Editor = doc.Editor
            Dim IDobjBL As ObjectId
            doc.LockDocument()
            Using tr As Transaction = db.TransactionManager.StartTransaction()        ' Start a transaction
                ' Test if block exists in the block table
                Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead), BlockTable)
                ' Get Model space
                Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), BlockTableRecord)
                ' Add the block reference to Model space
                IDobjBL = InsertBlockRef(PtOrig, btr, BlocName, BlocPathName, db).ObjectId
                tr.Commit() ' Commit the transaction
            End Using

            '   Exploser Block
            ExplodeBlock(IDobjBL)
            Zooming.ZoomExtents()
            doc.LockDocument.Dispose()
        End If

        'Attention unité ???

        'Importer l'espace Papier (-presenstation)
        Dim RVscript As New Revo.RevoScript
        If "paper" Like Space Then
            '   importer tout les espaces papier :   Importprop(2) Space
            RVscript.cmdCmd("#Cmd;filedia|0|_-Layout|_T|" & DataFile & "|*|filedia|1")

        ElseIf Space <> "" And "model" <> Space Then
            '   importer l'espace papier selon le nom :   Importprop(2) Space
            RVscript.cmdCmd("#Cmd;filedia|0|_-Layout|_T|" & DataFile & "|" & Space & "|filedia|1")
        End If

        Return True

    End Function
    'Get the database for 
    ''' <summary>
    ''' Inserts a block into the drawing
    ''' </summary>
    ''' <param name="insPoint">The location in the drawing at which to insert the block</param>
    ''' <param name="blockName">The name of the block to insert</param>
    ''' <param name="insAngle">The rotation to set the block to (in grade)</param>
    ''' <param name="layID">The ID of the layer to put the block on</param>
    ''' <param name="dblEchelle">Optional - the scale to set the block to</param>
    ''' <returns>The objectID of the inserted block if successful, otherwise, nothing</returns>
    ''' <remarks></remarks>
    Public Function InsertBlock(ByVal insPoint As Point3d, ByVal blockName As String, ByVal insAngle _
                                As Double, ByVal layID As ObjectId, Optional ByVal dblEchelle As Double = 1) As ObjectId
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim objId As ObjectId = Nothing 'Return value
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                'Get the modelspace blocktablerecord
                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                'See if the drawing contains the definition of the block we're trying to insert
                If bt.Has(blockName) Then
                    'Get the id for the block we're trying to insert
                    Dim btrSrc As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(blockName), OpenMode.ForRead), BlockTableRecord)
                    objId = btrSrc.Id
                End If
                'Check that we have the id for the block we're trying to insert
                If Not objId.IsNull Then
                    'Create a new block reference at the point specified using the ID we got from the blocktable
                    Dim blkRef As New BlockReference(insPoint, objId)
                    'Add the block reference to modelspace
                    btr.AppendEntity(blkRef)
                    'Let the transaction know about the new block
                    trans.AddNewlyCreatedDBObject(blkRef, True)
                    'Set the rotation of the block
                    blkRef.Rotation = ConvGisTopoTrigo(insAngle * (Math.PI / 200))
                    'and the scale
                    blkRef.ScaleFactors = New Scale3d(dblEchelle)
                    'and the layer
                    blkRef.LayerId = layID
                    'Save the changes
                    trans.Commit()
                    Return blkRef.ObjectId
                Else
                    'Notify the user that the block was not found
                    frms.MessageBox.Show("Drawing does not contain block " & blockName, "Insert Block", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                    Return Nothing
                End If
            Catch
                frms.MessageBox.Show(Err.Description, "Insert Block", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
        Return Nothing
    End Function

    ''' <summary>
    ''' Adds the attributes from the block definition to the block reference
    ''' </summary>
    ''' <param name="oid">The objectID of the block reference</param>
    ''' <remarks></remarks>
    Private Function AddAttributesToBlock(ByVal oid As ObjectId) As Boolean
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the blockreference
            Dim blkRef As BlockReference = DirectCast(trans.GetObject(oid, OpenMode.ForRead), BlockReference)
            'Store the block name
            Dim blkName As String = blkRef.Name
            'Get the block table record for ModelSpace
            Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
            'Check that the block is in the drawing
            If bt.Has(blkName) Then
                'Get the id of the block definition
                Dim id As ObjectId = bt(blkName)
                'Get the block table for the block (contains the entities in the block definition)
                Dim btBlkRec As BlockTableRecord = DirectCast(trans.GetObject(id, OpenMode.ForRead), BlockTableRecord)
                'Loop through the entities in the block definition
                For Each entId As ObjectId In btBlkRec
                    Dim ent As Entity = DirectCast(trans.GetObject(entId, OpenMode.ForRead), Entity)
                    'Is this entity an attribute definition?
                    If TypeOf ent Is AttributeDefinition Then
                        'Convert the entity to an attribute definition
                        Dim attdef As AttributeDefinition = DirectCast(ent, AttributeDefinition)
                        'Create a new attribute reference to add to the blockreference
                        Dim attRef As New AttributeReference
                        'Copy the properties from the definition to the new reference
                        attRef.SetAttributeFromBlock(attdef, blkRef.BlockTransform)
                        'attRef.Rotation = blkRef.Rotation
                        blkRef.AttributeCollection.AppendAttribute(attRef)
                        trans.AddNewlyCreatedDBObject(attRef, True)
                    End If
                Next
                trans.Commit()
                Return True
            Else
                Return False
            End If
        End Using

    End Function

    ''' <summary>
    ''' Sets the attributes in the block to the values in the recordset
    ''' </summary>
    ''' <param name="dr">The datarow containing the values for the attributes</param>
    ''' <param name="blkID">The objectId of the block to set the attributes on</param>
    ''' <remarks></remarks>
    Public Sub SetBlockAttributeValues(ByVal dr As DataRow, ByVal blkID As ObjectId)
        'Add the attribute references to the blockreference
        AddAttributesToBlock(blkID)
        'Get the database for this drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start as transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the blockreference from the objectID
            Dim blkRef As BlockReference = DirectCast(trans.GetObject(blkID, OpenMode.ForRead), BlockReference)
            'Get the block attributue collection
            Dim attCol As AttributeCollection = blkRef.AttributeCollection
            Dim attRef As AttributeReference = Nothing
            'Loop through the attribute collection
            For Each attId As ObjectId In attCol
                'Get this attribute reference
                attRef = DirectCast(trans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                'Store its name
                Dim attName As String = attRef.Tag
                Try
                    'Only set the attribute value if it's not a null string
                    If dr(attName).ToString.Replace("_", " ").Trim <> "" Then
                        attRef.TextString = dr(attName).ToString.Replace("_", " ")
                    End If
                    'Set the rotation of the attribute to that of the block
                    attRef.Rotation = blkRef.Rotation
                Catch 'ex As Exception
                    ' MsgBox("Erreur de chargement d'attribut")
                End Try

                'If the attribute name is "ALTITUDE" then set the value to "?" ' DESACTIVATE : 07.12.2015 THA
                If attName.ToUpper = "ALTITUDE" Then
                    ' attRef.TextString = "?"
                End If

            Next
            'commit the transaction to save the changes
            trans.Commit()
        End Using
    End Sub
    ''' <summary>
    ''' Sets the attributes in the block to the values in the recordset
    ''' </summary>
    ''' <param name="dr">The datarow containing the values for the attributes</param>
    ''' <param name="blkID">The objectId of the block to set the attributes on</param>
    ''' <remarks></remarks>
    Public Sub SetBlockAttributeValuesList(ByVal AttList As List(Of String), ValueList As List(Of String), ByVal blkID As ObjectId)
        'Add the attribute references to the blockreference
        AddAttributesToBlock(blkID)
        'Get the database for this drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start as transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the blockreference from the objectID
            Dim blkRef As BlockReference = DirectCast(trans.GetObject(blkID, OpenMode.ForRead), BlockReference)
            'Get the block attributue collection
            Dim attCol As AttributeCollection = blkRef.AttributeCollection
            Dim attRef As AttributeReference = Nothing
            'Loop through the attribute collection
            For Each attId As ObjectId In attCol
                'Get this attribute reference
                attRef = DirectCast(trans.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                'Store its name
                Dim attName As String = attRef.Tag.ToUpper
                Try
                    'Only set the attribute value if it's not a null string
                    If AttList.Contains(attName) Then
                        attRef.TextString = ValueList(AttList.IndexOf(attName))
                    End If
                    'Set the rotation of the attribute to that of the block
                    'attRef.Rotation = blkRef.Rotation
                Catch ex As Exception
                End Try
            Next
            'commit the transaction to save the changes
            trans.Commit()
        End Using
    End Sub

    ''' <summary>
    ''' Sets the colour of the block
    ''' </summary>
    ''' <param name="blkId">The objectID of the block</param>
    ''' <param name="col">The autodesk colour to set the block to</param>
    ''' <remarks></remarks>
    Public Sub SetBlockColour(ByVal blkId As ObjectId, ByVal col As Color)
        'Get the database for this drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the blockreference from the objectid
            Dim blkRef As Entity = DirectCast(trans.GetObject(blkId, OpenMode.ForWrite), Entity)
            'Set the colour
            blkRef.Color = col
            'Commit the transaction to save the change
            trans.Commit()
        End Using
    End Sub

    ''' <summary>
    ''' Sets the layer of an entity 
    ''' </summary>
    ''' <param name="oID">The objectid of the entity</param>
    ''' <param name="strLayer">The layer to set</param>
    ''' <remarks></remarks>
    Public Sub SetObjectLayer(ByVal oID As ObjectId, ByVal strLayer As String)
        'Get the database for this drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the entity
            Dim ent As Entity = DirectCast(trans.GetObject(oID, OpenMode.ForRead), Entity)
            'Get the layertable from the database
            Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            'If the layer we're looking for is in there
            If lt.Has(strLayer) Then
                'Get the objectid of the layer
                Dim LayID As ObjectId = lt.Item(strLayer)
                'Make the entity writeable
                ent.UpgradeOpen()
                'Set it's layer id
                ent.LayerId = LayID
                'Commit the transaction to save the changes
                trans.Commit()
            Else
                frms.MessageBox.Show("Layer " & strLayer & " not found", "Set Object Layer", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Gets the name of the current layer in the drawing
    ''' </summary>
    ''' <returns>The name of the current layer</returns>
    ''' <remarks></remarks>
    Public Function GetCurrentLayer() As String
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start as transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the objectid of the current layer
            Dim CurLayer As ObjectId = db.Clayer
            'Get the layertablerecord for the current layer
            Dim ltr As LayerTableRecord = DirectCast(trans.GetObject(CurLayer, OpenMode.ForRead), LayerTableRecord)
            'Return its name
            Return ltr.Name
        End Using

    End Function

    ''' <summary>
    ''' Selects all items in the drawing that pass the filter constructed from the values supplied
    ''' </summary>
    ''' <param name="values">The typedvalues to create the filter from </param>
    ''' <returns>A promptselectionresult based in the filter</returns>
    ''' <remarks></remarks>
    Public Function SelectAllItems(ByVal values() As TypedValue) As PromptSelectionResult
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        'Create a filter using the supplied values
        Dim sFilter As New SelectionFilter(values)
        'Select everything that passes the filter
        Dim pres As PromptSelectionResult = ed.SelectAll(sFilter)
        'and return the result
        Return pres
    End Function

    ''' <summary>
    ''' Return Layout name (paper space name)
    ''' </summary>
    ''' <param name="ent">object</param>
    ''' <param name="tr">Transaction</param>
    ''' <remarks></remarks>
    Public Function LayoutName(ent As Entity, tr As Transaction) As String
        Dim blockTableRecord As BlockTableRecord = DirectCast(tr.GetObject(ent.BlockId, OpenMode.ForRead), BlockTableRecord)
        Dim layout As Layout = DirectCast(tr.GetObject(blockTableRecord.LayoutId, OpenMode.ForRead), Layout)
        Return layout.LayoutName
    End Function

    ''' <summary>
    ''' Select ViewPort and return Value
    ''' </summary>
    ''' <param name="ent">object</param>
    ''' <param name="tr">Transaction</param>
    ''' <remarks></remarks>
    Public Function SelectViewPort(Optional ByVal BLid As String = "") As List(Of String)

        Dim ListV As New List(Of String)
        ListV.Add(0) '0 FenValide As Boolean = False
        ListV.Add(1) '1 FenScale As Double = 1
        ListV.Add("0,0,0") '2 FenCenter As Point3d '2
        ListV.Add("0,0,0") '3 FenCenterObj As Point3d '3
        ListV.Add(297) '4 FenHeight As Double = 297
        ListV.Add(420) '5 FenWidth As Double = 420
        ListV.Add(0) '6 ViewPortObjectID As String = 0

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor


        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            If BLid = "" Then 'Recherche du Fenêtre

                Try

                    Dim FenOK As Boolean = False

                    Dim filterValues As TypedValue() = New TypedValue() {New TypedValue(CInt(DxfCode.Start), "VIEWPORT"), New TypedValue(CInt(67), 1)}
                    Dim filter As New SelectionFilter(filterValues)
                    Dim foundBlocks As PromptSelectionResult = ed.SelectAll(filter)
                    If foundBlocks.Status <> PromptStatus.OK Then
                        'Erreur
                    Else

                    
                        For Each id As ObjectId In foundBlocks.Value.GetObjectIds()

                            ' Dim acEnt As Entity = acTrans.GetObject(id, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead)
                            '  MsgBox(TypeName(acEnt))

                            Dim Lay As Viewport = DirectCast(acTrans.GetObject(id, OpenMode.ForRead), Viewport)

                            Dim newlm As LayoutManager = LayoutManager.Current
                            If newlm.CurrentLayout.ToUpper = LayoutName(Lay, acTrans).ToUpper Then
                                FenOK = True
                                Fen = DirectCast(acTrans.GetObject(id, OpenMode.ForRead), Viewport)
                                Exit For

                            End If


                        Next
                    End If

                    If FenOK = False Then
                        MsgBox("Aucune fenêtre a été trouvée. Créer la fenêtre dans l'espace papier", vbInformation + vbOKOnly, "Interuption de la commande")
                        ListV(0) = 0 'FenValide
                        Return ListV
                        Exit Function
                    End If

                Catch ex As Exception
                    MsgBox(ex.Message)
                End Try





            Else 'Selection de la Fenêtre

                Dim acOPrompt As PromptSelectionOptions = New PromptSelectionOptions()

                acOPrompt.MessageForAdding = "Sélectionner une fenêtre" '"Sélectionner le périmètre de la parcelle (poyligne périmètrique)"
                acOPrompt.SingleOnly = True

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
                            Dim acEnt As Entity = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                            If Not IsDBNull(acEnt) Then
                                If TypeName(acEnt).ToLower Like "viewport" Then

                                    Fen = acTrans.GetObject(acSSObj.ObjectId, OpenMode.ForRead)
                                    Exit For

                                Else
                                    MsgBox("Aucune fenêtre a été sélectionnée", vbOKOnly + vbInformation, "Erreur de saisie")
                                End If

                            End If
                        End If
                    Next
                Else
                    If acSSPrompt.Status.ToString = "Cancel" Then
                    Else
                        ed.WriteMessage(vbLf & "Aucun éléments sélectionné ou élément sélectionné non valide")
                        ed.WriteMessage(vbLf & "('ECHAP' pour quitter)")
                    End If
                End If

            End If


            'Chargement des Valeurs
            Try


                '0  Dim FenValide As Boolean = False
                '1  Dim FenScale As Double = 1
                Dim FenCenter As Point3d '2
                Dim FenCenterObj As Point3d '3
                '4  Dim FenHeight As Double = 297
                '5  Dim FenWidth As Double = 420
                '6  Dim ViewPortObjectID As String = 0

                ListV(0) = 1 'FenValide
                ListV(1) = Fen.CustomScale   'FenScale
                FenCenter = Fen.CenterPoint ' FenCenter
                ListV(2) = FenCenter.X & "," & FenCenter.Y & "," & FenCenter.Z

                Revo.MyCommands.ActiveEventsCmd = False 'Desactive action
                acDoc.Editor.SwitchToModelSpace()
                Application.SetSystemVariable("CVPORT", Fen.Number)
                FenCenterObj = DirectCast(Application.GetSystemVariable("VIEWCTR"), Point3d)
                ' FenCenterObj = CType(Application.GetSystemVariable("VIEWCTR"), Point3d)


                acDoc.Editor.SwitchToPaperSpace()
                Revo.MyCommands.ActiveEventsCmd = True 'Active action

                ListV(3) = FenCenterObj.X & "," & FenCenterObj.Y & "," & FenCenterObj.Z

                ListV(4) = Fen.Height 'FenHeight
                ListV(5) = Fen.Width 'FenWidth
                ' ListV(6) = "" 'Fen.Handle.ToString '   ViewPortObjectID
                ListV(6) = Fen.TwistAngle 'Rotation


                'AddEventLayout(Fen, acTrans, acDoc.Database)
                ' AddHandler Fen.Modified, AddressOf acLayoutMod 'SUP  TH 15.10.2015

            Catch ex As Exception
                If ex.Message = "eCannotChangeActiveViewport" Then
                    MsgBox("Aucune fenêtre a été trouvée. Créer une fenêtre dans l'espace papier", vbInformation + vbOKOnly, "Interuption de la commande")
                    ListV(0) = 0 'FenValide
                    Return ListV
                    Exit Function
                End If
            End Try

        End Using

        Return ListV
    End Function

    Public Sub acLayoutMod(ByVal senderObj As Object, ByVal evtArgs As EventArgs)
        '  Application.ShowAlertDialog("The area of " & acPoly.ToString() & " is: " & acPoly.Area)

        If Revo.MyCommands.ActiveEventsCmd = True Then 'Desactive action

            Try
                Dim QuadriColl As New Collection
                QuadriColl = CheckExistQuadri()
                If QuadriColl.Count = 0 Then ' Création normal
                    '  Update = False
                ElseIf QuadriColl.Count = 1 Then ' Mise à jour


                    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                    Using docLock As DocumentLock = acDoc.LockDocument


                        ' RemoveHandler Fen.Modified, AddressOf acLayoutMod

                        ' Dim MyCmd As New Revo.MyCommands
                        'MyCmd.RevoQuadrillage()
                        acDoc.SendStringToExecute(String.Format("revoQuadrillage") & vbCr, True, False, False)

                        'AddHandler Fen.Modified, AddressOf acLayoutMod
                        ' Revo.MyCommands.ActiveEventsCmd = True 'Active action

                        docLock.Dispose()

                    End Using




                Else
                    'MsgBox("Plusieurs réseaux de coordonnées ont été trouvés, conserver uniquement un réseau par présentation", vbExclamation + vbOKOnly, "Interruption du traitement")
                    ' Exit Sub
                End If

            Catch ex As System.Exception
                MsgBox(ex.Message)
            End Try

            Revo.MyCommands.ActiveEventsCmd = False
        End If

    End Sub


    Friend Function MS2wPS(ByVal vp As Viewport) As Matrix3d
        Dim viewDirection As Vector3d = vp.ViewDirection
        Dim center As Point2d = vp.ViewCenter
        Dim viewCenter As Point3d = New Point3d(center.X, center.Y, 0)
        Dim viewTarget As Point3d = vp.ViewTarget
        Dim twistAngle As Double = (vp.TwistAngle * -1)
        Dim centerPoint As Point3d = vp.CenterPoint
        Dim viewHeight As Double = vp.ViewHeight
        Dim height As Double = vp.Height
        Dim width As Double = vp.Width
        Dim scaling As Double = (viewHeight / height)
        Dim lensLength As Double = vp.LensLength
        Dim zAxis As Vector3d = viewDirection.GetNormal
        Dim xAxis As Vector3d = Vector3d.ZAxis.CrossProduct(viewDirection)
        Dim yAxis As Vector3d
        If Not xAxis.IsZeroLength Then
            xAxis = NormalizeVector(xAxis)
            yAxis = zAxis.CrossProduct(xAxis)
            Dim tmp As Double = xAxis.X
            tmp = xAxis.Y
            tmp = xAxis.Z
        ElseIf (zAxis.Z < 0) Then
            xAxis = (Vector3d.XAxis * -1)
            yAxis = Vector3d.YAxis
            zAxis = (Vector3d.ZAxis * -1)
        Else
            xAxis = Vector3d.XAxis
            yAxis = Vector3d.YAxis
            zAxis = Vector3d.ZAxis
        End If

        Dim ps2dcs As Matrix3d = Matrix3d.Displacement((Point3d.Origin - centerPoint))
        ps2dcs = (ps2dcs * Matrix3d.Scaling(scaling, centerPoint))
        Dim dcs2wcs As Matrix3d = Matrix3d.Displacement((viewCenter - Point3d.Origin))
        Dim matCoords As Matrix3d = Matrix3d.AlignCoordinateSystem(Matrix3d.Identity.CoordinateSystem3d.Origin, Matrix3d.Identity.CoordinateSystem3d.Xaxis, Matrix3d.Identity.CoordinateSystem3d.Yaxis, Matrix3d.Identity.CoordinateSystem3d.Zaxis, Matrix3d.Identity.CoordinateSystem3d.Origin, xAxis, yAxis, zAxis)
        dcs2wcs = (matCoords * dcs2wcs)
        dcs2wcs = (Matrix3d.Displacement((viewTarget - Point3d.Origin)) * dcs2wcs)
        dcs2wcs = (Matrix3d.Rotation(twistAngle, zAxis, viewTarget) * dcs2wcs)
        Dim perspMat As Matrix3d = Matrix3d.Identity
        If vp.PerspectiveOn Then
            Dim viewsize As Double = viewHeight
            Dim aspectRatio As Double = (width / height)
            Dim adjustFactor As Double = (1 / 42)
            Dim adjustedLensLength As Double = (viewsize _
                        * (lensLength _
                        * (Math.Sqrt((1 _
                            + (aspectRatio * aspectRatio))) * adjustFactor)))
            Dim eyeDistance As Double = viewDirection.Length
            Dim lensDistance As Double = (eyeDistance - adjustedLensLength)
            Dim ed As Double = eyeDistance
            Dim ll As Double = adjustedLensLength
            Dim l As Double = lensDistance
            perspMat = New Matrix3d(New Double() {1, 0, 0, 0, 0, 1, 0, 0, 0, 0, ((ll - l) _
                                    / ll), (l _
                                    * ((ed - ll) _
                                    / ll)), 0, 0, ((1 / ll) _
                                    * -1), (ed / ll)})
        End If

        Return (ps2dcs.Inverse _
                    * (perspMat * dcs2wcs.Inverse))
    End Function

    Friend Function NormalizeVector(ByVal vec As Vector3d) As Vector3d
        Dim length As Double = Math.Sqrt(((vec.X * vec.X) _
                        + ((vec.Y * vec.Y) _
                        + (vec.Z * vec.Z))))
        Dim x As Double = (vec.X / length)
        Dim y As Double = (vec.Y / length)
        Dim z As Double = (vec.Z / length)
        Return New Vector3d(x, y, z)
    End Function

    ''' <summary>
    ''' Ensures that the polyline is closed and sets its layer 
    ''' </summary>
    ''' <param name="oid">the objectid of the polyline</param>
    ''' <param name="strLayer">The layer to set</param>
    ''' <remarks></remarks>
    Public Sub ClosePolyAndSetLayer(ByVal oid As ObjectId, ByVal strLayer As String)
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the polyline from the database
            Dim pLine As Polyline = DirectCast(trans.GetObject(oid, OpenMode.ForWrite), Polyline)
            'Close it
            pLine.Closed = True
            'Save the changes
            trans.Commit()
        End Using
        'Call the function to set the layer
        SetObjectLayer(oid, strLayer)
    End Sub

    ''' <summary>
    ''' Adds XData to the polyline
    ''' </summary>
    ''' <param name="oId">The objectId of the polyline</param>
    ''' <param name="resbuf">The resultbuffer containing the data to add</param>
    ''' <remarks></remarks>
    Public Sub SetPolyXData(ByVal oId As ObjectId, ByRef resbuf As ResultBuffer)
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Try
                'Get the polyline from the database
                Dim pLine As Polyline = DirectCast(trans.GetObject(oId, OpenMode.ForWrite), Polyline)
                'Set the XData
                pLine.XData = resbuf
                'Save the changes
                trans.Commit()
            Catch ex As Exception
                frms.MessageBox.Show(ex.Message, "SetObjectXData", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' Adds a regapptable record (required for adding XData)
    ''' </summary>
    ''' <param name="regAppName">The name required</param>
    ''' <remarks></remarks>
    Public Sub AddRegAppTableRecord(ByVal regAppName As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim ed As Editor = doc.Editor
        Dim db As Database = doc.Database
        Using trans As Transaction = doc.TransactionManager.StartTransaction()
            Dim rat As RegAppTable = DirectCast(trans.GetObject(db.RegAppTableId, OpenMode.ForRead, False), RegAppTable)
            If Not rat.Has(regAppName) Then
                rat.UpgradeOpen()
                Dim ratr As New RegAppTableRecord()
                ratr.Name = regAppName
                rat.Add(ratr)
                trans.AddNewlyCreatedDBObject(ratr, True)
            End If
            trans.Commit()
        End Using
    End Sub

    ''' <summary>
    ''' Restores the named layer state supplied
    ''' </summary>
    ''' <param name="strlayerStateName">The layerstatename to restore</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Public Function RestoreLayerState(ByVal strlayerStateName As String) As Boolean
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the layerstatemanager
            Dim lsm As LayerStateManager = db.LayerStateManager
            Try
                'Attempt to restore the layer state
                lsm.RestoreLayerState(strlayerStateName, ObjectId.Null, 1, LayerStateMasks.Color Or LayerStateMasks.CurrentViewport _
                                                                            Or LayerStateMasks.Frozen Or LayerStateMasks.LineType _
                                                                            Or LayerStateMasks.LineWeight Or LayerStateMasks.Locked _
                                                                            Or LayerStateMasks.NewViewport Or LayerStateMasks.On _
                                                                            Or LayerStateMasks.Plot Or LayerStateMasks.PlotStyle)
                'Save the changes
                trans.Commit()
                Return True
            Catch ex As Exception
                frms.MessageBox.Show("Layer State: " & strlayerStateName & " not found.", "Restore Layer State", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
                Return False
            End Try
        End Using

    End Function

    ''' <summary>
    ''' Deletes all entities in the required layer
    ''' </summary>
    ''' <param name="strLayer">The layer to delete items from</param>
    ''' <remarks></remarks>
    Public Sub DeleteHatchInLayer(ByVal strLayer As String)
        'Set up the TypedValues to filter for objects on the required layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "HATCH"), New TypedValue(DxfCode.LayerName, strLayer)}
        'Select all the items on the required layer
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then 'If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Loop through the found items
                For Each oid As ObjectId In pres.Value.GetObjectIds
                    Try
                        'Get the entity from the database
                        Dim ent As Entity = DirectCast(trans.GetObject(oid, OpenMode.ForWrite), Entity)
                        'Erase it
                        ent.Erase()
                    Catch ex As Exception
                        frms.MessageBox.Show(ex.Message, "Delete All In Layer " & strLayer, Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                    End Try
                Next
                'Save the changes
                trans.Commit()
            End Using
        End If
    End Sub
    ''' <summary>
    ''' Deletes all entities in the required layer
    ''' </summary>
    ''' <param name="strLayer">The layer to delete items from</param>
    ''' <remarks></remarks>
    Public Sub DeleteAllInLayer(ByVal strLayer As String)
        'Set up the TypedValues to filter for objects on the required layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.LayerName, strLayer)}
        'Select all the items on the required layer
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then 'If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Loop through the found items
                For Each oid As ObjectId In pres.Value.GetObjectIds
                    Try
                        'Get the entity from the database
                        Dim ent As Entity = DirectCast(trans.GetObject(oid, OpenMode.ForWrite), Entity)
                        'Erase it
                        ent.Erase()
                    Catch ex As Exception
                        frms.MessageBox.Show(ex.Message, "Delete All In Layer " & strLayer, Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
                    End Try
                Next
                'Save the changes
                trans.Commit()
            End Using
        End If
    End Sub
    ''' <summary>
    ''' Deletes a empty layer (with jocker)
    ''' </summary>
    ''' <param name="strLayer">The layer to be deleted</param>
    ''' <remarks></remarks>
    Public Function DeleteLayer(ByVal strLayer As String)
        '' Get the current document and database, and start a transaction

        If InStr(strLayer, "*") <> 0 Then 'joker *
            Dim ListLayer As New List(Of String)
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim acLyrTblinit As LayerTable = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)

                For Each LayerID As ObjectId In acLyrTblinit
                    Dim Layer As LayerTableRecord = LayerID.GetObject(OpenMode.ForRead)
                    If Layer.Name.ToLower Like strLayer.ToLower Then
                        ListLayer.Add(Layer.Name)
                    End If
                Next
            End Using

            For Each StrLA As String In ListLayer
                DeleteLayerByName(StrLA)
            Next

        Else 'Pas de joker *
            DeleteLayerByName(strLayer)

        End If
        Return True
    End Function

    ''' <summary>
    ''' Deletes a empty layer (unique)
    ''' </summary>
    ''' <param name="strLayer">The layer to be deleted</param>
    ''' <remarks></remarks>
    Public Function DeleteLayerByName(ByVal strLayer As String)
        '' Get the current document and database, and start a transaction

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
            '' Returns the layer table for the current database
            Dim acLyrTbl As LayerTable
            acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId, OpenMode.ForRead)


            '' Check to see if MyLayer exists in the Layer table
            If acLyrTbl.Has(strLayer) = True Then
                Dim acLyrTblRec As LayerTableRecord
                acLyrTblRec = acTrans.GetObject(acLyrTbl(strLayer), OpenMode.ForWrite)
                Try
                    acLyrTblRec.Erase()
                    acDoc.Editor.WriteMessage(vbLf & strLayer & " : a été supprimé")

                    '' Commit the changes
                    acTrans.Commit()
                Catch
                    acDoc.Editor.WriteMessage(vbLf & strLayer & " : ne peux être supprimé")
                End Try
            Else
                acDoc.Editor.WriteMessage(vbLf & strLayer & " : n'exsite pas")
            End If

            '' Dispose of the transaction
        End Using

        Return True
    End Function
    ''' <summary>
    ''' Copies all objects in the source layer to the destination layer
    ''' </summary>
    ''' <param name="strSourceLayer">The layer to copy objects from</param>
    ''' <param name="strDestLayer">The layer to copy objects to</param>
    ''' <remarks></remarks>
    Public Sub DuplicateAllObjectsInLayer(ByVal strSourceLayer As String, ByVal strDestLayer As String)
        'Set up the typedvalues to filter only objects on the source layer
        Dim Values() As TypedValue = {New TypedValue(DxfCode.LayerName, strSourceLayer)}
        'Select all items on the source layer
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then 'If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Get the modelspace blocktablerecord
                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                'Get the ID of the destination layer
                Dim LayID As ObjectId = GetLayerIDFromLayerName(strDestLayer)
                'Loopp through the found items copying them and setting the layer of the copied item to the destination layer
                For Each oid As ObjectId In pres.Value.GetObjectIds
                    Dim ent As Entity = DirectCast(trans.GetObject(oid, OpenMode.ForRead), Entity)
                    If TypeOf ent Is Polyline Then
                        Dim pLine As Polyline = DirectCast(ent.Clone, Polyline)
                        pLine.LayerId = LayID
                        btr.AppendEntity(pLine)
                        trans.AddNewlyCreatedDBObject(pLine, True)
                    ElseIf TypeOf ent Is Arc Then
                        Dim acArc As Arc = DirectCast(ent.Clone, Arc)
                        acArc.LayerId = LayID
                        btr.AppendEntity(acArc)
                        trans.AddNewlyCreatedDBObject(acArc, True)
                    Else
                        Dim entCopy As Entity = DirectCast(ent.Clone, Entity)
                        entCopy.LayerId = LayID
                        btr.AppendEntity(entCopy)
                        trans.AddNewlyCreatedDBObject(entCopy, True)
                    End If
                Next
                'Save the changes
                trans.Commit()
            End Using
        End If
    End Sub

    ''' <summary>
    ''' Gets the objectid of the layer from the layer name
    ''' </summary>
    ''' <param name="strLayerName">The name of the layer</param>
    ''' <returns>The objectid of the layer</returns>
    ''' <remarks></remarks>
    Public Function GetLayerIDFromLayerName(ByVal strLayerName As String) As ObjectId
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            If lt.Has(strLayerName) Then
                Return lt.Item(strLayerName)
            Else
                Return Nothing
            End If
        End Using
    End Function

    ''' <summary>
    ''' Brings the supplied entity types to the front in the drawing
    ''' </summary>
    ''' <param name="strEntity">The entity type to bring to the front</param>
    ''' <param name="strLayer">Optionally the layer that the entities should be on</param>
    ''' <remarks></remarks>
    Public Sub PlaceObjectsOnTop(ByVal strEntity As String, Optional ByVal strLayer As String = "")
        'Set up the typedvalues to filter out the objects we're looking for
        Dim Values() As TypedValue
        If strLayer = "" Then
            Values = New TypedValue() {New TypedValue(DxfCode.Start, strEntity)}
        Else
            Values = New TypedValue() {New TypedValue(DxfCode.Start, strEntity), New TypedValue(DxfCode.LayerName, strLayer)}
        End If
        'Select all items using the filter values
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then ' If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'Declare a objectidcollection (required for the MoveToTop method of the draw order table)
            Dim entCol As New ObjectIdCollection
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Get the blocktablerecord for modelspace
                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                'Loop through the found objects adding them to the objectidcollection
                For Each oid As ObjectId In pres.Value.GetObjectIds
                    entCol.Add(oid)
                Next


                'If to big - Stop !!!!! 5000 ????


                'Get the drawordertable for modelspace
                Dim dat As DrawOrderTable = DirectCast(trans.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite), DrawOrderTable)
                'Move the objects in the collection to the top
                Try
                    dat.MoveToTop(entCol)
                Catch
                End Try

                'Save the changes
                trans.Commit()

                Dim Connect As New Revo.connect
                Connect.RevoLog("Function 'PlaceObjectsOnTop', Collection count : " & entCol.Count())
            End Using
        End If
    End Sub

    ''' <summary>
    ''' Brings the supplied entity types to the Bottom in the drawing
    ''' </summary>
    ''' <param name="strEntity">The entity type to bring to the front</param>
    ''' <param name="strLayer">Optionally the layer that the entities should be on</param>
    ''' <remarks></remarks>
    Public Sub PlaceObjectsOnBottom(ByVal strEntity As String, Optional ByVal strLayer As String = "")
        'Set up the typedvalues to filter out the objects we're looking for
        Dim Values() As TypedValue
        If strLayer = "" Then
            Values = New TypedValue() {New TypedValue(DxfCode.Start, strEntity)}
        Else
            Values = New TypedValue() {New TypedValue(DxfCode.Start, strEntity), New TypedValue(DxfCode.LayerName, strLayer)}
        End If
        'Select all items using the filter values
        Dim pres As PromptSelectionResult = SelectAllItems(Values)
        If pres.Status = PromptStatus.OK Then ' If we found any...
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'Declare a objectidcollection (required for the MoveToTop method of the draw order table)
            Dim entCol As New ObjectIdCollection
            Using trans As Transaction = db.TransactionManager.StartTransaction
                'Get the blocktablerecord for modelspace
                Dim bt As BlockTable = DirectCast(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(bt.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite), BlockTableRecord)
                'Loop through the found objects adding them to the objectidcollection
                For Each oid As ObjectId In pres.Value.GetObjectIds
                    entCol.Add(oid)
                Next

                'If to big - Stop !!!!! 5000 ????

                'Get the drawordertable for modelspace
                Dim dat As DrawOrderTable = DirectCast(trans.GetObject(btr.DrawOrderTableId, OpenMode.ForWrite), DrawOrderTable)
                'Move the objects in the collection to the top
                Try
                    dat.MoveToBottom(entCol)
                Catch
                End Try

                'Save the changes
                trans.Commit()

                Dim Connect As New Revo.connect
                Connect.RevoLog("Function 'PlaceObjectsOnBottom', Collection count : " & entCol.Count())
            End Using
        End If
    End Sub

    ''' <summary>
    ''' Adds an XRecord to the object
    ''' </summary>
    ''' <param name="objID">The objectID of the object</param>
    ''' <param name="Location">The location (key) at which to store the data</param>
    ''' <param name="ResBuf">The data to store</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Public Function SetObjXRecord(ByVal objID As ObjectId, ByVal Location As String, ByVal ResBuf As ResultBuffer) As Boolean
        'Get the database for the current drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Try
                'Get the entity from the database
                Dim ent As Entity = DirectCast(trans.GetObject(objID, OpenMode.ForRead), Entity)
                'Check if it has an extensiondictionary
                If ent.ExtensionDictionary.IsNull Then
                    'If not, create one
                    ent.UpgradeOpen()
                    ent.CreateExtensionDictionary()
                End If
                'Get the extension dictionary for the object
                Dim extDict As DBDictionary = DirectCast(trans.GetObject(ent.ExtensionDictionary, OpenMode.ForRead), DBDictionary)
                'Create a new XRecord
                Dim myXrecord As New Xrecord
                'add data to the xrecord
                myXrecord.Data = ResBuf
                'create the entry
                extDict.UpgradeOpen()
                extDict.SetAt(Location, myXrecord)
                'tell the transaction about the newly created xrecord
                trans.AddNewlyCreatedDBObject(myXrecord, True)
                'Save the changes
                trans.Commit()
                Return True
            Catch ex As Exception
                MsgBox("Failed to Set Data : " & ex.Message, MsgBoxStyle.Critical)
                Return False
            End Try
        End Using
    End Function

    ''' <summary>
    ''' Gets the data from the XRecord on the supplied object at the supplied location
    ''' </summary>
    ''' <param name="objid">The objectID of the object from which to get the data</param>
    ''' <param name="Location">The Location (Key) where the data was stored on the object</param>
    ''' <returns>A list containing the data from the xrecord</returns>
    ''' <remarks></remarks>
    Public Function GetXRecord(ByVal objid As ObjectId, ByVal Location As String) As List(Of String)
        'Declare the list to return
        Dim retList As New List(Of String)
        'Get the database for the current drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction()
            'Get the entity from the database
            Dim ent As Entity = DirectCast(trans.GetObject(objid, OpenMode.ForRead, False, True), Entity)
            'Make sure it has an extension dictionary
            If Not ent.ExtensionDictionary.IsNull Then
                'Get the extension dictionary
                Dim Dict As DBDictionary = DirectCast(trans.GetObject(ent.ExtensionDictionary, OpenMode.ForRead, False, True), DBDictionary)
                Try
                    'Get the xrecord at the specified location (will error if the location does not exist)
                    Dim entID As ObjectId = Dict.GetAt(Location)
                    Dim xRec As Xrecord = DirectCast(trans.GetObject(entID, OpenMode.ForRead), Xrecord)
                    Dim XRecVal As String = ""
                    'Get the data from the xrecord
                    Dim rb As TypedValue() = xRec.Data.AsArray()
                    'Loop through the values in the data adding each one to the list
                    For i As Integer = 0 To rb.Count - 1
                        retList.Add(rb(i).Value.ToString)
                    Next
                    'Return the list
                    Return retList
                Catch ex As Exception
                    'The location does not exist on the object so return nothing
                    Return Nothing
                End Try
            Else
                'The entity does not have an extension dictionary so return nothing
                Return Nothing
            End If
        End Using
    End Function

    Public Sub UCSOn(ByVal switchOn As Boolean)
        If switchOn Then
            Application.SetSystemVariable("UCSICON", UCSVar)
        Else
            UCSVar = Convert.ToInt32(Application.GetSystemVariable("UCSICON"))
            Application.SetSystemVariable("UCSICON", 2)
        End If
    End Sub


    Public Sub UcsByLine(ByVal p1 As Point3d, ByVal p2 As Point3d) ' Fonction Horiz

        '' Get the current document and database, and start a transaction
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        'Dim pso As New PromptEntityOptions(vbCr & "Select a line : ")
        'pso.SetRejectMessage(vbLf & " Select a line only!")
        'pso.AddAllowedClass(GetType(Line), False)
        'Dim psr As PromptEntityResult = ed.GetEntity(pso)
        'If psr.Status <> PromptStatus.OK Then
        '    Return
        'End If
        'Dim id As ObjectId = psr.ObjectId


        Using tr As Transaction = db.TransactionManager.StartTransaction()
            'Dim ent As Entity = tr.GetObject(id, OpenMode.ForRead)
            Dim ln As Line = (New Line(p1, p2)) ' DirectCast(ent, Line)
            '' Open the UCS table for read
            Dim acUCSTbl As UcsTable
            acUCSTbl = tr.GetObject(db.UcsTableId, _
                                         OpenMode.ForRead)

            Dim acUCSTblRec As UcsTableRecord

            '' Check to see if the "New_UCS" UCS table record exists
            If acUCSTbl.Has("New_UCS") = False Then
                acUCSTblRec = New UcsTableRecord()
                acUCSTblRec.Name = "New_UCS"

                '' Open the UCSTable for write
                acUCSTbl.UpgradeOpen()

                '' Add the new UCS table record
                acUCSTbl.Add(acUCSTblRec)
                tr.AddNewlyCreatedDBObject(acUCSTblRec, True)
            Else
                acUCSTblRec = tr.GetObject(acUCSTbl("New_UCS"), _
                                                OpenMode.ForWrite)
            End If

            acUCSTblRec.Origin = ln.StartPoint
            acUCSTblRec.XAxis = ln.StartPoint.GetVectorTo(ln.EndPoint)
            acUCSTblRec.YAxis = acUCSTblRec.XAxis.GetPerpendicularVector

            '' Open the active viewport
            Dim acVportTblRec As ViewportTableRecord
            acVportTblRec = tr.GetObject(doc.Editor.ActiveViewportId, _
                                              OpenMode.ForWrite)

            '' Display the UCS Icon at the origin of the current viewport
            ' acVportTblRec.IconAtOrigin = True
            ' acVportTblRec.IconEnabled = True
            ' acVportTblRec.UcsFollowMode = True

            '' Set the UCS current
            acVportTblRec.SetUcs(acUCSTblRec.ObjectId)
            doc.Editor.UpdateTiledViewportsFromDatabase()
            Dim view As ViewTableRecord = ed.GetCurrentView()
            ed.SetCurrentView(view)
            '' Display the name of the current UCS
            '  Dim acUCSTblRecActive As UcsTableRecord
            ' acUCSTblRecActive = tr.GetObject(acVportTblRec.UcsName, _
            '                                      OpenMode.ForRead)

            'Application.ShowAlertDialog("The current UCS is: " & _
            '                            acUCSTblRecActive.Name)

            'Dim pPtRes As PromptPointResult
            'Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")

            ''' Prompt for a point
            'pPtOpts.Message = vbLf & "Enter a point: "
            'pPtRes = doc.Editor.GetPoint(pPtOpts)

            'Dim pPt3dWCS As Point3d
            'Dim pPt3dUCS As Point3d

            ''' If a point was entered, then translate it to the current UCS
            'If pPtRes.Status = PromptStatus.OK Then
            '    pPt3dWCS = pPtRes.Value
            '    pPt3dUCS = pPtRes.Value

            '    '' Translate the point from the current UCS to the WCS
            '    Dim newMatrix As Matrix3d = New Matrix3d()
            '    newMatrix = Matrix3d.AlignCoordinateSystem(Point3d.Origin, _
            '                                               Vector3d.XAxis, _
            '                                               Vector3d.YAxis, _
            '                                               Vector3d.ZAxis, _
            '                                               acVportTblRec.Ucs.Origin, _
            '                                               acVportTblRec.Ucs.Xaxis, _
            '                                               acVportTblRec.Ucs.Yaxis, _
            '                                               acVportTblRec.Ucs.Zaxis)

            '    pPt3dWCS = pPt3dWCS.TransformBy(newMatrix)

            '     Application.ShowAlertDialog("The WCS coordinates are: " & vbLf & _
            '                                pPt3dWCS.ToString() & vbLf & _
            '                                "The UCS coordinates are: " & vbLf & _
            '                                pPt3dUCS.ToString())
            'End If

            '' Save the new objects to the database
            tr.Commit()
        End Using

        doc.Editor.UpdateTiledViewportsFromDatabase()





        'Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        ' acDoc.Editor.CurrentUserCoordinateSystem = New Matrix3d(New Double(15) _
        '            {1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0}) 'SCU général

        'Dim zAxis As Vector3d = ed.CurrentUserCoordinateSystem.CoordinateSystem3d.Zaxis
        'Dim xAxis As Vector3d = p1.GetVectorTo(p2)
        'Dim yAxis As Vector3d = xAxis.GetPerpendicularVector 'zAxis.CrossProduct(xAxis)
        'Dim mat As Matrix3d = Matrix3d.AlignCoordinateSystem(Point3d.Origin, Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, p1, xAxis, yAxis, zAxis)
        'ed.CurrentUserCoordinateSystem = mat

        'acVportTblRec.UcsFollowMode = True

    End Sub

    '''///////////////////////////////////////////////////////////////
    ' Use: Sync UCS with Current view like command UCS V
    ' Author: Philippe Leefsma, September 2011
    '''///////////////////////////////////////////////////////////////
    'Function UcsV()
    '    Dim doc As Document = Application.DocumentManager.MdiActiveDocument
    '    Dim db As Database = doc.Database
    '    Dim ed As Editor = doc.Editor

    '    Dim mng As Autodesk.AutoCAD.GraphicsSystem.Manager = doc.GraphicsManager

    '    Dim cvport As Short = CShort(Application.GetSystemVariable("CVPORT"))

    '    Dim view As Autodesk.AutoCAD.GraphicsSystem.View = mng.GetGsView(cvport, True)

    '    Dim direction As Vector3d = (view.Target - view.Position)
    '    direction = direction.MultiplyBy(1 / direction.Length)

    '    Dim upVector As Vector3d = view.UpVector
    '    upVector = upVector.MultiplyBy(1 / upVector.Length)

    '    Dim xAxis As Vector3d = direction.CrossProduct(upVector)

    '    Dim ucs As Matrix3d = Matrix3d.AlignCoordinateSystem(New Point3d(0, 0, 0), Vector3d.XAxis, Vector3d.YAxis, Vector3d.ZAxis, New Point3d(0, 0, 0), xAxis, _
    '     upVector, direction)

    '    ed.CurrentUserCoordinateSystem = ucs

    '    Return True

    'End Function


    ''' <summary>
    ''' Sets all layers in the drawing to either visible or not depending on the value of the booVisible parameter
    ''' </summary>
    ''' <param name="booVisible">True to switch layers ON, False for OFF </param>
    ''' <remarks></remarks>
    Public Sub ShowAllLayers(ByVal booVisible As Boolean)
        'Get the database for this drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the layertable
            Dim lt As LayerTable = DirectCast(trans.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            Try
                'Loop through the layers switching them all On or Off depending of the booVisible parameter
                For Each id As ObjectId In lt
                    'Get the layertablerecord
                    Dim ltr As LayerTableRecord = DirectCast(trans.GetObject(id, OpenMode.ForRead), LayerTableRecord)
                    'Set it's open state to writeable
                    ltr.UpgradeOpen()
                    'Set the IsOff property to the inverse of the booVisible parameter
                    ltr.IsOff = Not (booVisible)
                Next
                'Commit the transaction to save the change
                trans.Commit()
            Catch ex As Exception
                'Something went wrong so notify the user
                frms.MessageBox.Show(ex.Message, "Show Layer", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
    End Sub

    ''' <summary>
    ''' Saves the current layer state to the supplied name
    ''' </summary>
    ''' <param name="LayerStateName">The name to svae the layer state to</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Public Function SaveLayerState(ByVal LayerStateName As String, Optional ByVal Overwrite As Boolean = True) As Boolean
        'Get the database
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the layerstatemanager
            Dim lsm As LayerStateManager = db.LayerStateManager
            Try
                'Check whether the layerstate name already exists
                If lsm.HasLayerState(LayerStateName) Then

                    If Overwrite Then ' Overwrite in any case
                        'They want to overwrite so delete the existing one
                        lsm.DeleteLayerState(LayerStateName)
                    Else
                        'Warn the user and ask if they want to overwrite it
                        If frms.MessageBox.Show("L'état de calques """ & LayerStateName & """ existe déjà  !" & vbCrLf & "Voulez-vous le remplacer ?", _
                                             "Enregistrement d'un état de calques", Windows.Forms.MessageBoxButtons.YesNoCancel, Windows.Forms.MessageBoxIcon.Question) _
                                             = Windows.Forms.DialogResult.Yes Then
                            'They want to overwrite so delete the existing one
                            lsm.DeleteLayerState(LayerStateName)
                        Else
                            'They don't want to overwrite so return false
                            Return False
                        End If
                    End If

                End If
                'Save the layerstate
                lsm.SaveLayerState(LayerStateName, LayerStateMasks.Frozen Or LayerStateMasks.Locked Or LayerStateMasks.On, ObjectId.Null)
                'Save the changes in the database
                trans.Commit()
                Return True
            Catch ex As Exception
                frms.MessageBox.Show(ex.Message, "SaveLayerState", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
        Return False
    End Function

    ''' <summary>
    ''' Removes the supplied layerstate from the layerstates in the drawing
    ''' </summary>
    ''' <param name="LayerStateName">The nameof the layerstate to remove</param>
    ''' <returns>True if successful</returns>
    ''' <remarks></remarks>
    Public Function RemoveLayerState(ByVal LayerStateName As String) As Boolean
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim lsm As LayerStateManager = db.LayerStateManager
            Try
                lsm.DeleteLayerState(LayerStateName)
                trans.Commit()
                lsm.Dispose()
                Return True
            Catch ex As Exception
                frms.MessageBox.Show(ex.Message, "Restore Layer State", Windows.Forms.MessageBoxButtons.OK, _
                                     Windows.Forms.MessageBoxIcon.Error)
            End Try
        End Using
        Return False
    End Function

    ''' <summary>
    ''' Gets an entity selection from the user
    ''' </summary>
    ''' <param name="Types">A list of types to restrict the selection to</param>
    ''' <param name="msg">The message to display to the user while selecting</param>
    ''' <returns>The objectid of the selected entity or nothing if the user cancels</returns>
    ''' <remarks></remarks>
    Public Function GetEntityFromUser(ByVal Types As List(Of Type), ByVal msg As String) As ObjectId
        'Get the document editor
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        'Set up the entity options
        Dim entOpts As New PromptEntityOptions(msg)
        entOpts.SetRejectMessage("L'élément sélectionné n'est pas valide")
        'Add the allowed entity types to the entity options
        For Each t As Type In Types
            entOpts.AddAllowedClass(t, True)
        Next
        'Ask the user to select an entity
        Dim entRes As PromptEntityResult = ed.GetEntity(entOpts)
        'Check that they selected something
        If entRes.Status = PromptStatus.OK Then
            'Return the objectid of the selected entity
            Return entRes.ObjectId
        Else
            'Nothing selected so return nothing
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Gets the polyline that contains the block
    ''' </summary>
    ''' <param name="ptCentroid">the position of the block</param>
    ''' <param name="dblRadius">The size of the window to search around</param>
    ''' <param name="strLayer">the layer that the polylines must be on to be selected</param>
    ''' <returns>The objectid of the polyline that the block is within</returns>
    ''' <remarks></remarks>
    Public Function GetLWPByCentroidAndRect(ByVal ptCentroid As Point3d, ByVal dblRadius As Double, ByVal strLayer As String, Optional SearchBiggest As Boolean = False) As ObjectId
        'Set the bottom left corner point for the selection window
        Dim p1 As New Point3d(ptCentroid.X - dblRadius, ptCentroid.Y - dblRadius, 0)
        'Set the top right corner point for the selection window
        Dim p2 As New Point3d(ptCentroid.X + dblRadius, ptCentroid.Y + dblRadius, 0)
        'Call the function to select the polylines within the window
        Dim colPolys As PlVertCol(Of PolylineVertexes) = GetPolylinesByRect(strLayer, p1, p2)
        'Check that we found some polylines
        If colPolys.Count > 0 Then
            'Collection to hold the polyline(s) that the centroid is within
            Dim FoundPolys As New PlVertCol(Of PolylineVertexes)
            Dim intLoc As Integer 'not used
            'Loop through the polylines checking if the centroid is within them
            For Each pl As PolylineVertexes In colPolys
                'If the polyline doesn't have arc segements
                If pl.WithArcs = False Then
                    'Use the mathematical algorithm to determine if the centroid is within it
                    If PointInsidePolygon(pl.Verts, New Point2d(ptCentroid.X, ptCentroid.Y), intLoc) Then
                        'If it is, then add it to the FoundPolys collection
                        FoundPolys.Add(pl)
                    End If
                Else 'The polyline contains arc segments so use the Ray in AutoCAD method
                    'to determine if the centroid is within it
                    If PtIsInsideLWP(New Point3d(ptCentroid.X, ptCentroid.Y, 0), pl.AcadObject) Then
                        'If it is, then add it to the FoundPolys collection
                        FoundPolys.Add(pl)
                    End If
                End If
            Next
            If FoundPolys.Count = 0 Then 'No polylines found surrounding the centroid so return nothing
                Return Nothing
            ElseIf FoundPolys.Count = 1 Then 'Only one polyline found surrounding the centroid so return that
                Return FoundPolys.Item(0).AcadObject
            Else

                If SearchBiggest = False Then ' Search de smallest

                    'Multiple polylines found surrounding the centroid so 
                    'call the function to find the smallest one
                    Dim Smallest As Integer = FindSmallestPoly(FoundPolys, ptCentroid)
                    'Check that the index of the smallest one is not -1 (shouldn't happen)
                    If Not Smallest = -1 Then
                        'return the objectid of the smallest poly from the collection
                        Return FoundPolys.Item(Smallest).AcadObject
                    Else
                        Return Nothing
                    End If

                Else ' Search de Biggeset

                    'Multiple polylines found surrounding the centroid so 
                    'call the function to find the smallest one
                    Dim Biggest As Integer = FindBiggestPoly(FoundPolys, ptCentroid)
                    'Check that the index of the smallest one is not -1 (shouldn't happen)
                    If Not Biggest = -1 Then
                        'return the objectid of the smallest poly from the collection
                        Return FoundPolys.Item(Biggest).AcadObject
                    Else
                        Return Nothing
                    End If

                End If

            End If
        End If
    End Function

    ''' <summary>
    ''' Gets the Com and Numero information from the block
    ''' </summary>
    ''' <param name="blkID">The ID of the block to get the information from</param>
    ''' <returns>An EDTListItem object containing the information</returns>
    ''' <remarks></remarks>
    Public Function GetInfoFromBlock(ByVal blkID As ObjectId) As EDTListItem
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Declare the return item
        Dim retItem As New EDTListItem
        'Start a transaction
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the blockreference
            Dim blkRef As BlockReference = CType(trans.GetObject(blkID, OpenMode.ForRead), BlockReference)
            'Get the attributecollection
            Dim attCol As AttributeCollection = blkRef.AttributeCollection
            'Loop through the attributes
            For Each attid As ObjectId In attCol
                'Get the attribute reference
                Dim attRef As AttributeReference = CType(trans.GetObject(attid, OpenMode.ForRead), AttributeReference)
                'Check which attribute this is
                Select Case attRef.Tag.ToUpper
                    Case "IDENTDN"
                        'Store this value in the Com property of the return class
                        If Len(attRef.TextString) >= 6 Then retItem.Com = attRef.TextString.Substring(3, 3)
                        If retItem.Com = "" Then retItem.Com = 0
                    Case "NUMERO"
                        'Store this value in the Num propety of the return class
                        retItem.Num = attRef.TextString
                End Select
            Next
            'Set the BlockID property in the return class to the objectid of the block
            retItem.BlockID = blkRef.ObjectId
            'Dispose the blockreference object as we don't need it any more
            blkRef.Dispose()
        End Using
        'Return the item
        Return retItem
    End Function

    ''' <summary>
    ''' Explode object
    ''' </summary>
    ''' <param name="BLid">The Block ID</param>
    ''' <remarks>only for block reference</remarks>
    Public Sub ExplodeBlock(BLid As ObjectId, Optional DeleteOriginal As Boolean = True)
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, _
                                         OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                            OpenMode.ForWrite)

            '' Create a lightweight polyline
            Dim ObjBL As BlockReference = acTrans.GetObject(BLid, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite)

            '' Explodes the polyline
            Dim acDBObjColl As DBObjectCollection = New DBObjectCollection()
            ObjBL.Explode(acDBObjColl)

            For Each acEnt As Entity In acDBObjColl
                '' Add the new object to the block table record and the transaction
                acBlkTblRec.AppendEntity(acEnt)
                acTrans.AddNewlyCreatedDBObject(acEnt, True)
            Next

            'Suppression si OK
            If DeleteOriginal Then ObjBL.Erase()

            '' Save the new objects to the database
            acTrans.Commit()
        End Using
    End Sub

    ''' <summary>
    ''' Convert string Linge Weight in Integer
    ''' </summary>
    ''' <param name="StrWeight">The Name or number of Line Weight</param>
    ''' <returns>Return the Weight Integer</returns>
    ''' <remarks></remarks>
    Public Function ConvertLineWeight(ByVal StrWeight As String) As Integer
        Dim TypeWeight As Integer = 0
        StrWeight = StrWeight.ToLower

        If StrWeight = "bylayer" Then
            TypeWeight = ACAD_LWEIGHT.acLnWtByLayer
        ElseIf StrWeight = "byblock" Then
            TypeWeight = ACAD_LWEIGHT.acLnWtByBlock
        ElseIf StrWeight = "bydefault" Then
            TypeWeight = ACAD_LWEIGHT.acLnWtByLwDefault

        Else 'Autre = numérique
            If IsNumeric(StrWeight) Then
                Select Case CDbl(StrWeight)
                    Case 0.0
                        TypeWeight = ACAD_LWEIGHT.acLnWt000
                    Case 0.05
                        TypeWeight = ACAD_LWEIGHT.acLnWt005
                    Case 0.09
                        TypeWeight = ACAD_LWEIGHT.acLnWt009
                    Case 0.13
                        TypeWeight = ACAD_LWEIGHT.acLnWt013
                    Case 0.15
                        TypeWeight = ACAD_LWEIGHT.acLnWt015
                    Case 0.18
                        TypeWeight = ACAD_LWEIGHT.acLnWt018
                    Case 0.2
                        TypeWeight = ACAD_LWEIGHT.acLnWt020
                    Case 0.25
                        TypeWeight = ACAD_LWEIGHT.acLnWt025
                    Case 0.3
                        TypeWeight = ACAD_LWEIGHT.acLnWt030
                    Case 0.35
                        TypeWeight = ACAD_LWEIGHT.acLnWt035
                    Case 0.4
                        TypeWeight = ACAD_LWEIGHT.acLnWt040
                    Case 0.5
                        TypeWeight = ACAD_LWEIGHT.acLnWt050
                    Case 0.53
                        TypeWeight = ACAD_LWEIGHT.acLnWt053
                    Case 0.6
                        TypeWeight = ACAD_LWEIGHT.acLnWt060
                    Case 0.7
                        TypeWeight = ACAD_LWEIGHT.acLnWt070
                    Case 0.8
                        TypeWeight = ACAD_LWEIGHT.acLnWt080
                    Case 0.9
                        TypeWeight = ACAD_LWEIGHT.acLnWt090
                    Case 1.0
                        TypeWeight = ACAD_LWEIGHT.acLnWt100
                    Case 0.18
                        TypeWeight = ACAD_LWEIGHT.acLnWt106
                    Case 0.18
                        TypeWeight = ACAD_LWEIGHT.acLnWt120
                    Case 0.18
                        TypeWeight = ACAD_LWEIGHT.acLnWt140
                    Case 0.18
                        TypeWeight = ACAD_LWEIGHT.acLnWt158
                    Case 0.18
                        TypeWeight = ACAD_LWEIGHT.acLnWt200
                    Case 0.18
                        TypeWeight = ACAD_LWEIGHT.acLnWt211
                End Select
            End If
        End If

        Return TypeWeight

    End Function
    ''' <summary>
    ''' Convert the supplied polyline to a region
    ''' </summary>
    ''' <param name="polyid">The objectid of the polyline</param>
    ''' <returns>The Region from the polyline</returns>
    ''' <remarks></remarks>
    Public Function ConvertPolylineToRegion(ByVal polyid As ObjectId) As Region
        Dim reg As Region = Nothing
        Try
            'Get the database
            Dim db As Database = HostApplicationServices.WorkingDatabase
            'Declare the polyline
            Dim pline As Polyline = Nothing
            'Get the polyline from the database 
            Using trans As Transaction = db.TransactionManager.StartTransaction
                pline = CType(trans.GetObject(polyid, OpenMode.ForRead), Polyline)
                pline.UpgradeOpen()
                pline.Closed = True
                pline.DowngradeOpen()
                'Declare the dbobjectcollection to 
                Dim objCol As New DBObjectCollection
                objCol.Add(pline)
                Try
                    Dim regCol As DBObjectCollection = Region.CreateFromCurves(objCol)
                    reg = CType(regCol(0), Region)
                    AddObjectToModelSpace(reg)
                    trans.Commit()
                Catch ex As Exception
                    'Error creating the region
                End Try
                pline.Dispose()
                Return reg
            End Using
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "ConvertPolylineToRegion", _
                                 Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
        End Try
        Return reg
    End Function

    ''' <summary>
    ''' Converts the supplied region to a polyline
    ''' </summary>
    ''' <param name="reg">The region to convert</param>
    ''' <param name="booKeepRegion">Optional parameter to specify whether to keep or delete the original region</param>
    ''' <returns>The objectid of the polyline if successful, otherwise nothing</returns>
    ''' <remarks></remarks>
    Public Function ConvertRegionToPolyline(ByVal reg As Region, Optional ByVal booKeepRegion As Boolean = False) As ObjectId
        Try
            'Declare collection to hold the bits of the exploded region
            Dim objCol As New DBObjectCollection
            'Explode the region
            Dim entCol As ObjectIdCollection = ExplodeRegion(reg)
            'Try to join the bits of the exploded region into a polyline
            CommandLine.CommandC("_.PEDIT", "_LAS", "_YES", "_JOI", "_ALL", "", "")
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            'Select the last created entity
            Dim pres As PromptSelectionResult = ed.SelectLast
            If pres.Status = PromptStatus.OK Then
                'Get the database and start a transaction
                Dim db As Database = HostApplicationServices.WorkingDatabase
                Using trans As Transaction = db.TransactionManager.StartTransaction
                    'Get the last entity from the database
                    Dim ent As Entity = CType(trans.GetObject(pres.Value.GetObjectIds(0), OpenMode.ForRead), Entity)
                    'Check whether it's a polyline
                    If TypeOf ent Is Polyline Then
                        'If it is then check whether we want to keep the original region
                        If Not booKeepRegion Then
                            'If not then get it from the database 
                            Dim r As Region = CType(trans.GetObject(reg.ObjectId, OpenMode.ForWrite), Region)
                            'and delete
                            r.Erase()
                            trans.Commit()
                        End If
                        'return the objectid of the polyline
                        Return ent.ObjectId
                    Else
                        'try to join it up a different way
                        CommandLine.CommandC("_.PEDIT", "m", "_all", "_YES", "_JOI")
                        'Select the last created entity
                        pres = ed.SelectLast()
                        If pres.Status = PromptStatus.OK Then
                            'Get the entity from the database
                            ent = CType(trans.GetObject(pres.Value.GetObjectIds(0), OpenMode.ForRead), Entity)
                            'Check if its a polyline
                            If TypeOf ent Is Polyline Then
                                'If it is then check whether we want to delete the original region
                                If Not booKeepRegion Then
                                    'Get the region from the database
                                    Dim r As Region = CType(trans.GetObject(reg.ObjectId, OpenMode.ForWrite), Region)
                                    'Delete it
                                    r.Erase()
                                    trans.Commit()
                                End If
                                Return ent.ObjectId
                            End If
                        End If
                    End If
                End Using
            End If
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "ConvertRegionToPolyline", _
                     Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
        End Try
        Return Nothing
    End Function
    ''' <summary>
    ''' Gets a value from the attributes in the block
    ''' </summary>
    ''' <param name="blkId">The ID of the block to get the attribute from</param>
    ''' <param name="strAttName">The name of the attribute to get the value of</param>
    ''' <returns>The value from the attribute</returns>
    ''' <remarks></remarks>
    Public Function GetBlockAttributes(ByRef objBlk As Autodesk.AutoCAD.Interop.Common.AcadBlock) As Collection

        Dim objEnt As Autodesk.AutoCAD.Interop.Common.AcadEntity
        Dim objAtt As Autodesk.AutoCAD.Interop.Common.AcadAttribute
        Dim coll As New Collection

        For Each objEnt In objBlk
            If objEnt.ObjectName = "AcDbAttributeDefinition" Then
                objAtt = objEnt
                coll.Add(objAtt, objAtt.TagString)
            End If

        Next objEnt
        GetBlockAttributes = coll

    End Function
    ''' <summary>
    ''' Gets a value from the attribute in the block
    ''' </summary>
    ''' <param name="blkId">The ID of the block to get the attribute from</param>
    ''' <param name="strAttName">The name of the attribute to get the value of</param>
    ''' <returns>The value from the attribute</returns>
    ''' <remarks></remarks>
    Public Function GetBlockAttribute(ByVal blkId As ObjectId, ByVal strAttName As String) As String
        Dim retVal As String = String.Empty
        Try
            'Get the database and start a transaction
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
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
            End Using
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetBlockAttribute", _
                                 Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Exclamation)
        End Try
        Return retVal
    End Function

    ''' <summary>
    ''' Selects objects within the supplied poly 
    ''' </summary>
    ''' <param name="values">The typedvalues to create the selection filter</param>
    ''' <param name="verts">the vertices of the polyline to select within</param>
    ''' <returns>The objectids of the found items if any</returns>
    ''' <remarks></remarks>
    Public Function SelectCrossingPoly(ByVal values() As TypedValue, ByVal verts As Point3dCollection) As ObjectId()
        Dim sFilter As New SelectionFilter(values)
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim pres As PromptSelectionResult = ed.SelectCrossingPolygon(verts, sFilter)
        If pres.Status = PromptStatus.OK Then
            Return pres.Value.GetObjectIds
        Else
            Return Nothing
        End If


    End Function

    ''' <summary>
    ''' Selects objects within the supplied poly 
    ''' </summary>
    ''' <param name="values">The typedvalues to create the selection filter</param>
    ''' <param name="verts">the vertices of the polyline to select within</param>
    ''' <returns>The objectids of the found items if any</returns>
    ''' <remarks></remarks>
    Public Function SelectWindowPoly(ByVal values() As TypedValue, ByVal verts As Point3dCollection) As ObjectId()
        Dim sFilter As New SelectionFilter(values)
        Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
        Dim pres As PromptSelectionResult = ed.SelectWindowPolygon(verts, sFilter)
        If pres.Status = PromptStatus.OK Then
            Return pres.Value.GetObjectIds
        Else
            Return Nothing
        End If
    End Function

    ''' <summary>
    ''' Adds the supplied object to modelspace
    ''' </summary>
    ''' <param name="obj">The object to add</param>
    ''' <returns>The objectid of the added object</returns>
    ''' <remarks></remarks>
    Public Function AddObjectToModelSpace(ByVal obj As Entity) As ObjectId
        Dim db As Database = Application.DocumentManager.MdiActiveDocument.Database
        Dim objId As ObjectId
        Using trans As Transaction = db.TransactionManager.StartTransaction
            Dim bt As BlockTable = CType(trans.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = CType(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
            btr.UpgradeOpen()
            objId = btr.AppendEntity(obj)
            trans.AddNewlyCreatedDBObject(obj, True)
            trans.Commit()
        End Using
        Return objId
    End Function


    ''' <summary>
    ''' Adds Hatch to surface object
    ''' </summary>
    ''' <param name="obj">The object to add</param>
    ''' <returns>The objectid of the added object</returns>
    ''' <remarks>obj is an entity that has already been added to the database.</remarks>
    Public Sub OLD_addHatch2(ByRef acObjIdColl As ObjectIdCollection)
        'Start(Transaction)
        'Step 1. Create the hatch object
        'Step 2. Setup how the hatch will look
        'Step 3. Add it to the Block Table record
        'Step 4. Now, assign the associativity
        'Step 5. Let the transaction know of the new object and commit
        'End Transaction
        '        Start(Transaction)
        'Step 6. Open the hatch object and append the loops
        'Step 7. Evaluate
        'Step 8. Commit
        'End Transaction


        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acDoc.Database
        Dim tm As Autodesk.AutoCAD.DatabaseServices.TransactionManager = db.TransactionManager
        Dim NewID As ObjectId = ObjectId.Null
        Using myT As Transaction = db.TransactionManager.StartTransaction()
            Try
                Using table As BlockTable = TryCast(myT.GetObject(db.BlockTableId, OpenMode.ForRead, False), BlockTable)
                    Using record As BlockTableRecord = TryCast(myT.GetObject(table(BlockTableRecord.ModelSpace), OpenMode.ForWrite, False), BlockTableRecord)
                        ' Dim o As Entity = DirectCast(myT.GetObject(obj, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, False), Entity)
                        Dim h As New Hatch()
                        h.SetDatabaseDefaults()
                        ' h.Layer = o.Layer
                        ' o.Dispose()
                        h.HatchStyle = Autodesk.AutoCAD.DatabaseServices.HatchStyle.Normal
                        h.SetHatchPattern(Autodesk.AutoCAD.DatabaseServices.HatchPatternType.PreDefined, "SOLID")
                        Dim oID As ObjectId = record.AppendEntity(h)
                        If oID.IsValid Then
                            NewID = oID
                            h.Associative = True
                            myT.AddNewlyCreatedDBObject(h, True)
                            myT.Commit()
                        Else
                            myT.Abort()
                        End If
                    End Using
                End Using
            Catch ex As Exception
                myT.Abort()
            End Try
        End Using

        If NewID <> ObjectId.Null Then
            Using myT As Transaction = tm.StartTransaction()
                Try
                    Dim h As Hatch = DirectCast(myT.GetObject(NewID, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, False), Hatch)
                    Dim mylist As ObjectIdCollection = acObjIdColl '(New ObjectId() {obj})
                    h.AppendLoop(HatchLoopTypes.Default, mylist)
                    h.EvaluateHatch(False)
                    h.Dispose()
                    myT.Commit()
                Catch ex As Exception
                    myT.Abort()
                End Try
            End Using
        End If
        tm.Dispose()
        db.Dispose()

    End Sub



    ''' <summary>
    ''' Adds Hatch to surface object
    ''' </summary>
    ''' <param name="acObjIdColl">The object Collection to add</param>
    ''' <param name="TypeHatch">The style Hatch (ANSI31) défault SOLID</param>
    ''' <returns>The objectid of the added object</returns>
    ''' <remarks>obj is an entity that has already been added to the database.</remarks>
    Public Function AddHatch(ObjIDColl As Collection, ObjOutLayers As String, HatchType As String, HatchScale As Double, acLayer As String, acLayerHa As String, Annotatif As Boolean)
        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        If ObjOutLayers <> "" Then ObjOutLayers = ObjOutLayers & ","
        ObjOutLayers = ObjOutLayers & acLayer
        If HatchType = "" Then HatchType = "SOLID"

        ''Tri de la collection (Revo.AcObjID : objectid, Area)
        Dim SortIdColl As Collection = SortAcObjIDColl(ObjIDColl, False)

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            '' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            '' Open the Block table record Model space for write
            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                            OpenMode.ForWrite)

            For u = 1 To SortIdColl.Count 'acObjId As ObjectId In acObjIdColl

                '' Recherche les surfaces localisée à l'interieur du périmètre
                Dim acObj As Revo.AcObjID = SortIdColl(u)
                Dim Verts As Point3dCollection = GetPolyVertices(acObj.objID)
                Dim Values() As TypedValue = {New TypedValue(DxfCode.Start, "LWPOLYLINE"), New TypedValue(DxfCode.LayerName, ObjOutLayers)}
                Dim IDs() As ObjectId = SelectWindowPoly(Values, Verts)
                Dim CurrObjOutIdColl As New ObjectIdCollection '= acObjOutIdColl

                If IDs IsNot Nothing Then
                    Dim PolyA As Polyline = acTrans.GetObject(acObj.objID, OpenMode.ForRead)

                    For i = 0 To IDs.Length - 1
                        Dim PolyB As Polyline = acTrans.GetObject(IDs(i), OpenMode.ForRead)

                        If IDs(i) <> acObj.objID Then 'Différent de l'objet utilisé 
                            CurrObjOutIdColl.Add(IDs(i))
                            For x = 1 To SortIdColl.Count
                                Dim xObj As Revo.AcObjID = SortIdColl(x)
                                If IDs(i) = xObj.objID Then SortIdColl.Remove(x) : Exit For
                            Next

                        End If
                    Next
                End If

                '' Create the hatch object and append it to the block table record
                Dim acHatch As Hatch = New Hatch()
                Try
                    If acLayerHa <> "" Then acHatch.Layer = acLayerHa
                Catch
                    'Erreur calques
                    Dim Connect As New Revo.connect
                    Connect.RevoLog(Connect.DateLog & "Cmd HA" & vbTab & False & vbTab & "Err layer : " & acLayerHa)
                End Try


                acBlkTblRec.AppendEntity(acHatch)
                acTrans.AddNewlyCreatedDBObject(acHatch, True)

                '' Set the properties of the hatch object
                '' Associative must be set after the hatch object is appended to the 
                '' block table record and before AppendLoop
                If HatchScale > 0 Then acHatch.PatternScale = HatchScale
                acHatch.HatchStyle = HatchStyle.Normal

                acHatch.SetHatchPattern(HatchPatternType.PreDefined, HatchType)
                acHatch.Associative = True
                acHatch.AppendLoop(HatchLoopTypes.Default, New ObjectIdCollection(New ObjectId() {acObj.objID}))

                If Annotatif Then
                    acHatch.Annotative = AnnotativeStates.True
                Else
                    acHatch.Annotative = AnnotativeStates.False
                End If

                For Each ObjOut As ObjectId In CurrObjOutIdColl
                    acHatch.AppendLoop(HatchLoopTypes.External, New ObjectIdCollection(New ObjectId() {ObjOut}))
                Next

                acHatch.EvaluateHatch(True)
                If u = SortIdColl.Count Then Exit For
            Next


            '' Save the new object to the database
            acTrans.Commit()


        End Using

        Return True
    End Function


    Public Class EDTListItem
        Private _ID As String
        Private _PolyLayer As String
        Private _PolyID As ObjectId
        Private _BlockID As ObjectId
        Private _strCom As String
        Private _strNum As String
        Private _DisplayVal As String

        Public Property ID() As String
            Get
                Return _ID
            End Get
            Set(ByVal value As String)
                _ID = value
            End Set
        End Property
        Public Property PolyLayer() As String
            Get
                Return _PolyLayer
            End Get
            Set(ByVal value As String)
                _PolyLayer = value
            End Set
        End Property
        Public Property PolyID() As ObjectId
            Get
                Return _PolyID
            End Get
            Set(ByVal value As ObjectId)
                _PolyID = value
            End Set
        End Property
        Public Property BlockID() As ObjectId
            Get
                Return _BlockID
            End Get
            Set(ByVal value As ObjectId)
                _BlockID = value
            End Set
        End Property
        Public Property Com() As String
            Get
                Return _strCom
            End Get
            Set(ByVal value As String)
                _strCom = value
            End Set
        End Property
        Public Property Num() As String
            Get
                Return _strNum
            End Get
            Set(ByVal value As String)
                _strNum = value
            End Set
        End Property
        Public ReadOnly Property ComNum() As String
            Get
                Return _strCom & " / " & _strNum
            End Get
        End Property
        Public Property DisplayVal() As String
            Get
                Return _DisplayVal
            End Get
            Set(ByVal value As String)
                _DisplayVal = value
            End Set
        End Property
    End Class

    Public Class PerListItem
        Private _ID As String
        Private _PolyLayer As String
        Private _PolyID As ObjectId
        Private _BlockID As ObjectId
        Private _strCom As String
        Private _strNum As String
        Private _DisplayVal As String
        Private _BFNum As String

        Public Property ID() As String
            Get
                Return _ID
            End Get
            Set(ByVal value As String)
                _ID = value
            End Set
        End Property
        Public Property PolyLayer() As String
            Get
                Return _PolyLayer
            End Get
            Set(ByVal value As String)
                _PolyLayer = value
            End Set
        End Property
        Public Property PolyID() As ObjectId
            Get
                Return _PolyID
            End Get
            Set(ByVal value As ObjectId)
                _PolyID = value
            End Set
        End Property
        Public Property BlockID() As ObjectId
            Get
                Return _BlockID
            End Get
            Set(ByVal value As ObjectId)
                _BlockID = value
            End Set
        End Property
        Public Property Com() As String
            Get
                Return _strCom
            End Get
            Set(ByVal value As String)
                _strCom = value
            End Set
        End Property
        Public Property Num() As String
            Get
                Return _strNum
            End Get
            Set(ByVal value As String)
                _strNum = value
            End Set
        End Property
        Public ReadOnly Property ComNum() As String
            Get
                Return _strCom & " / " & _strNum
            End Get
        End Property
        Public Property DisplayVal() As String
            Get
                Return _DisplayVal
            End Get
            Set(ByVal value As String)
                _DisplayVal = value
            End Set
        End Property
        Public Property BF() As String
            Get
                Return _BFNum
            End Get
            Set(ByVal value As String)
                _BFNum = value
            End Set
        End Property
    End Class

    Public Function GetRadiusFromBulge(ByVal dbleE As Double, ByVal dblN As Double, _
                                       ByVal dblE2 As Double, ByVal dblN2 As Double, ByVal dblBulge As Double) As Double
        If dblBulge = 0 Then
            Return 0
        Else
            'Calculate the chord length
            Dim dblC As Double = Dist(New Point2d(dbleE, dblN), New Point2d(dblE2, dblN2))
            Dim dblR As Double = dblC / (2 * Math.Sin(2 * Math.Atan(dblBulge)))
            Return dblR
        End If

    End Function


    Function ReplaceBlock(AcBlock As BlockReference, NewNameBL As String)


        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite, True, True)
            Dim btrSpace As BlockTableRecord = DirectCast(acTrans.GetObject(bt(BlockTableRecord.ModelSpace), Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite), BlockTableRecord)

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
                        NewBL.Layer = AcBlock.Layer 'Couche cible
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

                    End If
                Else

                End If

            Catch ex As Exception

            End Try


            acTrans.Commit() ' Commit the transaction

        End Using

        Return True
    End Function

    Function OLD_ReplaceBlock(ByVal OldBlock As String, ByVal NewBlock As String)

        Dim oldBLK, NewBLK As AcadBlockReference
        Dim oldLAY As String
        Dim oldSCL As Long
        Dim oldORIENT As Long

        Static newBLKname As String
        Dim retObj As AcadObject = Nothing
        Dim basePnt As New Object
        Dim Atts1, Atts2 As Object
        Dim i As Integer
        oldBLK = retObj
        basePnt = oldBLK.InsertionPoint
        oldSCL = oldBLK.XScaleFactor
        oldORIENT = oldBLK.Rotation
        oldLAY = oldBLK.Layer
        'SETTING DOC ref

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim ThisDWG As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim ThisDWG As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

        '-----------MAIN INSERTION SEQUENCE
        NewBLK = ThisDWG.ModelSpace.InsertBlock(basePnt, newBLKname, oldBLK.XScaleFactor, oldBLK.YScaleFactor, oldBLK.ZScaleFactor, oldBLK.Rotation) 'ThisDrawing.Utility.AngleToReal(oldORIENT, acDegrees))
        NewBLK.Layer = oldBLK.Layer
        NewBLK.color = oldBLK.color
        If NewBLK.HasAttributes And oldBLK.HasAttributes Then
            Atts1 = oldBLK.GetAttributes
            Atts2 = NewBLK.GetAttributes
            'This will Copy Attributes as the minimum count
            For i = 0 To Min(UBound(Atts1), UBound(Atts2)) - 1
                Atts2(i).TextString = Atts1(i).TextString
            Next i
        End If
        'Delete old Blk
        oldBLK.Delete()
        '-----------


        If False Then ' PROMT INFO, put True to prompt INFOS
            MsgBox("Old Block was: " & oldBLK.Name & " at " & basePnt(0) & "," & basePnt(1) & " on Layer " & oldLAY _
            & vbCr & "Scale is: " & oldSCL & " Orientation is: " & oldORIENT)
        End If

        Return True
    End Function

    Sub OLDReplaceBlock()
        'DECLARATIONS
        Dim oldBLK, NewBLK As AcadBlockReference
        Dim oldLAY As String
        Dim oldSCL As Long
        Dim oldORIENT As Long

        Static newBLKname As String
        Dim retObj As AcadObject = Nothing
        Dim basePnt As New Object
        Dim Atts1, Atts2 As Object
        Dim i As Integer

        'SETTING DOC ref
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim ThisDWG As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim ThisDWG As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

        'Get the Name of New BLOCK to insert
        '********** You MUST MATCH the Up/low Case on the name ************

        newBLKname = GetSameString(newBLKname, "Name of the BLOCK to insert:")

        'Test if that BLOCK is in the DWG
        If IsBlockIN(newBLKname) Then
            ThisDWG.Utility.GetEntity(retObj, basePnt, "Select a BLOCK to Replace")
            If retObj.ObjectName = "AcDbBlockReference" Then
                oldBLK = retObj
                basePnt = oldBLK.InsertionPoint
                oldSCL = oldBLK.XScaleFactor
                oldORIENT = oldBLK.Rotation
                oldLAY = oldBLK.Layer


                '-----------MAIN INSERTION SEQUENCE
                NewBLK = ThisDWG.ModelSpace.InsertBlock(basePnt, newBLKname, oldBLK.XScaleFactor, oldBLK.YScaleFactor, oldBLK.ZScaleFactor, oldBLK.Rotation) 'ThisDrawing.Utility.AngleToReal(oldORIENT, acDegrees))
                NewBLK.Layer = oldBLK.Layer
                NewBLK.color = oldBLK.color
                If NewBLK.HasAttributes And oldBLK.HasAttributes Then
                    Atts1 = oldBLK.GetAttributes
                    Atts2 = NewBLK.GetAttributes
                    'This will Copy Attributes as the minimum count
                    For i = 0 To Min(UBound(Atts1), UBound(Atts2)) - 1
                        Atts2(i).TextString = Atts1(i).TextString
                    Next i
                End If
                'Delete old Blk
                oldBLK.Delete()
                '-----------



                If False Then ' PROMT INFO, put True to prompt INFOS
                    MsgBox("Old Block was: " & oldBLK.Name & " at " & basePnt(0) & "," & basePnt(1) & " on Layer " & oldLAY _
                    & vbCr & "Scale is: " & oldSCL & " Orientation is: " & oldORIENT)
                End If
            Else
                MsgBox("The selection was NOT a BlockReference" & _
                vbCr & "It is a : " & retObj.ObjectName, vbExclamation, "Block Selection")
                GoTo GoOUT
            End If
        Else
            MsgBox("This BLOCK: " & newBLKname & " is NOT in this Drawing.", vbCritical, "Block Search")
        End If


GoOUT:

    End Sub



    '============================================ FUNCTIONS =======================================

    Public Function IsBlockIN(ByVal BlockN As String) As Boolean
        Dim i As Integer
        IsBlockIN = False
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim ThisDWG As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim ThisDWG As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
        For i = 0 To ThisDWG.Blocks.Count - 1
            If ThisDWG.Blocks.Item(i).Name.ToUpper = BlockN.ToUpper Then
                IsBlockIN = True
                Exit Function
            End If
        Next i
    End Function

    Function Min(ByVal n1 As Object, ByVal n2 As Object) As Object
        Min = IIf(n1 < n2, n1, n2)
    End Function

    Function Max(ByVal n1 As Object, ByVal n2 As Object) As Object
        Max = IIf(n1 > n2, n1, n2)
    End Function

    Function GetSameString(ByVal oldStr As String, ByVal Prmt As String) As String
        Dim inputSTR As String
        'If IsNull(oldStr) Then oldStr = ""
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim ThisDWG As AcadDocument = Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim ThisDWG As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
        inputSTR = ThisDWG.Utility.GetString(False, Prmt & ":<" & oldStr & "> :")
        If Len(inputSTR) = 0 Then
            GetSameString = oldStr
        Else
            GetSameString = inputSTR
        End If
    End Function

    Public Sub ClearXData(ByRef objEntity As Autodesk.AutoCAD.Interop.Common.AcadEntity)

        Dim xType, xData As New Object
        Dim xTypeClear(0) As Short
        Dim xDataClear(0) As Object
        Dim i As Integer

        Try
            xTypeClear(0) = 1001
            objEntity.GetXData("", xType, xData)

            For i = LBound(xType) To UBound(xType)
                If xType(i) = 1001 Then
                    xDataClear(0) = xData(i)
                    objEntity.SetXData(xTypeClear, xDataClear)
                End If
            Next

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try

    End Sub

    Public Sub ClearAllXData(ID As ObjectId)
        Dim db As Database = HostApplicationServices.WorkingDatabase

        Try
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim ent As Entity = DirectCast(tr.GetObject(ID, OpenMode.ForRead), Entity)
                If ent.XData Is Nothing Then
                    ' ed.WriteMessage(vbLf & "No XData in the entity.")
                    Return
                End If

                For Each tv As TypedValue In ent.XData.AsArray().Where(Function(e) e.TypeCode = 1001)
                    ent.UpgradeOpen()
                    ent.XData = New ResultBuffer(New TypedValue(1001, tv.Value))
                    ent.DowngradeOpen()
                Next

                tr.Commit()
            End Using

        Catch ex As System.Exception
            ' ed.WriteMessage(ex.ToString())
        End Try
    End Sub

    Public Function SelectPolyline(Optional ByVal controle_surf As Boolean = True, Optional ByVal SingleSelect As Boolean = True, Optional ByVal Message As String = "Sélectionner la poyligne") As Collection

        ''Selection of the polylign
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim Poly As Polyline
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
                            If TypeName(acEnt) Like "Polyline*" Then
                                If controle_surf Then
                                    Poly = acEnt
                                    If Poly.Area <> 0 Then
                                        collection_retour.Add(acEnt)
                                    Else
                                        ed.WriteMessage(vbLf & "Aucun éléments sélectionné non valide (surface nulle)")
                                        ed.WriteMessage(vbLf & "('ECHAP' pour quitter)")
                                        collection_retour = SelectPolyline(controle_surf)
                                    End If
                                Else
                                    collection_retour.Add(acEnt)
                                End If

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
                    collection_retour = SelectPolyline(controle_surf)
                End If
            End If
            ''End of the selection of the polylign 
        End Using
        Return collection_retour
    End Function

    Public Function CreateFilterListForBlocks(blkNames As List(Of String)) As TypedValue()
        ' If we don't have any block names, return null

        If blkNames.Count = 0 Then
            Return Nothing
        End If

        ' If we only have one, return an array of a single value

        If blkNames.Count = 1 Then
            Return New TypedValue() {New TypedValue(CInt(DxfCode.BlockName), blkNames(0))}
        End If

        ' We have more than one block names to search for...

        ' Create a list big enough for our block names plus
        ' the containing "or" operators

        Dim tvl As New List(Of TypedValue)(blkNames.Count + 2)

        ' Add the initial operator

        tvl.Add(New TypedValue(CInt(DxfCode.[Operator]), "<or"))

        ' Add an entry for each block name, prefixing the
        ' anonymous block names with a reverse apostrophe

        For Each blkName In blkNames
            tvl.Add(New TypedValue(CInt(DxfCode.BlockName), (If(blkName.StartsWith("*"), "`" & blkName, blkName))))
        Next

        ' Add the final operator

        tvl.Add(New TypedValue(CInt(DxfCode.[Operator]), "or>"))

        ' Return an array from the list

        Return tvl.ToArray()
    End Function



    Public Function SelectObj(Optional ByVal SingleSelect As Boolean = True, Optional ByVal Message As String = "Sélectionner les objets", Optional ByVal ActivePolyligne As Boolean = True, Optional ByVal ActiveTexte As Boolean = True, Optional ByVal ActiveBlock As Boolean = False) As Collection

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
                    collection_retour = SelectPolyline()
                End If
            End If
            ''End of the selection of the polylign 
        End Using
        Return collection_retour
    End Function

    Public Function BreakPolyline(ByVal PolyID As ObjectId, ByVal Pts As Point3d)

        Dim CollPolyID As New Collection

        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        '  Dim opt1 As New PromptEntityOptions(vbLf & "select a line:")
        '  opt1.SetRejectMessage(vbLf & "error!")
        ' opt1.AddAllowedClass(GetType(Line), True)
        '  Dim res1 As PromptEntityResult = ed.GetEntity(opt1)

        '   If res1.Status = PromptStatus.OK Then

        'Dim opt2 As New PromptPointOptions(vbLf & "select second point:")
        ' opt2.AllowNone = True
        ' Dim res2 As PromptPointResult = ed.GetPoint(opt2)

        Using tr As Transaction = db.TransactionManager.StartTransaction()

            Dim l As Polyline = DirectCast(tr.GetObject(PolyID, OpenMode.ForRead), Polyline)
            Dim pars As New List(Of Double)()
            Dim pt1 As Point3d = Pts 'res2.Value ' l.GetClosestPointTo(res1.PickedPoint, False)
            Dim pt2 As New Point3d()
            pars.Add(l.GetParameterAtPoint(pt1))


            Dim btr As BlockTableRecord = DirectCast(tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite, False), BlockTableRecord)

            ' If res2.Status = PromptStatus.OK Then

            pt2 = l.GetClosestPointTo(Pts, False)
            pars.Add(l.GetParameterAtPoint(pt2))
            pars.Sort()
            Dim objs As DBObjectCollection = l.GetSplitCurves(New DoubleCollection(pars.ToArray()))
            For Each ll As Polyline In objs

                If (ll.StartPoint <> pt1 AndAlso ll.StartPoint <> pt2) Xor (ll.EndPoint <> pt1 AndAlso ll.EndPoint <> pt2) Then

                    btr.AppendEntity(ll)

                    tr.AddNewlyCreatedDBObject(ll, True)
                    'CollPolyID.Add(ll.ObjectId)
                    Ncolor += 40 : If Ncolor > 255 Then Ncolor = 2
                    ll.ColorIndex = Ncolor
                    PolyID = (ll.ObjectId)
                End If


            Next
            'Else

            'Dim objs As DBObjectCollection = l.GetSplitCurves(New DoubleCollection(pars.ToArray()))
            'For Each ll As Line In objs

            '    btr.AppendEntity(ll)

            '    tr.AddNewlyCreatedDBObject(ll, True)
            'Next
            'End If

            l.UpgradeOpen()
            l.[Erase]()

            tr.Commit()
        End Using
        ' End If

        Return PolyID 'CollPolyID

    End Function


    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetObjectId(db As Database, handle As String) As ObjectId
        Dim h As New Handle(Int64.Parse(handle, Globalization.NumberStyles.AllowHexSpecifier))
        Dim id As ObjectId = ObjectId.Null
        db.TryGetObjectId(h, id)
        'TryGetObjectId﻿ method
        Return id
    End Function



    Public Function GetBlocksByAttributes(blkname As String, aTag As String, aVal As String, Optional LayerName As String = "")
        Dim blkcoll As New Collection 'As New ObjectIdCollection()
        'block name
        '  Dim blkname As String = "MyBlock" '<-- change the block name here
        ' tag name to search for 
        'Dim atag As String = "TAG1"  '<-- change the attribute tag here
        'new attribute value
        'Dim newval As String = "Value1"  '<-- change the new attribute value here
        ' get active drawing
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        ' get document editor
        Dim ed As Editor = doc.Editor
        'get document database
        Dim db As Database = doc.Database

        ' build selection filter 
        'Dim sfilter As New SelectionFilter(New TypedValue() {New TypedValue(-4, "<AND"), New TypedValue(410, "Model"), New TypedValue(CInt(DxfCode.BlockName), blkname), New TypedValue(CInt(DxfCode.HasSubentities), 1), New TypedValue(-4, "AND>"), ValuesLayer})

        ' request for blocks to be selected all in the  model space 
        ' Dim res As PromptSelectionResult = ed.SelectAll(sfilter


        Dim Values() As TypedValue
        If LayerName = "" Then
            Values = New TypedValue() {New TypedValue(-4, "<AND"), New TypedValue(410, "Model"), New TypedValue(CInt(DxfCode.BlockName), blkname), New TypedValue(CInt(DxfCode.HasSubentities), 1), New TypedValue(-4, "AND>")}
        Else
            Values = New TypedValue() {New TypedValue(-4, "<AND"), New TypedValue(410, "Model"), New TypedValue(CInt(DxfCode.BlockName), blkname), New TypedValue(CInt(DxfCode.HasSubentities), 1), New TypedValue(-4, "AND>"), New TypedValue(DxfCode.LayerName, LayerName)}
        End If
        'Select all items using the filter values
        Dim res As PromptSelectionResult = SelectAllItems(Values)
        Try
            'check on valid selection result
            If res.Status = PromptStatus.OK Then
                'display result
                ed.WriteMessage(vbLf & "--->  Selected {0} objects", res.Value.Count) ' debug only, maybe removed
                'get object transaction 
                Using tr As Transaction = doc.TransactionManager.StartTransaction()
                    Dim ids As ObjectId() = res.Value.GetObjectIds()

                    'iterate through the selection set
                    For Each id As ObjectId In ids
                        If id.IsValid AndAlso Not id.IsErased Then
                            Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead, False), Entity)
                            ed.WriteMessage(vbLf & "--->  {0}", ent.GetRXClass().DxfName) ' debug only, maybe removed
                            'cast entity as BlockReference
                            Dim bref As BlockReference = TryCast(ent, BlockReference)
                            If bref IsNot Nothing Then
                                For Each aid As ObjectId In bref.AttributeCollection
                                    Dim subent As Entity = TryCast(tr.GetObject(aid, OpenMode.ForRead, False), Entity)
                                    If TypeOf subent Is AttributeReference Then
                                        'cast entity as AttributeReference
                                        Dim atref As AttributeReference = TryCast(subent, AttributeReference)

                                        If atref IsNot Nothing Then
                                            ' check on desired tag name
                                            If atref.Tag.ToUpper Like aTag.ToUpper Then
                                                If atref.TextString.ToUpper Like aVal.ToUpper Then
                                                    '  MsgBox(atref.TextString)
                                                    blkcoll.Add(bref)
                                                    Exit For
                                                    ' atref.UpgradeOpen()
                                                    ' atref.TextString = aVal
                                                    ' atref.DowngradeOpen()
                                                End If

                                            End If
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    Next
                    tr.Commit()

                End Using
            End If
        Catch ex As System.Exception
            ed.WriteMessage(ex.Message)
        End Try

        Return blkcoll 'Collection BlockReference

        'Dim ss, ss2 As Autodesk.AutoCAD.Interop.AcadSelectionSet

    End Function


    Public Function PlotCurrentLayout(PC3Name As String, FormatName As String, ScaleNum As String, ScaleDenom As String, CTBName As String, SaveSettings As Boolean, Print As Boolean)
        '' Get the current document and database, and start a transaction
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Try

                'Formatage du nom du format de papier
                FormatName = Replace(FormatName, " ", "_")

                '' Reference the Layout Manager
                Dim acLayoutMgr As LayoutManager
                acLayoutMgr = LayoutManager.Current

                '' Get the current layout and output its name in the Command Line window
                Dim acLayout As Layout
                acLayout = acTrans.GetObject(acLayoutMgr.GetLayoutId(acLayoutMgr.CurrentLayout), _
                                             OpenMode.ForWrite)

                '' Get the PlotInfo from the layout
                Dim acPlInfo As PlotInfo = New PlotInfo()
                acPlInfo.Layout = acLayout.ObjectId

                '' Get a copy of the PlotSettings from the layout
                Dim acPlSet As PlotSettings = New PlotSettings(acLayout.ModelType)
                acPlSet.CopyFrom(acLayout)

                '' Update the PlotSettings object
                Dim acPlSetVdr As PlotSettingsValidator = PlotSettingsValidator.Current

                '' Set the plot type
                acPlSetVdr.SetPlotType(acPlSet, Autodesk.AutoCAD.DatabaseServices.PlotType.Layout)

                '' Set the plot scale
                If ScaleNum <> "" Or ScaleDenom <> "" Then
                    acPlSetVdr.SetUseStandardScale(acPlSet, False) 'acPlSetVdr.SetUseStandardScale(acPlSet, True)
                    If IsNumeric(ScaleNum) = False Then ScaleNum = 1
                    If IsNumeric(ScaleDenom) = False Then ScaleDenom = 1
                    Dim PersoScale As CustomScale = New CustomScale(CDbl(Trim(ScaleNum)), CDbl(Trim(ScaleDenom)))
                    acPlSetVdr.SetCustomPrintScale(acPlSet, PersoScale) ' StdScaleType.ScaleToFit
                End If

                '' Center the plot
                ' acPlSetVdr.SetPlotCentered(acPlSet, True)

                'Listing des CTB
                Dim CollStyle As System.Collections.Specialized.StringCollection = acPlSetVdr.GetPlotStyleSheetList()
                If CTBName <> "" And CTBName.ToLower <> "none" Then
                    If CollStyle.Contains(CTBName) = True Then 'Si CTB trouvé applique (sinon ne fais rien)

                        For Each StyleCTB In CollStyle
                            MsgBox(StyleCTB.ToString)
                        Next
                        ' acPlSetVdr.SetCurrentStyleSheet(acPlSet, CTBName)
                    End If

                ElseIf CTBName.ToLower = "none" Then
                    acPlSetVdr.SetCurrentStyleSheet(acPlSet, "")
                End If



                ' Set the plot device to use
                If PC3Name = "" Or FormatName = "" Then
                    'ne fait rien
                Else
                    acPlSetVdr.SetPlotConfigurationName(acPlSet, PC3Name, FormatName) ' "DWF6 ePlot.pc3", "ANSI_A_(8.50_x_11.00_Inches)"
                End If


                ''Listing des plotters
                'Dim CollPlot As System.Collections.Specialized.StringCollection = acPlSetVdr.GetPlotDeviceList()
                'For Each Plot In CollPlot
                '    MsgBox(Plot.ToString)
                'Next



                ''Listing des Format
                'Dim CollStyle As System.Collections.Specialized.StringCollection = acPlSetVdr.GetCanonicalMediaNameList(acPlSet)
                'For Each StyleCTB In CollStyle
                '    MsgBox(StyleCTB.ToString)
                'Next

                '' Set the plot info as an override since it will
                '' not be saved back to the layout
                acPlInfo.OverrideSettings = acPlSet



                '' Validate the plot info
                Dim acPlInfoVdr As PlotInfoValidator = New PlotInfoValidator()
                acPlInfoVdr.MediaMatchingPolicy = MatchingPolicy.MatchEnabled
                acPlInfoVdr.Validate(acPlInfo)


                If SaveSettings = True Then
                    ' Update the layout
                    acLayout.UpgradeOpen()
                    acLayout.CopyFrom(acPlSet)
                    ' Save the new objects to the database
                    acTrans.Commit()
                    'Regen the drawing
                    acDoc.Editor.Regen()
                End If


                'Print realy (False = just save)  -----------------------
                If Print = True Then

                    '' Check to see if a plot is already in progress
                    If PlotFactory.ProcessPlotState = Autodesk.AutoCAD.PlottingServices. _
                                                      ProcessPlotState.NotPlotting Then
                        Using acPlEng As PlotEngine = PlotFactory.CreatePublishEngine()

                            '' Track the plot progress with a Progress dialog
                            Dim acPlProgDlg As PlotProgressDialog = New PlotProgressDialog(False, _
                                                                                           1, _
                                                                                           True)

                            Using (acPlProgDlg)
                                '' Define the status messages to display when plotting starts
                                acPlProgDlg.PlotMsgString(PlotMessageIndex.DialogTitle) = _
                                                                           "Plot Progress"
                                acPlProgDlg.PlotMsgString(PlotMessageIndex.CancelJobButtonMessage) = _
                                                                           "Cancel Job"
                                acPlProgDlg.PlotMsgString(PlotMessageIndex.CancelSheetButtonMessage) = _
                                                          "Cancel Sheet"
                                acPlProgDlg.PlotMsgString(PlotMessageIndex.SheetSetProgressCaption) = _
                                                          "Sheet Set Progress"
                                acPlProgDlg.PlotMsgString(PlotMessageIndex.SheetProgressCaption) = _
                                                                           "Sheet Progress"

                                '' Set the plot progress range
                                acPlProgDlg.LowerPlotProgressRange = 0
                                acPlProgDlg.UpperPlotProgressRange = 100
                                acPlProgDlg.PlotProgressPos = 0

                                '' Display the Progress dialog
                                acPlProgDlg.OnBeginPlot()
                                acPlProgDlg.IsVisible = True

                                '' Start to plot the layout
                                acPlEng.BeginPlot(acPlProgDlg, Nothing)

                                '' Define the plot output
                                acPlEng.BeginDocument(acPlInfo, _
                                                      acDoc.Name, _
                                                      Nothing, _
                                                      1, _
                                                      True, _
                                                      "C:\temp\test.pdf")

                                '' Display information about the current plot
                                acPlProgDlg.PlotMsgString(PlotMessageIndex.Status) = _
                                                              "Plotting: " & acDoc.Name & _
                                                              " - " & acLayout.LayoutName

                                '' Set the sheet progress range
                                acPlProgDlg.OnBeginSheet()
                                acPlProgDlg.LowerSheetProgressRange = 0
                                acPlProgDlg.UpperSheetProgressRange = 100
                                acPlProgDlg.SheetProgressPos = 0

                                '' Plot the first sheet/layout
                                Dim acPlPageInfo As PlotPageInfo = New PlotPageInfo()
                                acPlEng.BeginPage(acPlPageInfo, _
                                                  acPlInfo, _
                                                  True, _
                                                  Nothing)

                                acPlEng.BeginGenerateGraphics(Nothing)
                                acPlEng.EndGenerateGraphics(Nothing)

                                '' Finish plotting the sheet/layout
                                acPlEng.EndPage(Nothing)
                                acPlProgDlg.SheetProgressPos = 100
                                acPlProgDlg.OnEndSheet()

                                '' Finish plotting the document
                                acPlEng.EndDocument(Nothing)

                                '' Finish the plot
                                acPlProgDlg.PlotProgressPos = 100
                                acPlProgDlg.OnEndPlot()
                                acPlEng.EndPlot(Nothing)
                            End Using
                        End Using
                    End If

                End If


            Catch ex As Exception
                If ex.Message = "eDeviceNotFound" Then MsgBox("Impossible de trouver l'imprimante, les configurations ne sont pas appliqués.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Configuration d'impression")
            End Try


        End Using

        Return True

    End Function












    'Mise à jour avril 2016

    Public Function ConvertRegionToPolyline2(objCSreg As Region, Optional ByVal booKeepRegion As Boolean = False)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        'Dim peo As New PromptEntityOptions(vbLf & "Select a region:")
        'peo.SetRejectMessage(vbLf & "Must be a region.")
        'peo.AddAllowedClass(GetType(Region), True)

        '  Dim per As PromptEntityResult = ed.GetEntity(peo)

        'If per.Status <> PromptStatus.OK Then
        'Return False
        'End If
        Dim NewPolyID As ObjectId = Nothing
        Dim ids As New ObjectIdCollection()

        Dim tr As Transaction = doc.TransactionManager.StartTransaction()
        Using tr
            Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim btr As BlockTableRecord = DirectCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

            Dim reg As Region = TryCast(tr.GetObject(objCSreg.ObjectId, OpenMode.ForRead), Region)


            If reg IsNot Nothing Then
                Dim objs As DBObjectCollection = PolylineFromRegion(reg)

                ' Append our new entities to the database

                btr.UpgradeOpen()

                For Each ent As Entity In objs
                    ids.Add(btr.AppendEntity(ent))
                    tr.AddNewlyCreatedDBObject(ent, True)
                Next

                ' Finally we erase the original region

                reg.UpgradeOpen()
                reg.[Erase]()
            End If
            tr.Commit()
        End Using

        If ids.Count <> 0 Then NewPolyID = ids(0)

        If NewPolyID <> Nothing Then
            Return NewPolyID
        Else
            Return Nothing
        End If

    End Function

    Private Function PolylineFromRegion(reg As Region) As DBObjectCollection
        ' We will return a collection of entities
        ' (should include closed Polylines and other
        ' closed curves, such as Circles)

        Dim res As New DBObjectCollection()

        ' Explode Region -> collection of Curves / Regions

        Dim cvs As New DBObjectCollection()
        reg.Explode(cvs)

        ' Create a plane to convert 3D coords
        ' into Region coord system

        Dim pl As New Plane(New Point3d(0, 0, 0), reg.Normal)

        Using pl
            Dim finished As Boolean = False

            While Not finished AndAlso cvs.Count > 0
                ' Count the Curves and the non-Curves, and find
                ' the index of the first Curve in the collection

                Dim cvCnt As Integer = 0, nonCvCnt As Integer = 0, fstCvIdx As Integer = -1

                For i As Integer = 0 To cvs.Count - 1
                    Dim tmpCv As Curve = TryCast(cvs(i), Curve)
                    If tmpCv Is Nothing Then
                        nonCvCnt += 1
                    Else
                        ' Closed curves can go straight into the
                        ' results collection, and aren't added
                        ' to the Curve count

                        If tmpCv.Closed Then
                            res.Add(tmpCv)
                            cvs.Remove(tmpCv)
                            ' Decrement, so we don't miss an item
                            i -= 1
                        Else
                            cvCnt += 1
                            If fstCvIdx = -1 Then
                                fstCvIdx = i
                            End If
                        End If
                    End If
                Next

                If fstCvIdx >= 0 Then
                    ' For the initial segment take the first
                    ' Curve in the collection

                    Dim fstCv As Curve = DirectCast(cvs(fstCvIdx), Curve)

                    ' The resulting Polyline

                    Dim p As New Polyline()

                    ' Set common entity properties from the Region

                    p.SetPropertiesFrom(reg)

                    ' Add the first two vertices, but only set the
                    ' bulge on the first (the second will be set
                    ' retroactively from the second segment)

                    ' We also assume the first segment is counter-
                    ' clockwise (the default for arcs), as we're
                    ' not swapping the order of the vertices to
                    ' make them fit the Polyline's order

                    p.AddVertexAt(p.NumberOfVertices, fstCv.StartPoint.Convert2d(pl), BulgeFromCurve(fstCv, False), 0, 0)

                    p.AddVertexAt(p.NumberOfVertices, fstCv.EndPoint.Convert2d(pl), 0, 0, 0)

                    cvs.Remove(fstCv)

                    ' The next point to look for

                    Dim nextPt As Point3d = fstCv.EndPoint

                    ' We no longer need the curve

                    fstCv.Dispose()

                    ' Find the line that is connected to
                    ' the next point

                    ' If for some reason the lines returned were not
                    ' connected, we could loop endlessly.
                    ' So we store the previous curve count and assume
                    ' that if this count has not been decreased by
                    ' looping completely through the segments once,
                    ' then we should not continue to loop.
                    ' Hopefully this will never happen, as the curves
                    ' should form a closed loop, but anyway...

                    ' Set the previous count as artificially high,
                    ' so that we loop once, at least.

                    Dim prevCnt As Integer = cvs.Count + 1
                    While cvs.Count > nonCvCnt AndAlso cvs.Count < prevCnt
                        prevCnt = cvs.Count
                        For Each obj As DBObject In cvs
                            Dim cv As Curve = TryCast(obj, Curve)

                            If cv IsNot Nothing Then
                                ' If one end of the curve connects with the
                                ' point we're looking for...

                                If cv.StartPoint = nextPt OrElse cv.EndPoint = nextPt Then
                                    ' Calculate the bulge for the curve and
                                    ' set it on the previous vertex

                                    Dim bulge As Double = BulgeFromCurve(cv, cv.EndPoint = nextPt)
                                    If bulge <> 0.0 Then
                                        p.SetBulgeAt(p.NumberOfVertices - 1, bulge)
                                    End If

                                    ' Reverse the points, if needed

                                    If cv.StartPoint = nextPt Then
                                        nextPt = cv.EndPoint
                                    Else
                                        ' cv.EndPoint == nextPt
                                        nextPt = cv.StartPoint
                                    End If

                                    ' Add out new vertex (bulge will be set next
                                    ' time through, as needed)

                                    p.AddVertexAt(p.NumberOfVertices, nextPt.Convert2d(pl), 0, 0, 0)

                                    ' Remove our curve from the list, which
                                    ' decrements the count, of course

                                    cvs.Remove(cv)
                                    cv.Dispose()

                                    Exit For
                                End If
                            End If
                        Next
                    End While

                    ' Once we have added all the Polyline's vertices,
                    ' transform it to the original region's plane

                    p.TransformBy(Matrix3d.PlaneToWorld(pl))
                    res.Add(p)

                    If cvs.Count = nonCvCnt Then
                        finished = True
                    End If
                End If

                ' If there are any Regions in the collection,
                ' recurse to explode and add their geometry

                If nonCvCnt > 0 AndAlso cvs.Count > 0 Then
                    For Each obj As DBObject In cvs
                        Dim subReg As Region = TryCast(obj, Region)
                        If subReg IsNot Nothing Then
                            Dim subRes As DBObjectCollection = PolylineFromRegion(subReg)

                            For Each o As DBObject In subRes
                                res.Add(o)
                            Next

                            cvs.Remove(subReg)
                            subReg.Dispose()
                        End If
                    Next
                End If
                If cvs.Count = 0 Then
                    finished = True
                End If
            End While
        End Using
        Return res
    End Function

    ' Helper function to calculate the bulge for arcs

    Private Function BulgeFromCurve(cv As Curve, clockwise As Boolean) As Double
        Dim bulge As Double = 0.0

        Dim a As Arc = TryCast(cv, Arc)
        If a IsNot Nothing Then
            Dim newStart As Double

            ' The start angle is usually greater than the end,
            ' as arcs are all counter-clockwise.
            ' (If it isn't it's because the arc crosses the
            ' 0-degree line, and we can subtract 2PI from the
            ' start angle.)

            If a.StartAngle > a.EndAngle Then
                newStart = a.StartAngle - 8 * Math.Atan(1)
            Else
                newStart = a.StartAngle
            End If

            ' Bulge is defined as the tan of
            ' one fourth of the included angle

            bulge = Math.Tan((a.EndAngle - newStart) / 4)

            ' If the curve is clockwise, we negate the bulge

            If clockwise Then
                bulge = -bulge
            End If
        End If
        Return bulge
    End Function


End Module



