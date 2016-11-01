Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry


Module modGeometry
    ''' <summary>
    ''' Converts the supplied entity to a polyline
    ''' </summary>
    ''' <param name="ent">The entity to convert</param>
    ''' <returns>The entity converted to a polyline</returns>
    ''' <remarks></remarks>
    Public Function ConvertEntityToLWPolyLine(ByVal ent As Entity) As Polyline
        Dim pLine As Polyline = Nothing
        If TypeOf (ent) Is Line Then
            pLine = ConvertLineToPolyline(DirectCast(ent, Line))
        ElseIf TypeOf (ent) Is Polyline Then
            Return DirectCast(ent, Polyline)
        ElseIf TypeOf (ent) Is Arc Then
            pLine = ConvertArcToPolyline(DirectCast(ent, Arc))
        Else
            Return Nothing
        End If
        Return pLine
    End Function

    ''' <summary>
    ''' Converts the supplied line to a polyline
    ''' </summary>
    ''' <param name="L">The line to convert</param>
    ''' <returns>The line converted to a polyline</returns>
    ''' <remarks></remarks>
    Private Function ConvertLineToPolyline(ByVal L As Line) As Polyline
        Dim pLine As Polyline = Nothing
        'Get the database for this drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Create a point2dcollection
            Dim Verts As New Point2dCollection
            'Add the start point to the Verts collection
            Verts.Add(New Point2d(L.StartPoint.X, L.StartPoint.Y))
            'Add the end point to the Verts collection
            Verts.Add(New Point2d(L.EndPoint.X, L.EndPoint.Y))
            'Call the function to create the polyline from the vertices
            Dim curLayer As String = GetCurrentLayer()
            pLine = InsertLWPolyline(Verts, curLayer)
            'Get the existing line from the database
            Dim lin As Line = DirectCast(trans.GetObject(L.ObjectId, OpenMode.ForWrite), Line)
            'Delete the original line
            lin.Erase()
            'Commit the transaction to save the changes
            trans.Commit()
        End Using
        'Return the polyline
        Return pLine
    End Function

    ''' <summary>
    ''' Converts the supplied Arc to a polyline
    ''' </summary>
    ''' <param name="A">The arc to convert</param>
    ''' <returns>The arc converted to a polyline</returns>
    ''' <remarks></remarks>
    Private Function ConvertArcToPolyline(ByVal A As Arc) As Polyline
        'Create a point2dcollection to hold the vertices of the polyline
        Dim Verts As New Point2dCollection
        'Add the arc start point to the Verts collection
        Verts.Add(New Point2d(A.StartPoint.X, A.StartPoint.Y))
        'Add the arc end point to the verts collection
        Verts.Add(New Point2d(A.EndPoint.X, A.EndPoint.Y))
        'Determine if the arc is clockwise
        Dim booClockWise As Boolean = (A.EndAngle - A.StartAngle >= 0)
        Dim curLayer As String = GetCurrentLayer()
        'Create the polyline using the start and end points of the arc
        Dim pLine As Polyline = InsertLWPolyline(Verts, curLayer)
        'Get the database for the drawing
        Dim db As Database = HostApplicationServices.WorkingDatabase
        'Start a transaction with the database
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'If the radius of the arc is greater than 0
            If A.Radius > 0 Then
                Dim dblBulge As Double = 0
                'Calculate the bulge factor
                Dim newStart As Double = 0
                'If start angle is greater than end angle then the arc crosses the 0 degree line
                'so subtract 2PI from the start angle
                If A.StartAngle > A.EndAngle Then
                    newStart = A.StartAngle - 8 * Math.Atan(1)
                Else
                    newStart = A.StartAngle
                End If
                'Bulge is the tan of one quarter of the included angle
                dblBulge = Math.Tan((A.EndAngle - newStart) / 4)
                If booClockWise Then dblBulge = -dblBulge
                SetPolyBulge(pLine.ObjectId, 0, dblBulge)
            End If
            'Delete the original arc
            Dim acArc As Arc = DirectCast(trans.GetObject(A.ObjectId, OpenMode.ForWrite), Arc)
            acArc.Erase()
            'Save the changes
            trans.Commit()
            'Return the polyline
            Return pLine
        End Using
    End Function

    ''' <summary>
    ''' Joins two polylines together
    ''' </summary>
    ''' <param name="fstPol">The first polyline</param>
    ''' <param name="NxtPol">The second polyline</param>
    ''' <param name="FuzVal">A join tolerance</param>
    ''' <returns>true if the lines can be and have been joined, otherwise false</returns>
    ''' <remarks></remarks>
    Public Function Join2Polylines(ByVal fstPol As Polyline, ByVal NxtPol As Polyline, ByVal FuzVal As Double) As Boolean
        Dim FstArr As Point2dCollection = Nothing
        Dim NxtArr As Point2dCollection = Nothing
        Dim tmpPt As Point2d = Nothing
        Dim fstLen As Integer
        Dim nxtLen As Integer
        Dim vtxCnt As Integer
        Dim fstCnt As Integer
        Dim nxtCnt As Integer
        Dim revFlg As Boolean
        Dim retVal As Boolean
        'Get the coordinates of the first line
        FstArr = GetPolyCoordinates(fstPol)
        'Get the coordinates of the second line
        NxtArr = GetPolyCoordinates(NxtPol)
        'Store the number of vertices in each
        fstLen = fstPol.NumberOfVertices
        nxtLen = NxtPol.NumberOfVertices
        'Call the function to test if the first point of the first line
        ' is the same as the last point of the second line
        If MePointsEqual(FstArr(1), NxtArr(nxtLen), FuzVal) Then
            'Reverse both lines
            fstPol = MeReversePline(fstPol)
            FstArr = GetPolyCoordinates(fstPol)
            NxtPol = MeReversePline(NxtPol)
            NxtArr = GetPolyCoordinates(NxtPol)
            revFlg = True
            retVal = True
        ElseIf MePointsEqual(FstArr(1), NxtArr(1), FuzVal) Then
            'First point of First line = First  point of second line
            'So reverse the first line
            fstPol = MeReversePline(fstPol)
            FstArr = GetPolyCoordinates(fstPol)
            revFlg = True
            retVal = True
        ElseIf MePointsEqual(FstArr(fstLen), NxtArr(nxtLen), FuzVal) Then
            'Last point of first line = last point of second line
            'So reverse the second line
            NxtPol = MeReversePline(NxtPol)
            NxtArr = GetPolyCoordinates(NxtPol)
            revFlg = True
            retVal = True
        ElseIf MePointsEqual(FstArr(fstLen), NxtArr(1), FuzVal) Then
            'Last point of first line = first point of second line
            'So no reversing required
            revFlg = False
            retVal = True
        Else
            retVal = False
        End If
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Now add the vertices of the second line to the first
            Dim finalLine As Polyline = DirectCast(trans.GetObject(fstPol.ObjectId, OpenMode.ForWrite), Polyline)
            If retVal Then
                fstCnt = fstPol.NumberOfVertices
                nxtCnt = 0
                finalLine.SetBulgeAt(fstCnt, NxtPol.GetBulgeAt(nxtCnt))
                For vtxCnt = 1 To nxtLen - 1
                    fstCnt += 1
                    nxtCnt += 1
                    tmpPt = New Point2d(NxtArr(vtxCnt).X, NxtArr(vtxCnt).Y)
                    finalLine.AddVertexAt(fstCnt, tmpPt, NxtPol.GetBulgeAt(nxtCnt), 0, 0)
                Next
            End If
            'Delete the second line
            Dim delLine As Polyline = DirectCast(trans.GetObject(NxtPol.ObjectId, OpenMode.ForWrite), Polyline)
            delLine.Erase()
            trans.Commit()
        End Using
    End Function

    ''' <summary>
    ''' Gets the vertices of the supplied polyline
    ''' </summary>
    ''' <param name="pLine">The polyline to get the vertices of</param>
    ''' <returns>a point2dcollection containing the vertices of the polyline</returns>
    ''' <remarks></remarks>
    Public Function GetPolyCoordinates(ByVal pLine As Polyline) As Point2dCollection
        Dim pCol As New Point2dCollection
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the number of vertices in the polyline
            Dim NumVerts As Integer = pLine.NumberOfVertices
            'Loop through the vertices adding each one to the return collection
            For i As Integer = 0 To NumVerts - 1
                pCol.Add(pLine.GetPoint2dAt(i))
            Next
        End Using
        Return pCol
    End Function

    ''' <summary>
    ''' Tests if two points are the same (within a tolerance)
    ''' </summary>
    '''<param name="fstPnt">The first point to test</param>
    ''' <param name="nxtPnt">The second point to test</param>
    ''' <param name="FuzVal">The tolerance value</param>
    ''' <returns>True if the points are the same (within the tolerance)</returns>
    ''' <remarks></remarks>
    Private Function MePointsEqual(ByVal fstPnt As Point2d, ByVal nxtPnt As Point2d, ByVal FuzVal As Double) As Boolean
        'Get the X distance bewteen the two points
        Dim XCoDst As Double = fstPnt.X - nxtPnt.X
        'Get the Y distance between the two points
        Dim YcoDst As Double = fstPnt.Y - nxtPnt.Y
        'Get the distance  between the two points and test if this value is less than the tolerance value
        Return (Math.Sqrt(XCoDst ^ 2 + YcoDst ^ 2) < FuzVal)
    End Function

    ''' <summary>
    ''' Reverses the coordinates of the polyline
    ''' </summary>
    ''' <param name="pLine">The polyline to reverse</param>
    ''' <returns>The reversed polyline</returns>
    ''' <remarks></remarks>
    Private Function MeReversePline(ByVal pLine As Polyline) As Polyline
        'Collection to hold the points of the polyline before being reversed
        Dim oldArr As Point2dCollection
        'Collection to hold the points of the polyline after being reversed
        Dim NewArr As New Point2dCollection
        Dim SegCnt As Integer
        'Store the original coordinates
        oldArr = GetPolyCoordinates(pLine)
        'Get the number of line segments
        SegCnt = pLine.NumberOfVertices - 1
        'Array to hold the bulge values
        Dim BlgArr(SegCnt + 1) As Double
        'Store the bulge values and vertices in reverse order
        For arrCnt As Integer = SegCnt To 0 Step -1
            BlgArr(arrCnt) = pLine.GetBulgeAt(SegCnt - arrCnt) * -1
            NewArr.Add(oldArr(arrCnt))
        Next
        Dim curLayer As String = GetCurrentLayer()
        'Add a new polyline to the drawing using the reversed vertices
        Dim newPline As Polyline = InsertLWPolyline(NewArr, curLayer)
        If newPline Is Nothing Then Return pLine
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the new polyline from the database
            Dim tmpLine As Polyline = DirectCast(trans.GetObject(newPline.ObjectId, OpenMode.ForWrite), Polyline)
            'Set its layer to that of the original 
            tmpLine.LayerId = pLine.LayerId
            'Set the bulge values
            For arrCnt As Integer = 0 To SegCnt
                tmpLine.SetBulgeAt(arrCnt, BlgArr(arrCnt + 1))
            Next
            'Get the original polyline
            Dim oldLine As Polyline = DirectCast(trans.GetObject(pLine.ObjectId, OpenMode.ForWrite), Polyline)
            'Then delete it
            oldLine.Erase()
            'Save the changes
            trans.Commit()
        End Using
        'Return the newly reversed polyline
        Return newPline
    End Function


    

End Module
