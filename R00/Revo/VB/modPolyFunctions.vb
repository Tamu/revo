Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.DatabaseServices
Imports frms = System.Windows.Forms
Module modPolyFunctions


    ''' <summary>
    ''' Finds the polyline with the smallest area in the supplied collection
    ''' </summary>
    ''' <param name="FoundPolys">An arra of PolylineVertexes objects</param>
    ''' <returns>The index in the array of the smallest polyline</returns>
    ''' <remarks></remarks>
    Public Function FindSmallestPoly(ByVal FoundPolys As PlVertCol(Of PolylineVertexes), ptCentroid As Point3d) As Integer
        Dim dblMinArea As Double = 0
        Dim pLine As Polyline
        Dim SmallestPolyIndx As Integer = -1
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Loop through the found polylines looking for one with the smallest area
            For j As Integer = 0 To FoundPolys.Count - 1
                pLine = DirectCast(trans.GetObject(FoundPolys.Item(j).AcadObject, OpenMode.ForRead), Polyline)

                'Check if pt is inside
                If IsPointInPolyLine(pLine, ptCentroid) Then 'If it is, then add it to the FoundPolys collection

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
                End If

            Next
        End Using
        Return SmallestPolyIndx
    End Function

    ''' <summary>
    ''' Finds the polyline with the biggest area in the supplied collection
    ''' </summary>
    ''' <param name="FoundPolys">An arra of PolylineVertexes objects</param>
    ''' <returns>The index in the array of the smallest polyline</returns>
    ''' <remarks></remarks>
    Public Function FindBiggestPoly(ByVal FoundPolys As PlVertCol(Of PolylineVertexes), ptCentroid As Point3d) As Integer
        Dim dblMaxArea As Double = 0
        Dim pLine As Polyline
        Dim BiggestPolyIndex As Integer = -1
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Loop through the found polylines looking for one with the biggest area
            For j As Integer = 0 To FoundPolys.Count - 1
                pLine = DirectCast(trans.GetObject(FoundPolys.Item(j).AcadObject, OpenMode.ForRead), Polyline)

                'Check if pt is inside
                If IsPointInPolyLine(pLine, ptCentroid) Then 'If it is, then add it to the FoundPolys collection

                    'If we haven't already store the maximumm
                    If dblMaxArea = 0 Then
                        'Set the minimum to the area of the polyline
                        dblMaxArea = pLine.Area
                        'Set the smallest poly index to 0
                        BiggestPolyIndex = 0
                    Else
                        'Test if this polyline is smaller than the previously stored smallest polyline
                        If pLine.Area > dblMaxArea Then
                            'If so, then store this poly area as the smallest one
                            dblMaxArea = pLine.Area
                            'Set the smallest poly index to j
                            BiggestPolyIndex = j
                        End If
                    End If

                End If
            Next
        End Using
        Return BiggestPolyIndex
    End Function

    ''' <summary>
    ''' Gets the vertices from the polyline
    ''' </summary>
    ''' <param name="polyId">The objectid of the polyline</param>
    ''' <returns>A point3dcollection containing the vertices of the polyline</returns>
    ''' <remarks></remarks>
    Public Function GetPolyVertices(ByVal polyId As ObjectId) As Point3dCollection
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Using trans As Transaction = db.TransactionManager.StartTransaction

            Dim Verts As New Point3dCollection

            Try
                If polyId.Handle.Value <> 0 Then

                    Dim pl As Polyline = CType(trans.GetObject(polyId, OpenMode.ForRead), Polyline)
                    For i As Integer = 0 To pl.NumberOfVertices - 1
                        Dim p As Point3d = pl.GetPoint3dAt(i)
                        If Not Verts.Contains(p) Then
                            Verts.Add(p)
                        End If
                    Next
                    pl.Dispose()

                End If

            Catch ex As Exception
                MsgBox(ex.Message)
            End Try

            Return Verts
        End Using
    End Function

    ''' <summary>
    ''' Gets the area of the polyline from the objectid
    ''' </summary>
    ''' <param name="polyID">The objectid of the polyline</param>
    ''' <returns>The area of the polyline</returns>
    ''' <remarks></remarks>
    Public Function GetPolyArea(ByVal polyID As ObjectId) As Double
        Dim retVal As Double = 0
        Try
            Dim db As Database = HostApplicationServices.WorkingDatabase
            Using trans As Transaction = db.TransactionManager.StartTransaction
                Dim pline As Polyline = CType(trans.GetObject(polyID, OpenMode.ForRead), Polyline)
                retVal = pline.Area
                pline.Dispose()
            End Using
        Catch ex As Exception
            frms.MessageBox.Show(ex.Message, "GetPolyArea", Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
        End Try
        Return retVal
    End Function

    ''' <summary>
    ''' Gets the vertices from the polyline after scaling it 
    ''' </summary>
    ''' <param name="polyId">The objectid of the polyline</param>
    ''' <param name="ScaleFactor">The scale factor to apply to the polyline</param>
    ''' <returns>A point3dcollection containing the vertices of the polyline</returns>
    ''' <remarks></remarks>
    Public Function GetScaledPolyVertices(ByVal polyId As ObjectId, ByVal ScaleFactor As Double, Optional ByVal ZoomObject As Boolean = False) As Point3dCollection
        Dim db As Database = HostApplicationServices.WorkingDatabase
        Dim MinPt As Point3d = Nothing
        Dim MaxPt As Point3d = Nothing
        Dim Verts As New Point3dCollection
        Using trans As Transaction = db.TransactionManager.StartTransaction
            'Get the polyline from the database
            Dim pl As Polyline = CType(trans.GetObject(polyId, OpenMode.ForWrite), Polyline)
            'Set up the scaling matrix getting the centre point from the GetGravityCentre function
            Dim ScaleMat As Matrix3d = Matrix3d.Scaling(ScaleFactor, GetGravityCentre(polyId))
            'Scale the polyline
            pl.TransformBy(ScaleMat)
            MinPt = pl.GeometricExtents.MinPoint
            MaxPt = pl.GeometricExtents.MaxPoint
            'Now get the vertices of the scaled polyline
            For i As Integer = 0 To pl.NumberOfVertices - 1
                Dim p As Point3d = pl.GetPoint3dAt(i)
                If Not Verts.Contains(p) Then
                    Verts.Add(p)
                End If
            Next
            pl.Dispose()
        End Using
        If ZoomObject Then
            Zooming.ZoomWindow(MinPt, MaxPt)
        End If
        Return Verts
    End Function
End Module
