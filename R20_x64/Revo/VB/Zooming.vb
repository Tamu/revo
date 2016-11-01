Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.ApplicationServices
Imports frms = System.Windows.Forms
Public Class Zooming

    Public Shared Sub Zoom(ByVal pMin As Point3d, ByVal pMax As Point3d, _
            ByVal pCenter As Point3d, ByVal dFactor As Double)
        'Get the current document and database
        Dim Doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim CurDb As Database = Doc.Database
        Dim nCurVport As Integer = System.Convert.ToInt32(Application.GetSystemVariable("CVPORT"))
        'Get the extents of the current space when no points 
        'or only a center point is provided
        'Check to see if Model space is current
        If CurDb.TileMode = True Then
            If pMin.Equals(New Point3d()) = True And _
                pMax.Equals(New Point3d()) = True Then
                pMin = CurDb.Extmin
                pMax = CurDb.Extmax
            End If
        Else
            'Check to see if Paper space is current
            If nCurVport = 1 Then
                If pMin.Equals(New Point3d()) = True And _
                   pMax.Equals(New Point3d()) = True Then
                    pMin = CurDb.Pextmin
                    pMax = CurDb.Pextmax
                End If
            Else
                'Get the extents of Model space
                If pMin.Equals(New Point3d()) = True And _
                   pMax.Equals(New Point3d()) = True Then
                    pMin = CurDb.Extmin
                    pMax = CurDb.Extmax
                End If
            End If
        End If
        'Start a transaction
        Using Trans As Transaction = CurDb.TransactionManager.StartTransaction()
            'Get the current view
            Using acView As ViewTableRecord = Doc.Editor.GetCurrentView()
                PrevExtents = CType(acView.Clone, ViewTableRecord)
                Dim eExtents As Extents3d
                'Translate WCS coordinates to DCS
                Dim matWCS2DCS As Matrix3d
                matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection)
                matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS
                matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist, _
                                               acView.ViewDirection, _
                                               acView.Target) * matWCS2DCS
                'If a center point is specified, define the 
                'min and max point of the extents
                'for Center and Scale modes
                If pCenter.DistanceTo(Point3d.Origin) <> 0 Then
                    pMin = New Point3d(pCenter.X - (acView.Width / 2), _
                                       pCenter.Y - (acView.Height / 2), 0)

                    pMax = New Point3d((acView.Width / 2) + pCenter.X, _
                                       (acView.Height / 2) + pCenter.Y, 0)
                End If
                'Create an extents object using a line
                Using acLine As Line = New Line(pMin, pMax)
                    'eExtents = New Extents3d(acLine.Bounds.Value.MinPoint, _
                    '                         acLine.Bounds.Value.MaxPoint)
                    eExtents = acLine.GeometricExtents
                End Using
                'Calculate the ratio between the width and height of the current view
                Dim dViewRatio As Double
                dViewRatio = (acView.Width / acView.Height)
                'Tranform the extents of the view
                matWCS2DCS = matWCS2DCS.Inverse()
                eExtents.TransformBy(matWCS2DCS)
                Dim dWidth As Double
                Dim dHeight As Double
                Dim pNewCentPt As Point2d
                'Check to see if a center point was provided (Center and Scale modes)
                If pCenter.DistanceTo(Point3d.Origin) <> 0 Then
                    dWidth = acView.Width
                    dHeight = acView.Height
                    If dFactor = 0 Then
                        pCenter = pCenter.TransformBy(matWCS2DCS)
                    End If
                    pNewCentPt = New Point2d(pCenter.X, pCenter.Y)
                Else 'Working in Window, Extents and Limits mode
                    'Calculate the new width and height of the current view
                    dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X
                    dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y
                    'Get the center of the view
                    pNewCentPt = New Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5), _
                                             ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5))
                End If
                'Check to see if the new width fits in current window
                If dWidth > (dHeight * dViewRatio) Then dHeight = dWidth / dViewRatio
                'Resize and scale the view
                If dFactor <> 0 Then
                    acView.Height = dHeight * dFactor
                    acView.Width = dWidth * dFactor
                End If
                'Set the center of the view
                acView.CenterPoint = pNewCentPt
                'Set the current view
                Doc.Editor.SetCurrentView(acView)
            End Using
            'Commit the changes
            Trans.Commit()
        End Using
    End Sub

    Public Shared Sub ZoomWindow(ByVal pMin As Point3d, ByVal pMax As Point3d)
        Zoom(pMin, pMax, New Point3d(), 1)
    End Sub


    ''' <summary>
    ''' Zooms the drawing to supplied object
    ''' </summary>
    ''' <param name="objid">The objectId of the object to zoom to</param>
    ''' <remarks></remarks>
    Public Shared Sub ZoomToObject(ByVal objid As ObjectId, Optional ByVal Padding As Int16 = 100)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        If Not objid.IsNull Then
            Try
                Using trans As Transaction = db.TransactionManager.StartTransaction()
                    Dim ent As Entity = CType(trans.GetObject(objid, OpenMode.ForRead), Entity)
                    Dim ext3d As Extents3d = ent.GeometricExtents
                    Dim minPt As New Point3d(ext3d.MinPoint.X - Padding, ext3d.MinPoint.Y - Padding, 0)
                    Dim maxPt As New Point3d(ext3d.MaxPoint.X + Padding, ext3d.MaxPoint.Y + Padding, 0)
                    Zoom(minPt, maxPt, New Point3d(), 1)
                    trans.Commit()
                End Using
            Catch ex As Exception
                frms.MessageBox.Show(ex.Message, "ZoomToObject", _
                                      Windows.Forms.MessageBoxButtons.OK, Windows.Forms.MessageBoxIcon.Error)
            End Try
        End If
    End Sub
    ''' <summary>
    ''' Zooms to the extents of the current space
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub ZoomExtents()
        Zoom(New Point3d(), New Point3d(), New Point3d(), 1.01075)
    End Sub

    Public Shared Sub ZoomPrevious()
        Application.DocumentManager.MdiActiveDocument.Editor.SetCurrentView(PrevExtents)
    End Sub
End Class


