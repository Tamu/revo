Imports System.Math
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry


Module modMaths

    ''' <summary>
    ''' Calcul de distance
    ''' </summary>
    ''' <param name="pt1">Point de base</param>
    ''' <param name="pt2">Point secondaire</param>
    Public Function Dist(ByVal p1 As Point2d, ByVal p2 As Point2d) As Double
        Return Sqrt((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2)
    End Function

    ''' <summary>
    ''' Calcul de distance 3D
    ''' </summary>
    ''' <param name="pt1">Point de base</param>
    ''' <param name="pt2">Point secondaire</param>
    Public Function Dist3D(ByVal p1 As Point3d, ByVal p2 As Point3d) As Double
        Return Sqrt((p1.X - p2.X) ^ 2 + (p1.Y - p2.Y) ^ 2 + (p1.Z - p2.Z) ^ 2)
    End Function


    Public Function ConvGisTopoTrigo(ByVal Angle As Double) As Double
        If Angle <= (PI / 2) Then
            Angle = (PI / 2) - (Angle)
        Else
            Angle = (2 * PI) - ((Angle) - (PI / 2))
        End If
        Return Angle
    End Function

    Public Function PointInsidePolygon(ByVal Verts As Point2dCollection, ByVal p As Point2d, ByRef Loc As Integer) As Boolean
        If Verts.Count < 3 Then
            Loc = 0
            Return False
        End If
        For Each v As Point2d In Verts
            If p.Equals(v) Then
                Loc = 2
                Return False
            End If
        Next
        'Check if poly is closed and close if not
        If Not Verts(0).Equals(Verts(Verts.Count - 1)) Then
            Verts.Add(Verts(0))
        End If

        Dim Angle As Double = 0
        Dim theta As Double = 0
        Dim Alfa As Double = 0
        Dim xNew As Double = 0
        Dim yNew As Double = 0
        For i As Integer = 0 To Verts.Count - 2
            theta = Atan2(Verts(i).Y - p.Y, Verts(i).X - p.X)
            xNew = (Verts(i + 1).X - p.X) * Cos(theta) + (Verts(i + 1).Y - p.Y) * Sin(theta)
            yNew = (Verts(i + 1).Y - p.Y) * Cos(theta) - (Verts(i + 1).X - p.X) * Sin(theta)
            Alfa = Atan2(yNew, xNew)
            If Alfa = PI Then
                Loc = 3 ' edge coincident
                Return False
            End If
            Angle = Angle + Alfa
        Next
        If Abs(Angle) < 0.0001 Then
            Loc = 1 ' point outside
            Return False
        ElseIf Abs(Angle) > PI Then
            Loc = 4
            Return True
        End If
    End Function

    Public Function ValueIsDouble(ByRef strValue As String, Optional ByRef booMsg As Boolean = True) As Boolean
        Dim dblValue As Object

        If strValue <> "" Then

            'On Error GoTo Erreur
            Try
                dblValue = CDbl(strValue)
                'On Error GoTo 0

                Return True
                Exit Function

            Catch
                'Erreur:
                If booMsg = True Then MsgBox("Saisir une valeur numérique !", MsgBoxStyle.Exclamation, "Saisie incorrecte")
                Return False
            End Try
        Else
            Return True
        End If

    End Function


    Public Function ValueIsInteger(ByRef strValue As String, Optional ByRef intValue As Short = 0, Optional ByRef booMsg As Boolean = True) As Boolean

        If strValue <> "" Then

            Try

                'On Error GoTo Erreur
                intValue = CShort(strValue)
                'On Error GoTo 0

                ValueIsInteger = True
                Exit Function

            Catch 'Erreur:

                If booMsg = True Then MsgBox("Saisir une valeur numérique !", MsgBoxStyle.Exclamation, "Saisie incorrecte")
                ValueIsInteger = False
            End Try
        Else
            ValueIsInteger = True
        End If

    End Function

    'Conversion d'un angle trigo (deg) en angle topo (grad) - (sens de rotation différent)
    Public Function ConvAngTrigoTopoG(ByRef Angle As Double) As Object

        'Conv Grad en rad > Modif TH input en Grad (31.10.2014)
        Angle = Angle * PI / 200

        If Angle <= (PI / 2) Then
            Angle = (PI / 2) - (Angle)
        Else
            Angle = (2 * PI) - ((Angle) - (PI / 2))
        End If

        'Con rad en grad
        Return Angle * 200 / PI

    End Function

    'Pts 2D, radians, Distance
    Public Function PolarPoints2D(ByVal pPt As Point2d, _
                                   ByVal dAng As Double, _
                                   ByVal dDist As Double)

        Return New Point2d(pPt.X + dDist * Math.Cos(dAng), _
                           pPt.Y + dDist * Math.Sin(dAng))
    End Function

    Public Function PolarPoints3D(ByVal pPt As Point3d, _
                                       ByVal dAng As Double, _
                                       ByVal dDist As Double)

        Return New Point3d(pPt.X + dDist * Math.Cos(dAng), _
                           pPt.Y + dDist * Math.Sin(dAng), _
                           pPt.Z)
    End Function

   
    'Angle modulo 400
    Public Function Modulo400(ByRef Angle As Double) As Object

        '< 0
        If Angle < 0 Then
            Do
                Angle = Angle + 400
            Loop Until (Angle >= 0)

            '> 0
        ElseIf Angle > 400 Then
            Do
                Angle = Angle - 400
            Loop Until (Angle < 400)
        End If

        Return Angle

    End Function


End Module
