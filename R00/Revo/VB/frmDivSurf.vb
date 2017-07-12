Imports frms = System.Windows.Forms
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Drawing
Imports System.Math

Public Class frmDivSurf

    Public N_parcelle As String = 1
    Public Somme_surf As Double = 0
    Public Poly As Polyline = Nothing
    Public meth As String = ""
    Public meth_surf As String
    Public orient As Boolean
    Public ptPic1_2d As Point2d
    Public ptPic2_2d As Point2d
    Public T_surf() As Double
    Public surf_affic As Double
    Public clic_calcul As Boolean = False

    Private Sub rbtn_methode_direction_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtn_methode_direction.CheckedChanged
        'affichage adapté au choix de la méthode et affichage du panel suivant
        If rbtn_methode_direction.Checked Then
            pan_orientation_dir.Visible = True
            pan_orientation_pt.Visible = False
            rbnt_dir_par.Checked = False
            rbnt_dir_perp.Checked = False
            btn_calcul.Enabled = False
        End If
    End Sub

    Private Sub rbtn_methode_pt_fixe_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtn_methode_pt_fixe.CheckedChanged
        'affichage adapté au choix de la méthode et affichage du panel suivant
        If rbtn_methode_pt_fixe.Checked Then
            pan_orientation_pt.Visible = True
            pan_orientation_dir.Visible = False
            rbnt_dir_perp.Checked = False
            rbnt_dir_par.Checked = False
            bnt_orientation_dir.Enabled = False
            btn_calcul.Enabled = False
            Refresh()
        End If
    End Sub

    'enregistrement du nombre de parcelle à créer
    Private Sub tbox_N_parcelle_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles tbox_N_parcelle.TextChanged
        'controle si le champ n'est pas vide
        If tbox_N_parcelle.Text <> "" Then
            'controle si le champ contient une valeur numérique
            If IsNumeric(tbox_N_parcelle.Text) Then
                N_parcelle = tbox_N_parcelle.Text

                'controle si N_parcelle est entier
                Dim l As Integer = N_parcelle.Length
                Dim i As Integer
                Dim est_entier As Boolean = True
                Dim Recup As String

                If l > 0 Then
                    For i = 1 To l
                        Recup = Mid(N_parcelle, i, 1)
                        If Recup = "," Or Recup = "." Then
                            est_entier = False
                        End If
                    Next i
                End If

                'controle si N_parcelle est entier et supérieur à 1
                If N_parcelle < 2 Or est_entier = False Then
                    tbox_N_parcelle.BackColor = Drawing.Color.Red
                    rbtn_calcul_surface_choisie.ForeColor = Color.Gray
                    rbtn_calcul_surface_egale.ForeColor = Color.Gray
                    rbtn_calcul_surface_choisie.Enabled = False
                    rbtn_calcul_surface_egale.Enabled = False
                    lbl_calcul_surf_1.ForeColor = Color.Gray
                    DataGridView_suface.Visible = False

                Else
                    'si tout est ok : passage au point suivant
                    DataGridView_suface.RowCount = N_parcelle - 1
                    rbtn_calcul_surface_egale.ForeColor = Color.White
                    rbtn_calcul_surface_choisie.ForeColor = Color.White
                    lbl_calcul_surf_1.ForeColor = Color.White
                    rbtn_calcul_surface_choisie.Enabled = True
                    rbtn_calcul_surface_egale.Enabled = True
                    tbox_N_parcelle.BackColor = Drawing.Color.White

                    If rbtn_calcul_surface_choisie.Checked Then
                        DataGridView_suface.Visible = True
                    End If
                End If
            Else
                tbox_N_parcelle.BackColor = Drawing.Color.Red
                rbtn_calcul_surface_choisie.ForeColor = Color.Gray
                rbtn_calcul_surface_egale.ForeColor = Color.Gray
                rbtn_calcul_surface_choisie.Enabled = False
                rbtn_calcul_surface_egale.Enabled = False
                lbl_calcul_surf_1.ForeColor = Color.Gray
                DataGridView_suface.Visible = False
            End If
        Else
            rbtn_calcul_surface_choisie.ForeColor = Color.Gray
            rbtn_calcul_surface_egale.ForeColor = Color.Gray
            DataGridView_suface.Visible = False
        End If
        If N_parcelle > 2 Then
            If DataGridView_suface.Rows(N_parcelle - 2).Cells(0).Value = "" Then
                btn_calcul.Enabled = False
            End If
        End If


    End Sub

    'sélection de la parcelle à diviser
    Private Sub btn_selection_parcelle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_selection_parcelle.Click
        ''Hide the frame
        Me.Hide()
        Dim collection_retour As New Collection
        ' Dim MyCmd As modDivSurf
        collection_retour = modDivSurf.selection_polyligne() ' MyCmd.selection_polyligne()

        If collection_retour.Count <> 0 Then
            lbl_methode.ForeColor = Color.White
            rbtn_methode_direction.ForeColor = Color.White
            rbtn_methode_pt_fixe.ForeColor = Color.White
            rbtn_methode_direction.Enabled = True
            rbtn_methode_pt_fixe.Enabled = True
            Poly = collection_retour.Item(1)
            Somme_surf = Poly.Area
            Me.Show()
        Else
            Me.Show()
        End If

    End Sub
    'sélection de l'orientation de la division - cas "suivant une direction"

    Private Sub bnt_orientation_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bnt_orientation_dir.Click
        meth = ""
        'Choisir l'orientation des droites via une sélection de 2 points
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim idAcLine2 As ObjectId

        Me.Hide()

        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            Dim pPtRes As PromptPointResult
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            '' Prompt for the start point
            pPtOpts.Message = vbLf & "Saisir le 1er point pour définir l'orientation des lignes de division :"
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)

            Dim ptPic1 As Point3d = pPtRes.Value


            '' Prompt for the start point
            pPtOpts.Message = vbLf & "Saisir le 2ème point :"
            pPtOpts.UseDashedLine = True
            pPtOpts.UseBasePoint = True
            pPtOpts.BasePoint = ptPic1
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            Dim ptPic2 As Point3d = pPtRes.Value

            ''Mettre le SCU GENERAL
            acDoc.Editor.CurrentUserCoordinateSystem = New Matrix3d(New Double(15) _
                              {1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, 0.0, 0.0, 0.0, 1.0}) 'SCU général

            Dim TransfMat As Matrix2d
            ptPic1_2d = New Point2d(ptPic1.X, ptPic1.Y)
            ptPic2_2d = New Point2d(ptPic2.X, ptPic2.Y)

            ''rotation du vecteur direction si méthode perpendiculaire0
            If rbnt_dir_perp.Checked Then
                Dim Pi As Double = Math.PI ' 3.1415926535897931
                TransfMat = Matrix2d.Rotation(Pi / 2, ptPic1_2d)
                ptPic2_2d = ptPic2_2d.TransformBy(TransfMat)
            End If

            ' Recherche des points Ymax et Ymin

            Dim VectDroiteProv As Vector2d = ptPic1_2d.GetVectorTo(New Point2d(ptPic2_2d.X, ptPic2_2d.Y))
            TransfMat = Matrix2d.Rotation(VectDroiteProv.Angle * (-1), ptPic1_2d)
            Dim iDmin As Integer
            Dim iDmax As Integer
            Dim CoordTest As Double = 0
            Dim CoordMin As Double
            Dim CoordMax As Double

            If Poly <> Nothing Then
                CoordMin = Poly.GetPoint2dAt(0).TransformBy(TransfMat).Y()
                CoordMax = Poly.GetPoint2dAt(0).TransformBy(TransfMat).Y()
                'Boucle dans les sommets des poylignes
                For nCnt As Integer = 0 To Poly.NumberOfVertices - 1
                    CoordTest = Poly.GetPoint2dAt(nCnt).TransformBy(TransfMat).Y()

                    'If Poly.GetSegmentType(nCnt) = 1 Then
                    ' MsgBox(Poly.GetArcSegment2dAt(nCnt).Radius.ToString)
                    ' MsgBox(Poly.GetArcSegment2dAt(nCnt).Center.Y.ToString)
                    ' End If

                    If CoordTest < CoordMin Then 'Détection du sommet le plus bas
                        CoordMin = CoordTest 'Coordonnée relative !!!
                        iDmin = nCnt
                    End If
                    If CoordTest > CoordMax Then 'Détection du sommet le plus haut
                        CoordMax = CoordTest
                        iDmax = nCnt
                    End If
                Next


                'Creation du point "moyen"
                Dim pt_moy As Point3d = New Point3d((Poly.GetPoint2dAt(iDmax).X + Poly.GetPoint2dAt(iDmin).X) * 0.5, _
                                                    (Poly.GetPoint2dAt(iDmax).Y + Poly.GetPoint2dAt(iDmin).Y) * 0.5, 0)
                Dim pt_moy_2d As Point2d = New Point2d(pt_moy.X, pt_moy.Y)
                Dim ptPic2_m As Point3d = New Point3d(ptPic2_2d.X + (pt_moy.X - ptPic1.X), ptPic2_2d.Y + (pt_moy.Y - ptPic1.Y), 0)


                ''Creation and drawing of the line
                Dim acline2 As Xline = New Xline()
                acline2.BasePoint = pt_moy
                acline2.SecondPoint = ptPic2_m
                Dim acBlkTbl As BlockTable
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim acBlkTblRec As BlockTableRecord
                acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                acline2.SetDatabaseDefaults()
                acBlkTblRec.AppendEntity(acline2)
                acTrans.AddNewlyCreatedDBObject(acline2, True)
                acTrans.Commit()
                idAcLine2 = acline2.ObjectId

                '' Regenerate the drawing
                Application.DocumentManager.MdiActiveDocument.Editor.Regen()


                'Orientation 
                '' Prompt for the start point
                pPtOpts.Message = vbLf & "Cliquez du côté de la ligne où la 1ère parcelle sera créée"
                pPtOpts.UseDashedLine = False
                pPtOpts.BasePoint = Nothing
                pPtRes = acDoc.Editor.GetPoint(pPtOpts)
                Dim pt_orient As Point3d = pPtRes.Value
                Dim pt_orient_2D As Point2d = New Point2d(pt_orient.X, pt_orient.Y)

                Dim Vect_orientation As Vector2d = pt_moy_2d.GetVectorTo(New Point2d(ptPic2_m.X, ptPic2_m.Y))
                TransfMat = Matrix2d.Rotation(Vect_orientation.Angle * (-1), pt_moy_2d)

                CoordTest = pt_orient_2D.TransformBy(TransfMat).Y()
                CoordMax = pt_moy_2d.TransformBy(TransfMat).Y()


                If CoordTest > CoordMax Then
                    orient = True
                    tbox_N_parcelle.Enabled = True
                    lbl_N_parcelle_1.ForeColor = Color.White
                    meth = "par"
                ElseIf CoordTest < CoordMax Then
                    orient = False
                    tbox_N_parcelle.Enabled = True
                    lbl_N_parcelle_1.ForeColor = Color.White
                    meth = "par"
                Else
                    MsgBox("Merci de cliquer de manière distincte d'un coté ou de l'autre de la ligne", vbInformation + vbOKOnly, "Erreur de saisie")
                    tbox_N_parcelle.Enabled = False
                    lbl_N_parcelle_1.ForeColor = Color.Gray
                End If


                'Dim MyCmd As New Autodesk.AutoCAD.REVO_Surf.MyCommands
            End If

        End Using

        EraseAc(idAcLine2)

        Me.Show()

    End Sub

    'sélection de l'orientation de la division - cas "point fixe"
    Private Sub bnt_orientation_pt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bnt_orientation_pt.Click

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database


        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

            'Choisir l'orientation des droites via une sélection de 2 points
            Dim pPtRes As PromptPointResult
            Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
            '' Prompt for the start point
            pPtOpts.Message = vbLf & "Saisir le point de base : (sommet du polygone)"
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)

            Dim ptPic1 As Point3d = pPtRes.Value
            ptPic1_2d = New Point2d(ptPic1.X, ptPic1.Y)


            '' Prompt for the start point
            pPtOpts.Message = vbLf & "Saisir le 2ème point qui définira l'orientation de la division : (sommet contigu du polygone)"
            pPtOpts.UseDashedLine = True
            pPtOpts.UseBasePoint = True
            pPtOpts.BasePoint = ptPic1
            pPtRes = acDoc.Editor.GetPoint(pPtOpts)
            Dim ptPic2 As Point3d = pPtRes.Value
            ptPic2_2d = New Point2d(ptPic2.X, ptPic2.Y)

            'CONTROLE QUE CE SONT DEUX POINTS CONTIGUS
            meth = ""
            For i As Integer = 0 To Poly.NumberOfVertices - 1
                If Poly.GetPoint2dAt(i).X = ptPic1.X And Poly.GetPoint2dAt(i).Y = ptPic1.Y Then
                    If i > 0 And i < Poly.NumberOfVertices - 1 Then
                        If Poly.GetPoint2dAt(i - 1).X = ptPic2.X And Poly.GetPoint2dAt(i - 1).Y = ptPic2.Y Then
                            meth = "pt"
                            orient = False
                        End If
                        If (Poly.GetPoint2dAt(i + 1).X = ptPic2.X And Poly.GetPoint2dAt(i + 1).Y = ptPic2.Y) Then
                            meth = "pt"
                            orient = True
                        End If
                    ElseIf i = 0 Then
                        If (Poly.GetPoint2dAt(1).X = ptPic2.X And Poly.GetPoint2dAt(1).Y = ptPic2.Y) Then
                            meth = "pt"
                            orient = True
                        End If
                        If (Poly.GetPoint2dAt(Poly.NumberOfVertices - 1).X = ptPic2.X And Poly.GetPoint2dAt(Poly.NumberOfVertices - 1).Y = ptPic2.Y) Then
                            meth = "pt"
                            orient = False
                        End If
                    ElseIf i = Poly.NumberOfVertices - 1 Then
                        If (Poly.GetPoint2dAt(0).X = ptPic2.X And Poly.GetPoint2dAt(0).Y = ptPic2.Y) Then
                            meth = "pt"
                            orient = True
                        End If
                        If (Poly.GetPoint2dAt(Poly.NumberOfVertices - 2).X = ptPic2.X And Poly.GetPoint2dAt(Poly.NumberOfVertices - 2).Y = ptPic2.Y) Then
                            meth = "pt"
                            orient = False
                        End If
                    End If
                End If
            Next

            'For i As Integer = 0 To Poly.EndPoint.X
            '    If Point(i).x = ptPic1.X And Point(i).Y = ptPic1.Y Then
            'Next

        End Using

        If meth = "pt" Then
            tbox_N_parcelle.Enabled = True
            lbl_N_parcelle_1.ForeColor = Color.White

        Else
            tbox_N_parcelle.Enabled = False
            lbl_N_parcelle_1.ForeColor = Color.Gray
        End If


    End Sub

    'affichage des élements si l'option "surfaces égales" est cochée
    Private Sub rbtn_calcul_surface_egale_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtn_calcul_surface_egale.CheckedChanged
        Dim mtd_surf As String
        If rbtn_calcul_surface_egale.Checked Then
            DataGridView_suface.Visible = False
            lbl_surf_reste.Visible = False
            mtd_surf = "egale"
            btn_calcul.Enabled = True

        End If
    End Sub

    'affichage des élements si l'option "surfaces choisies" est cochée
    Private Sub rbtn_calcul_surface_choisie_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbtn_calcul_surface_choisie.CheckedChanged
        Dim mtd_surf As String
        If rbtn_calcul_surface_choisie.Checked Then
            btn_calcul.Enabled = False
            mtd_surf = "choisie"
            DataGridView_suface.Visible = True
            lbl_surf_reste.Visible = True
        End If
    End Sub

    Private Sub CBox_enregistrement_CheckedChanged_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CBox_enregistrement.CheckedChanged
        If CBox_enregistrement.Checked Then
            btn_enregistrement.Enabled = True
            tbox_enregistrement.Enabled = True
        Else
            btn_enregistrement.Enabled = False
            tbox_enregistrement.Enabled = False
        End If
    End Sub

    Private Sub btn_calcul_Click(sender As System.Object, e As System.EventArgs) Handles btn_calcul.Click
        ' Dim myCdm As New Autodesk.AutoCAD.REVO_Surf.MyCommands
        clic_calcul = True
        Me.Hide()
    End Sub
   
    'remplissage du tableau de valeur de surface
    Private Sub DataGridView_suface_CellValueChanged(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView_suface.CellValueChanged

        Dim i As Integer 'variable de boucle
        Dim n As Integer = 0 'compteur
        surf_affic = Somme_surf


        If CInt(N_parcelle) > 1 Then

            For i = 0 To CInt(N_parcelle) - 2
                If DataGridView_suface.Rows.Count <> 0 Then
                    If DataGridView_suface.Rows(i).Cells(0).Value <> "" Then

                        If IsNumeric(DataGridView_suface.Rows(i).Cells(0).Value) And DataGridView_suface.Rows(i).Cells(0).Value > 0 Then
                            n = n + 1
                            surf_affic = surf_affic - DataGridView_suface.Rows(i).Cells(0).Value
                        End If

                    End If
                End If
                lbl_surf_reste.Text = "Surface de la dernière parcelle : " & Round(CDbl(surf_affic), 3)
            Next

            If surf_affic <= 0 Then
                lbl_surf_reste.ForeColor = Color.Red
                btn_calcul.Enabled = False
            Else
                lbl_surf_reste.ForeColor = Color.White
            End If

            If surf_affic > 0 And n = CInt(N_parcelle) - 1 Then
                lbl_surf_reste.ForeColor = Color.White
                btn_calcul.Enabled = True

                Dim T_surf(CInt(N_parcelle), 1)
                For i = 0 To CInt(N_parcelle) - 2
                    T_surf(i, 0) = DataGridView_suface.Rows(i).Cells(0).Value
                Next
                T_surf(CInt(N_parcelle) - 1, 0) = surf_affic
            Else
                btn_calcul.Enabled = False
            End If


        End If
    End Sub

    Private Sub rbnt_dir_par_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnt_dir_par.CheckedChanged
        If rbnt_dir_par.Checked Then
            bnt_orientation_dir.Enabled = True
            btn_calcul.Enabled = False
        End If
    End Sub

    Private Sub rbnt_dir_perp_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rbnt_dir_perp.CheckedChanged
        If rbnt_dir_perp.Checked Then
            bnt_orientation_dir.Enabled = True
            btn_calcul.Enabled = False
        End If
    End Sub


    Private Sub frmDivSurf_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        DataGridView_suface.AutoGenerateColumns = False

        Me.AcceptButton = btn_calcul
        'Me.CancelButton = cmdCancel

        chkLayerMO.Checked = True

    End Sub
End Class
