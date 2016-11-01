

Imports System
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.EditorInput
Imports System.Math
Imports Autodesk.AutoCAD.Colors

Module modDivSurf

    Public T_parcelle_id() As ObjectId
    Public T_parcelle_surf() As Double

    ''' <summary>
    ''' module de lancement du calcul
    ''' </summary>
    ''' <param name="T_surf">Tableau de surface</param>
    ''' <param name="meth">Méthode de division (par ou point fixe)</param>
    ''' <param name="orient">Orientation de la division</param>
    ''' <param name="N_parcelle">Nbr de parcelle à créer</param>
    ''' <param name="poly">Polyligne fermée</param>
    ''' <param name="ptPic1">Point 1 de la direction</param>
    ''' <param name="ptPic2">Point 2 de la direction</param>
    ''' <remarks></remarks>
    Public Sub Calcul(ByVal meth As String, ByVal orient As Boolean, ByVal N_parcelle As String, ByVal poly As Polyline, ByVal ptPic1 As Point2d, ByVal ptPic2 As Point2d, ByVal T_surf() As Double)

        ReDim T_parcelle_id(N_parcelle - 1)
        ReDim T_parcelle_surf(N_parcelle - 1)

        'appel de la méthode de calcul adapté 
        If meth = "par" Then
            Calcul_par(meth, orient, N_parcelle, poly, ptPic1, ptPic2, T_surf)
        ElseIf meth = "pt" Then
            Calcul_pt(meth, orient, N_parcelle, poly, ptPic1, ptPic2, T_surf)
        End If


    End Sub

    ''' <summary>
    ''' Module de calcul passant par un point
    ''' </summary>
    ''' <param name="T_surf">Tableau de surface</param>
    ''' <param name="meth">Méthode de division (par ou point fixe)</param>
    ''' <param name="orient">Orientation de la division</param>
    ''' <param name="N_par">Nbr de parcelle à créer</param>
    ''' <param name="poly">Polyligne fermée</param>
    ''' <param name="ptPic1">Point 1 de la direction</param>
    ''' <param name="ptPic2">Point 2 de la direction</param>
    ''' <remarks></remarks>
    Public Sub Calcul_par(ByVal meth As String, ByVal orient As Boolean, ByVal N_par As String, ByVal poly As Polyline, ByVal ptPic1 As Point2d, ByVal ptPic2 As Point2d, ByVal T_surf() As Double)
        ' Calcul d'une division parallèle à une direction 


        Dim N_parcelle As Integer = CInt(N_par) ''Nombre de parcelle à creer 

        Dim VectDroiteProv As Vector2d = ptPic1.GetVectorTo(New Point2d(ptPic2.X, ptPic2.Y))
        Dim TransfMat As Matrix2d = Matrix2d.Rotation(VectDroiteProv.Angle * (-1), ptPic1)
        Dim j As Integer

        'For each surface to create : creation of the surface and complementary surface
        For k = 0 To N_parcelle - 2
            Dim T_point(poly.NumberOfVertices - 1, 4) As Double
            Dim Tempid, Tempxt, Tempyt, Tempx, Tempy As Double    'Variable temporaire
            Dim inter_1 As Point2d = New Point2d '' Premier point d'intersection entre la polyligne et la droite passant par le sommet j 
            Dim inter_2 As Point2d = New Point2d '' Deuxième point d'intersection entre la polyligne et la droite passant par le sommet j 
            Dim inter_1_p As Point2d = New Point2d(T_point(0, 3), T_point(0, 4))  '' Premier point d'intersection entre la polyligne et la droite passant par le sommet j - 1
            Dim inter_2_p As Point2d = New Point2d(T_point(0, 3), T_point(0, 4))  '' Deuxième point d'intersection entre la polyligne et la droite passant par le sommet j - 1
            Dim inter_1_temp As Point2d = New Point2d(T_point(0, 3), T_point(0, 4))
            Dim inter_2_temp As Point2d = New Point2d(T_point(0, 3), T_point(0, 4))
            Dim p2 As Integer = 0  '' Variable d'incrémentation pour le parcours du polygone 
            Dim p1 As Integer = 0 '' Variable d'incrémentation pour le parcours du polygone
            Dim Cas_cplx As Boolean = False '' Si = true : cas complexe / Si = false : cas "normal" : cas cplx avec une pointe "sortante"
            Dim T_parcelle(,) As Double '' Table de point contenant les points de la polyligne de travail
            Dim T_point_nouv(1, 1) As Double '' Table de point contenant les points créés après la division
            Dim S_recherche As Double = 0 ' surface de la parcelle de recherche
            Dim S_recherche_p As Double = 0 ''Surface précédente 
            Dim sommet_p1 As Boolean = False
            Dim sommet_p2 As Boolean = False
            Dim cas_part As Boolean = False ' sera vrai lorsque la direction de la division sera // à un des cotés auxquels appartient le point de base 
            Dim cas_cplx_rentrant As Boolean = False '' Si = true : cas complexe / Si = false : cas "normal" : cas cplx avec une pointe "rentrante"
            Dim fin_cas_cplx_rentrant As Boolean = False '' est vrai quand il n'existe plus d'ambiguité dûe au cas cplx 
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor
            Dim calcul_ok As Boolean = True

            ''creation du tableau de point
            ''calcul des coordonnées dans le nouveau repère aligné avec la droite
            For i As Integer = 0 To poly.NumberOfVertices - 1
                T_point(i, 0) = i
                T_point(i, 1) = poly.GetPoint2dAt(i).TransformBy(TransfMat).X
                T_point(i, 2) = poly.GetPoint2dAt(i).TransformBy(TransfMat).Y
                T_point(i, 3) = poly.GetPoint2dAt(i).X
                T_point(i, 4) = poly.GetPoint2dAt(i).Y
            Next
            ''fin - creation des coordonnées 



            ''tri des coordonnees suivant la valeur de Y
            ''tri croissant si orient = false
            If orient = False Then
                'Routine de tri
                For i = 0 To poly.NumberOfVertices.ToString - 2    'Boucle externe
                    For j = 0 To poly.NumberOfVertices.ToString - 2   'Boucle interne
                        If T_point(j, 2) > T_point(j + 1, 2) Then
                            Tempid = T_point(j, 0) : Tempxt = T_point(j, 1) : Tempyt = T_point(j, 2) _
                              : Tempx = T_point(j, 3) : Tempy = T_point(j, 4)

                            T_point(j, 0) = T_point(j + 1, 0) : T_point(j, 1) = T_point(j + 1, 1) _
                                : T_point(j, 2) = T_point(j + 1, 2) : T_point(j, 3) = T_point(j + 1, 3) : T_point(j, 4) = T_point(j + 1, 4)

                            T_point(j + 1, 0) = Tempid : T_point(j + 1, 1) = Tempxt : T_point(j + 1, 2) = Tempyt _
                                : T_point(j + 1, 3) = Tempx : T_point(j + 1, 4) = Tempy 'Inverser si pas dans le bon ordre
                        End If
                    Next j
                Next i
            Else
                ''tri décroissant si orient = true
                For i = 0 To poly.NumberOfVertices.ToString - 2    'Boucle externe
                    For j = 0 To poly.NumberOfVertices.ToString - 2   'Boucle interne
                        If T_point(j, 2) < T_point(j + 1, 2) Then
                            Tempid = T_point(j, 0) : Tempxt = T_point(j, 1) : Tempyt = T_point(j, 2) _
                              : Tempx = T_point(j, 3) : Tempy = T_point(j, 4)

                            T_point(j, 0) = T_point(j + 1, 0) : T_point(j, 1) = T_point(j + 1, 1) _
                                : T_point(j, 2) = T_point(j + 1, 2) : T_point(j, 3) = T_point(j + 1, 3) : T_point(j, 4) = T_point(j + 1, 4)

                            T_point(j + 1, 0) = Tempid : T_point(j + 1, 1) = Tempxt : T_point(j + 1, 2) = Tempyt _
                              : T_point(j + 1, 3) = Tempx : T_point(j + 1, 4) = Tempy 'Inverser si pas dans le bon ordre
                        End If
                    Next j
                Next i
            End If
            ''fin - tri des coordonnees suivant la valeur de Y

            Console.Write(T_point)

            'initialisation des index de sommets 
            p1 = T_point(0, 0)
            p2 = T_point(0, 0)


            'initalisation des intersections dans la 1ère boucle 
            If k = 0 Then
                If T_point(0, 0) <> poly.NumberOfVertices.ToString - 1 Then : inter_1_temp = poly.GetPoint2dAt(T_point(0, 0) + 1)
                Else : inter_1_temp = poly.GetPoint2dAt(0)
                End If

                If T_point(0, 0) <> 0 Then : inter_2_temp = poly.GetPoint2dAt(T_point(0, 0) - 1)
                Else : inter_2_temp = poly.GetPoint2dAt(poly.NumberOfVertices - 1)
                End If
            End If


            j = 0 'numéro du sommet courant



            'Initialisation des valeurs de inter_1_temp et inter_2_temp
            If T_point(0, 0) <> poly.NumberOfVertices.ToString - 1 Then
                inter_1_temp = poly.GetPoint2dAt(T_point(0, 0) + 1)
            Else
                inter_1_temp = poly.GetPoint2dAt(poly.NumberOfVertices.ToString - 1)
            End If

            If T_point(0, 0) <> 0 Then
                inter_2_temp = poly.GetPoint2dAt(T_point(0, 0) - 1)
            Else
                inter_2_temp = poly.GetPoint2dAt(0)
            End If






            ''Tant que la surface calculée est insufisante - on crée la surface suivante
            ''Si la surfarce calculéé est suffisante alors division du polygone 
            While S_recherche <= T_surf(k) And j + 1 < poly.NumberOfVertices - 1
                j = j + 1

                'Création de la droite parallèle passant par le point j
                Dim pt_j As Point3d = New Point3d(T_point(j, 3), T_point(j, 4), 0) 'pt_j est le sommet de la polyligne courant où passe la ligne acLine
                Dim pt_j2 As Point3d = New Point3d(ptPic2.X + (T_point(j, 3) - ptPic1.X), ptPic2.Y + (T_point(j, 4) - ptPic1.Y), 0)
                Dim acLine As Xline = New Xline()
                acLine.BasePoint = pt_j
                acLine.SecondPoint = pt_j2





                ''recherche de la 1ère intersection "coté matricule ascendant"
                Dim intersec As Boolean = False
                sommet_p1 = False


                'En partant du point de base, on cherche la 1ère intersection entre un coté du polygone et la droite acLine
                While intersec = False
                    ''Si le compteur n'attiens pas l'iD du sommet courant alors on cherche l'intersection


                    If p1 + 1 <> T_point(j, 0) And Not (p1 + 1 = poly.NumberOfVertices.ToString And T_point(j, 0) = 0) Then

                        ''Création de la droite passant par deux sommets du polygone 
                        Dim acLine_p As Xline = New Xline()
                        If p1 < (CInt(poly.NumberOfVertices.ToString) - 1) Then
                            acLine_p.BasePoint = poly.GetPoint3dAt(p1)
                            acLine_p.SecondPoint = poly.GetPoint3dAt(p1 + 1)
                        Else
                            acLine_p.BasePoint = poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1)
                            acLine_p.SecondPoint = poly.GetPoint3dAt(0)
                        End If

                        '' Fin - Création de la droite passant par deux sommets du polygone 

                        '' Calcul de l'intersection 
                        Dim IntersectPts As Point3dCollection = New Point3dCollection
                        acLine.IntersectWith(acLine_p, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)
                        '' Fin - Calcul de l'intersection 

                        '' Vérifier que l'intersection est entre les deux sommets
                        ''Si oui alors intersection = True
                        If IntersectPts.Count = 1 Then
                            inter_1_temp = inter_1
                            inter_1 = pt_Intersect(IntersectPts)

                            If p1 + 1 < poly.NumberOfVertices.ToString Then
                                If Not (inter_1.X > poly.GetPoint3dAt(p1).X And inter_1.X > poly.GetPoint3dAt(p1 + 1).X) And _
                                    Not (inter_1.X < poly.GetPoint3dAt(p1).X And inter_1.X < poly.GetPoint3dAt(p1 + 1).X) And _
                                    Not (inter_1.Y > poly.GetPoint3dAt(p1).Y And inter_1.Y > poly.GetPoint3dAt(p1 + 1).Y) And _
                                    Not (inter_1.Y < poly.GetPoint3dAt(p1).Y And inter_1.Y < poly.GetPoint3dAt(p1 + 1).Y) Then
                                    intersec = True
                                End If
                            ElseIf Not (inter_1.X > poly.GetPoint3dAt(0).X And inter_1.X > poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).X) And _
                                    Not (inter_1.X < poly.GetPoint3dAt(0).X And inter_1.X < poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).X) And _
                                    Not (inter_1.Y > poly.GetPoint3dAt(0).Y And inter_1.Y > poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).Y) And _
                                    Not (inter_1.Y < poly.GetPoint3dAt(0).Y And inter_1.Y < poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).Y) Then
                                intersec = True
                            End If
                        End If
                        ''Fin - Vérifier que l'intersection est entre les deux sommets

                        ''Si intersect = false : incrémentation du compteur
                        If intersec = False Then
                            p1 = p1 + 1
                            inter_1 = inter_1_temp
                        Else
                            inter_1_p = inter_1_temp
                        End If

                        If p1 = poly.NumberOfVertices.ToString Then
                            p1 = 0
                        End If

                    Else
                        inter_1_p = inter_1
                        inter_1 = New Point2d(pt_j.X, pt_j.Y)
                        intersec = True
                        sommet_p1 = True
                    End If


                End While
                '' Fin - recherche de la 1ère intersection "coté matricule ascendant"


                'Dim colecIntersect1 As New Collection
                'colecIntersect1 = calcul_intersection(poly, ptPic1, ptPic2, inter_1, inter_1_p, p1, T_point, j)
                'inter_1_p = colecIntersect1.Item(1)
                'inter_1 = colecIntersect1.Item(2)
                'p1 = colecIntersect1.Item(3)











                ''recherche de la 2ème intersection "coté matricule descendant"
                intersec = False
                sommet_p2 = False
                p2 = T_point(0, 0)

                ''En partant du point de base, on cherche la 2ème intersection entre un coté du polygone et la droite acLine
                While intersec = False

                    ''Si le compteur n'attiens pas l'iD du sommet courant alors on cherche l'intersection



                    If p2 - 1 <> T_point(j, 0) And Not (p2 - 1 = -1 And T_point(j, 0) = poly.NumberOfVertices - 1) Then

                        ''Création de la droite passant par deux sommets du polygone 
                        Dim acLine_p As Xline = New Xline()

                        If p2 - 1 > -1 Then
                            acLine_p.BasePoint = poly.GetPoint3dAt(p2)
                            acLine_p.SecondPoint = poly.GetPoint3dAt(p2 - 1)
                        Else
                            acLine_p.BasePoint = poly.GetPoint3dAt(0)
                            acLine_p.SecondPoint = poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1)
                        End If
                        ''Fin - Création de la droite passant par deux sommets du polygone 

                        ''Calcul de l'intersection
                        Dim IntersectPts As Point3dCollection = New Point3dCollection
                        acLine.IntersectWith(acLine_p, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)
                        ''Fin - Calcul de l'intersection

                        ''Vérifier que l'intersection est entre les deux sommets
                        ''Si oui alors intersection = True
                        If IntersectPts.Count = 1 Then
                            inter_2_temp = inter_2
                            inter_2 = pt_Intersect(IntersectPts)

                            If p2 - 1 > -1 Then
                                If Not (inter_2.X > poly.GetPoint3dAt(p2).X And inter_2.X > poly.GetPoint3dAt(p2 - 1).X) And _
                                        Not (inter_2.X < poly.GetPoint3dAt(p2).X And inter_2.X < poly.GetPoint3dAt(p2 - 1).X) And _
                                        Not (inter_2.Y > poly.GetPoint3dAt(p2).Y And inter_2.Y > poly.GetPoint3dAt(p2 - 1).Y) And _
                                        Not (inter_2.Y < poly.GetPoint3dAt(p2).Y And inter_2.Y < poly.GetPoint3dAt(p2 - 1).Y) Then
                                    intersec = True

                                End If
                            ElseIf Not (inter_2.X > poly.GetPoint3dAt(0).X And inter_2.X > poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).X) And _
                                        Not (inter_2.X < poly.GetPoint3dAt(0).X And inter_2.X < poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).X) And _
                                        Not (inter_2.Y > poly.GetPoint3dAt(0).Y And inter_2.Y > poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).Y) And _
                                        Not (inter_2.Y < poly.GetPoint3dAt(0).Y And inter_2.Y < poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).Y) Then
                                intersec = True
                            End If
                        End If
                        ''Fin - Vérifier que l'intersection est entre les deux sommets

                        ''Si intersect = false : incrémentation du compteur
                        If intersec = False Then
                            p2 = p2 - 1
                            inter_2 = inter_2_temp
                        Else
                            inter_2_p = inter_2_temp
                        End If

                        ''Ajuster compteur s'il est négatif
                        If p2 - 1 < -1 And intersec = False Then
                            p2 = CInt(poly.NumberOfVertices.ToString) - 1
                        End If


                        ''Si le compteur attiens l'iD du sommet courant alors intersection = sommet courant
                    Else
                        inter_2_p = inter_2
                        inter_2 = New Point2d(pt_j.X, pt_j.Y)
                        intersec = True
                        sommet_p2 = True
                    End If



                End While
                ''fin - recherche de la 2ème intersection "coté matricule descendant"



                'Dim colecIntersect2 As New Collection
                'colecIntersect2 = calcul_intersection(poly, ptPic1, ptPic2, inter_2, inter_2_p, p2, T_point, j, 1)
                'inter_2_p = colecIntersect2.Item(1)
                'inter_2 = colecIntersect2.Item(2)
                'p2 = colecIntersect2.Item(3)


                ' Si on est dans le cas_cplx_rentrant 
                ' On teste si on est tj dans le cas cplx (la droite ne fait plus que 3 intersections avec le polygone
                ' Si on est dans le cap_cplx_rentrant on n'effectue pas de calcul 
                If cas_cplx_rentrant Then

                    'création de la droite
                    Dim xline_cplx_1 As New Xline
                    xline_cplx_1.BasePoint = New Point3d(inter_1.X, inter_1.Y, 0)
                    xline_cplx_1.SecondPoint = New Point3d(inter_2.X, inter_2.Y, 0)

                    Dim id_temp_cplx_1 As Double
                    Dim id_temp_cplx_2 As Double
                    Dim id_a_1, id_a_2, id_b_1, id_b_2 As Double

                    Dim inter_1_cplx As Point2d = New Point2d
                    Dim inter_2_cplx As Point2d = New Point2d

                    Dim compteur_inter As Double = 0

                    'Vérifie si la droite intersecte chaque coté du polygone 
                    For n_ligne = 0 To poly.NumberOfVertices - 1

                        'Création de la 2ème droite
                        Dim xline_cplx_2 As New Xline
                        If n_ligne <> poly.NumberOfVertices - 1 Then
                            xline_cplx_2.BasePoint = poly.GetPoint3dAt(n_ligne)
                            xline_cplx_2.SecondPoint = poly.GetPoint3dAt(n_ligne + 1)
                            id_temp_cplx_1 = n_ligne
                            id_temp_cplx_2 = n_ligne + 1
                        Else
                            xline_cplx_2.BasePoint = poly.GetPoint3dAt(n_ligne)
                            xline_cplx_2.SecondPoint = poly.GetPoint3dAt(0)
                            id_temp_cplx_1 = n_ligne
                            id_temp_cplx_2 = 0
                        End If

                        'Calcul de l'intersection 
                        Dim IntersectPts As Point3dCollection = New Point3dCollection
                        xline_cplx_1.IntersectWith(xline_cplx_2, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)

                        'S'il existe une intersection on vérifie qu'elle est sur le segment 
                        If IntersectPts.Count <> 0 Then
                            Dim inter_temp_cplx As Point2d
                            inter_temp_cplx = pt_Intersect(IntersectPts)

                            'test si l'intersection est sur le polygone
                            If dist(inter_temp_cplx, poly.GetPoint2dAt(id_temp_cplx_1)) <= dist(poly.GetPoint2dAt(id_temp_cplx_1), poly.GetPoint2dAt(id_temp_cplx_2)) _
                                And dist(inter_temp_cplx, poly.GetPoint2dAt(id_temp_cplx_2)) <= dist(poly.GetPoint2dAt(id_temp_cplx_1), poly.GetPoint2dAt(id_temp_cplx_2)) Then
                                If inter_temp_cplx.X <> T_point(j, 3) And inter_temp_cplx.Y <> T_point(j, 4) Then
                                    compteur_inter = compteur_inter + 1

                                    If compteur_inter = 1 Then
                                        inter_1_cplx = inter_temp_cplx
                                        id_a_1 = id_temp_cplx_1
                                        id_a_2 = id_temp_cplx_2
                                    End If

                                    If compteur_inter = 2 Then
                                        inter_2_cplx = inter_temp_cplx
                                        id_b_1 = id_temp_cplx_1
                                        id_b_2 = id_temp_cplx_2
                                    End If

                                End If
                            End If
                        End If
                    Next

                    'Si le cas_cplx_cplx n'est plus, on affecte les valeurs corrigées des index et des valeurs d'intersections 
                    If compteur_inter = 2 Then
                        cas_cplx_rentrant = False
                        fin_cas_cplx_rentrant = True

                        If Math.Round(inter_1.X, 10) = Math.Round(T_point(j, 3), 10) And Math.Round(inter_1.Y, 10) = Math.Round(T_point(j, 4), 10) Then
                            If Math.Round(inter_2.X, 10) = Math.Round(inter_1_cplx.X, 10) And Math.Round(inter_2.Y, 10) = Math.Round(inter_1_cplx.Y, 10) Then
                                inter_1 = inter_2_cplx
                                p1 = id_b_1
                            Else
                                inter_1 = inter_1_cplx
                                p1 = id_a_1
                            End If
                        Else
                            If Math.Round(inter_1.X, 10) = Math.Round(inter_1_cplx.X, 10) And Math.Round(inter_1.Y, 10) = Math.Round(inter_1_cplx.Y, 10) Then
                                inter_2 = inter_2_cplx
                                p2 = id_b_2
                            Else
                                inter_2 = inter_1_cplx
                                p2 = id_b_2
                            End If
                        End If

                    End If
                End If
                ' Fin - Si on est dans le cas_cplx_rentrant 
                ' Fin - On teste si on est tj dans le cas cplx (la droite ne fait plus que 3 intersections avec le polygone
                ' Fin - Si on est dans le cap_cplx_rentrant on n'effectue pas de calcul 



                ' Si le compteur p1 = -1 alors p1 = Nombre de sommets - 1 
                If p1 = -1 Then
                    p1 = poly.NumberOfVertices.ToString - 1
                End If







                '' Création de la polyligne provisoire & calcul de S_recherche suivant 
                Dim T_parcelle_X As New List(Of Double)
                Dim T_parcelle_Y As New List(Of Double)


                '' Création des listes de points (X et Y) des points de la parcelles suivant les valeurs de p1 et p2 
                '' Ajout des points entre p1 et le point de base
                If p1 >= T_point(0, 0) Then
                    For i = 0 To p1 - T_point(0, 0)
                        T_parcelle_X.Add(CDbl(poly.GetPoint2dAt(T_point(0, 0) + i).X))
                        T_parcelle_Y.Add(CDbl(poly.GetPoint2dAt(T_point(0, 0) + i).Y))
                    Next
                ElseIf p1 < T_point(0, 0) Then
                    For i = 0 To poly.NumberOfVertices - T_point(0, 0) + p1
                        If T_point(0, 0) + i < poly.NumberOfVertices.ToString Then
                            T_parcelle_X.Add(CDbl(poly.GetPoint2dAt(T_point(0, 0) + i).X))
                            T_parcelle_Y.Add(CDbl(poly.GetPoint2dAt(T_point(0, 0) + i).Y))
                        Else
                            T_parcelle_X.Add(CDbl(poly.GetPoint2dAt(T_point(0, 0) + i - CInt(poly.NumberOfVertices)).X))
                            T_parcelle_Y.Add(CDbl(poly.GetPoint2dAt(T_point(0, 0) + i - CInt(poly.NumberOfVertices)).Y))
                        End If

                    Next
                Else
                    T_parcelle_X.Add(T_point(0, 3))
                    T_parcelle_Y.Add(T_point(0, 4))
                End If

                T_parcelle_X.Add(inter_1.X)
                T_parcelle_Y.Add(inter_1.Y)
                T_parcelle_X.Add(inter_2.X)
                T_parcelle_Y.Add(inter_2.Y)



                '' Ajout des points entre p2 et le point de base
                If p2 <= T_point(0, 0) Then
                    For i = 0 To T_point(0, 0) - p2
                        T_parcelle_X.Add(CDbl(poly.GetPoint2dAt(p2 + i).X))
                        T_parcelle_Y.Add(CDbl(poly.GetPoint2dAt(p2 + i).Y))
                    Next
                    'ElseIf p2 > T_point(0, 0) Then
                    '    For i = 0 To poly.NumberOfVertices + T_point(0, 0) - p2 - 1
                    '        If p2 - i >= 0 Then
                    '            T_parcelle_X.Add(CDbl(poly.GetPoint2dAt(p2 - i).X))
                    '            T_parcelle_Y.Add(CDbl(poly.GetPoint2dAt(p2 - i).Y))
                    '        Else
                    '            T_parcelle_X.Add(CDbl(poly.GetPoint2dAt(p2 - i + CInt(poly.NumberOfVertices)).X))
                    '            T_parcelle_Y.Add(CDbl(poly.GetPoint2dAt(p2 - i + CInt(poly.NumberOfVertices)).Y))
                    '        End If
                    '    Next
                    'Else
                ElseIf p2 > T_point(0, 0) Then
                    For i = 0 To poly.NumberOfVertices + T_point(0, 0) - p2 - 1
                        If p2 + i <= poly.NumberOfVertices.ToString - 1 Then
                            T_parcelle_X.Add(CDbl(poly.GetPoint2dAt(p2 + i).X))
                            T_parcelle_Y.Add(CDbl(poly.GetPoint2dAt(p2 + i).Y))
                        Else
                            T_parcelle_X.Add(CDbl(poly.GetPoint2dAt(p2 + i - CInt(poly.NumberOfVertices)).X))
                            T_parcelle_Y.Add(CDbl(poly.GetPoint2dAt(p2 + i - CInt(poly.NumberOfVertices)).Y))
                        End If
                    Next
                Else
                    T_parcelle_X.Add(T_point(0, 3))
                    T_parcelle_Y.Add(T_point(0, 4))
                End If
                '' Fin - Création des listes de points (X et Y) des points de la parcelles suivant les valeurs de p1 et p2 


                ReDim T_parcelle(T_parcelle_X.Count - 1, 1)
                ''Création de la table de points 
                For i = 0 To (T_parcelle.Length / 2) - 1
                    T_parcelle(i, 0) = T_parcelle_X(i)
                    T_parcelle(i, 1) = T_parcelle_Y(i)
                Next



                ''Test si cas_part 
                ''Si oui : calcul des intersections adaptées
                '' note : cas_part : si le point base se trouve sur un sommet // à la direction 
                '' note : Dans ce cas les deux inter sont égales 
                cas_part = False
                If j = 2 And (New Point2d(inter_1_p.X, inter_1_p.Y) = New Point2d(inter_2_p.X, inter_2_p.Y)) Then
                    cas_part = True
                    If T_point(0, 0) = 0 And T_point(1, 0) = poly.NumberOfVertices.ToString - 1 Then
                        inter_2_p = New Point2d(T_point(1, 3), T_point(1, 4))
                    ElseIf T_point(1, 0) = 0 And T_point(0, 0) = poly.NumberOfVertices.ToString - 1 Then
                        inter_1_p = New Point2d(T_point(0, 3), T_point(0, 4))
                    ElseIf T_point(0, 0) < T_point(1, 0) Then
                        inter_2_p = New Point2d(T_point(0, 3), T_point(0, 4))
                    Else
                        inter_1_p = New Point2d(T_point(1, 3), T_point(1, 4))
                    End If
                End If
                '' Fin - Création de la polyligne provisoire & calcul de S_recherche suivant 



                ' Si aucun pb : on calcul 
                If Not cas_cplx_rentrant Then
                    ''Calcul de la surface
                    S_recherche_p = S_recherche
                    S_recherche = Surface_Parcelle(T_parcelle)
                End If


                ' Lorsque la surface calculée à la sortie du cas_cplx_rentrant est trop grande, on stope le calcul
                If fin_cas_cplx_rentrant Then
                    If S_recherche > T_surf(k) Then
                        ed.WriteMessage(vbLf & "La division pose problème : tentez de diviser en orientant de l'autre sens")
                        calcul_ok = False
                    End If
                    fin_cas_cplx_rentrant = False
                End If



                If inter_1 = inter_2 Then S_recherche = S_recherche_p


                'Test si l'on est dans le cas_cplx rentrant : le point courant se trouve (strictement) entre les deux intersections
                If dist(inter_1, inter_2) > dist(inter_1, New Point2d(T_point(j, 3), T_point(j, 4))) And _
                    dist(inter_1, inter_2) > dist(inter_2, New Point2d(T_point(j, 3), T_point(j, 4))) Then
                    If dist(inter_1, New Point2d(T_point(j, 3), T_point(j, 4))) <> 0 And _
                        dist(inter_2, New Point2d(T_point(j, 3), T_point(j, 4))) Then
                        If S_recherche < T_surf(k) Then
                            cas_cplx_rentrant = True
                        End If
                    End If
                End If



            End While
            ''Fin - Tant que la surface calculée est insufisante - on crée la surface suivante
            ''Fin - Si la surfarce calculéé est suffisante alors division du polygone 



















            If calcul_ok Then



                Dim cas_parcelle_triangle As Boolean = False
                Dim cas_triangle As Boolean = False
                ReDim T_parcelle(3, 1)

                'Si la surface est un triangle et la division est parrallèle à une base 
                If T_point(0, 2) - 0.00001 < T_point(1, 2) < T_point(0, 2) + 0.00001 And poly.NumberOfVertices = 3 Then
                    cas_parcelle_triangle = True

                    'Construction du tableau de point
                    ReDim T_parcelle(2, 1)
                    T_parcelle(0, 0) = T_point(0, 3)
                    T_parcelle(0, 1) = T_point(0, 4)
                    T_parcelle(1, 0) = T_point(1, 3)
                    T_parcelle(1, 1) = T_point(1, 4)
                    T_parcelle(2, 0) = T_point(2, 3)
                    T_parcelle(2, 1) = T_point(2, 4)


                    'Calcul de la surface à créer 
                    Dim S_a_creer As Double = T_surf(k)
                    'Calcul des points nouveaux
                    T_point_nouv = div_triangle(poly, S_a_creer, T_parcelle)

                    'Affectation des coordonnées de l'intersection au table de point 
                    '(parcelle calculée) 
                    ReDim T_parcelle(3, 1)
                    T_parcelle(0, 0) = T_point_nouv(0, 0)
                    T_parcelle(0, 1) = T_point_nouv(0, 1)
                    T_parcelle(1, 0) = T_point_nouv(1, 0)
                    T_parcelle(1, 1) = T_point_nouv(1, 1)
                    T_parcelle(3, 0) = T_point(0, 3)
                    T_parcelle(3, 1) = T_point(0, 4)
                    T_parcelle(2, 0) = T_point(1, 3)
                    T_parcelle(2, 1) = T_point(1, 4)

                    Draw(Parcelle(T_parcelle))

                    'Affectation des coordonnées de l'intersection au table de point 
                    '(parcelle suivante) 
                    ReDim T_parcelle(2, 1)
                    T_parcelle(1, 0) = T_point_nouv(0, 0)
                    T_parcelle(1, 1) = T_point_nouv(0, 1)
                    T_parcelle(0, 0) = T_point_nouv(1, 0)
                    T_parcelle(0, 1) = T_point_nouv(1, 1)
                    T_parcelle(2, 0) = T_point(2, 3)
                    T_parcelle(2, 1) = T_point(2, 4)

                    'Si la parcelle suivante est la dernière alors on la dessine 
                    If k = N_parcelle - 2 Then
                        Draw(Parcelle(T_parcelle))
                    End If


                    'Si la solution se trouve dans le 1er triangle alors on divise le triangle
                ElseIf j = 1 Then
                    ReDim T_parcelle(2, 1)
                    T_parcelle(0, 0) = inter_1.X
                    T_parcelle(0, 1) = inter_1.Y
                    T_parcelle(1, 0) = inter_2.X
                    T_parcelle(1, 1) = inter_2.Y
                    T_parcelle(2, 0) = T_point(0, 3)
                    T_parcelle(2, 1) = T_point(0, 4)

                    'Calcul de la surface à créer 
                    Dim S_a_creer As Double = Surface_Parcelle(T_parcelle) - T_surf(k)
                    'Calcul des points nouveaux
                    T_point_nouv = div_triangle(poly, S_a_creer, T_parcelle)

                    T_parcelle(1, 0) = T_point_nouv(0, 0)
                    T_parcelle(1, 1) = T_point_nouv(0, 1)
                    T_parcelle(0, 0) = T_point_nouv(1, 0)
                    T_parcelle(0, 1) = T_point_nouv(1, 1)


                    'Si le prochain sommet courant le dernier sommet, on résoud dans le dernier triangle
                ElseIf (j + 2 = poly.NumberOfVertices) And (S_recherche < T_surf(k)) Then

                    cas_triangle = True
                    If sommet_p1 Then
                        If p1 + 1 > poly.NumberOfVertices - 1 Then
                            p1 = 0
                        Else
                            p1 = p1 + 1
                        End If
                    End If

                    If sommet_p2 Then
                        If p2 - 1 < 0 Then
                            p2 = poly.NumberOfVertices - 1
                        Else
                            p2 = p2 - 1
                        End If
                    End If

                    ReDim T_parcelle(2, 1)
                    T_parcelle(0, 0) = inter_1.X
                    T_parcelle(0, 1) = inter_1.Y
                    T_parcelle(1, 0) = inter_2.X
                    T_parcelle(1, 1) = inter_2.Y
                    T_parcelle(2, 0) = T_point(poly.NumberOfVertices - 1, 3)
                    T_parcelle(2, 1) = T_point(poly.NumberOfVertices - 1, 4)

                    'Calcul de la surface à créer 
                    Dim S_a_creer As Double = T_surf(k) - S_recherche
                    'Calcul des points nouveaux
                    T_point_nouv = div_triangle(poly, S_a_creer, T_parcelle)

                    T_parcelle(1, 0) = T_point_nouv(0, 0)
                    T_parcelle(1, 1) = T_point_nouv(0, 1)
                    T_parcelle(0, 0) = T_point_nouv(1, 0)
                    T_parcelle(0, 1) = T_point_nouv(1, 1)

                    'Sinon la solution revient à diviser un trapèze 
                Else



                    ''Définir si l'on est dans un cas ambigue ou non 
                    T_parcelle(0, 0) = inter_1_p.X
                    T_parcelle(0, 1) = inter_1_p.Y
                    T_parcelle(1, 0) = inter_1.X
                    T_parcelle(1, 1) = inter_1.Y
                    T_parcelle(2, 0) = inter_2.X
                    T_parcelle(2, 1) = inter_2.Y
                    T_parcelle(3, 0) = inter_2_p.X
                    T_parcelle(3, 1) = inter_2_p.Y
                    Dim p_cas_complexe_1 As Double = -10
                    Dim p_cas_complexe_2 As Double = -10

                    Cas_cplx = False

                    'Vérifie que si interX_p est un sommet il est sur la droite (interX, pX +/- 1)
                    'Si non : cas complexe 
                    For i As Double = 0 To ((T_point.Length / 5) - 1)


                        '' Si inter_1_p est un sommet du polygone 
                        If New Point2d(inter_1_p.X, inter_1_p.Y) = New Point2d(T_point(i, 3), T_point(i, 4)) Then
                            If p1 + 1 < (CInt(poly.NumberOfVertices.ToString) - 1) Then
                                If New Point2d(poly.GetPoint2dAt(p1).X, poly.GetPoint2dAt(p1).Y) <> _
                                    New Point2d(inter_1_p.X, inter_1_p.Y) Then
                                    Cas_cplx = True
                                    p_cas_complexe_1 = T_point(i, 0)
                                    i = T_point.Length / 5 - 1
                                End If
                            Else
                                If New Point2d(poly.GetPoint2dAt(poly.NumberOfVertices.ToString - 1).X, _
                                               poly.GetPoint2dAt(poly.NumberOfVertices.ToString - 1).Y) <> _
                                           New Point2d(inter_1_p.X, inter_1_p.Y) Then
                                    Cas_cplx = True
                                    p_cas_complexe_1 = T_point(i, 0)
                                    i = T_point.Length / 5 - 1
                                End If
                            End If
                        End If


                        '' Si inter_2_p est un sommet du polygone
                        If p_cas_complexe_1 = -10 Then
                            If New Point2d(inter_2_p.X, inter_2_p.Y) = New Point2d(T_point(i, 3), T_point(i, 4)) Then
                                If p2 > -1 Then
                                    If New Point2d(poly.GetPoint2dAt(p2).X, poly.GetPoint2dAt(p2).Y) _
                                        <> New Point2d(inter_2_p.X, inter_2_p.Y) Then
                                        Cas_cplx = True
                                        p_cas_complexe_2 = T_point(i, 0)
                                        i = T_point.Length / 5 - 1
                                    End If
                                Else
                                    If New Point2d(poly.GetPoint2dAt(0).X, poly.GetPoint2dAt(0).Y) _
                                        <> New Point2d(inter_2_p.X, inter_2_p.Y) Then
                                        Cas_cplx = True
                                        p_cas_complexe_2 = T_point(i, 0)
                                        i = T_point.Length / 5 - 1
                                    End If
                                End If
                            End If
                        End If
                    Next
                    ' Fin - Définir si l'on est dans un cas ambigue ou non 

                    Console.Write(T_point)


                    If Cas_cplx = True Then

                        ''Calcul de la surface p_cas_complexe_i - pi - intersection 
                        '' Si la surface est trop grande alors stop
                        '' Sinon calcul dans le trapèze T_parcelle en remplacant inter_i_p par intersection

                        Dim collection_cas_cplx As Collection = New Collection
                        If p_cas_complexe_1 <> -10 Then
                            collection_cas_cplx = calcul_intersection(poly, ptPic1, ptPic2, New Point2d, New Point2d, p_cas_complexe_1 + 1, T_point, j - 1, 2)

                            ReDim T_parcelle(Abs(p_cas_complexe_1 - p1) + 2, 1) 'ReDim T_parcelle(Abs(p_cas_complexe_1 - p1) + 1, 1)
                            T_parcelle(0, 0) = collection_cas_cplx.Item(2).X
                            T_parcelle(0, 1) = collection_cas_cplx.Item(2).Y
                            T_parcelle(1, 0) = poly.GetPoint2dAt(p_cas_complexe_1).X
                            T_parcelle(1, 1) = poly.GetPoint2dAt(p_cas_complexe_1).Y
                            Dim pt As Double = 0
                            For i = 0 To Abs(p_cas_complexe_1 - p1)

                                If pt + p1 + 1 > poly.NumberOfVertices.ToString - 1 Then
                                    pt = -p1 - 1
                                End If
                                T_parcelle(i + 2, 0) = poly.GetPoint2dAt(pt + p1 + 1).X
                                T_parcelle(i + 2, 1) = poly.GetPoint2dAt(pt + p1 + 1).Y
                                pt = pt + 1
                            Next

                        ElseIf p_cas_complexe_2 <> -10 Then
                            collection_cas_cplx = calcul_intersection(poly, ptPic1, ptPic2, New Point2d, New Point2d, p_cas_complexe_2 - 1, T_point, j - 1, 1)

                            ReDim T_parcelle(Abs(p_cas_complexe_2 - p2) + 1, 1)
                            T_parcelle(0, 0) = collection_cas_cplx.Item(2).X
                            T_parcelle(0, 1) = collection_cas_cplx.Item(2).Y

                            Dim pt As Double = 0
                            For i = 0 To Abs(p_cas_complexe_2 - p2)

                                If pt + p2 > poly.NumberOfVertices.ToString - 1 Then
                                    pt = -p2
                                End If

                                T_parcelle(i + 1, 0) = poly.GetPoint2dAt(pt + p2).X
                                T_parcelle(i + 1, 1) = poly.GetPoint2dAt(pt + p2).Y
                                pt = pt + 1
                            Next
                        End If
                        Dim d As Double = Surface_Parcelle(T_parcelle)
                        '' Fin - Calcul de la surface p_cas_complexe_i - pi - intersection 


                        S_recherche_p = S_recherche_p + d

                        'Si la somme est trop grande on ne peut pas diviser sans créer deux surfaces 
                        If S_recherche_p > T_surf(k) Then
                            Dim connect As New Revo.connect
                            connect.Message("Calcul de surface", "Attention ! Ce calcul peut imposer de créer deux surfaces différentes.", False, 100, 100, "critical")
                            'Sinon transformation du tableau de point pour permettre le calcul et la division 
                        Else
                            ReDim T_parcelle(3, 1)
                            If p_cas_complexe_2 <> -10 Then
                                T_parcelle(0, 0) = inter_1_p.X
                                T_parcelle(0, 1) = inter_1_p.Y
                                T_parcelle(1, 0) = inter_1.X
                                T_parcelle(1, 1) = inter_1.Y
                                T_parcelle(2, 0) = inter_2.X
                                T_parcelle(2, 1) = inter_2.Y
                                T_parcelle(3, 0) = collection_cas_cplx.Item(2).X
                                T_parcelle(3, 1) = collection_cas_cplx.Item(2).Y
                                p2 = collection_cas_cplx.Item(3)
                                Cas_cplx = False
                            ElseIf p_cas_complexe_1 <> -10 Then
                                T_parcelle(0, 0) = collection_cas_cplx.Item(2).X
                                T_parcelle(0, 1) = collection_cas_cplx.Item(2).Y
                                T_parcelle(1, 0) = inter_1.X
                                T_parcelle(1, 1) = inter_1.Y
                                T_parcelle(2, 0) = inter_2.X
                                T_parcelle(2, 1) = inter_2.Y
                                T_parcelle(3, 0) = inter_2_p.X
                                T_parcelle(3, 1) = inter_2_p.Y
                                p1 = collection_cas_cplx.Item(3)
                                Cas_cplx = False
                            End If

                        End If

                    End If

                    ''Si cas_cplx = Faux : calcul des points de la division
                    If Cas_cplx = False Then

                        ''creation des nouveaux points
                        T_point_nouv = div_quadrilatère(poly, S_recherche, S_recherche_p, T_surf(k), T_parcelle)
                        T_parcelle(1, 0) = T_point_nouv(0, 0)
                        T_parcelle(1, 1) = T_point_nouv(0, 1)
                        T_parcelle(0, 0) = T_point_nouv(1, 0)
                        T_parcelle(0, 1) = T_point_nouv(1, 1)
                    End If

                End If
                ''FIN DES 4 CAS --> on peut dessiner 


                If Cas_cplx = False And cas_parcelle_triangle = False Then
                    'Dessin : parcelle calculée 

                    'Construction de la collection de point 
                    Dim T_parcelle_list As New Collection

                    'Ajout des points 
                    If p2 <= p1 Then
                        T_parcelle_list.Add(New Point2d(T_parcelle(1, 0), T_parcelle(1, 1)))
                        T_parcelle_list.Add(New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)))

                        For i As Integer = p2 To p1 Step 1
                            T_parcelle_list.Add(poly.GetPoint2dAt(i))
                        Next
                    Else
                        T_parcelle_list.Add(New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)))
                        T_parcelle_list.Add(New Point2d(T_parcelle(1, 0), T_parcelle(1, 1)))

                        For i As Integer = 0 To poly.NumberOfVertices - (p2 - p1)
                            If p1 - i < 0 Then
                                T_parcelle_list.Add(poly.GetPoint2dAt(poly.NumberOfVertices - i + p1))
                            Else
                                T_parcelle_list.Add(poly.GetPoint2dAt(p1 - i))
                            End If

                        Next
                    End If

                    ReDim T_parcelle(T_parcelle_list.Count - 1, 1)

                    For i = 0 To T_parcelle_list.Count - 1
                        T_parcelle(i, 0) = T_parcelle_list.Item(i + 1).X
                        T_parcelle(i, 1) = T_parcelle_list.Item(i + 1).Y
                    Next

                    'Dessin de la parcelle calculé
                    T_parcelle_id(k) = Draw(Parcelle(T_parcelle))
                    T_parcelle_surf(k) = Surface_Parcelle(T_parcelle)
                    T_parcelle_list.Clear()

                    'Dessin : parcelle suivante  
                    If cas_triangle Then
                        ReDim T_parcelle(2, 1)
                        T_parcelle(1, 0) = T_point_nouv(0, 0)
                        T_parcelle(1, 1) = T_point_nouv(0, 1)
                        T_parcelle(0, 0) = T_point_nouv(1, 0)
                        T_parcelle(0, 1) = T_point_nouv(1, 1)
                        T_parcelle(2, 0) = T_point(poly.NumberOfVertices - 1, 3)
                        T_parcelle(2, 1) = T_point(poly.NumberOfVertices - 1, 4)
                    Else
                        If p1 < p2 Then
                            T_parcelle_list.Add(New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)))
                            T_parcelle_list.Add(New Point2d(T_parcelle(1, 0), T_parcelle(1, 1)))

                            ' For i As Integer = 0 To poly.NumberOfVertices - p2 + p1 Step 1
                            For i As Integer = 0 To p2 - p1 - 2 Step 1

                                If poly.GetPoint2dAt(p1 + 1 + i) <> T_parcelle_list.Item(1) And poly.GetPoint2dAt(p1 + 1 + i) <> T_parcelle_list.Item(2) Then
                                    T_parcelle_list.Add(poly.GetPoint2dAt(p1 + 1 + i))
                                End If

                            Next

                        Else

                            T_parcelle_list.Add(New Point2d(T_parcelle(1, 0), T_parcelle(1, 1)))
                            T_parcelle_list.Add(New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)))

                            ' For i As Integer = 0 To poly.NumberOfVertices + p2 - p1 - 2 Step 1
                            For i As Integer = 0 To poly.NumberOfVertices.ToString - (p1 - p2 + 2)  ' test av + 2 au lieu de + 1 !!!! modification le 2 juillet !!!  
                                If p1 + 1 + i < poly.NumberOfVertices.ToString Then
                                    If poly.GetPoint2dAt(p1 + 1 + i) <> T_parcelle_list.Item(1) And poly.GetPoint2dAt(p1 + 1 + i) <> T_parcelle_list.Item(2) Then
                                        T_parcelle_list.Add(poly.GetPoint2dAt(p1 + 1 + i))
                                    End If
                                Else
                                    If poly.GetPoint2dAt((p1 + 1 + i) - poly.NumberOfVertices.ToString) <> T_parcelle_list.Item(1) And poly.GetPoint2dAt((p1 + 1 + i) - poly.NumberOfVertices.ToString) <> T_parcelle_list.Item(2) Then
                                        T_parcelle_list.Add(poly.GetPoint2dAt((p1 + 1 + i) - poly.NumberOfVertices.ToString))
                                    End If
                                End If

                            Next
                        End If



                        'dessin de la parcelle suivante 
                        ReDim T_parcelle(T_parcelle_list.Count - 1, 1)

                        For i = 0 To T_parcelle_list.Count - 1
                            T_parcelle(i, 0) = T_parcelle_list.Item(i + 1).X
                            T_parcelle(i, 1) = T_parcelle_list.Item(i + 1).Y
                        Next
                    End If

                    If k = N_parcelle - 2 Then
                        T_parcelle_id(k + 1) = Draw(Parcelle(T_parcelle))
                        T_parcelle_surf(k + 1) = Surface_Parcelle(T_parcelle)
                    End If

                    'Affectation de la parcelle suivant à la variable poly
                    poly = Parcelle(T_parcelle)

                ElseIf cas_parcelle_triangle = True Then
                    poly = Parcelle(T_parcelle)
                Else
                    k = N_parcelle - 2
                End If
            Else
                k = N_parcelle - 2
            End If
        Next

    End Sub

    ''' <summary>
    ''' Module de calcul passant par un point
    ''' </summary>
    ''' <param name="T_surf">Tableau de surface</param>
    ''' <param name="meth">Méthode de division (par ou point fixe)</param>
    ''' <param name="orient">Orientation de la division</param>
    ''' <param name="N_parcelle">Nbr de parcelle à créer</param>
    ''' <param name="poly_init">Polyligne fermée</param>
    ''' <param name="ptPic1">Point 1 de la direction</param>
    ''' <param name="ptPic2">Point 2 de la direction</param>
    ''' <remarks></remarks>
    Public Sub Calcul_pt(ByVal meth As String, ByVal orient As Boolean, ByVal N_parcelle As String, ByVal poly_init As Polyline, ByVal ptPic1 As Point2d, ByVal ptPic2 As Point2d, ByVal T_surf() As Double)

        Dim k As Double = 0 'Variable de la boucle "For 0 to N_parcelle - 1"
        Dim poly As Polyline = poly_init.Clone()

        For k = 0 To N_parcelle - 2 'Pour chaque parcelle à créer : 

            Draw(poly)

            '' Déclaration des variables
            Dim S_recherche As Double = 0  'Surface provisoire calculée
            Dim S_a_creer As Double = 0 'Surface à créer dans le triangle de travail
            Dim j As Integer = 0 'Variable d'incrémentation pour l'ID du sommet
            Dim j_p As Integer = 0
            Dim n As Integer = 1 'Variable qui sera égale à 1 ou 2 suivant que la parcelle sera parcourue dans le sens positif ou négatif
            Dim b As Integer = 1 'ID du point de base de la division
            Dim pt_1 As Integer  'ID du 1er point du triangle de travail
            Dim pt_2 As Integer  'ID du 2ème point du triangle de travail
            Dim pt_1_p As Integer  'ID du 1er point du précédent triangle de travail
            Dim pt_2_p As Integer  'ID du 2ème point du précédent triangle de travail
            Dim T_parcelle(2, 1) As Double  'Tableau de point qui contient les points du triangle de travail
            Dim pt_nouv As Point2d 'Nouveau point issu du calcul d'intersection
            Dim T_parcelle_X As Collection = New Collection 'Collection de coordonnées 
            Dim T_parcelle_Y As Collection = New Collection ' Collection de coordonnées
            Dim id_next As Double 'Variable de contrôle
            Dim id_prev As Double 'Variable de contrôle
            Dim Calcul_ok_p As Boolean = True 'Variable de contrôle
            Dim id_prov_1 As Double 'Variable de contrôle
            Dim id_prov_2 As Double 'Variable de contrôle
            Dim line2 As Line2d 'Ligne de contrôle d'intersection
            Dim line1 As Line2d 'Ligne de contrôle d'intersection
            Dim ed As Editor = Application.DocumentManager.MdiActiveDocument.Editor 'Message pour utilisateur


            ''Calcul de l'ID du point de base 
            For i As Integer = 0 To poly.NumberOfVertices - 1
                If poly.GetPoint2dAt(i).X = ptPic1.X And poly.GetPoint2dAt(i).Y = ptPic1.Y Then
                    b = i
                End If
            Next


            ''Initialisation de la valeur de b suivant l'orientation de la division
            If orient Then
                n = 2
            End If



            Dim gis_min, gis_max, gis_moy As Double 'Gis du point moy, gis max et gis min 
            Dim point_moy As New Point3d 'milieu du segment (b-1,b+1)
            Dim point_moy_2d As New Point2d 'milieu du segment (b-1,b+1)
            Dim interieur As Boolean 'défini si le point moyen est à l
            Dim est_compris_entre As Boolean = True 'défini si la fenetre de possibilité est valeur mini --> valeur maxi ou valeur maxi --> 400 = 0 --> valeur mini  

            ''Calcul des gisements max et min et du point moyen 
            If b = poly.NumberOfVertices - 1 Then
                gis_min = gisment(poly.GetPoint2dAt(b), poly.GetPoint2dAt(b - 1))
                gis_max = gisment(poly.GetPoint2dAt(b), poly.GetPoint2dAt(0))
                point_moy = New Point3d((poly.GetPoint2dAt(b - 1).X + poly.GetPoint2dAt(0).X) / 2, (poly.GetPoint2dAt(b - 1).Y + poly.GetPoint2dAt(0).Y) / 2, 0)
                id_next = 0
                id_prev = b - 1
            ElseIf b = 0 Then
                gis_min = gisment(poly.GetPoint2dAt(b), poly.GetPoint2dAt(poly.NumberOfVertices - 1))
                gis_max = gisment(poly.GetPoint2dAt(b), poly.GetPoint2dAt(b + 1))
                point_moy = New Point3d((poly.GetPoint2dAt(b + 1).X + poly.GetPoint2dAt(poly.NumberOfVertices - 1).X) / 2, (poly.GetPoint2dAt(b + 1).Y + poly.GetPoint2dAt(poly.NumberOfVertices - 1).Y) / 2, 0)
                id_next = 1
                id_prev = poly.NumberOfVertices - 1
            Else
                gis_min = gisment(poly.GetPoint2dAt(b), poly.GetPoint2dAt(b - 1))
                gis_max = gisment(poly.GetPoint2dAt(b), poly.GetPoint2dAt(b + 1))
                point_moy = New Point3d((poly.GetPoint2dAt(b - 1).X + poly.GetPoint2dAt(b + 1).X) / 2, (poly.GetPoint2dAt(b - 1).Y + poly.GetPoint2dAt(b + 1).Y) / 2, 0)
                id_next = b + 1
                id_prev = b - 1
            End If

            point_moy_2d = New Point2d(point_moy.X, point_moy.Y)
            gis_moy = gisment(poly.GetPoint2dAt(b), point_moy_2d)

            'Si les gisements sont < 0 alors + 400 gon 
            If gis_max < 0 Then
                gis_max = gis_max + 2 * Math.PI ' 3.141
            End If
            If gis_min < 0 Then
                gis_min = gis_min + 2 * Math.PI '3.141
            End If
            If gis_moy < 0 Then
                gis_moy = gis_moy + 2 * Math.PI '3.141
            End If

            'Si le gisement est supérieur au gisement max, on inverse 
            If gis_min > gis_max Then
                Dim gis_temp As Double = gis_max
                gis_max = gis_min
                gis_min = gis_temp
            End If

            'Determiner si le point_moy est dans le polygone ou à l'extérieur
            interieur = PtIsInsideLWP(point_moy, poly.ObjectId)




            ' Determine si la fenetre adaptée de gisement se trouve entre gismin et gismax ou entre gismax et 0 puis 0 et gismin
            If interieur Then
                If (gis_moy > gis_max And gis_moy > gis_min) Or (gis_moy < gis_max And gis_moy < gis_min) Then
                    est_compris_entre = False 'non compris 
                Else : est_compris_entre = True 'est compris
                End If
            Else
                If (gis_moy > gis_max And gis_moy > gis_min) Or (gis_moy < gis_max And gis_moy < gis_min) Then
                    est_compris_entre = True ' est compris 
                Else : est_compris_entre = False ' non compris
                End If
            End If









            'intialisation des ID de points
            pt_1 = b
            pt_2 = b + (-1) ^ n

            'pt1_p prend la 1ère valeur que prendra pt_1 pour le calcul 
            pt_1_p = pt_2
            pt_2_p = pt_2

            If pt_2_p < 0 Then
                pt_2_p = poly.NumberOfVertices.ToString - 1
            ElseIf pt_2_p > poly.NumberOfVertices.ToString - 1 Then
                pt_2_p = 0
            End If

            If poly.NumberOfVertices.ToString <> 3 Then
                'Tant que la surface calculée est insufisante 
                While S_recherche <= T_surf(k) And j < poly.NumberOfVertices - 2

                    Dim Calcul_ok As New Boolean 'définit si les calcul est possible ou non 

                    pt_1 = pt_1 + (-1) ^ n
                    pt_2 = pt_2 + (-1) ^ n

                    ''Contrôle et correction des valeurs des ID des sommets du triangle de travail
                    If pt_1 < 0 And pt_2 < 0 Then
                        If orient Then
                            pt_1 = poly.NumberOfVertices - 2
                            pt_2 = poly.NumberOfVertices - 1
                        Else
                            pt_1 = poly.NumberOfVertices - 1
                            pt_2 = poly.NumberOfVertices - 2
                        End If
                    ElseIf pt_1 > poly.NumberOfVertices - 1 And pt_2 > poly.NumberOfVertices - 1 Then
                        pt_1 = 0
                        pt_2 = 1
                    ElseIf pt_1 < 0 Then
                        pt_1 = poly.NumberOfVertices - 1
                    ElseIf pt_1 > poly.NumberOfVertices - 1 Then
                        pt_1 = 0
                    End If

                    If pt_2 < 0 Then
                        pt_2 = poly.NumberOfVertices - 1
                    ElseIf pt_2 > poly.NumberOfVertices - 1 Then
                        pt_2 = 0
                    End If

                 

                    '' Contrôles si le gisement du point pt_2 est dans l'intervale adapté  
                    Dim gis_b_pt2 As Double = gisment(poly.GetPoint2dAt(b), poly.GetPoint2dAt(pt_2)) 'gisement entre le point de base et le point courant 2 

                    If gis_b_pt2 < 0 Then
                        gis_b_pt2 = gis_b_pt2 + 2 * Math.PI '3.1415926535897931
                    End If

                    If est_compris_entre Then
                        If gis_min < gis_b_pt2 And gis_b_pt2 < gis_max Then
                            Calcul_ok = True 'calcul ok 
                        Else : Calcul_ok = False 'calcul non ok 
                        End If
                    Else
                        If gis_max < gis_b_pt2 Or gis_b_pt2 < gis_min Then
                            Calcul_ok = True 'calcul ok 
                        Else : Calcul_ok = False 'calcul non ok 
                        End If
                    End If


                    'Controle si la droite (point de base, sommet courant) intersecte le polygone ou non 
                    line2 = New Line2d(poly.GetPoint2dAt(b), poly.GetPoint2dAt(pt_2))

                    For pt_id = 0 To poly.NumberOfVertices - 1

                        If pt_id <> poly.NumberOfVertices - 1 Then
                            line1 = New Line2d(poly.GetPoint2dAt(pt_id), poly.GetPoint2dAt(pt_id + 1))
                            id_prov_1 = pt_id
                            id_prov_2 = pt_id + 1
                        Else
                            line1 = New Line2d(poly.GetPoint2dAt(pt_id), poly.GetPoint2dAt(0))
                            id_prov_1 = pt_id
                            id_prov_2 = 0
                        End If



                        Try
                            For n_inter = 0 To line2.IntersectWith(line1).Count - 1
                                Dim point_inter As Point2d = line2.IntersectWith(line1)(n_inter)
                                If id_prov_1 <> pt_2 And id_prov_1 <> pt_1 And id_prov_2 <> pt_1 And id_prov_2 <> pt_2 Then
                                    If point_inter.X <> poly.GetPoint2dAt(b).X And point_inter.Y <> poly.GetPoint2dAt(b).Y Then
                                        If 1.00001 * dist(poly.GetPoint2dAt(id_prov_1), poly.GetPoint2dAt(id_prov_2)) > _
                                            dist(poly.GetPoint2dAt(id_prov_1), point_inter) + dist(poly.GetPoint2dAt(id_prov_2), point_inter) Then
                                            If 1.001 * dist(poly.GetPoint2dAt(b), poly.GetPoint2dAt(pt_2)) > _
                                                dist(poly.GetPoint2dAt(b), point_inter) + dist(poly.GetPoint2dAt(pt_2), point_inter) Then
                                                Calcul_ok = False
                                            End If
                                        End If
                                    End If
                                End If
                            Next
                        Catch
                        End Try
                    Next


                    '' Si le gisement est dans l'intervale adapté
                    If Calcul_ok Then

                        ''Création du tableau de point 
                        T_parcelle_X = New Collection
                        T_parcelle_Y = New Collection

                        ''Remplissage du tableau de point
                        For i = 0 To j + 2
                            If b + i * (-1) ^ n < 0 Then
                                T_parcelle_X.Add(poly.GetPoint2dAt((b + i * (-1) ^ n) + (poly.NumberOfVertices)).X)
                                T_parcelle_Y.Add(poly.GetPoint2dAt((b + i * (-1) ^ n) + (poly.NumberOfVertices)).Y)
                            ElseIf b + i * (-1) ^ n > poly.NumberOfVertices - 1 Then
                                T_parcelle_X.Add(poly.GetPoint2dAt((b + i * (-1) ^ n) - (poly.NumberOfVertices)).X)
                                T_parcelle_Y.Add(poly.GetPoint2dAt((b + i * (-1) ^ n) - (poly.NumberOfVertices)).Y)
                            Else
                                T_parcelle_X.Add(poly.GetPoint2dAt(b + i * (-1) ^ n).X)
                                T_parcelle_Y.Add(poly.GetPoint2dAt(b + i * (-1) ^ n).Y)
                            End If
                        Next

                        ReDim T_parcelle(T_parcelle_X.Count - 1, 1)

                        For i = 0 To T_parcelle_X.Count - 1
                            T_parcelle(i, 0) = T_parcelle_X(i + 1)
                            T_parcelle(i, 1) = T_parcelle_Y(i + 1)
                        Next

                        ' If Not (Calcul_ok_p = False And Surface_Parcelle(T_parcelle) > T_surf(k)) Then
                        ''Calcul de la surface de recherche 
                        S_recherche = Surface_Parcelle(T_parcelle)
                        ''Calcul de la valeur de l'ID du pt_2 du triangle de travail
                        pt_2_p = pt_2
                        j_p = j
                        'End If
                    End If
                    Calcul_ok_p = Calcul_ok
                    j = j + 1

                End While
                'Fin - Tant que la surface calculée est insufisante 

            Else
                'Si la surface à diviser est un triangle 
                j = 1
                S_recherche = poly.Area
                pt_2 = 2
                pt_1 = 1
            End If


            'Si la S_recherche est trop petite on teste avec le dernier point qui était ok 
            If S_recherche < T_surf(k) Then

                pt_1 = pt_2_p
                pt_2 = pt_1 + (-1) ^ n

                If pt_2 < 0 Then
                    pt_2 = poly.NumberOfVertices - 1
                ElseIf pt_2 > poly.NumberOfVertices - 1 Then
                    pt_2 = 0
                End If

                ReDim T_parcelle(2, 1)

                T_parcelle(0, 0) = poly.GetPoint2dAt(b).X
                T_parcelle(0, 1) = poly.GetPoint2dAt(b).Y

                T_parcelle(1, 0) = poly.GetPoint2dAt(pt_1).X
                T_parcelle(1, 1) = poly.GetPoint2dAt(pt_1).Y
                T_parcelle(2, 0) = poly.GetPoint2dAt(pt_2).X
                T_parcelle(2, 1) = poly.GetPoint2dAt(pt_2).Y

                S_recherche = S_recherche + Surface_Parcelle(T_parcelle)
                j = j_p + 2
            End If



            If S_recherche > T_surf(k) Then

                ''Calcul de la surface à créer dans le triangle 
                S_a_creer = S_recherche - T_surf(k)




                ''Début - Calcul des points d'intersection
                Dim d_1 As Double = dist(poly.GetPoint2dAt(b), poly.GetPoint2dAt(pt_1)) 'Distance entre le point de base et le 1er point de la base du triangle
                Dim d_2 As Double = dist(poly.GetPoint2dAt(b), poly.GetPoint2dAt(pt_2)) 'Distance entre le point de base et le 2ème point de la base du triangle
                Dim v_2_1 As Vector2d = poly.GetPoint2dAt(pt_2).GetVectorTo(poly.GetPoint2dAt(pt_1)) 'Calcul de l'ange 1 - 2 - b 
                Dim v_2_b As Vector2d = poly.GetPoint2dAt(pt_2).GetVectorTo(poly.GetPoint2dAt(b)) 'Calcul de l'angle 1 - 2 - b 
                Dim a_2 As Double = Abs(v_2_1.Angle - v_2_b.Angle) 'Calcul de l'angle 2 - 1 - b 
                Dim d_x As Double = Abs(S_a_creer / (0.5 * d_2 * Sin(a_2))) 'Calcul de la distance entre pt_1 et le point nouveau
                Dim gis_2_1 As Double = gisment(poly.GetPoint2dAt(pt_2), poly.GetPoint2dAt(pt_1)) 'Calcul du gisement pt_1 vers pt_2
                pt_nouv = pt_lance(poly.GetPoint2dAt(pt_2), d_x, gis_2_1) 'Calcul du point nouveau
                ''Fin - Calcul des points d'intersection





                'Contrôle si le pt_nouv et compris entre pt_1 et pt_2
                'Si non alors gis_1_2 = gis_1_2 + 200 et recalcul
                If 1.00000001 * dist(poly.GetPoint2dAt(pt_1), poly.GetPoint2dAt(pt_2)) < (dist(poly.GetPoint2dAt(pt_1), pt_nouv) + dist(poly.GetPoint2dAt(pt_2), pt_nouv)) Then
                    gis_2_1 = gis_2_1 + Math.PI '3.1415926535897931
                    pt_nouv = pt_lance(poly.GetPoint2dAt(pt_2), d_x, gis_2_1) 'Calcul du point nouveau
                End If




                'Calcul de la parcelle calculée 
                T_parcelle_X = New Collection
                T_parcelle_Y = New Collection

                'Construction des collections de points 
                For i = 0 To j
                    If b + i * (-1) ^ n < 0 Then
                        T_parcelle_X.Add(poly.GetPoint2dAt((b + i * (-1) ^ n) + (poly.NumberOfVertices)).X)
                        T_parcelle_Y.Add(poly.GetPoint2dAt((b + i * (-1) ^ n) + (poly.NumberOfVertices)).Y)
                    ElseIf b + i * (-1) ^ n > poly.NumberOfVertices - 1 Then
                        T_parcelle_X.Add(poly.GetPoint2dAt((b + i * (-1) ^ n) - (poly.NumberOfVertices)).X)
                        T_parcelle_Y.Add(poly.GetPoint2dAt((b + i * (-1) ^ n) - (poly.NumberOfVertices)).Y)
                    Else
                        T_parcelle_X.Add(poly.GetPoint2dAt(b + i * (-1) ^ n).X)
                        T_parcelle_Y.Add(poly.GetPoint2dAt(b + i * (-1) ^ n).Y)
                    End If
                Next


                T_parcelle_X.Add(pt_nouv.X)
                T_parcelle_Y.Add(pt_nouv.Y)


                ReDim T_parcelle(T_parcelle_X.Count - 1, 1)

                For i = 0 To T_parcelle_X.Count - 1
                    T_parcelle(i, 0) = T_parcelle_X(i + 1)
                    T_parcelle(i, 1) = T_parcelle_Y(i + 1)
                Next

                ' Fin calcul de la parcelle calculée 

                ' Si la surface est fausse alors on recalcule la distance du point lancée
                If Abs(T_surf(k) - Surface_Parcelle(T_parcelle)) > 0.01 Then
                    d_x = dist(poly.GetPoint2dAt(pt_1), poly.GetPoint2dAt(pt_2)) - d_x
                    pt_nouv = pt_lance(poly.GetPoint2dAt(pt_2), d_x, gis_2_1)
                    T_parcelle((T_parcelle.Length / 2) - 1, 0) = pt_nouv.X
                    T_parcelle((T_parcelle.Length / 2) - 1, 1) = pt_nouv.Y
                End If


                'Vérification que la nouvelle parcelle est vraiment comprise dans la parcelle initiale
                Dim dessin_ok As Boolean = True

                line2 = New Line2d(poly.GetPoint2dAt(b), pt_nouv)
                For pt_id = 0 To poly.NumberOfVertices - 1

                    If pt_id <> poly.NumberOfVertices - 1 Then
                        line1 = New Line2d(poly.GetPoint2dAt(pt_id), poly.GetPoint2dAt(pt_id + 1))
                        id_prov_1 = pt_id
                        id_prov_2 = pt_id + 1
                    Else
                        line1 = New Line2d(poly.GetPoint2dAt(pt_id), poly.GetPoint2dAt(0))
                        id_prov_1 = pt_id
                        id_prov_2 = 0
                    End If

                    Try

                        'Verification qu'il existe une intersection entre les deux segments 
                        For n_inter = 0 To line2.IntersectWith(line1).Count - 1

                            Dim point_inter As Point2d = line2.IntersectWith(line1)(n_inter)

                            If id_prov_1 <> pt_2 And id_prov_1 <> pt_1 And id_prov_2 <> pt_1 And id_prov_2 <> pt_2 Then
                                If point_inter.X <> poly.GetPoint2dAt(b).X And point_inter.Y <> poly.GetPoint2dAt(b).Y Then
                                    If 1.00001 * dist(poly.GetPoint2dAt(id_prov_1), poly.GetPoint2dAt(id_prov_2)) > _
                                        dist(poly.GetPoint2dAt(id_prov_1), point_inter) + dist(poly.GetPoint2dAt(id_prov_2), point_inter) Then
                                        If 1.001 * dist(poly.GetPoint2dAt(b), poly.GetPoint2dAt(pt_2)) > _
                                            dist(poly.GetPoint2dAt(b), point_inter) + dist(poly.GetPoint2dAt(pt_2), point_inter) Then
                                            dessin_ok = False
                                            k = N_parcelle - 2
                                            ed.WriteMessage(vbLf & "ligne d'intersection est mauvaise")
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    Catch
                    End Try
                Next
                'Fin - Vérification que la nouvelle parcelle est vraiment comprise dans la parcelle initiale


                If dessin_ok Then

                    ' Dessin de la parcelle calculée 
                    T_parcelle_id(k) = Draw(Parcelle(T_parcelle))
                    T_parcelle_surf(k) = Surface_Parcelle(T_parcelle)

                    'Création de la parcelle suivante
                    T_parcelle_X = New Collection
                    T_parcelle_Y = New Collection

                    'Construction des collection de points
                    For i = 0 To poly.NumberOfVertices - 1 - j
                        If b + i * (-1) ^ (n + 1) < 0 Then
                            T_parcelle_X.Add(poly.GetPoint2dAt((b + i * (-1) ^ (n + 1)) + (poly.NumberOfVertices)).X)
                            T_parcelle_Y.Add(poly.GetPoint2dAt((b + i * (-1) ^ (n + 1)) + (poly.NumberOfVertices)).Y)
                        ElseIf b + i * (-1) ^ (n + 1) > poly.NumberOfVertices - 1 Then
                            T_parcelle_X.Add(poly.GetPoint2dAt((b + i * (-1) ^ (n + 1)) - (poly.NumberOfVertices)).X)
                            T_parcelle_Y.Add(poly.GetPoint2dAt((b + i * (-1) ^ (n + 1)) - (poly.NumberOfVertices)).Y)
                        Else
                            T_parcelle_X.Add(poly.GetPoint2dAt(b + i * (-1) ^ (n + 1)).X)
                            T_parcelle_Y.Add(poly.GetPoint2dAt(b + i * (-1) ^ (n + 1)).Y)
                        End If
                    Next



                    T_parcelle_X.Add(pt_nouv.X)
                    T_parcelle_Y.Add(pt_nouv.Y)



                    ReDim T_parcelle(T_parcelle_X.Count - 1, 1)

                    For i = 0 To T_parcelle_X.Count - 1
                        T_parcelle(i, 0) = T_parcelle_X(i + 1)
                        T_parcelle(i, 1) = T_parcelle_Y(i + 1)
                    Next

                    'Fin - Dessin de la parcelle suivante




                    EraseAc(poly.ObjectId)
                    poly = Parcelle(T_parcelle)
                    orient = False

                    'Si la parcelle suivante est la dernière parcelle on la dessine 
                    If k = N_parcelle - 2 Then
                        T_parcelle_id(k + 1) = Draw(Parcelle(T_parcelle))
                        T_parcelle_surf(k + 1) = Surface_Parcelle(T_parcelle)
                    End If


                End If
            Else
                ed.WriteMessage(vbLf & "Pas possible de créer " & N_parcelle & " parcelles avec les surfaces choisies")
                Exit For
            End If
        Next




    End Sub

    ''' <summary>
    ''' Crée les valeurs du tableau des surfaces à calculer 
    ''' </summary>
    ''' <param name="meth_surf">Méthode de calcul de surface : égale ou choisie</param>
    ''' <param name="N_parcelle">Nbr de parcelle à créer</param>
    ''' <param name="poly">Polygone à diviser</param>
    ''' <param name="T_Surf">Valeur de la surface à conserver ou non</param>
    Function Calcul_surf(ByVal meth_surf As String, ByVal poly As Polyline, ByVal N_parcelle As Integer, ByVal T_Surf() As Double) As Double()
        If meth_surf = "egale" Then
            For i As Integer = 0 To N_parcelle - 1
                T_Surf(i) = CDbl(poly.Area.ToString) / N_parcelle
            Next
            Return T_Surf
        Else
            Return T_Surf
        End If

    End Function


    ''' <summary>
    ''' Renvoi le 1er point d'une collection
    ''' </summary>
    ''' <param name="collection_inter"></param>
    Function pt_Intersect(ByVal collection_inter As Point3dCollection) As Point2d
        Dim X As Double
        Dim Y As Double

        X = collection_inter.Item(0).X
        Y = collection_inter.Item(0).Y

        Dim pt_Int As Point2d = New Point2d(X, Y)
        Return pt_Int
    End Function

    ''' <summary>
    ''' Crée une parcelle à partir d'un tableau de point 
    ''' </summary>
    ''' <param name="T_point">Tableau de points</param>
    Function Parcelle(ByVal T_point(,) As Double) As Polyline
        'Crée une parcelle à partir d'un tableau de coordonnées 
        Dim acPoly As Polyline = New Polyline()
        acPoly.SetDatabaseDefaults()
        If T_point.Length > 0 Then
            For i = 0 To (T_point.Length / 2) - 1
                acPoly.AddVertexAt(i, New Point2d(T_point(i, 0), T_point(i, 1)), 0, 0, 0)
            Next
        End If

        Return acPoly

    End Function

    ''' <summary>
    ''' Calcul de surface à partir de tableau de points 
    ''' </summary>
    ''' <param name="T_point">Tableau de points</param>
    Function Surface_Parcelle(ByVal T_point(,) As Double) As Double
        ' Calcul de la surface de la surface à partir d'un tableau de coordonnées
        Dim S As Double = CDbl(Parcelle(T_point).Area.ToString)
        Return S
    End Function

    ''' <summary>
    ''' Dessin d'élt Acad 
    ''' </summary>
    ''' <param name="acItem">Eléet à dessiner</param>
    ''' <param name="fermer_poly">Paramètre si la polyligne doit être fermée ou non</param>
    Public Function Draw(ByVal acItem As Global.Autodesk.AutoCAD.DatabaseServices.Entity, Optional ByRef fermer_poly As Boolean = True) As ObjectId
        'Dessine toute entité AutoCad 

        '' Get the current document and database
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Database = acDoc.Database
        Dim idobject As ObjectId
        '' Start a transaction
        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()


            ' Open the Block table for read
            Dim acBlkTbl As BlockTable
            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

            If fermer_poly And TypeName(acItem) = "Polyline" Then
                Dim acline_p As Polyline = acItem
                acline_p.Closed = True
                acItem = acline_p
            End If
            '' Open the Block table record Model space for write


            Dim acBlkTblRec As BlockTableRecord
            acBlkTblRec = acTrans.GetObject(acBlkTbl(BlockTableRecord.ModelSpace), _
                                            OpenMode.ForWrite)


            '' Add the new object to the block table record and the transaction

            acBlkTblRec.AppendEntity(acItem)
            acTrans.AddNewlyCreatedDBObject(acItem, True)

            idobject = acItem.ObjectId
            '' Save the new object to the database
            acTrans.Commit()

        End Using
        Return idobject
    End Function

   

    ''' <summary>
    ''' Calcul de gisement
    ''' </summary>
    ''' <param name="pt1">Point de base</param>
    ''' <param name="pt2">Point secondaire</param>
    Function gisment(ByVal pt1 As Point2d, ByVal pt2 As Point2d) As Double
        Dim gis As Double
        If pt1.X = pt2.X Then
            If pt1.Y >= pt2.Y Then
                gis = Math.PI ' 3.1415926535897931
            Else
                gis = 0
            End If
        ElseIf pt1.Y = pt2.Y Then
            If pt1.X >= pt2.X Then
                gis = Math.PI / 2 ' 3.1415926535897931 / 2
            Else
                gis = 1.5 * Math.PI  ' 3.1415926535897931
            End If
        Else
            gis = 2 * Atan((pt2.X - pt1.X) / (((pt2.X - pt1.X) ^ 2 + (pt2.Y - pt1.Y) ^ 2) ^ (1 / 2) + (pt2.Y - pt1.Y)))
        End If

        Return gis
    End Function

    ''' <summary>
    ''' Calcul de point lancé
    ''' </summary>
    ''' <param name="pt1">Point de base</param>
    ''' <param name="d">Distance</param>
    ''' <param name="g">Gisement</param>
    ''' <remarks></remarks>

    Function pt_lance(ByVal pt1 As Point2d, ByVal d As Double, ByVal g As Double) As Point2d
        Dim pt2 As Point2d = New Point2d(pt1.X + d * Sin(g), pt1.Y + d * Cos(g))
        Return pt2
    End Function

    ''' <summary>
    ''' Division de triangle
    ''' </summary>
    ''' <param name="poly">Polygone complet</param>
    ''' <param name="S_a_creer">Surface ¨à créer</param>
    ''' <param name="T_parcelle">Tableau des points du triangle</param>
    ''' <remarks></remarks>

    Function div_triangle(ByVal poly As Polyline, ByVal S_a_creer As Double, ByVal T_parcelle As Double(,)) As Double(,)
        Dim inter_1 As Point2d = New Point2d(T_parcelle(0, 0), T_parcelle(0, 1))
        Dim inter_2 As Point2d = New Point2d(T_parcelle(1, 0), T_parcelle(1, 1))
        Dim pt3 As Point2d = New Point2d(T_parcelle(2, 0), T_parcelle(2, 1))

        Dim d1 As Double = (((Surface_Parcelle(T_parcelle) - S_a_creer) * dist(inter_1, pt3) ^ 2) / Surface_Parcelle(T_parcelle)) ^ (1 / 2)
        Dim d2 As Double = (((Surface_Parcelle(T_parcelle) - S_a_creer) * dist(inter_2, pt3) ^ 2) / Surface_Parcelle(T_parcelle)) ^ (1 / 2)

        Dim g1 As Double = gisment(pt3, inter_1)
        Dim g2 As Double = gisment(pt3, inter_2)

        Dim pt_f_1 As Point2d = pt_lance(pt3, d1, g1)
        Dim pt_f_2 As Point2d = pt_lance(pt3, d2, g2)

        Dim T_retour(1, 1) As Double
        T_retour(0, 0) = pt_f_1.X
        T_retour(0, 1) = pt_f_1.Y
        T_retour(1, 0) = pt_f_2.X
        T_retour(1, 1) = pt_f_2.Y
        Return T_retour

    End Function

    ''' <summary>
    ''' Division de trapèze
    ''' </summary>
    ''' <param name="poly">Polygone complet</param>
    ''' <param name="S_parcelle">Surface calculée avec le trapèze</param>
    ''' <param name="S_parcelle_p">Surface calculée sans le trapéze</param>
    ''' <param name="T_parcelle">Tableau des points du trapèze</param>
    ''' <param name="T_surf">Surface de la parcelle à créer</param>
    ''' <remarks></remarks>

    Function div_quadrilatère(ByVal poly As Polyline, ByVal S_parcelle As Double, ByRef S_parcelle_p As Double, ByVal T_surf As Double, ByVal T_parcelle As Double(,)) As Double(,)
        Dim line_1 As Xline = New Xline
        line_1.BasePoint = New Point3d(T_parcelle(0, 0), T_parcelle(0, 1), 0)
        line_1.SecondPoint = New Point3d(T_parcelle(1, 0), T_parcelle(1, 1), 0)

        Dim line_2 As Xline = New Xline
        line_2.BasePoint = New Point3d(T_parcelle(2, 0), T_parcelle(2, 1), 0)
        line_2.SecondPoint = New Point3d(T_parcelle(3, 0), T_parcelle(3, 1), 0)

        Dim pt_sommet As Point2d
        Dim S_a_creer As Double
        Dim IntersectPts As Point3dCollection = New Point3dCollection
        line_1.IntersectWith(line_2, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)

        Dim S_trapeze As Double = Surface_Parcelle(T_parcelle)

        Dim pt_f_1 As Point2d
        Dim pt_f_2 As Point2d

        Dim d_a As Double
        Dim d_b As Double
        Dim g1 As Double
        Dim g2 As Double

        If IntersectPts.Count = 1 Then
            pt_sommet = pt_Intersect(IntersectPts)

            Dim T_triangle(2, 1) As Double
            T_triangle(0, 0) = pt_sommet.X
            T_triangle(0, 1) = pt_sommet.Y

            If dist(pt_sommet, New Point2d(T_parcelle(0, 0), T_parcelle(0, 1))) < dist(pt_sommet, New Point2d(T_parcelle(1, 0), T_parcelle(1, 1))) Then
                S_a_creer = S_trapeze - (T_surf - S_parcelle_p)
                T_triangle(1, 0) = T_parcelle(1, 0)
                T_triangle(1, 1) = T_parcelle(1, 1)
                T_triangle(2, 0) = T_parcelle(2, 0)
                T_triangle(2, 1) = T_parcelle(2, 1)

                d_a = dist(pt_sommet, New Point2d(T_parcelle(1, 0), T_parcelle(1, 1)))
                d_b = dist(pt_sommet, New Point2d(T_parcelle(2, 0), T_parcelle(2, 1)))

                g1 = gisment(pt_sommet, New Point2d(T_parcelle(1, 0), T_parcelle(1, 1)))
                g2 = gisment(pt_sommet, New Point2d(T_parcelle(2, 0), T_parcelle(2, 1)))
            Else
                S_a_creer = T_surf - S_parcelle_p
                T_triangle(1, 0) = T_parcelle(0, 0)
                T_triangle(1, 1) = T_parcelle(0, 1)
                T_triangle(2, 0) = T_parcelle(3, 0)
                T_triangle(2, 1) = T_parcelle(3, 1)

                d_a = dist(pt_sommet, New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)))
                d_b = dist(pt_sommet, New Point2d(T_parcelle(3, 0), T_parcelle(3, 1)))

                g1 = gisment(pt_sommet, New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)))
                g2 = gisment(pt_sommet, New Point2d(T_parcelle(3, 0), T_parcelle(3, 1)))
            End If

            Dim s As Double = Surface_Parcelle(T_triangle)
            Dim d1 As Double = (((Surface_Parcelle(T_triangle) - S_a_creer) * d_a ^ 2) / Surface_Parcelle(T_triangle)) ^ (1 / 2)
            Dim d2 As Double = (((Surface_Parcelle(T_triangle) - S_a_creer) * d_b ^ 2) / Surface_Parcelle(T_triangle)) ^ (1 / 2)


            pt_f_1 = pt_lance(pt_sommet, d1, g1)
            pt_f_2 = pt_lance(pt_sommet, d2, g2)

        Else
            ''on a affaire à un parallélogramme
            S_a_creer = T_surf - S_parcelle_p
            Dim rap_surf As Double = S_a_creer / Surface_Parcelle(T_parcelle)

            Dim d1 As Double = rap_surf * dist(New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)), New Point2d(T_parcelle(1, 0), T_parcelle(1, 1)))
            Dim d2 As Double = rap_surf * dist(New Point2d(T_parcelle(2, 0), T_parcelle(2, 1)), New Point2d(T_parcelle(3, 0), T_parcelle(3, 1)))

            g1 = gisment(New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)), New Point2d(T_parcelle(1, 0), T_parcelle(1, 1)))

            pt_f_1 = pt_lance(New Point2d(T_parcelle(0, 0), T_parcelle(0, 1)), d1, g1)
            pt_f_2 = pt_lance(New Point2d(T_parcelle(3, 0), T_parcelle(3, 1)), d2, g1)

        End If


        Dim T_retour(1, 1) As Double
        T_retour(0, 0) = pt_f_1.X
        T_retour(0, 1) = pt_f_1.Y
        T_retour(1, 0) = pt_f_2.X
        T_retour(1, 1) = pt_f_2.Y
        Return T_retour

    End Function

    ''' <summary>
    ''' selection_polyligne
    ''' </summary>
    ''' <param name="inter_1"></param>
    ''' <param name="inter_1_p"></param>
    ''' <param name="j"></param>
    ''' <param name="n">exposant : ^(-1)^n</param>
    ''' <param name="p1"></param>
    ''' <param name="poly">Polygone</param>
    ''' <param name="ptPic1">Point 1 de la direction</param>
    ''' <param name="ptPic2">Point 2 de la direction</param>
    ''' <param name="T_point">Tableau de point</param>
    ''' <remarks>Renvoi des coordonnées de l'intersection </remarks>
    Function calcul_intersection(ByVal poly As Polyline, ByVal ptPic1 As Point2d, ByVal ptPic2 As Point2d, ByVal inter_1 As Point2d, ByVal inter_1_p As Point2d, ByVal p1 As Integer, ByVal T_point As Double(,), ByVal j As Integer, Optional ByVal n As Integer = 2) As Collection

        Dim pt_j As Point3d = New Point3d(T_point(j, 3), T_point(j, 4), 0)
        Dim pt_j2 As Point3d = New Point3d(ptPic2.X + (T_point(j, 3) - ptPic1.X), ptPic2.Y + (T_point(j, 4) - ptPic1.Y), 0)
        Dim acLine As Xline = New Xline()
        acLine.BasePoint = pt_j
        acLine.SecondPoint = pt_j2
        Dim intersec As Boolean = False
        Dim inter_1_temp As Point2d = New Point2d(T_point(0, 3), T_point(0, 4))



        'En partant du point de base, on cherche la 1ère intersection entre un coté du polygone et la droite acLine
        While intersec = False
            ''Si le compteur n'attiens pas l'iD du sommet courant alors on cherche l'intersection
            If (p1 + (-1) ^ n) <> T_point(j, 0) Then

                ''Création de la droite passant par deux sommets du polygone 
                Dim acLine_p As Xline = New Xline()
                If n = 2 Then
                    If p1 < (CInt(poly.NumberOfVertices.ToString) - 1) Then
                        acLine_p.BasePoint = poly.GetPoint3dAt(p1)
                        acLine_p.SecondPoint = poly.GetPoint3dAt(p1 + (-1) ^ n)
                    Else
                        acLine_p.BasePoint = poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1)
                        acLine_p.SecondPoint = poly.GetPoint3dAt(0)
                    End If
                Else
                    If p1 > 0 Then
                        acLine_p.BasePoint = poly.GetPoint3dAt(p1)
                        acLine_p.SecondPoint = poly.GetPoint3dAt(p1 + (-1) ^ n)
                    Else
                        acLine_p.BasePoint = poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1)
                        acLine_p.SecondPoint = poly.GetPoint3dAt(0)
                    End If
                End If

                '' Fin - Création de la droite passant par deux sommets du polygone 

                '' Calcul de l'intersection 
                Dim IntersectPts As Point3dCollection = New Point3dCollection
                acLine.IntersectWith(acLine_p, Intersect.OnBothOperands, IntersectPts, IntPtr.Zero, IntPtr.Zero)
                '' Fin - Calcul de l'intersection 

                '' Vérifier que l'intersection est entre les deux sommets
                ''Si oui alors intersection = True
                If IntersectPts.Count = 1 Then
                    inter_1_temp = inter_1
                    inter_1 = pt_Intersect(IntersectPts)

                    If n = 2 Then
                        If p1 + (-1) ^ n < poly.NumberOfVertices.ToString Then
                            If Not (inter_1.X > poly.GetPoint3dAt(p1).X And inter_1.X > poly.GetPoint3dAt(p1 + (-1) ^ n).X) And _
                                Not (inter_1.X < poly.GetPoint3dAt(p1).X And inter_1.X < poly.GetPoint3dAt(p1 + (-1) ^ n).X) And _
                                Not (inter_1.Y > poly.GetPoint3dAt(p1).Y And inter_1.Y > poly.GetPoint3dAt(p1 + (-1) ^ n).Y) And _
                                Not (inter_1.Y < poly.GetPoint3dAt(p1).Y And inter_1.Y < poly.GetPoint3dAt(p1 + (-1) ^ n).Y) Then
                                intersec = True
                            End If
                        ElseIf Not (inter_1.X > poly.GetPoint3dAt(0).X And inter_1.X > poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).X) And _
                                Not (inter_1.X < poly.GetPoint3dAt(0).X And inter_1.X < poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).X) And _
                                Not (inter_1.Y > poly.GetPoint3dAt(0).Y And inter_1.Y > poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).Y) And _
                                Not (inter_1.Y < poly.GetPoint3dAt(0).Y And inter_1.Y < poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).Y) Then
                            intersec = True
                        End If
                    Else
                        If p1 + (-1) ^ n > -1 Then
                            If Not (inter_1.X > poly.GetPoint3dAt(p1).X And inter_1.X > poly.GetPoint3dAt(p1 + (-1) ^ n).X) And _
                                Not (inter_1.X < poly.GetPoint3dAt(p1).X And inter_1.X < poly.GetPoint3dAt(p1 + (-1) ^ n).X) And _
                                Not (inter_1.Y > poly.GetPoint3dAt(p1).Y And inter_1.Y > poly.GetPoint3dAt(p1 + (-1) ^ n).Y) And _
                                Not (inter_1.Y < poly.GetPoint3dAt(p1).Y And inter_1.Y < poly.GetPoint3dAt(p1 + (-1) ^ n).Y) Then
                                intersec = True
                            End If
                        ElseIf Not (inter_1.X > poly.GetPoint3dAt(0).X And inter_1.X > poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).X) And _
                                Not (inter_1.X < poly.GetPoint3dAt(0).X And inter_1.X < poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).X) And _
                                Not (inter_1.Y > poly.GetPoint3dAt(0).Y And inter_1.Y > poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).Y) And _
                                Not (inter_1.Y < poly.GetPoint3dAt(0).Y And inter_1.Y < poly.GetPoint3dAt(CInt(poly.NumberOfVertices.ToString) - 1).Y) Then
                            intersec = True
                        End If
                    End If
                End If
                ''Fin - Vérifier que l'intersection est entre les deux sommets

                ''Si intersect = false : incrémentation du compteur
                If intersec = False Then
                    p1 = p1 + (-1) ^ n
                    inter_1 = inter_1_temp
                Else
                    inter_1_p = inter_1_temp
                End If

                If n = 2 And p1 > poly.NumberOfVertices.ToString - 1 Then
                    p1 = 0
                ElseIf n = 1 And p1 < 1 Then
                    p1 = poly.NumberOfVertices.ToString - 1
                End If

            ElseIf p1 + (-1) ^ n = T_point(j, 0) Then
                inter_1_p = inter_1
                inter_1 = New Point2d(pt_j.X, pt_j.Y)
                intersec = True
            End If
        End While
        '' Fin - recherche de la 1ère intersection "coté matricule ascendant"

        Dim resultat As New Collection
        resultat.Add(inter_1_p)
        resultat.Add(inter_1)
        resultat.Add(p1)

        Return resultat

    End Function

    ''' <summary>
    ''' selection_polyligne
    ''' </summary>
    ''' <param name="controle_surf">Active le contrôle de la surface (diff de 0 ou non)</param>
    ''' <remarks>Permet la sélection d'un polygone</remarks>
    Public Function selection_polyligne(Optional ByVal controle_surf As Boolean = True) As Collection

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

            acOPrompt.MessageForAdding = "Sélectionner le périmètre de la parcelle (poyligne périmètrique)"
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
                                        collection_retour = selection_polyligne(controle_surf)
                                    End If
                                Else
                                    collection_retour.Add(acEnt)
                                End If

                                Exit For
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
                    collection_retour = selection_polyligne(controle_surf)
                End If
            End If
            ''End of the selection of the polylign 
        End Using
        Return collection_retour
    End Function

    ''' <summary>
    ''' EraseAc
    ''' </summary>
    ''' <param name="oid">Titel</param>
    ''' <remarks>Supprime un élément Acad</remarks>
    Function EraseAc(ByVal oId As ObjectId)
        'Supprime 
        Try
            '' Get the current document and database
            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
            Dim acCurDb As Database = acDoc.Database

            '' Start a transaction
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()

                Dim acItem As Entity = DirectCast(acTrans.GetObject(oId, OpenMode.ForWrite), Entity)
                If Not IsDBNull(acItem) Then
                    acItem.Erase()
                End If

                '' Save the new object to the database
                acTrans.Commit()
            End Using
            Return True
        Catch
            Return False
        End Try

    End Function

    ''' <summary>
    ''' Controle des surfaces crées 
    ''' </summary>
    ''' <param name="T_surf">Tableau de surface</param>
    ''' <param name="meth">Méthode de division (par ou point fixe)</param>
    ''' <param name="orient">Orientation de la division</param>
    ''' <param name="N_parcelle">Nbr de parcelle à créer</param>
    ''' <param name="Poly">Polyligne fermée</param>
    ''' <param name="ptPic1_2d">Point 1 de la direction</param>
    ''' <param name="ptPic2_2d">Point 2 de la direction</param>
    ''' <remarks>Contrôle si le calcul est juste</remarks>
    Sub ControleSurf(ByVal T_surf() As Double, ByVal meth As String, ByVal orient As Boolean, ByVal N_parcelle As Integer, ByVal Poly As Polyline, ByVal ptPic1_2d As Point2d, ByVal ptPic2_2d As Point2d)
        Dim nbrErreur As Double = 0
        Dim connect As New Revo.connect

        ' connect.Message("Calcul de surface", "Contrôle des surfaces en cours ...", False, 80, 100)

        'Vérifie si toute les surfaces ont été créees 
        If T_surf.Length <> T_parcelle_surf.Length Then
            MsgBox(T_surf.Length - T_parcelle_surf.Length & " parcelles non créées")
        Else

            Try
                'Pour chaque parcelle on vérifie si la surface est juste
                'Sinon on la colore en rose moche 
                For k = 0 To T_surf.Length - 1
                    If Round(T_surf(k), 5) <> Round(CDbl(T_parcelle_surf(k)), 5) Then
                        nbrErreur = nbrErreur + 1

                        '' Get the current document and database
                        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                        Dim acCurDb As Database = acDoc.Database

                        ' Start a transaction
                        Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                            Dim idOjt As ObjectId = T_parcelle_id(k)
                            Dim polyEr As Polyline = DirectCast(acTrans.GetObject(idOjt, OpenMode.ForWrite),  _
                                Polyline)
                            polyEr.ColorIndex = 21

                            acTrans.Commit()
                        End Using

                    End If
                Next

                'Information de l'utilisateur 
                If nbrErreur = 0 Then
                    connect.Message("Calcul de surface", "Calcul terminé", False, 100, 100)
                    connect.Message("", "", True, 100, 100, "info")
                Else

                    ''Si la méthode est "par" on recommence le calcul dans l'autre sens et on re-contrôle 
                    If meth = "par" Then

                        For k = 0 To T_surf.Length - 1
                            '' Get the current document and database
                            Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                            Dim acCurDb As Database = acDoc.Database
                            ' Start a transaction
                            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                Dim idOjt As ObjectId = T_parcelle_id(k)
                                Dim polyEr As Polyline = DirectCast(acTrans.GetObject(idOjt, OpenMode.ForWrite), Polyline)
                                polyEr.Erase()
                                acTrans.Commit()
                            End Using
                        Next


                        If orient = False Then
                            orient = True
                        Else : orient = False
                        End If

                        Dim T_surf_inverse(T_surf.Length - 1) As Double
                        Dim sommeSurf As Double = 0
                        Dim T_surf0 As Double = T_surf(0)

                        If T_surf.Length > 1 Then
                            For i = 0 To T_surf.Length - 2
                                T_surf_inverse(i + 1) = T_surf(T_surf.Length - 2 - i)
                                sommeSurf = sommeSurf + T_surf(T_surf.Length - 2 - i)
                            Next
                        End If

                        T_surf_inverse(0) = Poly.Area - sommeSurf
                        Calcul_par(meth, orient, N_parcelle, Poly, ptPic1_2d, ptPic2_2d, T_surf_inverse)



                        nbrErreur = 0
                        If T_surf.Length <> T_parcelle_surf.Length Then
                            MsgBox(T_surf.Length - T_parcelle_surf.Length & " parcelles non créées")
                        Else
                            For k = 0 To T_surf.Length - 1
                                If Round(T_surf_inverse(k), 5) <> Round(CDbl(T_parcelle_surf(k)), 5) Then
                                    nbrErreur = nbrErreur + 1

                                    '' Get the current document and database
                                    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                                    Dim acCurDb As Database = acDoc.Database

                                    ' Start a transaction
                                    Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                                        Dim idOjt As ObjectId = T_parcelle_id(k)
                                        Dim polyEr As Polyline = DirectCast(acTrans.GetObject(idOjt, OpenMode.ForWrite),  _
                                            Polyline)
                                        polyEr.ColorIndex = 21

                                        acTrans.Commit()
                                    End Using

                                End If
                            Next

                            If nbrErreur = 0 Then
                                connect.Message("Calcul de surface", "Calcul terminé", False, 100, 100, "info")
                                connect.Message("", "", True, 100, 100, "info")
                            Else
                                MsgBox(nbrErreur & " surfaces fausses... Merci de nous signaler le problème")
                            End If
                        End If

                    Else
                        MsgBox(nbrErreur & " surfaces fausses... Merci de nous signaler le problème")
                    End If
                End If
            Catch
            End Try
        End If
    End Sub



End Module
