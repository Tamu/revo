Option Strict Off
Option Explicit On

Imports Autodesk.AutoCAD.DatabaseServices

Friend Class frmPointsFiables
    Inherits System.Windows.Forms.Form

    Private Sub btnCancel_Click(sender As System.Object, e As System.EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub


    Private Sub cmdDeleteGrid_Click(sender As System.Object, e As System.EventArgs) Handles cmdDeleteGrid.Click

        'On efface les objets du quadrillage existants
        If MsgBox("Voulez-vous effacer définitivement tous les symboles de fiabilité planimétrique existants ?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Réinitialisation des symboles existants") = MsgBoxResult.Yes Then
            Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
            DeleteAllInLayer("MO_VALEUR_PTS")
            Me.Cursor = System.Windows.Forms.Cursors.Default
        End If

    End Sub

    Private Sub cmdNext_Click(sender As System.Object, e As System.EventArgs) Handles cmdNext.Click

        Dim Connect As New Revo.connect

        'Au moins une couche doit être sélectionnée
        Dim lngObjCount As Double = 0
        ' Dim booDelete As Boolean
        Dim repmsg As MsgBoxResult
        Dim objSelSet As Collection 'As Autodesk.AutoCAD.Interop.AcadSelectionSet
        'Dim objAcadBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference
        ' Dim varAttributes As Object
        Dim booAccessible As Boolean

        'Attributs géométriques
        Dim dblPt As Autodesk.AutoCAD.Geometry.Point3d = Nothing
        ' Dim dblOri As Double
        Dim dblScale As Double
        Dim layID As Autodesk.AutoCAD.DatabaseServices.ObjectId = GetLayerIDFromLayerName("MO_VALEUR_PTS")

        
        If chkPF.Checked = True Or Me.chkCS.Checked = True Or Me.chkOD.Checked = True Or Me.chkBF.Checked = True Or Me.chkCOM.Checked = True Then

            'Recherche si des symboles existent
            Try
                Dim Values = New TypedValue() {New TypedValue(DxfCode.Start, "INSERT"), New TypedValue(DxfCode.LayerName, "MO_VALEUR_PTS")}
                Dim ids As ObjectId() = SelectAllItems(Values).Value.GetObjectIds
                lngObjCount = ids.Length
            Catch
            End Try

            'Si des syboles existent, on demande à l'utilisateur s'il doivent être effacés
            If lngObjCount > 0 Then
                repmsg = MsgBox("Voulez-vous supprimer les symboles de fiabilité planimétrique existants ?", MsgBoxStyle.Question + MsgBoxStyle.YesNoCancel, "Supprimer les objets existants")
                If repmsg = MsgBoxResult.Yes Then  'On efface les objets du quadrillage existants
                    Me.Cursor = System.Windows.Forms.Cursors.WaitCursor
                    DeleteAllInLayer("MO_VALEUR_PTS")
                ElseIf repmsg = MsgBoxResult.No Then
                    'On continue normalement
                Else
                    Exit Sub
                End If
            End If

            'Masque la feuille
            Me.Hide()

            'Message
            Connect.Message("Points fiables", "Analyse de la fiabilité planimétrique des entités sélectionnées, veuillez patienter...", False, 6, 100)

            'Active le calque devant contenir les symboles
            FreezeLayer("MO_VALEUR_PTS", False)
            ShowLayer("MO_VALEUR_PTS", True)
            ActivateLayer("MO_VALEUR_PTS")

            'Echelle d'insertion
            dblScale = Val(GetCurrentScale()) / 1000


            'Création des symboles de valeur ------------------------------------------------------------------------

            'Points fixes
            If Me.chkPF.Checked = True Then
                Connect.Message("Points fiables", " Traitement du thème : Points fixes", False, 1, 6)

                'MO
                For i = 1 To 2 'PFP 1/2

                    'Sélectionne les PFP1 ou 2 fiables
                    ' objSelSet = GetBlocksByAttribAndLayer("MO_PFP" & i, "", "FIABPLAN", "oui", False, False)
                    objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MO_PFP" & i)


                    'Y a-t'il des résultats ?
                    If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)

                        For Each objAcadBlock As BlockReference In objSelSet

                            'Point d'insertion
                            dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                            'Point accessible ou non ?
                            Dim ValAtt As String = GetBlockAttribute(objAcadBlock.ObjectId, "ACCESSIBILITE").ToLower
                            If ValAtt = "oui" Or ValAtt = "accessible" Then
                                booAccessible = True
                            Else
                                booAccessible = False
                            End If

                            If booAccessible = True Then
                                InsertBlock(dblPt, "VAL_PFP12A", GetCurrentRotation, layID, dblScale)
                            Else
                                InsertBlock(dblPt, "VAL_PFP12I", GetCurrentRotation, layID, dblScale)
                            End If
                        Next
                    End If

                    'Effacement de la sélection
                    objSelSet.Clear()

                Next

                'MUT
                For i = 1 To 2 'PFP 1/2

                    'Sélectionne les PFP1 ou 2 fiables
                    'objSelSet = GetBlocksByAttribAndLayer("MUT_PFP" & i, "", "FIABPLAN", "oui", False, False)
                    objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MUT_PFP" & i)


                    'Y a-t'il des résultats ?
                    If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)

                        For Each objAcadBlock As BlockReference In objSelSet

                            'Point d'insertion
                            dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                            'Insertion d'un nouveau bloc (différent si accessible ou non)

                            'Point accessible ou non ?
                            'Dim varAttributes As Variant
                            Dim ValAtt As String = GetBlockAttribute(objAcadBlock.ObjectId, "ACCESSIBILITE").ToLower
                            If ValAtt = "oui" Or ValAtt = "accessible" Then
                                booAccessible = True
                            Else
                                booAccessible = False
                            End If

                            If booAccessible = True Then
                                InsertBlock(dblPt, "VAL_PFP12A", GetCurrentRotation, layID, dblScale)
                            Else
                                InsertBlock(dblPt, "VAL_PFP12I", GetCurrentRotation, layID, dblScale)
                            End If
                        Next objAcadBlock
                    End If

                    'Effacement de la sélection
                    objSelSet.Clear()

                Next


                'Sélectionne les PFP3 fiables
                'MO
                'objSelSet = GetBlocksByAttribAndLayer("MO_PFP3", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MO_PFP3")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_PFP3", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

                'MUT
                'objSelSet = GetBlocksByAttribAndLayer("MUT_PFP3", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MUT_PFP3")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_PFP3", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

            End If

            'CS
            If Me.chkCS.Checked = True Then
                Connect.Message("Points fiables", " Traitement du thème : Couverture du Sol", False, 2, 6)

                'MO
                'Sélectionne les points CS
                'objSelSet = GetBlocksByAttribAndLayer("MO_CS_PTS", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MO_CS_PTS")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_FIABLE", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

                'MUT
                'Sélectionne les points CS
                'objSelSet = GetBlocksByAttribAndLayer("MUT_CS_PTS", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MUT_CS_PTS")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_FIABLE", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

            End If

            'OD
            If Me.chkOD.Checked = True Then
                Connect.Message("Points fiables", " Traitement du thème : Objets divers", False, 3, 6)

                'MO
                'Sélectionne les points OD
                '  objSelSet = GetBlocksByAttribAndLayer("MO_OD_PTS", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MO_OD_PTS")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_FIABLE", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'MUT
                'Effacement de la sélection
                objSelSet.Clear()

                'Sélectionne les points OD
                'objSelSet = GetBlocksByAttribAndLayer("MUT_OD_PTS", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MUT_OD_PTS")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_FIABLE", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

            End If

            'BF
            If Me.chkBF.Checked = True Then
                Connect.Message("Points fiables", " Traitement du thème : Bien-fonds", False, 4, 6)

                'MO
                'Sélectionne les points BF
                'objSelSet = GetBlocksByAttribAndLayer("MO_BF_PTS", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MO_BF_PTS")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_FIABLE", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

                'MUT
                'Sélectionne les points BF
                ' objSelSet = GetBlocksByAttribAndLayer("MUT_BF_PTS", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MUT_BF_PTS")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_FIABLE", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

            End If

            'COM
            If Me.chkBF.Checked = True Then
                Connect.Message("Points fiables", " Traitement du thème : Commune", False, 5, 6)

                'MO
                'Sélectionne les points COM
                'objSelSet = GetBlocksByAttribAndLayer("MO_COM_PTS", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MO_COM_PTS")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_FIABLE", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

                'MUT
                'Sélectionne les points COM
                'objSelSet = GetBlocksByAttribAndLayer("MUT_COM_PTS", "", "FIABPLAN", "oui", False, False)
                objSelSet = GetBlocksByAttributes("*", "FIABPLAN", "oui", "MUT_COM_PTS")

                'Y a-t'il des résultats ?
                If objSelSet.Count > 0 Then 'Erreur si sélection inexistante (= pas de bloc trouvé)
                    For Each objAcadBlock As BlockReference In objSelSet

                        'Point d'insertion
                        dblPt = New Autodesk.AutoCAD.Geometry.Point3d(objAcadBlock.Position.X, objAcadBlock.Position.Y, 0)

                        'Insertion d'un nouveau bloc (différent si accessible ou non)
                        InsertBlock(dblPt, "VAL_FIABLE", GetCurrentRotation, layID, dblScale)

                    Next objAcadBlock
                End If

                'Effacement de la sélection
                objSelSet.Clear()

            End If

            'Met à jour le dessin
            ' ThisDrawing.Regen(Autodesk.AutoCAD.Interop.Common.AcRegenType.acAllViewports)

            Dim RVscript As New Revo.RevoScript
            RVscript.cmdCmd("#Cmd;ANNOAUTOSCALE|4|_CANNOSCALE|1:500|_CANNOSCALE|1:1000|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1")
            Connect.Message("Points fiables", " Traitement terminé avec succès !", False, 100, 100, "info")
            Connect.Message("Fin", "... ", True, 0, 0, "hide") 'Fin
            Me.Close()

        Else

            'Message d'erreur
            MsgBox("Vous devez sélectionner au moins une entité !", MsgBoxStyle.Exclamation, "Aucune sélection")

        End If

    End Sub

    Private Sub frmPointsFiables_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        Me.AcceptButton = cmdNext
        Me.CancelButton = btnCancel

    End Sub
End Class