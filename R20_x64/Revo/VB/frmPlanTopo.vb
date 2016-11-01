Imports System.Windows.Forms

Public Class frmPlanTopo


    Private Sub frmPlanTopo_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        With Me.cboUnite
            .Items.Add("0.000 - millimètres")
            .Items.Add("0.00 - centimètres")
            .Items.Add("0.0 - décimètres")
            .Items.Add("0 - mètres")
            .SelectedIndex = 1
        End With

    End Sub

    Private Sub btnNext_Click(sender As System.Object, e As System.EventArgs) Handles btnNext.Click

      
        Me.Hide()
        Dim Conn As New Revo.connect
        Conn.Message("Création plan topo", "Ecriture / mise à jour des altitudes des points (blocs)", False, 33, 100)
        System.Windows.Forms.Application.DoEvents()


        'Affiche les couches "_ALT"
        Dim RVscript As New Revo.RevoScript
        RVscript.cmdLA("#LA;*_ALT;[[LayerOn]]1;[[Freeze]]0;")


        'Sélection des blocs

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

        Dim objSelSet As Autodesk.AutoCAD.Interop.AcadSelectionSet
        Dim objAcadBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference

        'Restrictions pour la sélection
        Dim adDXFCode(0) As Short
        Dim adDXFGroup(0) As Object

        objSelSet = acDoc.SelectionSets.Add("REVO" & acDoc.SelectionSets.Count) 'Nom de la sélection
        adDXFCode(0) = 0
        adDXFGroup(0) = "INSERT" 'Type d'objet
        
        'Effectue la sélection
        objSelSet.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , adDXFCode, adDXFGroup)

       
        'Boucle pour tous les éléments sélectionnés (blocs)
        Dim varAttributes As Object
        Dim strNomAttr As String = ""
        Dim dblZ As Double = 0
        Dim strZ As String
        Dim NbreUnite As Double = 2
        NbreUnite = CDbl(Me.cboUnite.Tag)

        If objSelSet.Count <> 0 Then

            'Déclarations

            For Each objAcadBlock In objSelSet


                'Récupère l'altitude
                dblZ = objAcadBlock.InsertionPoint(2)

                'Ignore le point si 0 et option cochée
                If Not (chkIgnorer0.Checked = True And dblZ = 0) Then

                    'Arrondi à 0 ou 5
                    If Me.chkArrondi05.Checked = True Then

                        '=ARRONDI(C16 /5;1)*5
                        dblZ = Math.Round((dblZ / 5), CInt(Me.cboUnite.Tag)) * 5

                    End If

                    'Formattage
                    If Val(Me.cboUnite.Tag) = 0 Then
                        strZ = Format(dblZ, "0")
                    Else
                        strZ = Format(dblZ, "0." & New String("0", Val(Me.cboUnite.Tag)))
                    End If

                Else
                    strZ = ""

                End If

                'Lit / écrit les attributs
                varAttributes = objAcadBlock.GetAttributes

                For j = LBound(varAttributes) To UBound(varAttributes)
                    'Cas particulier de l'altitude prévue mais pas toujours dans le modèle
                    If UCase(varAttributes(j).TagString) = "ALTITUDE" Then varAttributes(j).TextString = strZ
                Next

            Next objAcadBlock

        End If

        'Effacement de la sélection
        objSelSet.Delete()


        'Fin
        Conn.Message("Création plan topo", "Ecriture / mise à jour des altitudes des points (blocs)", False, 100, 100)
        Conn.Message("Création plan topo", "Ecriture / mise à jour des altitudes des points (blocs)", True, 100, 100)

    End Sub

    Private Sub cboUnite_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboUnite.SelectedIndexChanged

        Select Case Me.cboUnite.SelectedIndex
            Case 0 : Me.cboUnite.Tag = "3" 'mm
            Case 1 : Me.cboUnite.Tag = "2" 'cm
            Case 2 : Me.cboUnite.Tag = "1" 'dm
            Case 3 : Me.cboUnite.Tag = "0" 'm
        End Select

    End Sub
End Class