Option Strict Off
Option Explicit On
Imports VB = Microsoft.VisualBasic
Friend Class frmInsertionPts
    Inherits System.Windows.Forms.Form
    Private strSelPrec As String
    Private booLoading As Boolean
    Public Shared SelectedNature As String = ""

    Private Sub frmInsertionPts_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        booLoading = True

        Me.AcceptButton = cmdNext
        Me.CancelButton = cmdCancel

        'Liste d�roulante des th�mes (% � la cat�gorie s�lectionn�e)
        strSelPrec = ""

        'Oui/non
        Me.cboFiabPlan.Items.Add("?") : Me.cboFiabPlan.Items.Add("oui") : Me.cboFiabPlan.Items.Add("non")
        Me.cboFiabAlt.Items.Add("?") : Me.cboFiabAlt.Items.Add("oui") : Me.cboFiabAlt.Items.Add("non")
        Me.cboExact.Items.Add("oui") : Me.cboExact.Items.Add("non")

        'Liste d�roulante des cat�gories/classes
        With Me.cboCat
            .Items.Add("MO")
            .Items.Add("MUT")
            .Items.Add("IMPL")
            .Items.Add("TOPO")
            .Items.Add("SOUT")
        End With


        'Valeurs par d�faut = derni�res valeurs saisies
        Me.cboCat.SelectedIndex = intInsPtCat
        Me.cboTheme.SelectedIndex = intInsPtTheme : Me.cboType.SelectedIndex = intInsPtType
        If Me.cboSigne.Visible = True Then Me.cboSigne.SelectedIndex = intInsPtSigne
        Me.cboExact.SelectedIndex = intInsPtDefini
        If Me.cboAttrVariable.Visible = True Then Me.cboAttrVariable.SelectedIndex = intInsPtVar
        Me.txtCom.Text = GetFileProperty("Commune") : Me.txtPlan.Text = GetFileProperty("Plan")

        Me.cboFiabPlan.SelectedIndex = intInsPtFiabPlan : Me.cboFiabAlt.SelectedIndex = intInsPtFiabAlt
        Me.txtPrecPlan.Text = strInsPtPrecPlan : Me.txtPrecAlt.Text = strInsPtPrecAlt

        'Orientation des blocs selon fonction rotation des blocs
        'Me.txtOri.Text = Format(GetCurrentRotation, "0.0") --> Current rotation not used anymore as of https://github.com/Tamu/revo/issues/16

        'Dernier num�ro de point + 1 (si num�rique)
        Dim strNr As String : strNr = strInsPtNum
        If strNr = "0" Then
            Me.txtNo.Text = "1"
        ElseIf Val(strNr) <> 0 Then
            Me.txtNo.Text = CStr(Val(strNr) + 1)
        Else
            Me.txtNo.Text = ""
        End If


        'Entit�s administratives pr�sentes dans le fichier
        'Entit�s administratives pr�sentes dans le fichier
        Dim j As Object
        Dim strPlan As String = ""
        Dim strCom As String = ""

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

        'S�lection
        Dim objSelSet As Autodesk.AutoCAD.Interop.AcadSelectionSet
        Dim objAcadBlock As Autodesk.AutoCAD.Interop.Common.AcadBlockReference

        'Restrictions pour la s�lection
        Dim adDXFCode(2) As Short
        Dim adDXFGroup(2) As Object

        objSelSet = acDoc.SelectionSets.Add("JourCAD" & acDoc.SelectionSets.Count) 'Nom de la s�lection
        adDXFCode(0) = 0
        adDXFGroup(0) = "INSERT" 'Type d'objet
        adDXFCode(1) = 2
        adDXFGroup(1) = "PLAN" 'Nom de bloc
        adDXFCode(2) = 8
        adDXFGroup(2) = "MO_RPL_PLAN" 'Calque

        'Effectue la s�lection
        objSelSet.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , adDXFCode, adDXFGroup)

        'Boucle pour tous les �l�ments s�lectionn�s (polylignes)
        Dim strCode As String = ""
        Dim varAttributes As Object
        If objSelSet.Count <> 0 Then


            'Boucle pour chaque bloc de la couche "MO_RPL_PLAN"
            For Each objAcadBlock In objSelSet

                '  MsgBox(TypeName(objAcadBlock))
                'Dim BL As  = objAcadBlock

                'Lecture des attributs
                varAttributes = objAcadBlock.GetAttributes

                For j = LBound(varAttributes) To UBound(varAttributes)
                    Select Case varAttributes(j).TagString
                        Case "IDENTDN"
                            strCom = Mid(varAttributes(j).TextString, 4, 3)
                        Case "NUMERO"
                            strPlan = varAttributes(j).TextString
                        Case "CODEPLAN"
                            strCode = varAttributes(j).TextString
                    End Select
                Next

                'Ajout de l'info � la ListView
                'La fonction "d�cortique" une ligne "OBJE" Interlis => mise en forme bidon pour r�utiliser la fctn
                PlanToList("0 1 2 VD0" & strCom & "000000 " & strPlan & " 5 6 7 " & strCode & " 9", Me.LVPlans)

            Next objAcadBlock

        End If

        'Effacement de la s�lection
        objSelSet.Delete()


        'MUT
        'objSelSet = ThisDrawing.SelectionSets.Add("JourCAD" & ThisDrawing.SelectionSets.Count) 'Nom de la s�lection
        adDXFCode(0) = 0
        'UPGRADE_WARNING: Impossible de r�soudre la propri�t� par d�faut de l'objet adDXFGroup(0). Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        adDXFGroup(0) = "INSERT" 'Type d'objet
        adDXFCode(1) = 2
        adDXFGroup(1) = "MUT_PLAN" 'Nom de bloc
        adDXFCode(2) = 8
        adDXFGroup(2) = "MUT_RPL_PLAN" 'Calque

        'Effectue la s�lection
        ' objSelSet.Select(Autodesk.AutoCAD.Interop.Common.AcSelect.acSelectionSetAll, , , adDXFCode, adDXFGroup)

        'Boucle pour tous les �l�ments s�lectionn�s (polylignes)
        Try

            If objSelSet.Count <> 0 Then

                'Boucle pour chaque bloc de la couche "MO_RPL_PLAN"
                For Each objAcadBlock In objSelSet

                    'Lecture des attributs
                    varAttributes = objAcadBlock.GetAttributes

                    For j = LBound(varAttributes) To UBound(varAttributes)
                        Select Case varAttributes(j).TagString
                            Case "IDENTDN"
                                strCom = Mid(varAttributes(j).TextString, 4, 3)
                            Case "NUMERO"
                                strPlan = varAttributes(j).TextString
                            Case "CODEPLAN"
                                strCode = varAttributes(j).TextString
                        End Select
                    Next

                    'Ajout de l'info � la ListView
                    'La fonction "d�cortique" une ligne "OBJE" Interlis => mise en forme bidon pour r�utiliser la fctn
                    PlanToList("0 1 2 VD0" & strCom & "000000 " & strPlan & " 5 6 7 " & strCode & " 9", Me.LVPlans)

                Next objAcadBlock

            End If


            'Effacement de la s�lection
            objSelSet.Delete()

        Catch ex As Runtime.InteropServices.COMException
            '  MsgBox(ex.Message)
        End Try


        'Trie les entit�s par num�ro de plan
        'Me.LVPlans.SortKey = 6
        'UPGRADE_WARNING: Impossible de r�soudre la propri�t� par d�faut de l'objet Me.LVPlans.Sorted. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        'Me.LVPlans.Sorted = True


        booLoading = False
    End Sub



    Private Sub cmdCancel_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdCancel.Click

        Me.Close()

    End Sub
    'UPGRADE_WARNING: L'�v�nement cboCat.SelectedIndexChanged peut se d�clencher lorsque le formulaire est initialis�. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboCat_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboCat.SelectedIndexChanged

        'S�lection pr�c�dente <> "" (= changement de cat�gorie)
        If strSelPrec <> "" And strSelPrec <> Me.cboCat.Text Then
            'Changement de cat�gorie (MO <=> MUT)
            If (strSelPrec = "MO" And Me.cboCat.Text = "MUT") Or (strSelPrec = "MUT" And Me.cboCat.Text = "MO") Then
                If Me.cboTheme.Text <> "01 PF" Then
                    strSelPrec = Me.cboCat.Text
                    Exit Sub
                End If
            End If
        End If


        With Me.cboTheme
            .Items.Clear()

            Select Case Me.cboCat.Text

                Case "MO", "MUT"
                    .Visible = True : Me.lblTheme.Visible = True
                    .Items.Add("01 PF")
                    .Items.Add("02 CS")
                    .Items.Add("03 OD")
                    .Items.Add("06 BF")
                    .Items.Add("09 COM")

                Case "TOPO", "IMPL"
                    .Visible = False : Me.lblTheme.Visible = False
                    .Items.Add("Points")

                Case "SOUT"
                    .Visible = True : Me.lblTheme.Visible = True
                    .Items.Add("Eau potable")
                    .Items.Add("E. P. (projet)")
                    .Items.Add("Eaux claires")
                    .Items.Add("E. C. (projet)")
                    .Items.Add("Eaux unitaires")
                    .Items.Add("Eaux us�es")
                    .Items.Add("E. U. (projet)")
                    .Items.Add("Eclairage public")
                    .Items.Add("Electricit�")
                    .Items.Add("Gaz")
                    .Items.Add("T�l�phone")
                    .Items.Add("TV")

            End Select

            'Valeur par d�faut
            'On Error Resume Next '.ListIndex si derni�re valeur m�moris�e hors plage (autre cat�gorie)

            '>> modRecentSettings
            If intInsPtTheme >= .Items.Count Then
                .SelectedIndex = 0
            Else
                .SelectedIndex = intInsPtTheme
            End If

            On Error GoTo 0


            If booLoading = False Then strSelPrec = Me.cboCat.Text

        End With

    End Sub

    'UPGRADE_WARNING: L'�v�nement cboSigne.SelectedIndexChanged peut se d�clencher lorsque le formulaire est initialis�. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboSigne_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboSigne.SelectedIndexChanged

        'Defini exactement ?
        'Valable pour MO, MUT, TOPO_PTS et IMPL
        Me.lblExact.Visible = (Me.cboSigne.Text = "non materialise" Or Me.cboSigne.Text = "non materialise non defini" Or Me.cboSigne.Text = "point terrain" Or Me.cboCat.Text = "IMPL")
        Me.cboExact.Visible = Me.lblExact.Visible

        If Me.cboSigne.Text = "non materialise" Then
            Me.cboExact.Text = "oui"
        ElseIf Me.cboSigne.Text = "non materialise non defini" Then
            Me.cboExact.Text = "non"
        End If

    End Sub

    Private Sub cboExact_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles cboExact.SelectedIndexChanged

        If Me.cboExact.Text = "oui" Then
            If Me.cboSigne.Text = "non materialise non defini" Then Me.cboSigne.Text = "non materialise"
        ElseIf Me.cboExact.Text = "non" Then
            If Me.cboSigne.Text = "non materialise" Then Me.cboSigne.Text = "non materialise non defini"
        End If

    End Sub

    'UPGRADE_WARNING: L'�v�nement cboTheme.SelectedIndexChanged peut se d�clencher lorsque le formulaire est initialis�. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboTheme_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboTheme.SelectedIndexChanged

        On Error Resume Next 'ListIndex si hors plage

        With Me.cboType
            .Items.Clear()

            Select Case Me.cboCat.Text

                Case "MO"

                    Select Case VB.Left(Me.cboTheme.Text, 2)

                        Case "01"
                            .Visible = True : Me.lblType.Visible = True
                            .Items.Add("PFA1")
                            .Items.Add("PFA2")
                            .Items.Add("PFA3")
                            .Items.Add("PFP1")
                            .Items.Add("PFP2")
                            .Items.Add("PFP3")
                            .SelectedIndex = intInsPtType

                        Case "02", "03"
                            .Visible = False : Me.lblType.Visible = False
                            .Items.Add("Point particulier")
                            .SelectedIndex = 0

                        Case "06"
                            .Visible = False : Me.lblType.Visible = False
                            .Items.Add("Point-limite")
                            .SelectedIndex = 0

                        Case "09"
                            .Visible = False : Me.lblType.Visible = False
                            .Items.Add("Point-limite territorial")
                            .SelectedIndex = 0

                    End Select


                Case "MUT"

                    Select Case VB.Left(Me.cboTheme.Text, 2)

                        Case "01"
                            .Visible = False : Me.lblType.Visible = False
                            .Items.Add("PFP3")
                            .SelectedIndex = 0

                        Case "02", "03"
                            .Visible = False : Me.lblType.Visible = False
                            .Items.Add("Point particulier")
                            .SelectedIndex = 0

                        Case "06"
                            .Visible = False : Me.lblType.Visible = False
                            .Items.Add("Point-limite")
                            .SelectedIndex = 0

                        Case "09"
                            .Visible = False : Me.lblType.Visible = False
                            .Items.Add("Point-limite territorial")
                            .SelectedIndex = 0

                    End Select


                Case "IMPL", "TOPO", "SOUT"

                    'Pas de type
                    .Visible = False : Me.lblType.Visible = False
                    .Items.Add("Point particulier")
                    .SelectedIndex = 0


            End Select

        End With

        On Error GoTo 0

    End Sub

    'UPGRADE_WARNING: L'�v�nement cboType.SelectedIndexChanged peut se d�clencher lorsque le formulaire est initialis�. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub cboType_SelectedIndexChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles cboType.SelectedIndexChanged

        'Liste des propri�t�s pour le type de points s�lectionn�
        Select Case Me.cboType.Text

            Case "PFA1", "PFA2"

                Me.lblCom.Enabled = False
                Me.txtCom.Enabled = False
                Me.lblPlan.Enabled = False : Me.lblPlan.Text = "Carte:"
                Me.txtPlan.Enabled = False
                Me.LVPlans.Enabled = False
                Me.lblSigne.Visible = False
                Me.cboSigne.Visible = False
                Me.lblFiabAlt.Visible = True
                Me.cboFiabAlt.Visible = True : If Me.Tag <> "EDIT" Then Me.cboFiabAlt.SelectedIndex = 1 'Oui par d�faut
                Me.cboFiabPlan.SelectedIndex = 1 'Oui par d�faut
                Me.lblPrecAlt.Visible = True
                Me.txtPrecAlt.Visible = True : If Me.Tag <> "EDIT" Then Me.txtPrecAlt.Text = ""
                Me.lblExact.Visible = False
                Me.cboExact.Visible = False
                Me.lblAttrVariable.Visible = False
                Me.cboAttrVariable.Visible = False
                Me.lblComment.Visible = False
                Me.txtComment.Visible = False

                'IdentDN
                If Me.cboType.Text = "PFA1" Then
                    Me.txtIdentDN.Text = "CH0200000000"
                ElseIf Me.cboType.Text = "PFA2" Then
                    Me.txtIdentDN.Text = "VD0100000000"
                End If


            Case "PFA3"

                Me.lblCom.Enabled = True
                Me.txtCom.Enabled = True
                Me.lblPlan.Enabled = True : Me.lblPlan.Text = "Plan:"
                Me.txtPlan.Enabled = True
                Me.LVPlans.Enabled = True
                Me.lblSigne.Visible = False
                Me.cboSigne.Visible = False
                Me.lblFiabAlt.Visible = True
                Me.cboFiabAlt.Visible = True : Me.cboFiabAlt.SelectedIndex = 1 'Oui par d�faut
                Me.cboFiabPlan.SelectedIndex = 1 'Oui par d�faut
                Me.lblPrecAlt.Visible = True
                Me.txtPrecAlt.Visible = True
                Me.lblExact.Visible = False
                Me.cboExact.Visible = False
                Me.lblAttrVariable.Visible = False
                Me.cboAttrVariable.Visible = False
                Me.lblComment.Visible = False
                Me.txtComment.Visible = False
                'IdentDN
                Me.txtIdentDN.Text = "VD0CCC00PPPP"
                Me.txtIdentDN.Tag = "VD0CCC00PPPP"


            Case "PFP1", "PFP2"

                Me.lblCom.Enabled = False
                Me.txtCom.Enabled = False
                Me.lblPlan.Enabled = True : Me.lblPlan.Text = "Carte:"
                Me.txtPlan.Enabled = True
                Me.LVPlans.Enabled = True
                Me.lblSigne.Visible = True : Me.lblSigne.Enabled = True
                Me.cboSigne.Visible = True : Me.cboSigne.Enabled = True
                Me.lblFiabAlt.Visible = True
                Me.cboFiabAlt.Visible = True : Me.cboFiabAlt.SelectedIndex = 1 'Oui par d�faut
                Me.cboFiabPlan.SelectedIndex = 1 'Oui par d�faut
                Me.lblPrecAlt.Visible = True
                Me.txtPrecAlt.Visible = True
                Me.lblExact.Visible = False
                Me.cboExact.Visible = False
                Me.lblAttrVariable.Visible = True : Me.lblAttrVariable.Text = "Accessibilit�:"
                Me.cboAttrVariable.Items.Clear() : Me.cboAttrVariable.Items.Add("accessible") : Me.cboAttrVariable.Items.Add("inaccessible") : Me.cboAttrVariable.SelectedIndex = 0
                Me.cboAttrVariable.Visible = True : Me.cboAttrVariable.SelectedIndex = 0 '? par d�faut
                Me.lblComment.Visible = False
                Me.txtComment.Visible = False

                'Liste d�roulante des signes (natures)
                With Me.cboSigne
                    .Items.Clear()
                    .Items.Add("?")
                    .Items.Add("borne")
                    .Items.Add("cheville")
                    .Items.Add("pieu")
                    .Items.Add("non materialise")
                    .Items.Add("non materialise non defini")
                    'Valeur par d�faut
                    .SelectedIndex = intInsPtSigne
                End With

                'IdentDN
                Me.txtIdentDN.Text = "CH030000PPPP"
                Me.txtIdentDN.Tag = "CH030000PPPP"


            Case "PFP3"

                Me.lblCom.Enabled = True
                Me.txtCom.Enabled = True
                Me.lblPlan.Enabled = True : Me.lblPlan.Text = "Plan:"
                Me.txtPlan.Enabled = True
                Me.LVPlans.Enabled = True
                Me.lblSigne.Visible = True : Me.lblSigne.Enabled = True
                Me.cboSigne.Visible = True : Me.cboSigne.Enabled = True
                Me.lblFiabAlt.Visible = True
                Me.cboFiabAlt.Visible = True : Me.cboFiabAlt.SelectedIndex = 1 'Oui par d�faut
                Me.cboFiabPlan.SelectedIndex = 1 'Oui par d�faut
                Me.lblPrecAlt.Visible = True
                Me.txtPrecAlt.Visible = True
                Me.lblExact.Visible = False
                Me.cboExact.Visible = False
                Me.lblAttrVariable.Visible = True : Me.lblAttrVariable.Text = "Fiche:"
                Me.cboAttrVariable.Items.Clear() : Me.cboAttrVariable.Items.Add("?") : Me.cboAttrVariable.Items.Add("oui") : Me.cboAttrVariable.Items.Add("non") : Me.cboAttrVariable.SelectedIndex = 0
                Me.cboAttrVariable.Visible = True
                Me.lblComment.Visible = False
                Me.txtComment.Visible = False

                'Liste d�roulante des signes (natures)
                With Me.cboSigne
                    .Items.Clear()
                    .Items.Add("borne")
                    .Items.Add("cheville")
                    .Items.Add("croix")
                    .Items.Add("pieu")
                    '.AddItem "non materialise"
                    'Valeur par d�faut
                    If intInsPtSigne < 4 Then .SelectedIndex = intInsPtSigne
                End With

                'IdentDN
                Me.txtIdentDN.Text = "VD0CCC00PPPP"
                Me.txtIdentDN.Tag = "VD0CCC00PPPP"


            Case "Point-limite"

                Me.lblCom.Enabled = True
                Me.txtCom.Enabled = True
                Me.lblPlan.Enabled = True : Me.lblPlan.Text = "Plan:"
                Me.txtPlan.Enabled = True
                Me.LVPlans.Enabled = True
                Me.lblSigne.Visible = True : Me.lblSigne.Enabled = True
                Me.cboSigne.Visible = True : Me.cboSigne.Enabled = True
                Me.lblFiabAlt.Visible = False
                Me.cboFiabAlt.Visible = False
                Me.cboFiabAlt.SelectedIndex = 0 '? par d�faut
                Me.cboFiabPlan.SelectedIndex = 1 'Oui par d�faut
                Me.lblPrecAlt.Visible = False
                Me.txtPrecAlt.Visible = False : Me.txtPrecAlt.Text = ""
                Me.lblExact.Visible = False
                Me.cboExact.Visible = False : Me.cboExact.SelectedIndex = 1 'Par d�faut: non
                Me.lblAttrVariable.Visible = True : Me.lblAttrVariable.Text = "Anc borne:"
                Me.cboAttrVariable.Items.Clear() : Me.cboAttrVariable.Items.Add("?") : Me.cboAttrVariable.Items.Add("oui") : Me.cboAttrVariable.Items.Add("non") : Me.cboAttrVariable.SelectedIndex = 0
                Me.cboAttrVariable.Visible = True
                Me.lblComment.Visible = False
                Me.txtComment.Visible = False

                'Liste d�roulante des signes (natures)
                With Me.cboSigne
                    .Items.Clear()
                    .Items.Add("borne")
                    .Items.Add("cheville")
                    .Items.Add("croix")
                    .Items.Add("pieu")
                    .Items.Add("non materialise")
                    .Items.Add("non materialise non defini")
                    'Valeur par d�faut
                    .SelectedIndex = intInsPtSigne
                End With

                'IdentDN
                Me.txtIdentDN.Text = "VD0CCC00PPPP"
                Me.txtIdentDN.Tag = "VD0CCC00PPPP"


            Case "Point-limite territorial"

                Me.lblCom.Enabled = True
                Me.txtCom.Enabled = True
                Me.lblPlan.Enabled = False : Me.lblPlan.Text = "Plan:"
                Me.txtPlan.Enabled = False
                Me.LVPlans.Enabled = True
                Me.lblSigne.Visible = True : Me.lblSigne.Enabled = True
                Me.cboSigne.Visible = True : Me.cboSigne.Enabled = True
                Me.lblFiabAlt.Visible = False
                Me.cboFiabAlt.Visible = False
                Me.cboFiabAlt.SelectedIndex = 0 '? par d�faut
                Me.cboFiabPlan.SelectedIndex = 1 'Oui par d�faut
                Me.lblPrecAlt.Visible = False
                Me.txtPrecAlt.Visible = False : Me.txtPrecAlt.Text = ""
                'Me.lblExact.Visible = True: Me.lblExact.Enabled = False
                'Me.cboExact.Visible = True: Me.cboExact.Enabled = False: Me.cboExact.ListIndex = 0 'Par d�faut: oui
                Me.lblExact.Visible = False
                Me.cboExact.Visible = False : Me.cboExact.SelectedIndex = 1 'Par d�faut: oui
                Me.lblAttrVariable.Visible = True : Me.lblAttrVariable.Text = "Anc borne:"
                Me.cboAttrVariable.Items.Clear() : Me.cboAttrVariable.Items.Add("?") : Me.cboAttrVariable.Items.Add("oui") : Me.cboAttrVariable.Items.Add("non") : Me.cboAttrVariable.SelectedIndex = 0
                Me.cboAttrVariable.Visible = True
                Me.lblComment.Visible = False
                Me.txtComment.Visible = False

                'Liste d�roulante des signes (natures)
                With Me.cboSigne
                    .Items.Clear()
                    .Items.Add("borne")
                    .Items.Add("cheville")
                    .Items.Add("croix")
                    .Items.Add("pieu")
                    .Items.Add("non materialise")
                    .Items.Add("non materialise non defini")
                    'Valeur par d�faut
                    .SelectedIndex = intInsPtSigne
                End With

                'IdentDN
                Me.txtIdentDN.Text = "VD0CCC000000"
                Me.txtIdentDN.Tag = "VD0CCC000000"


            Case "Point particulier" 'CS / OD / TOPO / IMPL / SOUT

                Me.lblCom.Enabled = True
                Me.txtCom.Enabled = True
                Me.lblPlan.Enabled = True : Me.lblPlan.Text = "Plan:"
                Me.txtPlan.Enabled = True
                Me.LVPlans.Enabled = True
                Me.lblFiabAlt.Visible = False
                Me.cboFiabAlt.Visible = False
                Me.cboFiabAlt.SelectedIndex = 0 '? par d�faut
                Me.cboFiabPlan.SelectedIndex = 2 'Non par d�faut
                Me.lblPrecAlt.Visible = False
                Me.txtPrecAlt.Visible = False : Me.txtPrecAlt.Text = ""
                Me.lblExact.Visible = False
                Me.cboExact.Visible = False : Me.cboExact.SelectedIndex = 1 'Par d�faut: oui
                Me.lblAttrVariable.Visible = False
                Me.cboAttrVariable.Visible = False

                'Liste d�roulante des signes (natures)
                With Me.cboSigne
                    .Items.Clear()

                    'CS/OD
                    If Me.cboCat.Text = "MO" Or Me.cboCat.Text = "MUT" Then
                        Me.lblComment.Visible = False
                        Me.txtComment.Visible = False

                        Me.lblSigne.Visible = False
                        .Visible = False
                        .Items.Add("non materialise")
                        .Items.Add("non materialise non defini")
                        'Valeur par d�faut
                        .SelectedIndex = 0

                        'TOPO
                    ElseIf Me.cboCat.Text = "TOPO" Then
                        Me.lblComment.Visible = False
                        Me.txtComment.Visible = False

                        Me.lblSigne.Visible = True
                        .Visible = True

                        .Items.Add("point terrain")
                        .Items.Add("arbre")

                        .Items.Add("pied de mur") '88	Pied de mur
                        .Items.Add("haut de mur") '89	Haut de mur
                        .Items.Add("bord de route") '93	Bord de route
                        .Items.Add("trottoir") '94 Trottoir
                        .Items.Add("escalier") '96 Escalier
                        .Items.Add("point cot�") '97 Point cot�
                        .Items.Add("point de ligne de rupture") '98 Point de ligne de rupture
                        .Items.Add("fa�te, corniche") '99 Fa�te, corniche


                        'Valeur par d�faut
                        On Error Resume Next
                        If intInsPtSigne >= .Items.Count Then
                            .SelectedIndex = 0
                        Else
                            .SelectedIndex = intInsPtSigne
                        End If
                        On Error GoTo 0

                        'IMPL
                    ElseIf Me.cboCat.Text = "IMPL" Then
                        Me.lblComment.Visible = False
                        Me.txtComment.Visible = False

                        Me.lblSigne.Visible = True
                        .Visible = True

                        .Items.Add("controle")
                        .Items.Add("projet")
                        'Valeur par d�faut
                        If intInsPtSigne >= .Items.Count Then
                            .SelectedIndex = 0
                        Else
                            .SelectedIndex = intInsPtSigne
                        End If

                        'SOUT
                    ElseIf Me.cboCat.Text = "SOUT" Then
                        Me.lblComment.Visible = True
                        Me.txtComment.Visible = True

                        Me.lblSigne.Visible = True
                        .Visible = True

                        'Objet
                        Select Case Me.cboTheme.Text

                            Case "Electricit�"
                                .Items.Add("Cabine distribution")
                                .Items.Add("Point de rep�rage")
                                .Items.Add("Poteau")

                            Case "T�l�phone"
                                .Items.Add("Chambre � dalle")
                                .Items.Add("Chambre � regard")
                                .Items.Add("Point de rep�rage")
                                .Items.Add("Poteau")

                            Case "Eclairage public"
                                .Items.Add("Cabine distribution")
                                .Items.Add("Cand�labre")
                                .Items.Add("Chambre �clairage")
                                .Items.Add("Point de rep�rage")

                            Case "Gaz"
                                .Items.Add("Manchon")
                                .Items.Add("Point de rep�rage")
                                .Items.Add("R�duction")
                                .Items.Add("Vanne")

                            Case "Eau potable", "E. P. (projet)"
                                .Items.Add("Borne hydrante")
                                .Items.Add("Chambre captage")
                                .Items.Add("Chambre vannes")
                                .Items.Add("Compteur")
                                .Items.Add("Manchon")
                                .Items.Add("Point de rep�rage")
                                .Items.Add("Purge")
                                .Items.Add("R�duction")
                                .Items.Add("R�servoir, fontaine")
                                .Items.Add("Vanne")

                            Case "TV"
                                .Items.Add("Antenne")
                                .Items.Add("Armoire")
                                .Items.Add("Chambre")
                                .Items.Add("Point de rep�rage")

                            Case "Eaux claires", "E. C. (projet)"
                                .Items.Add("Chambre")
                                .Items.Add("Chambre enterr�e")
                                .Items.Add("Cheneau")
                                .Items.Add("D�versoir")
                                .Items.Add("Exutoire")
                                .Items.Add("Grille")
                                .Items.Add("Gueulard")
                                .Items.Add("Ouvrage sp�cial")
                                .Items.Add("Point de rep�rage")
                                .Items.Add("Relevage")
                                .Items.Add("S�parateur")

                            Case "Eaux us�es", "E. U. (projet)"
                                .Items.Add("Chambre")
                                .Items.Add("Chambre enterr�e")
                                .Items.Add("Fosse septique")
                                .Items.Add("Ouvrage sp�cial")
                                .Items.Add("Point de rep�rage")
                                .Items.Add("Relevage")
                                .Items.Add("S�parateur")

                            Case "Eaux unitaires"
                                .Items.Add("Chambre")
                                .Items.Add("Chambre enterr�e")
                                .Items.Add("Ouvrage sp�cial")
                                .Items.Add("Point de rep�rage")

                        End Select

                        'Valeur par d�faut
                        On Error Resume Next
                        If intInsPtSigne >= .Items.Count Then
                            .SelectedIndex = 0
                        Else
                            .SelectedIndex = intInsPtSigne
                        End If
                        On Error GoTo 0
                    End If

                End With

                'IdentDN
                Me.txtIdentDN.Text = "VD0CCC00PPPP"
                Me.txtIdentDN.Tag = "VD0CCC00PPPP"

        End Select


        'Mise � jour de l'IdentDN
        Call txtCom_TextChanged(txtCom, New System.EventArgs())
        Call txtPlan_TextChanged(txtPlan, New System.EventArgs())

    End Sub



    Private Sub cmdDigit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdDigit.Click

        Dim acDoc As Autodesk.AutoCAD.ApplicationServices.Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Me.Hide()

        'Digitalisation en s�rie
        'If Me.chkDigitSerie.CheckState = True Then
        '    frmInsertionEnSerie.lblNo.Text = Me.txtNo.Text
        '    frmInsertionEnSerie.Show()
        'End If

        'R�cup�rer un point cliqu� � l'�cran
        'On Error GoTo Cancel 'Erreur si annulation ("ESC")

        Dim pPtRes As Autodesk.AutoCAD.EditorInput.PromptPointResult
        Dim pPtOpts As Autodesk.AutoCAD.EditorInput.PromptPointOptions = New Autodesk.AutoCAD.EditorInput.PromptPointOptions("")

        '' Prompt for the start point
        pPtOpts.Message = vbLf & "D�finir le point d'insertion du bloc: "
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim returnPnt As Autodesk.AutoCAD.Geometry.Point3d = pPtRes.Value

        'returnPnt = ThisDrawing.Utility.GetPoint(, "D�finir le point d'insertion du bloc: ")
        ' On Error GoTo 0

        'Affiche les coordonn�es du point s�lectionn�
        Me.txtX.Text = Format(returnPnt(0), "0.000")
        Me.txtY.Text = Format(returnPnt(1), "0.000")
        Me.txtZ.Text = Format(returnPnt(2), "0.000")

        'Digitalisation en s�rie
        If Me.chkDigitSerie.CheckState = True Then
            Call cmdNext_Click(cmdNext, New System.EventArgs())
            Exit Sub
        End If

        'Cancel:
        Me.Show()
        Me.cmdNext.Focus()
    End Sub



    Private Sub cmdListeNat_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdListeNat.Click

        Dim frmCodeNat As New frmCodesNatures

        CodesNatures(frmCodeNat.tvNatures)

        Dim SelectedName As String = ""
        Dim SelectSelectedName As String = ""

        With frmCodeNat
            .Tag = "InsertionPoint"

            'Active la s�lection courante
            Select Case Me.cboCat.Text
                Case "MO"
                    If VB.Left(Me.cboTheme.Text, 2) = "01" Then
                        SelectedName = "MO_PF"
                        Select Case Me.cboType.Text
                            Case "PFA1", "PFA2", "PFA3"
                                'Rien
                            Case "PFP1", "PFP2"
                                SelectedName = "MO_PF_PFP1-2"
                                Select Case Me.cboSigne.Text
                                    Case "borne"
                                        SelectSelectedName = "NAT16"
                                    Case "cheville"
                                        SelectSelectedName = "NAT17"
                                    Case "croix", "pieu"
                                        SelectSelectedName = "NAT18"
                                    Case "non materialise"
                                        SelectSelectedName = "NAT19"
                                End Select
                            Case "PFP3"
                                SelectedName = "MO_PF_PFP3"
                                Select Case Me.cboSigne.Text
                                    Case "borne"
                                        SelectSelectedName = "NAT11"
                                    Case "cheville"
                                        SelectSelectedName = "NAT12"
                                    Case "croix"
                                        SelectSelectedName = "NAT13"
                                    Case "pieu"
                                        SelectSelectedName = "NAT14"
                                End Select
                        End Select

                    ElseIf VB.Left(Me.cboTheme.Text, 2) = "02" Then
                        SelectedName = "MO_CS"
                        If Me.cboExact.SelectedIndex = 0 Then
                            SelectSelectedName = "NAT25"
                        Else
                            SelectSelectedName = "NAT26"
                        End If
                    ElseIf VB.Left(Me.cboTheme.Text, 2) = "03" Then
                        SelectedName = "MO_OD"
                        If Me.cboExact.SelectedIndex = 0 Then
                            SelectSelectedName = "NAT35"
                        Else
                            SelectSelectedName = "NAT36"
                        End If
                    Else 'If left(Me.cboTheme.Text, 2) = "06" Then
                        SelectedName = "MO_BF"
                        Select Case Me.cboSigne.Text
                            Case "borne"
                                SelectSelectedName = "NAT61"
                            Case "cheville"
                                SelectSelectedName = "NAT62"
                            Case "croix"
                                SelectSelectedName = "NAT63"
                            Case "pieu"
                                SelectSelectedName = "NAT64"
                            Case "non materialise"
                                ' If Me.cboExact.SelectedIndex = 0 Then
                                SelectSelectedName = "NAT65"
                                ' Else
                                'SelectSelectedName = "NAT66"
                                'End If
                            Case "non materialise non defini"
                                'If Me.cboExact.SelectedIndex = 0 Then
                                'SelectSelectedName = "NAT65"
                                'Else
                                SelectSelectedName = "NAT66"
                                'End If
                        End Select
                    End If

                Case "MUT"
                    'Hors liste des natures
                Case "IMPL"
                    SelectedName = "IMPL"
                    If Me.cboSigne.Text = "projet" Then
                        SelectSelectedName = "NAT38"
                    Else 'If Me.cboSigne.Text = "controle" Then
                        SelectSelectedName = "NAT39"
                    End If

                Case "TOPO"


                    SelectedName = "TOPO" 'Domaine:    Topo+   EFA+C ----------  +100
                    Select Case Me.cboSigne.Text

                        Case "point terrain"
                            SelectSelectedName = "NAT8"
                        Case "arbre"
                            SelectSelectedName = "NAT9"

                        Case "pied de mur"
                            SelectSelectedName = "NAT88" '88	Pied de mur
                        Case "haut de mur"
                            SelectSelectedName = "NAT89" '89	Haut de mur
                        Case "bord de route"
                            SelectSelectedName = "NAT93" '93	Bord de route
                        Case "trottoir"
                            SelectSelectedName = "NAT94" '94 Trottoir
                        Case "escalier"
                            SelectSelectedName = "NAT96" '96 Escalier
                        Case "point cot�"
                            SelectSelectedName = "NAT97" '97 Point cot�
                        Case "point de ligne de rupture"
                            SelectSelectedName = "NAT98" '98 Point de ligne de rupture
                        Case "fa�te, corniche"
                            SelectSelectedName = "NAT99" '99 Fa�te, corniche
                    End Select

                Case "SOUT"
                    Select Case Me.cboTheme.Text
                        Case "Eau potable"
                            SelectedName = "SOUT_EP"
                            Select Case Me.cboSigne.Text
                                Case "Borne hydrante"
                                    SelectSelectedName = "NAT43"
                                Case "Chambre captage"
                                    SelectSelectedName = "NAT48"
                                Case "Chambre vannes"
                                    SelectSelectedName = "NAT42"
                                Case "Compteur"
                                    SelectSelectedName = "NAT46"
                                Case "Manchon"
                                    SelectSelectedName = "NAT44"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT40"
                                Case "Purge"
                                    SelectSelectedName = "NAT45"
                                Case "R�duction"
                                    SelectSelectedName = "NAT47"
                                Case "R�servoir, fontaine"
                                    SelectSelectedName = "NAT49"
                                Case "Vanne"
                                    SelectSelectedName = "NAT41"
                            End Select

                        Case "Electricit�"
                            SelectedName = "SOUT_EL"
                            Select Case Me.cboSigne.Text
                                Case "Cabine distribution"
                                    SelectSelectedName = "NAT2"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT1"
                                Case "Poteau"
                                    SelectSelectedName = "NAT3"
                            End Select

                        Case "T�l�phone"
                            SelectedName = "SOUT_TT"
                            Select Case Me.cboSigne.Text
                                Case "Chambre � dalle"
                                    SelectSelectedName = "NAT5"
                                Case "Chambre � regard"
                                    SelectSelectedName = "NAT6"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT4"
                                Case "Poteau"
                                    SelectSelectedName = "NAT7"
                            End Select

                        Case "Eclairage public"
                            SelectedName = "SOUT_EG"
                            Select Case Me.cboSigne.Text
                                Case "Cabine distribution"
                                    SelectSelectedName = "NAT21"
                                Case "Cand�labre"
                                    SelectSelectedName = "NAT22"
                                Case "Chambre �clairage"
                                    SelectSelectedName = "NAT23"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT20"
                            End Select

                        Case "Gaz"
                            SelectedName = "SOUT_GA"
                            Select Case Me.cboSigne.Text
                                Case "Manchon"
                                    SelectSelectedName = "NAT32"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT30"
                                Case "R�duction"
                                    SelectSelectedName = "NAT33"
                                Case "Vanne"
                                    SelectSelectedName = "NAT31"
                            End Select

                        Case "TV"
                            SelectedName = "SOUT_TV"
                            Select Case Me.cboSigne.Text
                                Case "Antenne"
                                    SelectSelectedName = "NAT52"
                                Case "Armoire"
                                    SelectSelectedName = "NAT51"
                                Case "Chambre"
                                    SelectSelectedName = "NAT53"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT50"
                            End Select

                        Case "Eaux claires" ', "E. C. (projet)"
                            SelectedName = "SOUT_EC"
                            Select Case Me.cboSigne.Text
                                Case "Chambre"
                                    SelectSelectedName = "NAT71"
                                Case "Chambre enterr�e"
                                    SelectSelectedName = "NAT72"
                                    'Case "Cheneau": .tvNatures.Nodes("NAT79").Selected = True
                                Case "D�versoir"
                                    SelectSelectedName = "NAT79"
                                Case "Exutoire"
                                    SelectSelectedName = "NAT77"
                                Case "Grille"
                                    SelectSelectedName = "NAT73"
                                Case "Gueulard"
                                    SelectSelectedName = "NAT78"
                                Case "Ouvrage sp�cial"
                                    SelectSelectedName = "NAT75"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT70"
                                Case "Relevage"
                                    SelectSelectedName = "NAT76"
                                Case "S�parateur"
                                    SelectSelectedName = "NAT74"
                            End Select

                        Case "Eaux us�es" ', "E. U. (projet)"
                            SelectedName = "SOUT_EU"
                            Select Case Me.cboSigne.Text
                                Case "Chambre"
                                    SelectSelectedName = "NAT81"
                                Case "Chambre enterr�e"
                                    SelectSelectedName = "NAT82"
                                Case "Fosse septique"
                                    SelectSelectedName = "NAT83"
                                Case "Ouvrage sp�cial"
                                    SelectSelectedName = "NAT85"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT7480"
                                Case "Relevage"
                                    SelectSelectedName = "NAT86"
                                Case "S�parateur"
                                    SelectSelectedName = "NAT84"
                            End Select

                        Case "Eaux unitaires"
                            SelectedName = "SOUT_EM"
                            Select Case Me.cboSigne.Text
                                Case "Chambre"
                                    SelectSelectedName = "NAT91"
                                Case "Chambre enterr�e"
                                    SelectSelectedName = "NAT92"
                                Case "Ouvrage sp�cial"
                                    SelectSelectedName = "NAT95"
                                Case "Point de rep�rage"
                                    SelectSelectedName = "NAT90"
                            End Select
                    End Select


            End Select

            Dim NodeExp() As System.Windows.Forms.TreeNode = frmCodeNat.tvNatures.Nodes.Find(SelectedName, True)
            NodeExp(0).Expand()
            NodeExp = frmCodeNat.tvNatures.Nodes.Find(SelectSelectedName, True)
            .tvNatures.SelectedNode = NodeExp(0)

            .ShowDialog()

        End With

        If SelectedNature <> "0" And SelectedNature <> "" Then NatureToPointEdition(SelectedNature)



    End Sub



    Private Sub cmdNext_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmdNext.Click
        Dim j As Object

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
        Dim acDoc As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If


        Dim SaisieValide As Boolean = True
        Dim MessageErr As String = "Les �l�ments suivant ne sont pas d�finis correctement :"

        'Test de la saisie
        '-----------------

        'Coordonn�e Est saisie et num�rique ?
        If Me.txtX.Text <> "" Then
            If ValueIsDouble((Me.txtX.Text)) = False Then Me.txtX.Focus() : SaisieValide = False
        Else
            MessageErr += vbCrLf & " - La coordonn�e Est" : Me.txtX.Focus() : SaisieValide = False
        End If
        'Coordonn�e Nord saisie et num�rique ?
        If Me.txtY.Text <> "" Then
            If ValueIsDouble((Me.txtY.Text)) = False Then Me.txtY.Focus() : SaisieValide = False
        Else
            MessageErr += vbCrLf & " - La coordonn�e Nord" : Me.txtY.Focus() : SaisieValide = False
        End If
        'Altitude num�rique ?
        If Me.txtZ.Text <> "" Then
            If ValueIsDouble((Me.txtZ.Text)) = False Then Me.txtZ.Focus() : SaisieValide = False
        Else
            'Non saisie => 0
            Me.txtZ.Text = "0.000"
        End If

        'Orientation num�rique ?
        If Me.txtOri.Text <> "" Then
            If ValueIsDouble((Me.txtOri.Text)) = False Then Me.txtOri.Focus() : SaisieValide = False
        Else
            'Non saisie => 100
            Me.txtOri.Text = "100.0"
        End If


        'Commune saisie et num�rique ?
        If Me.txtCom.Visible = True And Me.txtCom.Enabled = True Then
            If Me.txtCom.Text <> "" Then
                If ValueIsInteger((Me.txtCom.Text)) = False Then Me.txtCom.Focus() : SaisieValide = False
            Else
                MessageErr += vbCrLf & " - Le num�ro de commune" : Me.txtCom.Focus() : SaisieValide = False
            End If
        End If
        'Plan saisi et num�rique ?
        If Me.txtPlan.Visible = True And Me.txtPlan.Enabled = True Then
            If Me.txtPlan.Text <> "" Then
                If ValueIsInteger((Me.txtPlan.Text)) = False Then Me.txtPlan.Focus() : SaisieValide = False
            Else
                If Me.lblPlan.Text = "Plan:" Then
                    MessageErr += vbCrLf & " - Le num�ro de plan" : Me.txtPlan.Focus() : SaisieValide = False
                Else
                    MessageErr += vbCrLf & " - Le num�ro de carte nationale" : Me.txtPlan.Focus() : SaisieValide = False
                End If
            End If
        End If
        'Num�ro de point saisi ?
        If Me.txtNo.Text = "" Then
            MessageErr += vbCrLf & " - Le num�ro de point" : Me.txtNo.Focus() : SaisieValide = False
        End If

        'PrecPlan num�rique ?
        If Me.txtPrecPlan.Text <> "" Then
            If ValueIsDouble((Me.txtPrecPlan.Text)) = False Then Me.txtPrecPlan.Focus() : SaisieValide = False
        End If

        'PrecAlt num�rique ?
        If Me.txtPrecAlt.Visible = True And Me.txtPrecAlt.Enabled = True Then
            If Me.txtPrecAlt.Text <> "" Then
                If ValueIsDouble((Me.txtPrecAlt.Text)) = False Then Me.txtPrecAlt.Focus() : SaisieValide = False
            End If
        End If



        '----------------------------
        If SaisieValide Then

            'D�termination du calque et du bloc concern�s
            Dim strLayer As String = ""
            Dim strBlock As String = ""

            'Selon cat�gorie
            Select Case Me.cboCat.Text

                Case "MO", "MUT" 'MO_xxx ou MUT_xxx

                    'Selon Th�me
                    Select Case Me.cboTheme.SelectedIndex

                        Case 0 '1 PF
                            strLayer = Me.cboCat.Text & "_" & Me.cboType.Text

                            'Bloc selon type du point et accessibilite
                            Select Case Me.cboType.Text
                                Case "PFP1", "PFP2"
                                    If Me.cboAttrVariable.Text = "accessible" Then
                                        strBlock = Me.cboType.Text & "_stationnable"
                                    Else
                                        strBlock = Me.cboType.Text & "_inaccessible"
                                    End If
                                Case "PFP3"
                                    'Bloc selon nature du point
                                    Select Case Me.cboSigne.Text
                                        Case "borne" : strBlock = "PFP3_Borne"
                                        Case "cheville" : strBlock = "PFP3_Cheville"
                                        Case "croix" : strBlock = "PFP3_Croix"
                                        Case "pieu" : strBlock = "PFP3_Pieu"
                                    End Select
                                Case "PFA1", "PFA2", "PFA3"
                                    strBlock = Me.cboType.Text
                            End Select

                        Case 1 '2 CS
                            strLayer = Me.cboCat.Text & "_CS_PTS"
                            strBlock = "PTS_CS"

                        Case 2 '3 OD
                            strLayer = Me.cboCat.Text & "_OD_PTS"
                            strBlock = "PTS_OD"

                        Case 3 '6 BF
                            strLayer = Me.cboCat.Text & "_BF_PTS"

                            'Bloc selon nature du point
                            Select Case Me.cboSigne.Text
                                Case "borne" : strBlock = "Pts_BF_Borne"
                                Case "cheville" : strBlock = "Pts_BF_Cheville"
                                Case "croix" : strBlock = "Pts_BF_Croix"
                                Case "pieu" : strBlock = "Pts_BF_Pieu"
                                Case "non materialise"
                                    'If Me.cboExact.Text = "oui" Then
                                    strBlock = "Pts_BF_Non_mate_defini"
                                    'Else
                                    '    strBlock = "Pts_BF_Non_mate_nondefini"
                                    'End If
                                Case "non materialise non defini"
                                    'If Me.cboExact.Text = "oui" Then
                                    '    strBlock = "Pts_BF_Non_mate_defini"
                                    'Else
                                    strBlock = "Pts_BF_Non_mate_nondefini"
                                    'End If
                            End Select

                        Case 4 '9 COM
                            strLayer = Me.cboCat.Text & "_COM_PTS"

                            'Bloc selon nature du point
                            Select Case Me.cboSigne.Text
                                Case "borne" : strBlock = "Pts_COM_Borne"
                                Case "cheville" : strBlock = "Pts_COM_Cheville"
                                Case "croix" : strBlock = "Pts_COM_Croix"
                                Case "pieu" : strBlock = "Pts_COM_Pieu"
                                Case "non materialise"
                                    'If Me.cboExact.Text = "oui" Then
                                    strBlock = "Pts_COM_Non_mate_defini"
                                    'Else
                                    '    strBlock = "Pts_COM_Non_mate_nondefini"
                                    'End If
                                Case "non materialise non defini"
                                    'If Me.cboExact.Text = "oui" Then
                                    '    strBlock = "Pts_COM_Non_mate_defini"
                                    'Else
                                    strBlock = "Pts_COM_Non_mate_nondefini"
                                    'End If
                            End Select

                    End Select

                    'Pour le bloc, on ajoute "MUT_" pour la cat�gorie mutation, sinon c'est MO
                    If Me.cboCat.Text = "MUT" Then strBlock = "MUT_" & strBlock


                Case "TOPO"
                    'Bloc et calque selon signe
                    Select Case Me.cboSigne.Text
                        Case "point terrain"
                            strBlock = "TOPO_B_PTS"
                            strLayer = "TOPO_PTS"
                        Case "arbre"
                            strBlock = "TOPO_B_ARBRE"
                            strLayer = "TOPO_ARBRE"

                            'Topographie +
                        Case "pied de mur" '88	Pied de mur
                            strBlock = "TOPO_B_PIED_MUR"
                            strLayer = "TOPO_PTS"
                        Case "haut de mur" '89	Haut de mur
                            strBlock = "TOPO_B_HAUT_MUR"
                            strLayer = "TOPO_PTS"
                        Case "bord de route" '93	Bord de route
                            strBlock = "TOPO_B_BORD_ROUTE"
                            strLayer = "TOPO_PTS"
                        Case "trottoir" '94 Trottoir
                            strBlock = "TOPO_B_TROTTOIR"
                            strLayer = "TOPO_PTS"
                        Case "escalier" '96 Escalier
                            strBlock = "TOPO_B_ESCALIER"
                            strLayer = "TOPO_PTS"
                        Case "point cot�" '97 Point cot�
                            strBlock = "TOPO_B_PTS_COTE"
                            strLayer = "TOPO_PTS"
                        Case "point de ligne de rupture" '98 Point de ligne de rupture
                            strBlock = "TOPO_B_PTS_LIGNE_RUPTURE"
                            strLayer = "TOPO_PTS"
                        Case "fa�te, corniche" '99 Fa�te, corniche
                            strBlock = "TOPO_B_FAITE_CORNICHE"
                            strLayer = "TOPO_PTS"
                    End Select

                Case "IMPL"
                    'Bloc et calque selon signe
                    Select Case Me.cboSigne.Text
                        Case "controle"
                            strBlock = "IMPL_B_PTS_CONTROL"
                            strLayer = "IMPL_CONT_PTS"
                        Case "projet"
                            strBlock = "IMPL_B_PTS_PROJET"
                            strLayer = "IMPL_PROJ_PTS"
                    End Select

                Case "SOUT"
                    'Bloc et calque selon th�me et signe
                    Select Case Me.cboTheme.SelectedIndex

                        Case 8 'Electricit�
                            Select Case Me.cboSigne.SelectedIndex
                                Case 1
                                    strBlock = "SOUT_B_EL_PR"
                                    strLayer = "SOUT_EL_PR"
                                Case 0
                                    strBlock = "SOUT_B_EL_CD"
                                    strLayer = "SOUT_EL_CD"
                                Case 2
                                    strBlock = "SOUT_B_EL_POT"
                                    strLayer = "SOUT_EL_POT"
                            End Select

                        Case 10 'T�l�phone
                            Select Case Me.cboSigne.SelectedIndex
                                Case 2
                                    strBlock = "SOUT_B_TT_PR"
                                    strLayer = "SOUT_TT_PR"
                                Case 0
                                    strBlock = "SOUT_B_TT_DTT"
                                    strLayer = "SOUT_TT_DTT"
                                Case 1
                                    strBlock = "SOUT_B_TT_RTT"
                                    strLayer = "SOUT_TT_RTT"
                                Case 3
                                    strBlock = "SOUT_B_TT_POT"
                                    strLayer = "SOUT_TT_POT"
                            End Select

                        Case 7 'Eclairage public
                            Select Case Me.cboSigne.SelectedIndex
                                Case 3
                                    strBlock = "SOUT_B_EG_PR"
                                    strLayer = "SOUT_EG_PR"
                                Case 0
                                    strBlock = "SOUT_B_EG_CD"
                                    strLayer = "SOUT_EG_CD"
                                Case 1
                                    strBlock = "SOUT_B_EG_CB"
                                    strLayer = "SOUT_EG_CB"
                                Case 2
                                    strBlock = "SOUT_B_EG_CE"
                                    strLayer = "SOUT_EG_CE"
                            End Select

                        Case 9 'Gaz
                            Select Case Me.cboSigne.SelectedIndex
                                Case 1
                                    strBlock = "SOUT_B_GA_PR"
                                    strLayer = "SOUT_GA_PR"
                                Case 3
                                    strBlock = "SOUT_B_GA_VA"
                                    strLayer = "SOUT_GA_VA"
                                Case 0
                                    strBlock = "SOUT_B_GA_MA"
                                    strLayer = "SOUT_GA_MA"
                                Case 2
                                    strBlock = "SOUT_B_GA_RE"
                                    strLayer = "SOUT_GA_RE"
                            End Select

                        Case 0 'Eau potable
                            Select Case Me.cboSigne.SelectedIndex
                                Case 5
                                    strBlock = "SOUT_B_EP_PR"
                                    strLayer = "SOUT_EP_PR"
                                Case 9
                                    strBlock = "SOUT_B_EP_VA"
                                    strLayer = "SOUT_EP_VA"
                                Case 2
                                    strBlock = "SOUT_B_EP_CVA"
                                    strLayer = "SOUT_EP_CVA"
                                Case 0
                                    strBlock = "SOUT_B_EP_BH"
                                    strLayer = "SOUT_EP_BH"
                                Case 4
                                    strBlock = "SOUT_B_EP_MA"
                                    strLayer = "SOUT_EP_MA"
                                Case 6
                                    strBlock = "SOUT_B_EP_PU"
                                    strLayer = "SOUT_EP_PU"
                                Case 3
                                    strBlock = "SOUT_B_EP_CO"
                                    strLayer = "SOUT_EP_CO"
                                Case 7
                                    strBlock = "SOUT_B_EP_RE"
                                    strLayer = "SOUT_EP_RE"
                                Case 1
                                    strBlock = "SOUT_B_EP_CC"
                                    strLayer = "SOUT_EP_CC"
                                Case 8
                                    strBlock = "SOUT_B_EP_FO"
                                    strLayer = "SOUT_EP_FO"
                            End Select

                        Case 1 'Eau potable projet
                            Select Case Me.cboSigne.SelectedIndex
                                Case 5
                                    strBlock = "SOUT_B_PROJ_EP_PR"
                                    strLayer = "SOUT_PROJ_EP_PR"
                                Case 9
                                    strBlock = "SOUT_B_PROJ_EP_VA"
                                    strLayer = "SOUT_PROJ_EP_VA"
                                Case 2
                                    strBlock = "SOUT_B_PROJ_EP_CVA"
                                    strLayer = "SOUT_PROJ_EP_CVA"
                                Case 0
                                    strBlock = "SOUT_B_PROJ_EP_BH"
                                    strLayer = "SOUT_PROJ_EP_BH"
                                Case 4
                                    strBlock = "SOUT_B_PROJ_EP_MA"
                                    strLayer = "SOUT_PROJ_EP_MA"
                                Case 6
                                    strBlock = "SOUT_B_PROJ_EP_PU"
                                    strLayer = "SOUT_PROJ_EP_PU"
                                Case 3
                                    strBlock = "SOUT_B_PROJ_EP_CO"
                                    strLayer = "SOUT_PROJ_EP_CO"
                                Case 7
                                    strBlock = "SOUT_B_PROJ_EP_RE"
                                    strLayer = "SOUT_PROJ_EP_RE"
                                Case 1
                                    strBlock = "SOUT_B_PROJ_EP_CC"
                                    strLayer = "SOUT_PROJ_EP_CC"
                                Case 8
                                    strBlock = "SOUT_B_PROJ_EP_FO"
                                    strLayer = "SOUT_PROJ_EP_FO"
                            End Select

                        Case 11 'TV
                            Select Case Me.cboSigne.SelectedIndex
                                Case 3
                                    strBlock = "SOUT_B_TV_PR"
                                    strLayer = "SOUT_TV_PR"
                                Case 1
                                    strBlock = "SOUT_B_TV_ART"
                                    strLayer = "SOUT_TV_ART"
                                Case 0
                                    strBlock = "SOUT_B_TV_ATT"
                                    strLayer = "SOUT_TV_ATT"
                                Case 2
                                    strBlock = "SOUT_B_TV_CT"
                                    strLayer = "SOUT_TV_CT"
                            End Select

                        Case 2 'Eaux claires
                            Select Case Me.cboSigne.SelectedIndex
                                Case 8
                                    strBlock = "SOUT_B_EC_PR"
                                    strLayer = "SOUT_EC_PR"
                                Case 0
                                    strBlock = "SOUT_B_EC_CH"
                                    strLayer = "SOUT_EC_CH"
                                Case 1
                                    strBlock = "SOUT_B_EC_CE"
                                    strLayer = "SOUT_EC_CE"
                                Case 5
                                    strBlock = "SOUT_B_EC_GR"
                                    strLayer = "SOUT_EC_GR"
                                Case 10
                                    strBlock = "SOUT_B_EC_SE"
                                    strLayer = "SOUT_EC_SE"
                                Case 7
                                    strBlock = "SOUT_B_EC_OS"
                                    strLayer = "SOUT_EC_OS"
                                Case 9
                                    strBlock = "SOUT_B_EC_RE"
                                    strLayer = "SOUT_EC_RE"
                                Case 4
                                    strBlock = "SOUT_B_EC_EX"
                                    strLayer = "SOUT_EC_EX"
                                Case 6
                                    strBlock = "SOUT_B_EC_GU"
                                    strLayer = "SOUT_EC_GU"
                                Case 3
                                    strBlock = "SOUT_B_EC_DE"
                                    strLayer = "SOUT_EC_DE"
                                Case 2
                                    strBlock = "SOUT_B_EC_CU"
                                    strLayer = "SOUT_EC_CU"
                            End Select

                        Case 3 'Eaux claires projet
                            Select Case Me.cboSigne.SelectedIndex
                                Case 8
                                    strBlock = "SOUT_B_PROJ_EC_PR"
                                    strLayer = "SOUT_PROJ_EC_PR"
                                Case 0
                                    strBlock = "SOUT_B_PROJ_EC_CH"
                                    strLayer = "SOUT_PROJ_EC_CH"
                                Case 1
                                    strBlock = "SOUT_B_PROJ_EC_CE"
                                    strLayer = "SOUT_PROJ_EC_CE"
                                Case 5
                                    strBlock = "SOUT_B_PROJ_EC_GR"
                                    strLayer = "SOUT_PROJ_EC_GR"
                                Case 10
                                    strBlock = "SOUT_B_PROJ_EC_SE"
                                    strLayer = "SOUT_PROJ_EC_SE"
                                Case 7
                                    strBlock = "SOUT_B_PROJ_EC_OS"
                                    strLayer = "SOUT_PROJ_EC_OS"
                                Case 9
                                    strBlock = "SOUT_B_PROJ_EC_RE"
                                    strLayer = "SOUT_PROJ_EC_RE"
                                Case 4
                                    strBlock = "SOUT_B_PROJ_EC_EX"
                                    strLayer = "SOUT_PROJ_EC_EX"
                                Case 6
                                    strBlock = "SOUT_B_PROJ_EC_GU"
                                    strLayer = "SOUT_PROJ_EC_GU"
                                Case 3
                                    strBlock = "SOUT_B_PROJ_EC_DE"
                                    strLayer = "SOUT_PROJ_EC_DE"
                                Case 2
                                    strBlock = "SOUT_B_PROJ_EC_CU"
                                    strLayer = "SOUT_PROJ_EC_CU"
                            End Select

                        Case 5 'Eaux us�es
                            Select Case Me.cboSigne.SelectedIndex
                                Case 4
                                    strBlock = "SOUT_B_EU_PR"
                                    strLayer = "SOUT_EU_PR"
                                Case 0
                                    strBlock = "SOUT_B_EU_CH"
                                    strLayer = "SOUT_EU_CH"
                                Case 1
                                    strBlock = "SOUT_B_EU_CE"
                                    strLayer = "SOUT_EU_CE"
                                Case 2
                                    strBlock = "SOUT_B_EU_FS"
                                    strLayer = "SOUT_EU_FS"
                                Case 6
                                    strBlock = "SOUT_B_EU_SE"
                                    strLayer = "SOUT_EU_SE"
                                Case 3
                                    strBlock = "SOUT_B_EU_OS"
                                    strLayer = "SOUT_EU_OS"
                                Case 5
                                    strBlock = "SOUT_B_EU_RE"
                                    strLayer = "SOUT_EU_RE"
                            End Select

                        Case 6 'Eaux us�es projet�es
                            Select Case Me.cboSigne.SelectedIndex
                                Case 4
                                    strBlock = "SOUT_B_PROJ_EU_PR"
                                    strLayer = "SOUT_PROJ_EU_PR"
                                Case 0
                                    strBlock = "SOUT_B_PROJ_EU_CH"
                                    strLayer = "SOUT_PROJ_EU_CH"
                                Case 1
                                    strBlock = "SOUT_B_PROJ_EU_CE"
                                    strLayer = "SOUT_PROJ_EU_CE"
                                Case 2
                                    strBlock = "SOUT_B_PROJ_EU_FS"
                                    strLayer = "SOUT_PROJ_EU_FS"
                                Case 6
                                    strBlock = "SOUT_B_PROJ_EU_SE"
                                    strLayer = "SOUT_PROJ_EU_SE"
                                Case 3
                                    strBlock = "SOUT_B_PROJ_EU_OS"
                                    strLayer = "SOUT_PROJ_EU_OS"
                                Case 5
                                    strBlock = "SOUT_B_PROJ_EU_RE"
                                    strLayer = "SOUT_PROJ_EU_RE"
                            End Select

                        Case 4 'Eaux unitaires
                            Select Case Me.cboSigne.SelectedIndex
                                Case 3
                                    strBlock = "SOUT_B_EM_PR"
                                    strLayer = "SOUT_EM_PR"
                                Case 0
                                    strBlock = "SOUT_B_EM_CH"
                                    strLayer = "SOUT_EM_CH"
                                Case 1
                                    strBlock = "SOUT_B_EM_CE"
                                    strLayer = "SOUT_EM_CE"
                                Case 2
                                    strBlock = "SOUT_B_EM_OS"
                                    strLayer = "SOUT_EM_OS"
                            End Select

                    End Select

            End Select


            'Lib�ration et activation du calque cible
            Call FreezeLayer(strLayer, False)
            Call ActivateLayer(strLayer)

            'Attributs g�om�triques
            '  Dim dblInsertPt As Autodesk.AutoCAD.Geometry.Point3d
            Dim dblOri, dblScale As Double

            Dim dblInsertPt(2) As Double 'New Autodesk.AutoCAD.Geometry.Point3d(Val(Me.txtX.Text), Val(Me.txtY.Text), Val(Me.txtZ.Text))
            dblInsertPt(0) = CDbl(Me.txtX.Text)
            dblInsertPt(1) = CDbl(Val(Me.txtY.Text))
            If Alt3D.Checked = True Then
                dblInsertPt(2) = CDbl(Val(Me.txtZ.Text))
            Else
                dblInsertPt(2) = 0
            End If



            ''Traitement de l'altitude
            'If objBlock.InsertionPoint(2) <> 0 Then '  Si Z<>0 -> 3D 
            '    .Alt3D.Checked = True : .Alt2D.Checked = False
            '    .txtZ.Text = Format(objBlock.InsertionPoint(2), "0.000")
            'Else ' Si(Z = 0 -> 2D)
            '    If Altinfo <> 0 Then .txtZ.Text = Format(Altinfo, "0.000") '(Altinfo())
            'End If

            dblOri = Val(Me.txtOri.Text)
            dblScale = Val(GetCurrentScale()) / 1000


            '----------------------

            'Bloc � cr�er ou �diter
            Dim objBlockRef As Autodesk.AutoCAD.Interop.Common.AcadBlockReference

            'EDITION
            Try
                If Me.Tag = "EDIT" Then

                    'Obtention d'objet actuel en fonction du Handle
                    objBlockRef = acDoc.HandleToObject(Me.txtBlockID.Text)

                    'Cas particulier: si il y a un changement de couche, on recr�e le boc, on ne l'�dite pas (les infos dans les autres couches, par ex. NUM, ne suivent pas)
                    If strLayer.ToUpper <> objBlockRef.Layer.ToString.ToUpper Then

                        Dim RotBL As Double = objBlockRef.Rotation
                        'Suppression de l'ancien bloc              
                        objBlockRef.Delete()
                        'Insertion d'un nouveau bloc
                        'objBlockRef = InsertBlock(dblInsertPt, strBlock, dblOri, 0, dblScale)

#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                        Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                        Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
                        'objBlockRef = acDocX.ModelSpace.InsertBlock(dblInsertPt, strBlock, dblScale, dblScale, dblScale, ConvGisTopoTrigo((dblOri) * (Math.PI / 200)))
                        objBlockRef = acDocX.ModelSpace.InsertBlock(dblInsertPt, strBlock, dblScale, dblScale, dblScale, RotBL)
                        objBlockRef.Layer = strLayer 'Calque

                    Else
                        'Mise � jour des attributs g�n�raux du bloc
                        objBlockRef.Name = strBlock 'Type de bloc 
                        objBlockRef.Layer = strLayer 'Calque
                        objBlockRef.Rotation = ConvGisTopoTrigo((dblOri) * (Math.PI / 200)) 'Rotation
                        objBlockRef.XScaleFactor = dblScale 'Echelle
                        objBlockRef.YScaleFactor = dblScale
                        objBlockRef.ZScaleFactor = dblScale

                        'Dim NewInsertPts As Object

                        '  objBlockRef.InsertionPoint = dblInsertPt 'Point d'insertion
                        objBlockRef.InsertionPoint = dblInsertPt ' dblInsertPt(0)
                        ' objBlockRef.InsertionPoint(1) = dblInsertPt.Y 'Point d'insertion
                        'objBlockRef.InsertionPoint(2) = dblInsertPt.Z 'Point d'insertion

                    End If


                Else 'CREATION

                    'Controle de la version du dessin
                    Dim RVinfo As New Revo.RevoInfo

                    If InterlisProjectTest() = False Then 'Format du fichier non compatible
                        Dim RVscript As New Revo.RevoScript
                        If IO.File.Exists(RVinfo.Template) Then
                            RVscript.cmdBL("#BL;>" & RVinfo.Template & ";") ' "c:\test.dwg"
                        Else
                            MsgBox("Impossible de charger la biblioth�que de bloc, il est n�cessaire de disposer d'un gabarit compatible" & vbCrLf _
                                                                           & "Relancer le flux de travail dans un dessin compatible. (cr�er un nouveau projet)", MsgBoxStyle.Information + MsgBoxStyle.OkOnly, "Impossible de lancer le module")
                        End If
                    End If


                    'Insertion du bloc
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                    Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                    Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If
                    objBlockRef = acDocX.ModelSpace.InsertBlock(dblInsertPt, strBlock, dblScale, dblScale, dblScale, ConvGisTopoTrigo((dblOri) * (Math.PI / 200)))
                End If


                '----------------------


                'Ecriture des attributs
                Dim varAttributes As Object
                Dim strNomAttr As String = ""
                varAttributes = objBlockRef.GetAttributes

                For j = LBound(varAttributes) To UBound(varAttributes)

                    'En fonction du nom de l'attribut
                    'UPGRADE_WARNING: Impossible de r�soudre la propri�t� par d�faut de l'objet varAttributes().TagString. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                    Select Case UCase(varAttributes(j).TagString)

                        Case "ALTITUDE" 'Reste vide
                            Dim Alt As Double = CDbl(Me.txtZ.Text)
                            If Alt = 0 Then
                                varAttributes(j).TextString = "?"
                            Else
                                varAttributes(j).TextString = Format(Alt, "0.000")
                            End If
                        Case "IDENTDN"
                            'UPGRADE_WARNING: Impossible de r�soudre la propri�t� par d�faut de l'objet varAttributes().TextString. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                            varAttributes(j).TextString = Me.txtIdentDN.Text
                        Case "NUMERO"
                            varAttributes(j).TextString = Me.txtNo.Text
                        Case "SIGNE"
                            varAttributes(j).TextString = Me.cboSigne.Text
                        Case "ACCESSIBILITE"
                            varAttributes(j).TextString = Me.cboAttrVariable.Text
                        Case "FICHE"
                            varAttributes(j).TextString = Me.cboAttrVariable.Text
                        Case "ANC_BORNE_SPECIAL"
                            varAttributes(j).TextString = Me.cboAttrVariable.Text
                        Case "DEFINI_EXACTEMENT"
                            varAttributes(j).TextString = Me.cboExact.Text
                        Case "BORNE_TERRITORIALE"
                            varAttributes(j).TextString = Me.cboAttrVariable.Text
                        Case "PRECPLAN"
                            If Me.txtPrecPlan.Text <> "" Then varAttributes(j).TextString = Me.txtPrecPlan.Text
                        Case "PRECALT"
                            If Me.txtPrecAlt.Text <> "" Then varAttributes(j).TextString = Me.txtPrecAlt.Text
                        Case "FIABPLAN"
                            varAttributes(j).TextString = Me.cboFiabPlan.Text
                        Case "FIABALT"
                            varAttributes(j).TextString = Me.cboFiabAlt.Text
                        Case "COMMENTAIRE"
                            If Me.txtComment.Text <> "" Then varAttributes(j).TextString = Me.txtComment.Text
                    End Select

                Next


            Catch ex As Exception
                MsgBox(ex.Message & " : Si le probl�me est persistant fermez et re-ouvrez le dessin.", MsgBoxStyle.Critical)
            End Try


            'Point ins�r�
            'Beep

            Dim Connect As New Revo.connect
            Dim Ass As New Revo.RevoInfo
            'M�morise les param�tres pour la prochaine ouverture du formulaire (valeurs par d�faut)
            intInsPtCat = Me.cboCat.SelectedIndex : Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "intinsptcat", Me.cboCat.SelectedIndex)
            intInsPtTheme = Me.cboTheme.SelectedIndex : Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "intinspttheme", Me.cboTheme.SelectedIndex)
            intInsPtType = Me.cboType.SelectedIndex : Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "intinspttype", Me.cboType.SelectedIndex)
            intInsPtSigne = Me.cboSigne.SelectedIndex : Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "intinsptsigne", Me.cboSigne.SelectedIndex)
            intInsPtDefini = Me.cboExact.SelectedIndex : Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "intinsptdefini", Me.cboExact.SelectedIndex)
            intInsPtVar = Me.cboAttrVariable.SelectedIndex : Connect.XMLwriteConfig("/" & Ass.xProduct.ToLower & "/config", "intinsptvar", Me.cboAttrVariable.SelectedIndex)
            SetFileProperty("Commune", (Me.txtCom.Text)) : SetFileProperty("Plan", (Me.txtPlan.Text))
            strInsPtNum = Me.txtNo.Text

            intInsPtFiabPlan = Me.cboFiabPlan.SelectedIndex : intInsPtFiabAlt = Me.cboFiabAlt.SelectedIndex
            strInsPtPrecPlan = Me.txtPrecPlan.Text : strInsPtPrecAlt = Me.txtPrecAlt.Text


            'EDITION => fermeture de la fen�tre
            If Me.Tag = "EDIT" Then

                Me.Close()

                'INSERTION => 'R�initialise le formulaire pour la saisie suivante
            Else

                ' Beep()

                'Masque et r�affiche la fen�tre pour que l'utilisateur voie qu'il s'est pass� quelque chose !!
                Me.Hide()
                System.Windows.Forms.Application.DoEvents()

                'Num�ro + 1
                If Me.txtNo.Text = "0" Then '0 => 1
                    Me.txtNo.Text = "1"
                ElseIf CStr(Val(Me.txtNo.Text)) = Me.txtNo.Text Then  'Valeur num�rique
                    Me.txtNo.Text = CStr(Val(Me.txtNo.Text) + 1)
                Else 'Valeur alphanum�rique
                    Me.txtNo.Text = IncrementID((Me.txtNo.Text))
                End If

                Me.txtX.Text = ""
                Me.txtY.Text = ""
                Me.txtZ.Text = ""

                Me.txtNo.Focus()

                'Saisie standard
                If Me.chkDigitSerie.CheckState = False Then
                    'Affichage de la bo�te de dialogue pour la saisie suivante
                    'Me.ShowDialog()

                    'Saisie en s�rie
                Else
                    'Continue la digitalisation
                    Call cmdDigit_Click(cmdDigit, New System.EventArgs())
                End If

            End If

            Exit Sub

        Else
            MsgBox(MessageErr, MsgBoxStyle.Information + vbOKOnly, "Saisie incompl�te")
        End If 'Fin de test : DonneeValide

            'Active le formulaire si saisie en s�rie
            If Me.chkDigitSerie.CheckState = True Then
                'frmInsertionEnSerie.Close()
                'On Error Resume Next
                Me.ShowDialog()
                'On Error GoTo 0
            End If
    End Sub


    Private Function IncrementID(ByRef strNum As String) As String

        Dim strChar, strNewChar As String
        Dim booAddChar As Boolean
        Dim intLen As Short : intLen = Len(strNum)
        Dim intPos As Short : intPos = Len(strNum)
        Dim strNewNum As String
        Dim strCharType As String = ""
        Dim strCharTypePrec As String

        Do
            'Caract�re � analyser
            strChar = Mid(strNum, intPos, 1)

            'Type de caract�re (0-9, A-Z, a-z, autre)
            strCharTypePrec = strCharType
            strCharType = GetCharType(strChar)

            'Si ajout d'une dizaine/centaine/milier/etc, on regarder si c'est le m�me type de donn�e ou s'il faut ajouter un caract�re
            ' EX: A23 -> A24, A9 -> A10, AB -> AC, 9Z->9AA
            If strCharTypePrec <> "" Then 'Pas de test pour le dernier caract�re (premi�re analyse)

                If strCharTypePrec <> strCharType Then
                    'Type diff�rent => on ins�re un caract�re
                    intLen = intLen + 1
                    intPos = intPos + 1
                    booAddChar = False
                    Select Case strCharTypePrec
                        Case "09"
                            strNewChar = "1"
                        Case "AZ"
                            strNewChar = "A"
                        Case "az"
                            strNewChar = "a"
                        Case Else
                            strNewChar = "0" 'Ne devrait pas se produire
                    End Select
                Else
                    'M�me type => on incr�mente la dizaine/centaine/etc
                    strNewChar = IncrementChar(strChar, booAddChar)
                End If

            Else
                '+1
                strNewChar = IncrementChar(strChar, booAddChar)
            End If


            'Remplacement du caract�re
IncrementOk:
            If intPos = 1 Then
                strNewNum = strNewChar & VB.Right(strNum, intLen - 1)
            ElseIf intPos = intLen Then
                strNewNum = VB.Left(strNum, intLen - 1) & strNewChar
            Else
                strNewNum = VB.Left(strNum, intPos - 1) & strNewChar & VB.Right(strNum, intLen - intPos)
            End If

            'Ajouter un chiffre des dizaines ?
            If booAddChar = True Then
                intPos = intPos - 1
                strNum = strNewNum
            Else
                Exit Do
            End If
        Loop

        IncrementID = strNewNum

    End Function

    Private Function GetCharType(ByRef strChar As String) As String

        Dim intChar As Short
        intChar = Asc(strChar)

        If intChar >= 48 And intChar <= 57 Then
            GetCharType = "09"
        ElseIf intChar >= 65 And intChar <= 90 Then
            GetCharType = "AZ"
        ElseIf intChar >= 97 And intChar <= 122 Then
            GetCharType = "az"
        Else
            GetCharType = "**"
        End If

    End Function

    Private Function IncrementChar(ByRef strChar As String, ByRef booAddChar As Boolean) As String

        Dim intChar As Short 'Dernier caract�re
        intChar = Asc(strChar)

        Dim strNewChar As String
        strNewChar = Chr(intChar + 1)

        If strNewChar = ":" Then ': apr�s 9 (chiffre)
            strNewChar = "0"
            booAddChar = True
        ElseIf strNewChar = "[" Then  '[ apr�s Z
            strNewChar = "A"
            booAddChar = True
        ElseIf strNewChar = "{" Then  '{ apr�s z
            strNewChar = "a"
            booAddChar = True
        Else
            booAddChar = False
        End If

        IncrementChar = strNewChar

    End Function


    'UPGRADE_WARNING: L'�v�nement txtCom.TextChanged peut se d�clencher lorsque le formulaire est initialis�. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtCom_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtCom.TextChanged

        'IdentDN
        If Me.txtCom.Enabled = True Then
            If Me.txtCom.Text <> "" Then
                Me.txtIdentDN.Text = VB.Left(Me.txtIdentDN.Text, 3) & Format(Val(Me.txtCom.Text), "000") & VB.Right(Me.txtIdentDN.Text, 6)
            Else
                Me.txtIdentDN.Text = VB.Left(Me.txtIdentDN.Text, 3) & "CCC" & VB.Right(Me.txtIdentDN.Text, 6)
            End If
        End If

    End Sub

    'UPGRADE_WARNING: L'�v�nement txtNo.TextChanged peut se d�clencher lorsque le formulaire est initialis�. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtNo_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtNo.TextChanged

        'PFP1/2 => valeur par d�fatu de l'accessibilit� selon dernier chiffre du num�ro (norme 1001.1)
        If Me.lblAttrVariable.Text = "Accessible:" Then

            Select Case VB.Right(Me.txtNo.Text, 1)

                Case "0", "1", "2", "3", "4", "5", "6"
                    Me.cboAttrVariable.SelectedIndex = 1
                Case "7", "8", "9"
                    Me.cboAttrVariable.SelectedIndex = 2
                Case Else
                    Me.cboAttrVariable.SelectedIndex = 0

            End Select

        End If

    End Sub

    'UPGRADE_WARNING: L'�v�nement txtPlan.TextChanged peut se d�clencher lorsque le formulaire est initialis�. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="88B12AE1-6DE0-48A0-86F1-60C0686C026A"'
    Private Sub txtPlan_TextChanged(ByVal eventSender As System.Object, ByVal eventArgs As System.EventArgs) Handles txtPlan.TextChanged

        'IdentDN
        If Me.txtPlan.Enabled = True Then
            If Me.txtPlan.Text <> "" Then
                Me.txtIdentDN.Text = VB.Left(Me.txtIdentDN.Text, 8) & Format(Val(Me.txtPlan.Text), "0000")
            Else
                Me.txtIdentDN.Text = VB.Left(Me.txtIdentDN.Text, 8) & "PPPP"
            End If
        End If

    End Sub

    Private Sub LVPlans_MouseClick(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles LVPlans.MouseClick
        Try
            'Entit� s�lectionn�e => valeurs par d�faut
            Me.txtCom.Text = LVPlans.SelectedItems(0).Text.ToString
            Me.txtPlan.Text = LVPlans.SelectedItems(0).SubItems(1).Text.ToString
        Catch
        End Try

    End Sub

    Private Sub LVPlans_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LVPlans.SelectedIndexChanged
        'Entit� s�lectionn�e => valeurs par d�faut
        Try
            Me.txtCom.Text = LVPlans.SelectedItems(0).Text.ToString
            Me.txtPlan.Text = LVPlans.SelectedItems(0).SubItems(1).Text.ToString
        Catch
        End Try
    End Sub






    Private Sub CodesNatures(ByVal tvNatures As System.Windows.Forms.TreeView)

        'Images des les ListView
        'Me.tvNatures.ImageList = Me.imgList
        tvNatures.Nodes.Clear()

        'D�clarations
        Dim objNode(15) As System.Windows.Forms.TreeNode
        'Dim strLine As String = ""
        'Dim arrLine() As String
        'Dim intLevel As Short



        'Cr�ation de l'arborescence: liste des calques
        'UPGRADE_WARNING: Impossible de r�soudre la propri�t� par d�faut de l'objet tvNatures.Nodes. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        objNode(0) = tvNatures.Nodes.Add("")
        objNode(0).Text = "Revo"
        'objNode(0).Image = "Fichier"
        objNode(0).Expand()

        'Electricit�
        'If GetJourCADVersion >= 2.5 Then
        objNode(1) = AddNode(tvNatures, objNode(0), "ELECTRICITE", True, RGB(153, 51, 102), "SOUT_EL")
        Call AddNode(tvNatures, objNode(1), "1. Point de rep�rage sur conduite", , RGB(153, 51, 102))
        Call AddNode(tvNatures, objNode(1), "2. Cabine de distribition", , RGB(153, 51, 102))
        Call AddNode(tvNatures, objNode(1), "3. Poteau", , RGB(153, 51, 102))

        ''T�l�phone
        objNode(2) = AddNode(tvNatures, objNode(0), "TELEPHONE", , RGB(153, 204, 0), "SOUT_TT")
        Call AddNode(tvNatures, objNode(2), "4. Point de rep�rage sur conduite", , RGB(153, 204, 0))
        Call AddNode(tvNatures, objNode(2), "5. Chambre � dalle", , RGB(153, 204, 0))
        Call AddNode(tvNatures, objNode(2), "6. Chambre � regard", , RGB(153, 204, 0))
        Call AddNode(tvNatures, objNode(2), "7. Poteau", , RGB(153, 204, 0))
        ' End If

        'Topo
        objNode(3) = AddNode(tvNatures, objNode(0), "TOPOGRAPHIE", , , "TOPO")
        Call AddNode(tvNatures, objNode(3), "8. Point terrain")
        Call AddNode(tvNatures, objNode(3), "9. Arbre")

        'Topo+
        objNode(3) = AddNode(tvNatures, objNode(0), "TOPOGRAPHIE+", , , "TOPO")
        Call AddNode(tvNatures, objNode(3), "88. Pied de mur")
        Call AddNode(tvNatures, objNode(3), "89. Haut de mur")
        Call AddNode(tvNatures, objNode(3), "93. Bord de route")
        Call AddNode(tvNatures, objNode(3), "94. Trottoir")
        'Call AddNode(tvNatures, objNode(3), "94. Arbre", , , "NAT194")
        Call AddNode(tvNatures, objNode(3), "96. Escalier")
        Call AddNode(tvNatures, objNode(3), "97. Point cot�")
        'Call AddNode(tvNatures, objNode(3), "97. Point terrain", , , "NAT197")
        Call AddNode(tvNatures, objNode(3), "98. Point de ligne de rupture")
        Call AddNode(tvNatures, objNode(3), "99. Fa�te, corniche")


        'Mensuration PF
        objNode(4) = AddNode(tvNatures, objNode(0), "MENSURATION - POINTS FIXES (Th�me 1 PF)", , , "MO_PF")
        objNode(5) = AddNode(tvNatures, objNode(4), "PFP3", , , "MO_PF_PFP3")
        Call AddNode(tvNatures, objNode(5), "11. Borne")
        Call AddNode(tvNatures, objNode(5), "12. Cheville")
        Call AddNode(tvNatures, objNode(5), "13. Croix")
        Call AddNode(tvNatures, objNode(5), "14. Pieu")
        objNode(5) = AddNode(tvNatures, objNode(4), "PFP1/2", , , "MO_PF_PFP1-2")
        Call AddNode(tvNatures, objNode(5), "16. Borne")
        Call AddNode(tvNatures, objNode(5), "17. Cheville")
        Call AddNode(tvNatures, objNode(5), "18. Croix/Pieu")
        Call AddNode(tvNatures, objNode(5), "19. Non mat�rialis� (point inaccessible)")

        'Eclairage public
        'If GetJourCADVersion >= 2.5 Then
        objNode(5) = AddNode(tvNatures, objNode(0), "ECLAIRAGE PUBLIC", , RGB(255, 102, 0), "SOUT_EG")
        Call AddNode(tvNatures, objNode(5), "20. Point de rep�rage sur conduite", , RGB(255, 102, 0))
        Call AddNode(tvNatures, objNode(5), "21. Cabine de distribution", , RGB(255, 102, 0))
        Call AddNode(tvNatures, objNode(5), "22. Cand�labre", , RGB(255, 102, 0))
        Call AddNode(tvNatures, objNode(5), "23. Chambre d'�clairage", , RGB(255, 102, 0))
        'End If

        'Mensuration CS
        objNode(6) = AddNode(tvNatures, objNode(0), "MENSURATION - COUVERTURE DU SOL (Th�me 2 CS)", , , "MO_CS")
        Call AddNode(tvNatures, objNode(6), "25. Point non mat�rialis�, d�fini exactement")
        Call AddNode(tvNatures, objNode(6), "26. Point non mat�rialis�, non d�fini exactement")

        'Gaz
        'If GetJourCADVersion >= 2.5 Then
        objNode(7) = AddNode(tvNatures, objNode(0), "GAZ", , RGB(255, 153, 0), "SOUT_GA")
        Call AddNode(tvNatures, objNode(7), "30. Point de rep�rage sur conduite", , RGB(255, 153, 0))
        Call AddNode(tvNatures, objNode(7), "31. Vanne", , RGB(255, 153, 0))
        Call AddNode(tvNatures, objNode(7), "32. Manchon", , RGB(255, 153, 0))
        Call AddNode(tvNatures, objNode(7), "33. R�duction", , RGB(255, 153, 0))
        'End If

        'Mensuration OD
        objNode(8) = AddNode(tvNatures, objNode(0), "MENSURATION - OBJETS DIVERS (Th�me 3 OD)", , , "MO_OD")
        Call AddNode(tvNatures, objNode(8), "35. Point non mat�rialis�, d�fini exactement")
        Call AddNode(tvNatures, objNode(8), "36. Point non mat�rialis�, non d�fini exactement")

        'Implantation
        objNode(9) = AddNode(tvNatures, objNode(0), "IMPLANTATION", , , "IMPL")
        Call AddNode(tvNatures, objNode(9), "38. Point � implanter issu du projet")
        Call AddNode(tvNatures, objNode(9), "39. Point issu de l'implantation sur le terrain (contr�le)")

        'Eau potable
        ' If GetJourCADVersion >= 2.5 Then
        objNode(10) = AddNode(tvNatures, objNode(0), "EAU POTABLE", , RGB(51, 153, 102), "SOUT_EP")
        Call AddNode(tvNatures, objNode(10), "40. Point de rep�rage sur conduite", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "41. Vanne", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "42. Chambre de vannes", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "43. Borne hydrante", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "44. Manchon", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "45. Purge", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "46. Compteur", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "47. R�duction", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "48. Chambre de captage", , RGB(51, 153, 102))
        Call AddNode(tvNatures, objNode(10), "49. R�servoir, fontaine", , RGB(51, 153, 102))

        'T�l�vision
        objNode(11) = AddNode(tvNatures, objNode(0), "TELEVISON", , RGB(0, 255, 0), "SOUT_TV")
        Call AddNode(tvNatures, objNode(11), "50. Point de rep�rage sur conduite", , RGB(0, 255, 0))
        Call AddNode(tvNatures, objNode(11), "51. Armoire", , RGB(0, 255, 0))
        Call AddNode(tvNatures, objNode(11), "52. Antenne", , RGB(0, 255, 0))
        Call AddNode(tvNatures, objNode(11), "53. Chambre", , RGB(0, 255, 0))
        ' End If

        'Mensuration BF
        objNode(12) = AddNode(tvNatures, objNode(0), "MENSURATION - BIEN-FONDS (Th�me 6 BF)", , , "MO_BF")
        Call AddNode(tvNatures, objNode(12), "61. Borne")
        Call AddNode(tvNatures, objNode(12), "62. Cheville")
        Call AddNode(tvNatures, objNode(12), "63. Croix")
        Call AddNode(tvNatures, objNode(12), "64. Pieu")
        Call AddNode(tvNatures, objNode(12), "65. Point non mat�rialis�, d�fini exactement")
        Call AddNode(tvNatures, objNode(12), "66. Point non mat�rialis�, non d�fini exactement")

        'Eaux claires
        'If GetJourCADVersion >= 2.5 Then
        objNode(13) = AddNode(tvNatures, objNode(0), "EAUX CLAIRES", , RGB(0, 204, 255), "SOUT_EC")
        Call AddNode(tvNatures, objNode(13), "70. Point de rep�rage sur conduite", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "71. Chambre", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "72. Chambre enterr�e", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "73. Grille", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "74. S�parateur", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "75. Ouvrage sp�cial", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "76. Relevage", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "77. Exutoire", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "78. Gueulard", , RGB(0, 204, 255))
        Call AddNode(tvNatures, objNode(13), "79. D�versoir", , RGB(0, 204, 255))

        'Eaux us�es
        objNode(14) = AddNode(tvNatures, objNode(0), "EAUX USEES", , RGB(255, 0, 0), "SOUT_EU")
        Call AddNode(tvNatures, objNode(14), "80. Point de rep�rage sur conduite", , RGB(255, 0, 0))
        Call AddNode(tvNatures, objNode(14), "81. Chambre", , RGB(255, 0, 0))
        Call AddNode(tvNatures, objNode(14), "82. Chambre enterr�e", , RGB(255, 0, 0))
        Call AddNode(tvNatures, objNode(14), "83. Fosse sceptique", , RGB(255, 0, 0))
        Call AddNode(tvNatures, objNode(14), "84. S�parateur", , RGB(255, 0, 0))
        Call AddNode(tvNatures, objNode(14), "85. Ouvrage sp�cial", , RGB(255, 0, 0))
        Call AddNode(tvNatures, objNode(14), "86. Relevage", , RGB(255, 0, 0))

        'Eaux unitaires
        objNode(14) = AddNode(tvNatures, objNode(0), "EAUX UNITAIRES", , RGB(153, 51, 0), "SOUT_EM")
        Call AddNode(tvNatures, objNode(14), "90. Point de rep�rage sur conduite", , RGB(153, 51, 0))
        Call AddNode(tvNatures, objNode(14), "91. Chambre", , RGB(153, 51, 0))
        Call AddNode(tvNatures, objNode(14), "92. Chambre enterr�e", , RGB(153, 51, 0))
        Call AddNode(tvNatures, objNode(14), "95. Ouvrage sp�cial", , RGB(153, 51, 0))



        'End If

        'S�lectionner le premier �l�ment (afficher tout)
        'Call tvNatures_NodeClick(Me.tvNatures.Nodes(1))
        'UPGRADE_WARNING: Impossible de r�soudre la propri�t� par d�faut de l'objet Me.tvNatures.Nodes. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
        ' Me.tvNatures.Nodes(1).Selected = True

    End Sub


    Private Function AddNode(ByRef objTreeView As System.Windows.Forms.TreeView, ByRef objParent As System.Windows.Forms.TreeNode, ByRef strText As String, Optional ByRef booExpanded As Boolean = False, Optional ByRef lngColor As Integer = 0, Optional ByRef strKey As String = "") As System.Windows.Forms.TreeNode

        'Ajout du noeud
        Dim objNode As System.Windows.Forms.TreeNode
        objNode = objParent.Nodes.Add(strText)


        'If booExpanded = True Then
        '    objNode.Image = "Calques"  'Image
        objNode.ForeColor = System.Drawing.ColorTranslator.FromOle(lngColor)
        'Else
        '    objNode.Image = "Calque"  'Image
        'End If
        'UPGRADE_ISSUE: MSComctlLib.Node propri�t� objNode.Expanded - Mise � niveau non effectu�e. Cliquez ici�: 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="CC4C7EC0-C903-48FC-ACCC-81861D12DA4A"'
        'objNode.Expand() '.Expanded = booExpanded 'Noeud "d�velopp�"

        'Key bas� sur nature
        Dim intNat As Short
        If strKey = "" Then
            intNat = Val(Replace(VB.Left(strText, 2), ".", ""))
            If intNat > 0 Then
                strKey = "NAT" & CStr(intNat)
            End If
        End If

        If strKey <> "" Then objNode.Name = strKey 'Key

        'Retourne le noeud cr��
        AddNode = objNode

    End Function

    Private Sub NatureToPointEdition(ByRef intNat As Short)

        With Me

            Select Case intNat

                'SOUT EL
                Case 1 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 8
                    .cboSigne.SelectedIndex = 1
                Case 2 'CD
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 8
                    .cboSigne.SelectedIndex = 0
                Case 3 'POT
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 8
                    .cboSigne.SelectedIndex = 2

                    'SOUT TT
                Case 4 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 10
                    .cboSigne.SelectedIndex = 2
                Case 5 'DTT
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 10
                    .cboSigne.SelectedIndex = 0
                Case 6 'RTT
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 10
                    .cboSigne.SelectedIndex = 1
                Case 7 'POT
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 10
                    .cboSigne.SelectedIndex = 3

                    'TOPO
                Case 8 'PTS
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 0
                Case 9 'ARBRE
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 1

                    'TOPO+
                Case 88 '88. Pied de mur
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 2
                Case 89 '89. Haut de mur
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 3
                Case 93 '93. Bord de route
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 4
                Case 94 '94. Trottoir
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 5
                Case 96 '96. Escalier
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 6
                Case 97 '97. Point cot�
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 7
                Case 98 '98. Point de ligne de rupture
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 8
                Case 99 '99. Fa�te, corniche
                    .cboCat.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 9



                    'MO PFP3
                Case 11 'Borne
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 0
                    .cboType.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 0
                Case 12 'Cheville
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 0
                    .cboType.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 1
                Case 13 'Croix
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 0
                    .cboType.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 2
                Case 14 'Pieu
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 0
                    .cboType.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 3
                    'MO PFP1/2
                Case 16 'Borne
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 0
                    .cboType.SelectedIndex = 3 'PFP1
                    .cboSigne.SelectedIndex = 1
                    .cboAttrVariable.SelectedIndex = 0
                Case 17 'Cheville
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 0
                    .cboType.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 2
                    .cboAttrVariable.SelectedIndex = 0
                Case 18 'Croix/Pieu
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 0
                    .cboType.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 3
                    .cboAttrVariable.SelectedIndex = 0
                Case 19 'Non mat�rialis�
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 0
                    .cboType.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 4
                    .cboAttrVariable.SelectedIndex = 1

                    'SOUT EG
                Case 20 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 7
                    .cboSigne.SelectedIndex = 3
                Case 21 'CD
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 7
                    .cboSigne.SelectedIndex = 0
                Case 22 'CB
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 7
                    .cboSigne.SelectedIndex = 1
                Case 23 'CE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 7
                    .cboSigne.SelectedIndex = 2

                    'MO CS
                Case 25 'Exact
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 1
                    .cboExact.SelectedIndex = 0
                Case 26 'Non exact
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 1
                    .cboExact.SelectedIndex = 1

                    'SOUT GA
                Case 30 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 9
                    .cboSigne.SelectedIndex = 1
                Case 31 'VA
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 9
                    .cboSigne.SelectedIndex = 3
                Case 32 'MA
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 9
                    .cboSigne.SelectedIndex = 0
                Case 33 'RE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 9
                    .cboSigne.SelectedIndex = 2

                    'MO OD
                Case 35 'Exact
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 2
                    .cboExact.SelectedIndex = 0
                Case 36 'Non exact
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 2
                    .cboExact.SelectedIndex = 1

                    'IMPL
                Case 38 'Projet
                    .cboCat.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 1
                Case 39 'Contr�le
                    .cboCat.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 0

                    'SOUT EP
                Case 40 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 5
                Case 41 'VA
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 9
                Case 42 'CH
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 2
                Case 43 'BH
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 0
                Case 44 'MA
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 4
                Case 45 'PU
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 6
                Case 46 'CO
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 3
                Case 47 'RE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 7
                Case 48 'CC
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 1
                Case 49 'FO
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 0
                    .cboSigne.SelectedIndex = 8

                    'SOUT TV
                Case 50 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 11
                    .cboSigne.SelectedIndex = 3
                Case 51 'ART
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 11
                    .cboSigne.SelectedIndex = 1
                Case 52 'ATT
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 11
                    .cboSigne.SelectedIndex = 0
                Case 53 'CT
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 11
                    .cboSigne.SelectedIndex = 2

                    'MO BF
                Case 61 'Borne
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 0
                Case 62 'Cheville
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 1
                Case 63 'Croix
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 2
                Case 64 'Pieu
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 3
                Case 65 'Exact
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 4
                    .cboExact.SelectedIndex = 0
                Case 66 'Non exact
                    .cboCat.SelectedIndex = 0
                    .cboTheme.SelectedIndex = 3
                    .cboSigne.SelectedIndex = 4
                    .cboExact.SelectedIndex = 1

                    'SOUT EC
                Case 70 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 8
                Case 71 'CH
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 0
                Case 72 'CE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 1
                Case 73 'GR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 5
                Case 74 'SE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 10
                Case 75 'OS
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 7
                Case 76 'RE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 9
                Case 77 'EX
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 4
                Case 78 'GU
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 6
                Case 79 'DE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 2
                    .cboSigne.SelectedIndex = 3

                    'SOUT EU
                Case 80 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 4
                Case 81 'CH
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 0
                Case 82 'CE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 1
                Case 83 'FO
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 2
                Case 84 'SE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 6
                Case 85 'OS
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 3
                Case 86 'RE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 5
                    .cboSigne.SelectedIndex = 5

                    'SOUT EM
                Case 90 'PR
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 4
                    .cboSigne.SelectedIndex = 3
                Case 91 'CH
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 4
                    .cboSigne.SelectedIndex = 0
                Case 92 'CE
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 4
                    .cboSigne.SelectedIndex = 1
                Case 93 'OS
                    .cboCat.SelectedIndex = 4
                    .cboTheme.SelectedIndex = 4
                    .cboSigne.SelectedIndex = 2

            End Select

        End With

    End Sub


End Class