Imports System.Runtime
Imports Autodesk.AutoCAD.Runtime
Imports VB = Microsoft.VisualBasic
Imports Autodesk.AutoCAD
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry


Public Class FrmPaletteMO
    'auto-enable our toolpalette for AutoCAD
    
    Implements Autodesk.AutoCAD.Runtime.IExtensionApplication

    Private docXML As New System.Xml.XmlDocument
    Private ListField As New List(Of String)
    Private ListTopic As New List(Of String)
    Private ListObject As New List(Of String)
    Private PaletteName As String = "insertiontxt"
    Private PaletteXML As System.Xml.XmlNode
    Private PaletteXMLfield As System.Xml.XmlNode
    Private PaletteXMLTopic As System.Xml.XmlNode
    Private ActiveLayer As String = ""
    Private ActiveSymbol As String = ""
    Private ActiveCommune As String = ""
    Private ActivePlan As String = ""
    Private NbCreated As Double = 0
    Private CurrNum As Double = 0
    Public objBlockModif As Autodesk.AutoCAD.Interop.Common.AcadBlockReference

    Public intInsTxtTheme, intInsTxtCat, intInsTxtType As Short


    Public Sub Initialize() Implements Autodesk.AutoCAD.Runtime.IExtensionApplication.Initialize
    End Sub
    Public Sub Terminate() Implements Autodesk.AutoCAD.Runtime.IExtensionApplication.Terminate
    End Sub

    Private Sub LoadVariable()

        Dim RVinfo As New Revo.RevoInfo
        If IO.File.Exists(RVinfo.PalletteXML) Then
            Try
                docXML.Load(RVinfo.PalletteXML)

                Dim Palettes As System.Xml.XmlNodeList
                Palettes = docXML.SelectNodes("/revo/palettes/palette")
                For Each Palette As System.Xml.XmlNode In Palettes
                    If Palette.Attributes.ItemOf(0).InnerText.ToLower = PaletteName Then

                        PaletteXML = Palette

                        Dim Fields As System.Xml.XmlNodeList
                        Fields = Palette.SelectNodes("field")
                        For Each Field As System.Xml.XmlNode In Fields
                            cboCat.Items.Add(Field.Attributes.ItemOf(0).InnerText) 'MO, MUT, CHECK_ITF (IMPL,TOPO )
                        Next
                    End If
                Next

            Catch ex As Xml.XmlException
                MsgBox("Erreur de lecture XML : " & ex.Message & vbCrLf & RVinfo.PalletteXML)
            Catch
                MsgBox("Erreur de lecture XML")
            End Try
        End If

        Try
            If cboCat.Text = "" Then cboCat.Text = cboCat.Items(0).ToString
            intInsTxtCat = cboCat.Text
        Catch
        End Try



    End Sub


    Private Structure jcBlockType
        Dim Cat As String
        Dim Caption As String
        Dim Theme As String
        Dim Layer As String
        Dim Block As String
    End Structure

    Private arrTypes() As jcBlockType
    Private lngTypesCount As Integer
    Private Ass As New Revo.RevoInfo
    Private strSelPrec As String


    Private Sub cboCat_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboCat.SelectedIndexChanged


        Dim ValCat As String = cboCat.Text

        Try
            cboTheme.Items.Clear()
            Dim Fields As System.Xml.XmlNodeList
            Fields = PaletteXML.SelectNodes("field")
            For Each Field As System.Xml.XmlNode In Fields
                If Field.Attributes.ItemOf(0).InnerText.ToLower = ValCat.ToLower Then

                    PaletteXMLfield = Field

                    Dim Topics As System.Xml.XmlNodeList
                    Topics = Field.SelectNodes("topic")
                    For Each Topic As System.Xml.XmlNode In Topics
                        cboTheme.Items.Add(Topic.Attributes.ItemOf(0).InnerText)
                    Next

                    Exit For
                End If
            Next

        Catch ex As Xml.XmlException
            MsgBox("Erreur de lecture XML : " & ex.Message)
        Catch
            MsgBox("Erreur de lecture XML")
        End Try

        Try
            If cboTheme.Text = "" Then cboTheme.Text = cboTheme.Items(0).ToString
            intInsTxtTheme = cboTheme.Text
        Catch
        End Try



    End Sub

    Private Sub cboTheme_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboTheme.SelectedIndexChanged

        Dim ValCat As String = cboTheme.Text

        Try
            cboType.Items.Clear()
            Dim Topics As System.Xml.XmlNodeList
            Topics = PaletteXMLfield.SelectNodes("topic")
            For Each Topic As System.Xml.XmlNode In Topics
                If Topic.Attributes.ItemOf(0).InnerText.ToLower = ValCat.ToLower Then

                    PaletteXMLTopic = Topic

                    Dim ObjectMOs As System.Xml.XmlNodeList
                    ObjectMOs = Topic.SelectNodes("object")
                    For Each ObjectMO As System.Xml.XmlNode In ObjectMOs
                        Dim NameOs As System.Xml.XmlNodeList
                        NameOs = ObjectMO.SelectNodes("name")
                        cboType.Items.Add(NameOs.Item(0).InnerText)
                    Next

                    Exit For
                End If
            Next

        Catch ex As Xml.XmlException
            MsgBox("Erreur de lecture XML : " & ex.Message)
        Catch
            MsgBox("Erreur de lecture XML")
        End Try

        Try
            If cboType.Text = "" Then cboType.Text = cboType.Items(0).ToString
            intInsTxtType = cboType.Text
        Catch
        End Try



    End Sub


    'Public Class strOldAttr
    '    Public Title As String
    '    Public TxtValue As String
    'End Class
    Private Function SelectBoxClear()

        cbo01.Items.Clear()
        cbo02.Items.Clear()
        cbo03.Items.Clear()
        cbo04.Items.Clear()
        cbo05.Items.Clear()
        cbo06.Items.Clear()
        cbo07.Items.Clear()
        cbo08.Items.Clear()
        cbo09.Items.Clear()
        cbo10.Items.Clear()
        cbo11.Items.Clear()
        cbo12.Items.Clear()
        cbo13.Items.Clear()
        cbo14.Items.Clear()
        cbo15.Items.Clear()

        lbl01.Text = "[-Attribut-]"
        lbl02.Text = "[-Attribut-]"
        lbl03.Text = "[-Attribut-]"
        lbl04.Text = "[-Attribut-]"
        lbl05.Text = "[-Attribut-]"
        lbl06.Text = "[-Attribut-]"
        lbl07.Text = "[-Attribut-]"
        lbl08.Text = "[-Attribut-]"
        lbl09.Text = "[-Attribut-]"
        lbl10.Text = "[-Attribut-]"
        lbl11.Text = "[-Attribut-]"
        lbl12.Text = "[-Attribut-]"
        lbl13.Text = "[-Attribut-]"
        lbl14.Text = "[-Attribut-]"
        lbl15.Text = "[-Attribut-]"


        Return True
    End Function
    Private Function SelectBoxHide()

        Dim DivHight As Double = 26
        
        '''If lbl01.Text = "[-Attribut-]" Then Div01.Visible = False Else Div01.Visible = True : DivHight = 28
        Dim DistDep As Double = 52
        Dim Disti As Double = 28

        If lbl02.Text = "[-Attribut-]" Then Div02.Visible = False Else Div02.Visible = True : DivHight = DistDep '0
        If lbl03.Text = "[-Attribut-]" Then Div03.Visible = False Else Div03.Visible = True : DivHight = DistDep + Disti '1
        If lbl04.Text = "[-Attribut-]" Then Div04.Visible = False Else Div04.Visible = True : DivHight = DistDep + (Disti * 2) '2
        If lbl05.Text = "[-Attribut-]" Then Div05.Visible = False Else Div05.Visible = True : DivHight = DistDep + (Disti * 3) '3
        If lbl06.Text = "[-Attribut-]" Then Div06.Visible = False Else Div06.Visible = True : DivHight = DistDep + (Disti * 4) '4
        If lbl07.Text = "[-Attribut-]" Then Div07.Visible = False Else Div07.Visible = True : DivHight = DistDep + (Disti * 5) '5
        If lbl08.Text = "[-Attribut-]" Then Div08.Visible = False Else Div08.Visible = True : DivHight = DistDep + (Disti * 6) '6
        If lbl09.Text = "[-Attribut-]" Then Div09.Visible = False Else Div09.Visible = True : DivHight = DistDep + (Disti * 7) '7
        If lbl10.Text = "[-Attribut-]" Then Div10.Visible = False Else Div10.Visible = True : DivHight = DistDep + (Disti * 8) '8
        If lbl11.Text = "[-Attribut-]" Then Div11.Visible = False Else Div11.Visible = True : DivHight = DistDep + (Disti * 9) '9
        If lbl12.Text = "[-Attribut-]" Then Div12.Visible = False Else Div12.Visible = True : DivHight = DistDep + (Disti * 10) '10
        If lbl13.Text = "[-Attribut-]" Then Div13.Visible = False Else Div13.Visible = True : DivHight = DistDep + (Disti * 11) '11
        If lbl14.Text = "[-Attribut-]" Then Div14.Visible = False Else Div14.Visible = True : DivHight = DistDep + (Disti * 12) '12
        If lbl15.Text = "[-Attribut-]" Then Div15.Visible = False Else Div15.Visible = True : DivHight = DistDep + (Disti * 13) '13

        LayoutAtt.Height = DivHight + 3

        Return True
    End Function
    Private Function SelectBoxFill(Nb As Double, Titre As String, StrMO() As String)

        Try

            Dim Dstyle As System.Windows.Forms.ComboBoxStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
            If StrMO.Length = 1 And StrMO(0) = "" Then Dstyle = System.Windows.Forms.ComboBoxStyle.DropDown

            If Titre.ToUpper = "IDENTDN" Then 'And cbo01.Text <> "" And cbo01.Text <> "?" Then
                If ActiveCommune = "" Or ActivePlan = "" Then ActiveCommune = GetFileProperty("Commune") : ActivePlan = GetFileProperty("Plan")
                StrMO(0) = "VD0" & Format(Val((ActiveCommune)), "000") & Format(Val((ActivePlan)), "000000")
            End If


            Select Case Nb
                Case 1 'Div01.Location.Y
                    lbl01.Text = Titre
                    cbo01.Items.AddRange(StrMO)
                    cbo01.DropDownStyle = Dstyle
                    cbo01.Text = cbo01.Items(0)

                Case 2
                    lbl02.Text = Titre
                    cbo02.Items.AddRange(StrMO)
                    cbo02.DropDownStyle = Dstyle
                    cbo02.Text = cbo02.Items(0)

                Case 3
                    lbl03.Text = Titre
                    cbo03.Items.AddRange(StrMO)
                    cbo03.DropDownStyle = Dstyle
                    cbo03.Text = cbo03.Items(0)

                Case 4
                    lbl04.Text = Titre
                    cbo04.Items.AddRange(StrMO)
                    cbo04.DropDownStyle = Dstyle
                    cbo04.Text = cbo04.Items(0)

                Case 5
                    lbl05.Text = Titre
                    cbo05.Items.AddRange(StrMO)
                    cbo05.DropDownStyle = Dstyle
                    cbo05.Text = cbo05.Items(0)

                Case 6
                    lbl06.Text = Titre
                    cbo06.Items.AddRange(StrMO)
                    cbo06.DropDownStyle = Dstyle
                    cbo06.Text = cbo06.Items(0)

                Case 7
                    lbl07.Text = Titre
                    cbo07.Items.AddRange(StrMO)
                    cbo07.DropDownStyle = Dstyle
                    cbo07.Text = cbo07.Items(0)

                Case 8
                    lbl08.Text = Titre
                    cbo08.Items.AddRange(StrMO)
                    cbo08.DropDownStyle = Dstyle
                    cbo08.Text = cbo08.Items(0)

                Case 9
                    lbl09.Text = Titre
                    cbo09.Items.AddRange(StrMO)
                    cbo09.DropDownStyle = Dstyle
                    cbo09.Text = cbo09.Items(0)

                Case 10
                    lbl10.Text = Titre
                    cbo10.Items.AddRange(StrMO)
                    cbo10.DropDownStyle = Dstyle
                    cbo10.Text = cbo10.Items(0)

                Case 11
                    lbl11.Text = Titre
                    cbo11.Items.AddRange(StrMO)
                    cbo11.DropDownStyle = Dstyle
                    cbo11.Text = cbo11.Items(0)

                Case 12
                    lbl12.Text = Titre
                    cbo12.Items.AddRange(StrMO)
                    cbo12.DropDownStyle = Dstyle
                    cbo12.Text = cbo12.Items(0)

                Case 13
                    lbl13.Text = Titre
                    cbo13.Items.AddRange(StrMO)
                    cbo12.DropDownStyle = Dstyle
                    cbo13.Text = cbo13.Items(0)

                Case 14
                    lbl14.Text = Titre
                    cbo14.Items.AddRange(StrMO)
                    cbo14.DropDownStyle = Dstyle
                    cbo14.Text = cbo14.Items(0)

                Case 15
                    lbl15.Text = Titre
                    cbo15.Items.AddRange(StrMO)
                    cbo15.DropDownStyle = Dstyle
                    cbo15.Text = cbo15.Items(0)
            End Select

        Catch

        End Try


        Return True

    End Function

    Private Sub cboType_SelectedIndexChanged(sender As System.Object, e As System.EventArgs) Handles cboType.SelectedIndexChanged


        Dim ValCat As String = cboType.Text
        Dim IDnum As Integer = -1

        Try
            SelectBoxClear()
            Dim ObjectMOs As System.Xml.XmlNodeList
            ObjectMOs = PaletteXMLTopic.SelectNodes("object")
            For Each ObjectMO As System.Xml.XmlNode In ObjectMOs
                Dim NameOs As System.Xml.XmlNodeList
                NameOs = ObjectMO.SelectNodes("name")
                If NameOs.Item(0).InnerText.ToLower = ValCat.ToLower Then
                    ActiveLayer = ObjectMO.SelectNodes("layer").Item(0).InnerText
                    ActiveSymbol = ObjectMO.SelectNodes("symbol").Item(0).InnerText

                    Dim TypeMOs As System.Xml.XmlNodeList
                    TypeMOs = ObjectMO.SelectNodes("type")
                    Dim Nb As Double = 1
                    IDnum = -1


                    Dim CurrID As Integer = 0
                    For Each TypeMO As System.Xml.XmlNode In TypeMOs
                        CurrID += 1

                        Dim Titre As String = TypeMO.Attributes.ItemOf(0).InnerText.ToLower

                        If Titre.ToUpper = "NUMERO" Then
                            IDnum = CurrID
                        End If

                        Dim ValueMOs As System.Xml.XmlNodeList
                        ValueMOs = TypeMO.SelectNodes("value")
                        Dim StrMO(0 To ValueMOs.Count - 1) As String

                        For i = 0 To ValueMOs.Count - 1
                            StrMO(i) = ValueMOs.Item(i).InnerText
                        Next

                        SelectBoxFill(Nb, Titre, StrMO)
                        Nb += 1
                    Next

                    Exit For
                End If
            Next

        Catch ex As Xml.XmlException
            MsgBox("Erreur de lecture XML : " & ex.Message)
        Catch
            MsgBox("Erreur de lecture XML")
        End Try

        Try
            SelectBoxHide()
            If cboType.Text = "" Then cboType.Text = cboType.Items(0).ToString
            intInsTxtType = cboType.Text
        Catch
        End Try


        Try
            ' incr numéro
            IDnum -= 1
            For u = 0 To TLayoutAtt.Controls(IDnum).Controls.Count - 1
                If TLayoutAtt.Controls(IDnum).Controls(u).Name Like "cbo*" Then
                    TLayoutAtt.Controls(IDnum).Controls(u).Text = CurrNum
                End If
            Next
        Catch
        End Try


    End Sub


    Private Function WriteBL(strBlock As String, strLayer As String, insertPt As Point3d, dblorient As Double, dblFactEch As Double, AttList As List(Of String), ValueList As List(Of String))

        Dim ActiveAnnotative As Boolean = False
        Try

            Using trans As Transaction = HostApplicationServices.WorkingDatabase.TransactionManager.StartTransaction



                Dim objBlockRef As Autodesk.AutoCAD.Interop.Common.AcadBlockReference
                'Insertion du bloc
#If _AcadVer_ < 19 Then ' Ancienne Version 2010-2011-2012
                Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.AcadDocument
#Else 'Versio AutoCAD 2013 et +
                    Dim acDocX As Autodesk.AutoCAD.Interop.AcadDocument = Autodesk.AutoCAD.ApplicationServices.DocumentExtension.GetAcadDocument(Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument)
#End If

                Dim dblInsertPt(2) As Double 'New Autodesk.AutoCAD.Geometry.Point3d(Val(Me.txtX.Text), Val(Me.txtY.Text), Val(Me.txtZ.Text))
                dblInsertPt(0) = insertPt.X
                dblInsertPt(1) = insertPt.Y
                dblInsertPt(2) = insertPt.Z
                objBlockRef = acDocX.ModelSpace.InsertBlock(dblInsertPt, strBlock, dblFactEch, dblFactEch, dblFactEch, ConvGisTopoTrigo((dblorient * Math.PI) / 200))
                objBlockRef.Layer = strLayer


                'Ecriture des attributs
                Dim varAttributes As Object
                Dim strNomAttr As String = ""
                varAttributes = objBlockRef.GetAttributes

                For j = LBound(varAttributes) To UBound(varAttributes)

                    For i = 0 To AttList.Count - 1
                        If AttList(i).ToUpper = varAttributes(j).TagString.ToString.ToUpper Then
                            varAttributes(j).TextString = ValueList(i)
                        End If
                    Next

                Next

                trans.Commit()
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


        If ActiveAnnotative Then   'Ajout de l'échelle annotative
            Dim RVscript As New Revo.RevoScript 'Mise à jour des échelles
            RVscript.cmdCmd("#Cmd;ANNOAUTOSCALE|4|_CANNOSCALE|1:500|_CANNOSCALE|1:1000|ANNOAUTOSCALE|-4|ANNOALLVISIBLE|1")
        End If

        Return True

    End Function




    Private Sub cmdDigit_Click(sender As System.Object, e As System.EventArgs) Handles cmdDigit.Click

        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        '' Prompt for the start point
        ' pPtOpts.UseBasePoint = True
        pPtOpts.Message = vbLf & "Saisir le 1er point"
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)

        Dim ptPic1 As Point3d = pPtRes.Value
        Me.txtX.Text = Format(ptPic1.X, "0.000")
        Me.txtY.Text = Format(ptPic1.Y, "0.000")
        Me.txtZ.Text = Format(ptPic1.Z, "0.000")

    End Sub

    Private Sub btnRotation_Click(sender As System.Object, e As System.EventArgs) Handles btnRotation.Click

        Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()

        'Orientation (rotation)
        Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim pPtRes As PromptPointResult
        Dim pPtOpts As PromptPointOptions = New PromptPointOptions("")
        '' Prompt for the start point
        pPtOpts.Message = vbLf & "Définir l'orientation (rotation) du bloc. Saisir le 1er point"
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptPic1 As Point3d = pPtRes.Value

        pPtOpts.Message = vbLf & "Saisir le 2ème point"
        pPtOpts.BasePoint = ptPic1
        pPtOpts.UseBasePoint = True
        pPtOpts.UseDashedLine = True
        pPtRes = acDoc.Editor.GetPoint(pPtOpts)
        Dim ptPic2 As Point3d = pPtRes.Value
        Dim Ligne As Line = New Line(ptPic1, ptPic2)
        Dim returnVal As Double = Ligne.Angle
        returnVal = CDbl(Format(ConvAngTrigoTopoG(returnVal * 200 / Math.PI), "0.0"))
        Me.txtOri.Text = Format(returnVal, "0.0")


    End Sub

    Private Sub cmdNext_Click(sender As System.Object, e As System.EventArgs) Handles cmdNext.Click

        'Vérification de la saisie
        Dim verifOK As Boolean = True
        Dim msgErr As String = ""

        If txtX.Text = "" Or txtY.Text = "" Or IsNumeric(txtX.Text) = False Or IsNumeric(txtY.Text) = False Then
            verifOK = False : msgErr += vbCrLf & "- saisie des coordonnées incomplète"
        End If
        If ActiveSymbol = "" Or ActiveLayer = "" Then
            verifOK = False : msgErr += vbCrLf & "- bloc ou calque non définit"
        End If
        If txtZ.Text = "" Then txtZ.Text = "0.000"
        If txtOri.Text = "" Or IsNumeric(txtOri.Text) = False Then txtOri.Text = GetCurrentRotation() '"100.0"


        If verifOK Then

            Dim AttList As New List(Of String)
            Dim ValueList As New List(Of String)
            Dim IDnum As Integer = -1
            Try
                For i = 0 To TLayoutAtt.Controls.Count - 1
                    If TLayoutAtt.Controls(i).Name Like "Div*" And TLayoutAtt.Controls(i).Visible = True Then
                        For u = 0 To TLayoutAtt.Controls(i).Controls.Count - 1
                            If TLayoutAtt.Controls(i).Controls(u).Name Like "lbl*" Then
                                AttList.Add(TLayoutAtt.Controls(i).Controls(u).Text.ToUpper)
                                If TLayoutAtt.Controls(i).Controls(u).Text.ToUpper = "NUMERO" Then IDnum = i
                            ElseIf TLayoutAtt.Controls(i).Controls(u).Name Like "cbo*" Then
                                ValueList.Add(TLayoutAtt.Controls(i).Controls(u).Text)
                            End If
                        Next
                    End If
                Next


                If cmdNext.Text.ToUpper <> "MODIFIER" Then


                    ' Création d'un nouveau bloc
                    Dim insertpt As Point3d = New Point3d(txtX.Text, txtY.Text, txtZ.Text)
                    Autodesk.AutoCAD.Internal.Utils.SetFocusToDwgView()
                    Dim acDoc As Document = Application.DocumentManager.MdiActiveDocument
                    Using docLock As DocumentLock = acDoc.LockDocument
                        Dim dblScale As Double = Val(GetCurrentScale()) / 1000
                        WriteBL(ActiveSymbol, ActiveLayer, insertpt, CDbl(txtOri.Text), dblScale, AttList, ValueList)
                    End Using


                    ' incr numéro
                    For u = 0 To TLayoutAtt.Controls(IDnum).Controls.Count - 1
                        If TLayoutAtt.Controls(IDnum).Controls(u).Name Like "cbo*" And IsNumeric(TLayoutAtt.Controls(IDnum).Controls(u).Text) Then
                            TLayoutAtt.Controls(IDnum).Controls(u).Text = CDbl(TLayoutAtt.Controls(IDnum).Controls(u).Text) + 1
                            CurrNum = CDbl(TLayoutAtt.Controls(IDnum).Controls(u).Text)
                        End If
                    Next

                    NbCreated += 1
                    If NbCreated < 2 Then lblStatut.Text = NbCreated & " objet crée"
                    If NbCreated > 1 Then lblStatut.Text = NbCreated & " objets crées"


                Else ' Mise de jour du bloc ------- MODIFIER

                    If objBlockModif IsNot Nothing Then

                        Dim dblInsertPt(2) As Double
                        dblInsertPt(0) = txtX.Text
                        dblInsertPt(1) = txtY.Text
                        dblInsertPt(2) = txtZ.Text
                        objBlockModif.InsertionPoint = dblInsertPt
                        objBlockModif.Rotation = ConvGisTopoTrigo((CDbl(txtOri.Text) * Math.PI) / 200)


                        'Mise à jour des attributs

                        'Ecriture des attributs
                        Dim varAttributes As Object
                        Dim strNomAttr As String = ""
                        varAttributes = objBlockModif.GetAttributes

                        For j = LBound(varAttributes) To UBound(varAttributes)
                            For i = 0 To AttList.Count - 1
                                If AttList(i).ToUpper = varAttributes(j).TagString.ToString.ToUpper Then
                                    varAttributes(j).TextString = ValueList(i)
                                End If
                            Next
                        Next




                        cmdNext.Text = "Créer"
                        lblTitle.Text = "Insertion de texte ou symbole"

                        cboCat.Enabled = True
                        cboTheme.Enabled = True
                        cboType.Enabled = True

                        lblStatut.Text = "Objet modifié avec succès"
                    End If

                End If


                ' Action final (refrech + reset value)

                Dim dwgDoc As Document = Application.DocumentManager.MdiActiveDocument
                dwgDoc.Editor.Regen()

                If IsNumeric(txtOri.Text) And txtOri.Text <> "" Then SetFileProperty("Rotation", Format(Val(txtOri.Text), "0.0"))

                'Vide le coord
                txtX.Text = "" : txtY.Text = "" : txtZ.Text = "0.000" : txtOri.Text = GetCurrentRotation() '"100.0"


            Catch

            End Try



        Else
            MsgBox("Saisie incomplète : " & msgErr, MsgBoxStyle.Information, "Insertion non validée")
        End If



    End Sub

    Private Sub cmdNext_MouseHover(sender As Object, e As System.EventArgs) Handles cmdNext.MouseHover
        cmdNext.BackgroundImage = My.Resources.bouton_revo_app_active
    End Sub

    Private Sub cmdNext_MouseLeave(sender As Object, e As System.EventArgs) Handles cmdNext.MouseLeave
        cmdNext.BackgroundImage = My.Resources.bouton_revo_app
    End Sub

    Private Sub txtX_LostFocus(sender As Object, e As System.EventArgs) Handles txtX.LostFocus
        '   MsgBox("super")
    End Sub


End Class
