Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.LayerManager
Imports System.Collections.Generic


Public Class modAcadLayer

    Public Shared Sub ListLayerFilters()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' List the nested layer filters

        Dim lfc As LayerFilterCollection = db.LayerFilters.Root.NestedFilters

        For i As Integer = 0 To lfc.Count - 1
            Dim lf As LayerFilter = lfc(i)
            ed.WriteMessage(vbLf & "{0} - {1} (can{2} be deleted)", i + 1, lf.Name, (If(lf.AllowDelete, "", "not")))

            Dim lfcX As LayerFilterCollection = lf.NestedFilters
            For x As Integer = 0 To lfcX.Count - 1
                Dim lfX As LayerFilter = lfcX(x)
                ed.WriteMessage(vbLf & " - " & "{0} - {1} (can{2} be deleted)", x + 1, lfX.Name, (If(lfX.AllowDelete, "", "not")))

                Dim lfcY As LayerFilterCollection = lfX.NestedFilters
                For y As Integer = 0 To lfcY.Count - 1
                    Dim lfy As LayerFilter = lfcY(y)
                    ed.WriteMessage(vbLf & " - - " & "{0} - {1} (can{2} be deleted)", y + 1, lfy.Name, (If(lfy.AllowDelete, "", "not")))
                Next
            Next

        Next
    End Sub

    Public Function CreateLayerFilters(FilterName As String, LayerName As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Try
             If InStr(FilterName, "/") <> 0 Then   'Avec structure
                Dim FilterStructure() As String = Split(FilterName, "/")
                Dim lyrTree As LayerFilterTree = HostApplicationServices.WorkingDatabase.LayerFilters
                Dim lyrFilt As LayerFilter = lyrTree.Root

                For u = 0 To FilterStructure.Length - 2
                    If lyrFilt.NestedFilters.Count > 0 Then lyrFilt = FindFilterLayer(FilterStructure(u), lyrFilt)
                    If lyrFilt = Nothing Then Return False ' Exit Function
                Next

                'Charge le nom du filtre (fin du tableau)
                FilterName = FilterStructure(FilterStructure.Length - 1)
                Dim lyrTree2 As New LayerFilterTree(lyrTree.Root, lyrFilt)
                HostApplicationServices.WorkingDatabase.LayerFilters = lyrTree2
            Else
                'Pas de séléction
                Dim lyrTree As LayerFilterTree = HostApplicationServices.WorkingDatabase.LayerFilters
                Dim lyrFilt As LayerFilter = lyrTree.Root
                Dim lyrTree2 As New LayerFilterTree(lyrTree.Root, lyrFilt)
                HostApplicationServices.WorkingDatabase.LayerFilters = lyrTree2
            End If


            Dim lft As LayerFilterTree = db.LayerFilters
            Dim lfc As LayerFilterCollection = lft.Current.NestedFilters

            ' Create three new layer filters
            Dim lf1 As New LayerFilter()
            lf1.Name = FilterName '"Unlocked Layers"
            If InStr(LayerName, "|") <> 0 Then
                Dim SplitLayerName() As String = Split(LayerName, "|")
                Dim FilterExp As String = ""
                For i = 0 To SplitLayerName.Count - 1
                    If FilterExp <> "" Then FilterExp += " OR "
                    FilterExp += "NAME==""" & SplitLayerName(i) & """"
                Next
                lf1.FilterExpression = FilterExp ' NAME==""" & LayerName & """ OR NAME==""" & "MUT_*" & """
            Else
                lf1.FilterExpression = "NAME==""" & LayerName & """"
            End If

            '"NAME==""WBLayer*"""

            If ExisteFilterLayer(FilterName, lfc) = False Then

                'Dim lf2 As New LayerFilter()
                'lf2.Name = "White Layers"
                'lf2.FilterExpression = "COLOR==""7"""

                'Dim lf3 As New LayerFilter()
                'lf3.Name = "Visible Layers"
                'lf3.FilterExpression = "OFF==""False"" AND FROZEN==""False"""

                ' Add them to the collection

                lfc.Add(lf1)
                'lfc.Add(lf2)

                db.LayerFilters = lft

                ' List the layer filters, to see the new ones
                ' ListLayerFilters()

            Else 'Filtre déjà existant
                ' ed.WriteMessage(vbLf & "Exception: {0}", "Filtre existant : pas de création")
            End If
        Catch ex As Exception
            ed.WriteMessage(vbLf & "Exception: {0}", ex.Message)
        End Try

        Return True

    End Function

    Private Function ExisteFilterLayer(NameFilter As String, FilterColl As LayerFilterCollection) As Boolean

        Dim ExistLF As Boolean = False
        For Each LF As LayerFilter In FilterColl
            If LF.Name.ToUpper = NameFilter.ToUpper Then ExistLF = True : Exit For
        Next

        Return ExistLF
    End Function

    Private Function FindFilterLayer(NameFilter As String, Filter As LayerFilter) As LayerFilter

        Dim lfc As LayerFilterCollection = Filter.NestedFilters
        For y As Integer = 0 To lfc.Count - 1
            Dim lfy As LayerFilter = lfc(y)
            If lfy.Name.ToUpper = NameFilter.ToUpper Then Return lfy
        Next

        Return Nothing
    End Function

    Public Function CreateLayerGroup(FilterName As String, LayerName As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        If InStr(FilterName, "/") <> 0 Then   'Avec structure
            Dim FilterStructure() As String = Split(FilterName, "/")
            Dim lyrTree As LayerFilterTree = HostApplicationServices.WorkingDatabase.LayerFilters
            Dim lyrFilt As LayerFilter = lyrTree.Root

            For u = 0 To FilterStructure.Length - 2
                If lyrFilt.NestedFilters.Count > 0 Then lyrFilt = FindFilterLayer(FilterStructure(u), lyrFilt)
                If lyrFilt = Nothing Then Return False ' Exit Function
            Next

            'Charge le nom du filtre (fin du tableau)
            FilterName = FilterStructure(FilterStructure.Length - 1)
            Dim lyrTree2 As New LayerFilterTree(lyrTree.Root, lyrFilt)
            HostApplicationServices.WorkingDatabase.LayerFilters = lyrTree2
        Else
            'Pas de séléction
            Dim lyrTree As LayerFilterTree = HostApplicationServices.WorkingDatabase.LayerFilters
            Dim lyrFilt As LayerFilter = lyrTree.Root
            Dim lyrTree2 As New LayerFilterTree(lyrTree.Root, lyrFilt)
            HostApplicationServices.WorkingDatabase.LayerFilters = lyrTree2
        End If


        ' A list of the layers' names & IDs contained
        ' in the current database, sorted by layer name
        Dim ld As New SortedList(Of String, ObjectId)()
        ' A list of the selected layers' IDs
        Dim lids As New ObjectIdCollection()

        ' Start by populating the list of names/IDs
        ' from the LayerTable
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim lt As LayerTable = DirectCast(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            For Each lid As ObjectId In lt
                Dim ltr As LayerTableRecord = DirectCast(tr.GetObject(lid, OpenMode.ForRead), LayerTableRecord)
                If ltr.Name.ToUpper Like LayerName.ToUpper Then
                    ' ld.Add(ltr.Name, lid)
                    lids.Add(ltr.ObjectId)
                End If
            Next
        End Using

        ' Display a numbered list of the available layers
        ' ed.WriteMessage(vbLf & LayerName) '"Layers available for group:")
        ' SelectLayers(ed, ld, lids)
        ' Now we've selected our layers, let's create the group
        Try
            If lids.Count > 0 Then
                ' Get the existing layer filters
                ' (we will add to them and set them back)
                Dim lft As LayerFilterTree = db.LayerFilters
                Dim lfc As LayerFilterCollection = lft.Current.NestedFilters

                ' Create a new layer group
                Dim lg As New LayerGroup()
                lg.Name = FilterName ' "My Layer Group"

                If ExisteFilterLayer(FilterName, lfc) = False Then

                    ' Add our layers' IDs to the list
                    For Each id As ObjectId In lids
                        lg.LayerIds.Add(id)
                    Next

                    ' Add the group to the collection
                    lfc.Add(lg)

                    ' Set them back on the Database
                    db.LayerFilters = lft

                    ed.WriteMessage(vbLf & """{0}"" group created containing {1} layers." & vbLf, lg.Name, lids.Count)

                    ' List the layer filters, to see the new group
                    ListLayerFilters()
                End If
            End If
        Catch ex As Exception
            ed.WriteMessage(vbLf & "Exception: {0}", ex.Message)
        End Try

        Return True
    End Function


    Public Shared Sub CreateNestedLayerGroup()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        ' A list of the layers' names & IDs contained
        ' in the current database, sorted by layer name
        Dim ld As New SortedList(Of String, ObjectId)()

        ' A list of the selected layers' IDs
        Dim lids As New ObjectIdCollection()

        ' Start by populating the list of names/IDs
        ' from the LayerTable
        Dim tr As Transaction = db.TransactionManager.StartTransaction()
        Using tr
            Dim lt As LayerTable = DirectCast(tr.GetObject(db.LayerTableId, OpenMode.ForRead), LayerTable)
            For Each lid As ObjectId In lt
                Dim ltr As LayerTableRecord = DirectCast(tr.GetObject(lid, OpenMode.ForRead), LayerTableRecord)
                ld.Add(ltr.Name, lid)
            Next
        End Using

        ' Display a numbered list of the available layers
        ed.WriteMessage(vbLf & "Layers available for nested group:")
        SelectLayers(ed, ld, lids)

        ' Now we've selected our layers, let's create the group
        Try
            If lids.Count > 0 Then
                ' Get the existing layer filters
                ' (we will add to them and set them back)
                Dim lft As LayerFilterTree = db.LayerFilters
                Dim lfc As LayerFilterCollection = lft.Root.NestedFilters
                Dim idxMap As New Dictionary(Of Integer, Integer)()

                ' List the nested layer filters
                Dim lfNum As Integer = 1

                For i As Integer = 0 To lfc.Count - 1
                    Dim lf As LayerFilter = lfc(i)
                    If lf.AllowNested Then
                        ed.WriteMessage(vbLf & "{0} - {1}", lfNum, lf.Name)
                        idxMap.Add(System.Math.Max(System.Threading.Interlocked.Increment(lfNum), lfNum - 1), i)
                    End If
                Next

                If lfNum = 1 Then
                    ed.WriteMessage(vbLf & "No filters found that allow nesting.")
                Else
                    ' We will ask the user to select from the list

                    Dim pio As New PromptIntegerOptions(vbLf & "Enter number of parent filter: ")
                    pio.LowerLimit = 1
                    pio.UpperLimit = lfNum
                    pio.AllowNone = False

                    Dim pir As PromptIntegerResult = ed.GetInteger(pio)

                    If pir.Status <> PromptStatus.OK Then
                        Return
                    End If

                    Dim parent As LayerFilter = lfc(idxMap(pir.Value))

                    ' Create a new layer group

                    Dim lg As New LayerGroup()
                    lg.Name = "My Nested Layer Group"

                    ' Add the group to the collection

                    parent.NestedFilters.Add(lg)

                    ' Add our layers' IDs to the list

                    For Each id As ObjectId In lids
                        lg.LayerIds.Add(id)
                    Next

                    ' Set them back on the Database

                    db.LayerFilters = lft

                    ed.WriteMessage(vbLf & """{0}"" nested group added to parent ""{1}"" " & "containing {2} layers." & vbLf, lg.Name, parent.Name, lids.Count)
                End If
            End If
        Catch ex As Exception
            ed.WriteMessage(vbLf & "Exception: {0}", ex.Message)
        End Try
    End Sub

    Private Shared Sub SelectLayers(ed As Editor, ld As SortedList(Of String, ObjectId), lids As ObjectIdCollection)
        Dim i As Integer = 1
        For Each kv As KeyValuePair(Of String, ObjectId) In ld
            ed.WriteMessage(vbLf & "{0} - {1}", System.Math.Max(System.Threading.Interlocked.Increment(i), i - 1), kv.Key)
        Next

        ' We will ask the user to select from the list

        Dim pio As New PromptIntegerOptions(vbLf & "Enter number of layer to add: ")
        pio.LowerLimit = 1
        pio.UpperLimit = ld.Count
        pio.AllowNone = True

        ' And will do so in a loop, waiting for
        ' Escape or Enter to terminate

        Dim pir As PromptIntegerResult
        Do
            ' Select one from the list

            pir = ed.GetInteger(pio)

            If pir.Status = PromptStatus.OK Then
                ' Get the layer's name

                Dim ln As String = ld.Keys(pir.Value - 1)

                ' And then its ID

                Dim lid As ObjectId
                ld.TryGetValue(ln, lid)

                ' Add the layer'd ID to the list, is it's not
                ' already on it

                If lids.Contains(lid) Then
                    ed.WriteMessage(vbLf & "Layer ""{0}"" has already been selected.", ln)
                Else
                    lids.Add(lid)
                    ed.WriteMessage(vbLf & "Added ""{0}"" to selected layers.", ln)
                End If
            End If
        Loop While pir.Status = PromptStatus.OK
    End Sub


    Public Function DeleteLayerFilter(FilterName As String)
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor


        Try
            If InStr(FilterName, "/") <> 0 Then   'Avec structure
                Dim FilterStructure() As String = Split(FilterName, "/")
                Dim lyrTree As LayerFilterTree = HostApplicationServices.WorkingDatabase.LayerFilters
                Dim lyrFilt As LayerFilter = lyrTree.Root

                For u = 0 To FilterStructure.Length - 2
                    If lyrFilt.NestedFilters.Count > 0 Then lyrFilt = FindFilterLayer(FilterStructure(u), lyrFilt)
                    If lyrFilt = Nothing Then Return False ' Exit Function
                Next

                'Charge le nom du filtre (fin du tableau)
                FilterName = FilterStructure(FilterStructure.Length - 1)
                Dim lyrTree2 As New LayerFilterTree(lyrTree.Root, lyrFilt)
                HostApplicationServices.WorkingDatabase.LayerFilters = lyrTree2
            Else
                'Pas de séléction
                Dim lyrTree As LayerFilterTree = HostApplicationServices.WorkingDatabase.LayerFilters
                Dim lyrFilt As LayerFilter = lyrTree.Root
                Dim lyrTree2 As New LayerFilterTree(lyrTree.Root, lyrFilt)
                HostApplicationServices.WorkingDatabase.LayerFilters = lyrTree2
            End If

            '''Check strucutre
            ''Dim lyrTree As LayerFilterTree = HostApplicationServices.WorkingDatabase.LayerFilters
            ''If InStr(FilterName, "/") <> 0 Then   'Avec structure
            ''    Dim FilterStructure() As String = Split(FilterName, "/")
            ''    For i As Integer = 0 To lyrTree.Root.NestedFilters.Count - 1
            ''        Dim lyrFilt As LayerFilter = lyrTree.Root.NestedFilters(i)
            ''        If lyrFilt.Name.ToUpper = FilterStructure(FilterStructure.Length - 2).ToUpper Then
            ''            '  ed.WriteMessage(vbLf & "MyFilter already exists, making it current")
            ''            ' This is the way I found to
            ''            ' make a Layer Filter current
            ''            'use LayerFilterTree constructor
            ''            Dim lyrTree2 As New LayerFilterTree(lyrTree.Root, lyrFilt)
            ''            HostApplicationServices.WorkingDatabase.LayerFilters = lyrTree2
            ''            Exit For
            ''        End If
            ''    Next

            ''Charge le nom du filtre (fin du tableau)
            'FilterName = FilterStructure(FilterStructure.Length - 1)
            'Else
            ''Pas de séléction
            'Dim lyrFilt As LayerFilter = lyrTree.Root
            'Dim lyrTree2 As New LayerFilterTree(lyrTree.Root, lyrFilt)
            'HostApplicationServices.WorkingDatabase.LayerFilters = lyrTree2
            'End If

            ' Get the existing layer filters
            ' (we will add to them and set them back)

            Dim lft As LayerFilterTree = db.LayerFilters
            Dim lfc As LayerFilterCollection = lft.Current.NestedFilters
            Dim NumFilter As Double = -1
            ' Prompt for the index of the filter to delete

            'Dim pio As New PromptIntegerOptions(vbLf & vbLf & "Enter index of filter to delete")
            'pio.LowerLimit = 1
            'pio.UpperLimit = lfc.Count

            For i = 0 To lfc.Count - 1
                If lfc(i).Name.ToUpper = FilterName.ToUpper Then
                    NumFilter = i
                    Exit For
                End If
            Next

            'Dim pir As PromptIntegerResult = ed.GetInteger(pio)

            'If pir.Status <> PromptStatus.OK Then
            '    Return True
            'End If

            ' Get the selected filter
            If NumFilter <> -1 Then
                Dim lf As LayerFilter = lfc(NumFilter) 'pir.Value - 1)

                ' If it's possible to delete it, do so

                If Not lf.AllowDelete Then
                    ed.WriteMessage(vbLf & "Layer filter cannot be deleted.")
                Else
                    lfc.Remove(lf)
                    db.LayerFilters = lft

                    ListLayerFilters()
                End If
            Else
                ed.WriteMessage(vbLf & "Layer filter : not found.")
            End If
        Catch ex As Exception
            ed.WriteMessage(vbLf & "Exception: {0}", ex.Message)
        End Try

        Return True
    End Function
End Class