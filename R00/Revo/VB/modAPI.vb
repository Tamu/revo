Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Windows.Forms
Imports System.Collections.Generic


Module modAPI
    Public PrevExtents As ViewTableRecord
    ''' <summary>
    ''' Displays a folder browser dialog allowing the user to browse and select a folder
    ''' </summary>
    ''' <param name="msg">The message to display to the user</param>
    ''' <returns>The path to the selected folder if the user selected one or an empty string if not.</returns>
    ''' <remarks></remarks>
    Public Function BrowseForFolder(ByVal msg As String) As String
        Dim fb As New FolderBrowserDialog
        fb.Description = msg
        If fb.ShowDialog = DialogResult.OK Then
            Return fb.SelectedPath
        Else
            Return String.Empty
        End If
    End Function

    Public Function GetCurrentWinLogin() As String
        Dim UserName As String = Environment.UserName
        If UserName.Contains("\") Then
            Return UserName.Substring(UserName.LastIndexOf("\") + 1)
        Else
            Return UserName
        End If
    End Function



    ' modVariousFunctions.vb <- Copié depuis ce fichier VB
    'Imports System.Windows.Forms.DataGridView

    ' Sort by the indicated column.
    Public Sub SortMSFlexByColumn(ByRef msFlexG As DataGridView, ByVal sort_column As Short, ByRef sort_Order As System.ComponentModel.ListSortDirection) ' MSFlexGridLib.SortSettings)

        ' Hide the FlexGrid.
        msFlexG.Visible = False
        msFlexG.Refresh()

        ' Sort using the clicked column.
        msFlexG.Sort(msFlexG.Columns(sort_column), sort_Order)
        'msFlexG.col = sort_column
        'msFlexG.ColSel = sort_column
        'msFlexG.Row = 0
        'msFlexG.RowSel = 0

        'msFlexG.Sort = sort_Order

        ' Display the FlexGrid.
        msFlexG.Visible = True

    End Sub

    Public Function SortCollection(ByVal c As Collection) As Collection
        Dim N As Integer : N = c.Count()
        If N = 0 Then SortCollection = New Collection : Exit Function
        Dim Index(N - 1) As Integer ' allocate index array
        Dim i, m As Integer
        For i = 0 To N - 1 : Index(i) = i + 1 : Next  ' fill index array
        For i = N \ 2 - 1 To 0 Step -1 ' generate ordered heap
            Heapify(c, Index, i, N)
        Next
        For m = N To 2 Step -1 ' sort the index array
            Exchange(Index, 0, m - 1) ' move highest element to top
            Heapify(c, Index, 0, m - 1)
        Next
        Dim c2 As New Collection
        For i = 0 To N - 1 : c2.Add(c.Item(Index(i))) : Next  ' fill output collection
        SortCollection = c2
    End Function

    ' --------------------------------------------------------------------------------------------------
    ' Effectue un trie du contenu de la Collection passé en argument,
    ' avec précision de la position de l'argument de trie :
    '       la valeur de chaque item de la Collection est composée de plusieurs chaines séparées, dont
    '       l'élément qui servira au trie
    '
    Public Function SortAcObjIDColl(ByRef theList As Collection, Optional ByVal smallToHigh As Boolean = True) As Collection
        Dim ct As Integer
        Dim theSortedList As New SortedList(Of Double, Revo.AcObjID)
        Dim SortColl As Collection = theList

        Try
            ' If Not theList Is Nothing Then
            If Not theList Is Nothing Then

                ' Boucle sur la Collection
                For ct = 1 To theList.Count
                    Dim ObjID As Revo.AcObjID = theList(ct)

                    ' Ajout dans la liste automatiquement triée, de la valeur de trie
                    ' et de la chaine originale
                    theSortedList.Add(CDbl(ObjID.Value), ObjID)
                Next

                ' Recréation de la liste pour ajout des valeurs dans l'ordre
                theList.Clear()

                If smallToHigh = False Then
                    ' Du plus grand au plus petit
                    For ct = theSortedList.Count - 1 To 0 Step -1
                        theList.Add(theSortedList.Values(ct))
                    Next
                Else
                    ' Du plus petit au plus grand : par défaut
                    For ct = 0 To theSortedList.Count - 1
                        theList.Add(theSortedList.Values(ct))
                    Next
                End If

                'SortThisCollByArgu = True
                Return theList
            End If


        Catch Err As Exception
            Dim Conn As New Revo.connect
            Conn.RevoLog(Err.Message)
        End Try

        Return SortColl

    End Function

    ' --------------------------------------------------------------------------------------------------
    ' Effectue un trie du contenu de la Collection passé en argument,
    ' avec précision de la position de l'argument de trie :
    '       la valeur de chaque item de la Collection est composée de plusieurs chaines séparées, dont
    '       l'élément qui servira au trie
    '
    Public Function SortThisCollByArgu(ByRef theList As Collection, ByVal thePos As Integer, _
                                            Optional ByVal theSeparator As String = "|", _
                                            Optional ByVal smallToHigh As Boolean = True) As Collection
        Dim ct As Integer, theInfo() As String
        Dim theSortedList As New SortedList(Of Double, String)
        Dim SortColl As Collection = theList

        'SortThisCollByArgu = False

        Try
            ' If Not theList Is Nothing Then
            If Not theList Is Nothing Then
                ' Test nombre d'argument en cohérence avec la position de trie
                theInfo = Split(theList(1).ToString, theSeparator)
                If thePos > theInfo.GetUpperBound(0) Then
                    Return SortColl : Exit Function
                End If

                ' Boucle sur la Collection
                For ct = 1 To theList.Count
                    ' Récupération formatée des informations
                    theInfo = Split(theList(ct).ToString, theSeparator)

                    ' Ajout dans la liste automatiquement triée, de la valeur de trie
                    ' et de la chaine originale
                    theSortedList.Add(CDbl(theInfo(thePos)), theList(ct).ToString)
                Next

                ' Recréation de la liste pour ajout des valeurs dans l'ordre
                theList.Clear()

                If smallToHigh = False Then
                    ' Du plus grand au plus petit
                    For ct = theSortedList.Count - 1 To 0 Step -1
                        theList.Add(theSortedList.Values(ct))
                    Next
                Else
                    ' Du plus petit au plus grand : par défaut
                    For ct = 0 To theSortedList.Count - 1
                        theList.Add(theSortedList.Values(ct))
                    Next
                End If

                'SortThisCollByArgu = True
                Return theList
            End If


        Catch Err As Exception
            Dim Conn As New Revo.connect
            Conn.RevoLog(Err.Message)
        End Try

        Return SortColl

    End Function


    Private Sub Heapify(ByVal c As Collection, ByRef Index() As Integer, ByVal i1 As Integer, ByVal N As Integer)
        ' Heap order rule: a[i] >= a[2*i+1] and a[i] >= a[2*i+2]
        Dim nDiv2 As Integer : nDiv2 = N \ 2
        Dim i As Integer : i = i1
        Dim k As Integer
        Do While i < nDiv2 : k = 2 * i + 1
            If k + 1 < N Then
                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet c.Item(Index(k + 1)). Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet c.Item(Index(k)). Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
                If c.Item(Index(k)) < c.Item(Index(k + 1)) Then k = k + 1
            End If
            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet c.Item(Index(k)). Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            'UPGRADE_WARNING: Impossible de résoudre la propriété par défaut de l'objet c.Item(Index(i)). Cliquez ici : 'ms-help://MS.VSCC.v90/dv_commoner/local/redirect.htm?keyword="6A50421D-15FE-4896-8A1B-2EC21E9037B2"'
            If c.Item(Index(i)) >= c.Item(Index(k)) Then Exit Do
            Exchange(Index, i, k)
            i = k
        Loop
    End Sub

    Private Sub Exchange(ByRef Index() As Integer, ByVal i As Integer, ByVal j As Integer)
        Dim Temp As Integer : Temp = Index(i)
        Index(i) = Index(j)
        Index(j) = Temp
    End Sub
End Module
