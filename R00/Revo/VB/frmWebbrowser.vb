Imports Microsoft.VisualBasic.ApplicationServices

Public Class frmWebbrowser

    'Dim flashanimation As ShockwaveFlashObjects.IShockwaveFlash


    Public xWebFormStat As String = "-"
    Dim init As Boolean = False

    Public Property WebStat() As String
        Get
            Return xWebFormStat
        End Get
        Set(ByVal value As String)
            xWebFormStat = value
        End Set
    End Property

    Private Sub frmWebbrowser_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'WebBrowser2.Visible = False
        'flashanimation = New ShockwaveFlashObjects.ShockwaveFlash
        'flashanimation.LoadMovie(1, "c:\\easydoc.swf")
        'flashanimation.Play()
        'flashanimation.A()
    
        Dim ass As New Revo.RevoInfo
        Me.Icon = ass.Icon
        WebBrowser2.ScrollBarsEnabled = False
        init = False
    End Sub


    Private Sub WebBrowser2_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowser2.DocumentCompleted
        If init = False Then
            init = True
        Else ' True
        End If

        If e.Url.ToString.ToLower <> "about:blank" Then
            xWebFormStat = "LOADED"
            Me.Hide()
        Else
            xWebFormStat = "-"
        End If

    End Sub

    
    Public Function LoadPage(url)

        If url <> "" Then WebBrowser2.Navigate(url) 'WebForm.

        Do While WebBrowser2.ReadyState <> Windows.Forms.WebBrowserReadyState.Complete
            If WebStat = "LOADED" Then
                Exit Do
            End If
            System.Windows.Forms.Application.DoEvents()
        Loop

        Return True
    End Function

End Class