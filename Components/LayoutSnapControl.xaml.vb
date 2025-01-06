Imports System.Diagnostics
Imports System.Windows
Imports Microsoft.Web.WebView2.Core
Imports System.Web
Imports System.Collections.Specialized
Imports System.IO
Imports System.Windows.Media.Imaging
Imports System.Windows.Media

Public Class LayoutSnapControl

    Private _layout As String
    Private _data As String
    Private _baseUrl As String
    Private _initDone As Boolean

    Public Property Layout As String
        Get
            Return _layout
        End Get
        Set(value As String)

            _layout = value
            Try
                Dim uriBuilder As New UriBuilder(_baseUrl)

                Dim queryParams As NameValueCollection = HttpUtility.ParseQueryString(String.Empty)
                queryParams("layout") = _layout

                uriBuilder.Query = queryParams.ToString()

                Dim htmlPath As String = uriBuilder.ToString()
                Debug.WriteLine($"Generate URL Successful : {htmlPath}")
                If _initDone Then
                    wv2LayoutSnapView.Source = New Uri(htmlPath)

                End If

            Catch ex As Exception
                Debug.WriteLine($"Error:{ex}")

            End Try

        End Set
    End Property

    Public Property Data As String
        Get
            Return _data
        End Get
        Set(value As String)
            _data = value
        End Set
    End Property

    Public Sub New(Data As String, LayoutId As String)
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        InitializeWebView(Data, LayoutId)
        Debug.WriteLine("Initilizing the webview")
    End Sub

    Private Async Sub InitializeWebView(Data As String, LayoutId As String)
        Debug.WriteLine("Setting up the webview")
        _baseUrl = "http://localhost:4000/placeholder"
        _initDone = False

        Try
            Dim webview2RuntimePath As String = "C:\Runtimes\EdgeWebviewRuntime"
            Dim environment As CoreWebView2Environment = Await CoreWebView2Environment.CreateAsync(webview2RuntimePath, "C:\Temp\WebView2UserData")
            Await wv2LayoutSnapView.EnsureCoreWebView2Async(environment)

            _initDone = True
            Me.Layout = LayoutId
            Me.Data = Data


        Catch ex As Exception
            MessageBox.Show($"Error initilizaing webView2: {ex.Message}")
        End Try
    End Sub

    Private Async Sub UpdatePreviewImage()
        Try
            If _initDone Then
                Using memoryStream As New MemoryStream()
                    wv2LayoutSnapView.Visibility = True
                    Await wv2LayoutSnapView.CoreWebView2.CapturePreviewAsync(CoreWebView2CapturePreviewImageFormat.Png, memoryStream)

                    memoryStream.Seek(0, SeekOrigin.Begin)

                    Dim bitmapImage As New BitmapImage()
                    bitmapImage.BeginInit()
                    bitmapImage.StreamSource = memoryStream
                    bitmapImage.CacheOption = BitmapCacheOption.OnLoad
                    bitmapImage.EndInit()

                    'imgWebView.Source = bitmapImage

                End Using
            Else
                Throw New Exception("Initilization of memory failed")
            End If
        Catch ex As Exception
            Debug.WriteLine($"Expeception error while getting preview ${ex}")
        End Try
    End Sub


End Class
