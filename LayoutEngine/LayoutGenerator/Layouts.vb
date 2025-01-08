
Imports SlideComponentType = System.Collections.Generic.Dictionary(Of String, Object)
Imports SlideLayoutComponentType = System.Collections.Generic.Dictionary(Of String, Object)
Imports SlideElementListType = System.Collections.Generic.List(Of System.Collections.Generic.Dictionary(Of String, Object))
Imports SlideContentKeyValuePairType = System.Collections.Generic.Dictionary(Of String, String)
Imports System.Drawing.Imaging
Imports System.Diagnostics

Namespace SlideTemplates
    Public Class Layouts
        Private _layoutDescriptor As SlideLayoutComponentType

        Private _elements As SlideElementListType
        Private _content As SlideContentKeyValuePairType
        Private _rerender As Boolean = False

        Public Property Layout As SlideLayoutComponentType
            Set(value As SlideLayoutComponentType)
                _rerender = True
                _layoutDescriptor = value
            End Set
            Get
                Return _layoutDescriptor
            End Get
        End Property

        Public Property Content As SlideContentKeyValuePairType
            Set(value As SlideContentKeyValuePairType)
                _rerender = True
                _content = Content
            End Set
            Get
                Return _content
            End Get
        End Property


        Public Sub New(content As SlideContentKeyValuePairType,
                       layout As SlideLayoutComponentType)
            Try
                If Not Test_ContentAndLayout(content, layout) Then
                    Throw New Exception("Key mismatch; Content and Layout")
                End If

                Me.Layout = layout
                Me.Content = content

            Catch ex As Exception
                Debug.WriteLine("Failed to create the layout")
            End Try
        End Sub

        Public Sub Render(Slide As PowerPoint.Slide)
            ' First check the integrity of content and layout
            Dim componentKey = Layout.Keys
            _elements = New SlideElementListType

            For Each Key As String In componentKey
                Select Case Processor.LayoutComponents.EnumToString(Key)
                    Case LayoutComponents.Title
                    Case LayoutComponents.Description
                    Case LayoutComponents.Points
                    Case LayoutComponents.Image
                    Case LayoutComponents.Cosmetic
                End Select

            Next
        End Sub

        Private Function GenerateTitle(slide As PowerPoint.Slide,
                                       content As String,
                                       description As Dictionary(Of String, Object))
            Dim title = New Data.Content.Text(Data.Content.TextTypes.Header,
                                              content,
                                              description)
            title.Render(slide)
            Return title

        End Function
        Private Function GenerateDescription(slide As PowerPoint.Slide,
                                             Content As String,
                                             description As Dictionary(Of String, Object))

            Dim Desc = New Data.Content.Text(Data.Content.TextTypes.Text,
                                              Content,
                                              description)
            Desc.Render(slide)
            Return Desc
        End Function
        Private Function GeneratePoints(content As List(Of String), description As Dictionary(Of String, Object))

        End Function
        Private Function GenerateImage(tempPath As String, description As Dictionary(Of String, Object))
        End Function
        Private Function GenerateCosmetic(description As Dictionary(Of String, Object))
        End Function

        Private Function Test_ContentAndLayout(content As SlideContentKeyValuePairType, layout As SlideLayoutComponentType)
            ' TODO : Customize the check by removing the Cosmetic Keys from the layout

            Dim layoutKeys = layout.Keys
            Dim contentKeys = content.Keys
            Dim areKeysMatching As Boolean = layoutKeys.Count = contentKeys.Count AndAlso layoutKeys.All(Function(key) contentKeys.Contains(key))

            Return areKeysMatching
        End Function

    End Class
End Namespace
