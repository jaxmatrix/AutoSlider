
Imports System.Diagnostics
Imports Newtonsoft.Json.Linq
Imports Microsoft.Core
Imports Microsoft.Office.Core

Namespace Data.Content
    Public Enum ImageTypes
        Image
        ShapeImage
    End Enum

    Public Class ImageElement
        Private _imagePath As String
        Private _rerender As Boolean = False
        Private _type As ImageTypes
        Private _description As JObject

        Public Property Rerender As Boolean
            Get
                Return _rerender
            End Get
            Set(value As Boolean)
                _rerender = value
            End Set
        End Property

        Public Property ImagePath As String
            Set(value As String)
                _imagePath = value
                Rerender = True
            End Set
            Get
                Return _imagePath
            End Get
        End Property

        Public Property Description As JObject
            Set(value As JObject)
                _description = value
            End Set
            Get
                Return _description
            End Get
        End Property

        Public Sub New(imagePath As String, description As JObject)
            Me.ImagePath = imagePath
            Me.Description = description
            Me._type = ImageTypes.Image
        End Sub

        Public Sub Render(slide As PowerPoint.Slide)
            If _type = ImageTypes.Image Then
                Try
                    If ImagePath = "" Then
                        Throw New Exception("Image Path not given")
                    End If

                    Dim top As Single = Description("Top")
                    Dim left As Single = Description("Left")
                    Dim width As Single = Description("Width")
                    Dim height As Single = Description("Height")

                    slide.Shapes.AddPicture(
                        _imagePath,
                        MsoTriState.msoFalse,
                        MsoTriState.msoTrue,
                        left,
                        top,
                        width,
                        height
                    )
                    'TODO : Add Image Style 

                Catch ex As Exception
                    Debug.WriteLine("Missing Image Path")
                End Try


            Else
                'TODO : Implement the shape rendered for intersection and 
                ' special shapes in the template 
                _type = ImageTypes.ShapeImage
                Debug.WriteLine("Shape Image not implemented ")
            End If
        End Sub

    End Class
End Namespace