Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports Shape = Microsoft.Office.Interop.PowerPoint.Shape

Namespace Data.Content
    Public Enum TextTypes
        Header
        SubHeader
        Text
        Highlight
    End Enum

    Public Class Text
        Private _content As String
        Private _type As TextTypes
        Public Property Rerender As Boolean

        Public Property content As String
            Get
                Return _content
            End Get
            Set(value As String)
                _content = value
                Rerender = True
            End Set
        End Property

        Public Property type As TextTypes
            Get
                Return _type
            End Get
            Set(value As TextTypes)
                _type = value
                Rerender = True
            End Set
        End Property

        Public Sub New(type As TextTypes, content As String)
            Me.type = type
            Me.content = content
        End Sub

        Public Sub Render(NewSlide As Slide, Orientation As MsoTextOrientation, left As Integer, top As Integer, width As Integer, height As Integer)
            Dim TextBox As Shape = NewSlide.Shapes.AddTextbox(Orientation, left, top, width, height)
            TextBox.TextFrame.TextRange.Text = _content

            'TODO: Add the style for the shape based on enum 

            Rerender = False
        End Sub
    End Class

End Namespace
