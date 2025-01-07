Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core

Namespace Data.Content
    Public Enum PointTypes
        Ordered
        Unordered
        Symbols
    End Enum

    Public Class Points
        Private _content As List(Of String)
        Private _type As PointTypes
        Public Property Rerender As Boolean

        Public Property content As List(Of String)
            Get
                Return _content
            End Get
            Set(value As List(Of String))
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

        Public Sub New(type As PointTypes, content As List(Of String))
            Me.type = type
            Me.content = content
        End Sub

        Public Sub Render(slide As Slide, orientation As MsoTextOrientation, left As Integer, top As Integer, width As Integer, height As Integer)
            Dim TextBox = slide.Shapes.AddTextbox(orientation, left, top, width, height)
            TextBox.TextFrame.TextRange.Text = ""

            For Each point As String In _content
                With TextBox.TextFrame.TextRange
                    .Text &= point & vbCrLf
                End With
            Next

            With TextBox.TextFrame.TextRange
                .ParagraphFormat.Bullet.Type = PpBulletType.ppBulletUnnumbered
                .Font.Name = "Arial"
                .Font.Size = 18
            End With

            Rerender = False
        End Sub
    End Class

End Namespace
