Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports Shape = Microsoft.Office.Interop.PowerPoint.Shape
Imports Newtonsoft.Json.Linq

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
        Private _description As JObject
        Private _rerender As Boolean = False
        Public Property Rerender As Boolean
            Set(value As Boolean)
                'Add Event to Rerender the slide 
                _rerender = True
            End Set
            Get
                Return _rerender

            End Get
        End Property

        Public Property Content As String
            Get
                Return _content
            End Get
            Set(value As String)
                _content = value
                Rerender = True
            End Set
        End Property

        Public Property Type As TextTypes
            Get
                Return _type
            End Get
            Set(value As TextTypes)
                _type = value
                Rerender = True
            End Set
        End Property

        Public Property Description As JObject
            Set(value As JObject)
                Rerender = True
                _description = value
            End Set
            Get
                Return _description

            End Get
        End Property

        Public Sub New(type As TextTypes, content As String, description As JObject)
            Me._type = type
            Me._content = content
            Me._description = description
            Me.Rerender = True
        End Sub

        Public Sub Render(NewSlide As Slide)
            If (CType(Description("Type").ToString(), MsoShapeType) = MsoShapeType.msoTextBox) Then
                Dim Orientation = MsoOrientation.msoOrientationHorizontal

                Dim Left As Integer = Description("Left")
                Dim Top As Integer = Description("Top")
                Dim Width As Integer = Description("Width")
                Dim Height As Integer = Description("Left")

                Dim TextBox As Shape = NewSlide.Shapes.AddTextbox(Orientation, Left, Top, Width, Height)
                TextBox.TextFrame.TextRange.Text = _content

                TextBox.Rotation = Description("Rotation")
                ' ReadOnlyProperty TextBox.ZOrderPosition = Description("ZOrderPosition")
                TextBox.Rotation = Description("Rotation")
                TextBox.Visible = CType(Description("Visible").ToString(), MsoTriState)
                TextBox.TextFrame.TextRange.Font.Name = Description("TextFontName")
                TextBox.TextFrame.TextRange.Font.Size = Description("TextFontSize")

                ' TODO : Add Correct Implementation of  Shadows, Glows and Reflection


                ' TODO : Add Correct Implementation of Fill, Line Style 


                ' TODO : Add Shape Specific Style to the Shape 

            End If

            'TODO: Add the style for the shape based on enum 

            Rerender = False
        End Sub
    End Class

End Namespace
