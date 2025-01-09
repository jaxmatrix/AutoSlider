Imports Microsoft.Office.Core
Imports Newtonsoft.Json.Linq
Imports AutoSlider.SlideTemplates.Processor.General
Imports System.Drawing.Imaging

Namespace Data.Content
    Public Class Cosmetic
        Private _rerender As Boolean = False
        Private _description As JArray

        Public Property Rerender As Boolean
            Set(value As Boolean)
                _rerender = value
            End Set
            Get
                Return _rerender
            End Get
        End Property

        Public Property Description As JArray
            Set(value As JArray)
                _description = value
                Rerender = True
            End Set
            Get
                Return _description
            End Get
        End Property

        Public Sub New(description As JArray)
            Me.Description = description
        End Sub


        Public Sub Render(slide As PowerPoint.Slide)
            Dim cosmeticData = Description
            For Each item In cosmeticData
                Dim shapeTypeString As String = item("Type").ToString()
                Dim shapeType As Office.MsoShapeType = TextToEnum(Of Office.MsoShapeType)(shapeTypeString)

                ' Handle different types of shapes
                If shapeType = Office.MsoShapeType.msoAutoShape Then
                    ' Create an AutoShape
                    Dim shape = slide.Shapes.AddShape(TextToEnum(Of Office.MsoAutoShapeType)(item("AutoShapeType").ToString()),
                                              CDbl(item("Left")),
                                              CDbl(item("Top")),
                                              CDbl(item("Width")),
                                              CDbl(item("Height")))

                    ' Apply additional properties
                    shape.Name = item("Name").ToString()
                    shape.Rotation = CDbl(item("Rotation"))
                    'shape.ZOrderPosition = CInt(item("ZOrderPosition"))
                    shape.Visible = TextToEnum(Of Office.MsoTriState)(item("Visible").ToString())
                    shape.Fill.ForeColor.RGB = CInt(item("FillColor"))
                    shape.Line.ForeColor.RGB = CInt(item("LineColor"))
                    shape.Line.Weight = CSng(item("LineWeight"))

                    ' Text Handling
                    If CBool(item("HasText")) Then
                        shape.TextFrame.TextRange.Text = item("Text").ToString()
                    End If

                ElseIf shapeType = Office.MsoShapeType.msoGroup Then
                    ' Handle Group Shape
                    Dim groupItemsData = item("GroupItems")
                    Dim tempShapes As New List(Of PowerPoint.Shape)

                    ' Add group items
                    For Each groupItem In groupItemsData
                        Dim groupShapeType As Office.MsoShapeType = TextToEnum(Of MsoShapeType)(groupItem("Type").ToString())

                        Dim groupShape As PowerPoint.Shape
                        If groupShapeType = Office.MsoShapeType.msoAutoShape Then
                            groupShape = slide.Shapes.AddShape(TextToEnum(Of MsoAutoShapeType)(groupItem("AutoShapeType").ToString()),
                                                       CDbl(groupItem("RelativeLeft")),
                                                       CDbl(groupItem("RelativeTop")),
                                                       CDbl(groupItem("Width")),
                                                       CDbl(groupItem("Height")))
                        Else
                            groupShape = slide.Shapes.AddShape(TextToEnum(Of MsoShapeType)(groupItem("Type").ToString()),
                                                       CDbl(groupItem("RelativeLeft")),
                                                       CDbl(groupItem("RelativeTop")),
                                                       CDbl(groupItem("Width")),
                                                       CDbl(groupItem("Height")))

                        End If


                        groupShape.Name = groupItem("Name").ToString()
                        groupShape.Rotation = CDbl(groupItem("Rotation"))
                        groupShape.Fill.ForeColor.RGB = CInt(groupItem("FillColor"))
                        groupShape.Line.ForeColor.RGB = CInt(groupItem("LineColor"))
                        groupShape.Line.Weight = CSng(groupItem("LineWeight"))

                        ' Text Handling for group items
                        If CBool(groupItem("HasText")) Then
                            groupShape.TextFrame.TextRange.Text = groupItem("Text").ToString()
                        End If

                        tempShapes.Add(groupShape)
                    Next

                    ' Group the shapes
                    Dim shapeNames = tempShapes.Select(Function(s) s.Name).ToArray()
                    Dim shapeRange = slide.Shapes.Range(shapeNames)
                    Dim groupedShape = shapeRange.Group()

                    ' Apply additional properties to the group
                    groupedShape.Name = item("Name").ToString()
                    groupedShape.Left = CDbl(item("Left"))
                    groupedShape.Top = CDbl(item("Top"))
                    groupedShape.Width = CDbl(item("Width"))
                    groupedShape.Height = CDbl(item("Height"))
                    groupedShape.Rotation = CDbl(item("Rotation"))
                    groupedShape.Visible = TextToEnum(Of Office.MsoTriState)(item("Visible").ToString())

                Else
                    Console.WriteLine($"Unsupported shape type: {shapeTypeString}")
                End If
            Next
        End Sub


    End Class
End Namespace