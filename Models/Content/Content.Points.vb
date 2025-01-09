Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports System.Drawing.Imaging
Imports Newtonsoft.Json.Linq
Imports AutoSlider.SlideTemplates.Processor.General
Imports System.Diagnostics

Namespace Data.Content
    Public Enum PointTypes
        Ordered
        Unordered
        Symbols
    End Enum

    Public Class Points
        Private _content As JArray
        Private _type As PointTypes
        Private _description As JObject
        Private _rerender As Boolean
        Private _gaps As Integer = 20
        Private _cols As Integer = 1
        Private _direction As MsoOrientation = MsoOrientation.msoOrientationVertical

        Public Property Rerender As Boolean
            Set(value As Boolean)
                _rerender = True
                'Create a event to trigger rerender
            End Set
            Get
                Return _rerender
            End Get
        End Property

        Public Property Content As JArray
            Get
                Return _content
            End Get
            Set(value As JArray)
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

        Public Property description As JObject
            Get
                Return _description
            End Get
            Set(value As JObject)
                _description = value
                Rerender = True
            End Set
        End Property

        Public Sub New(content As JArray, description As JObject)
            If TextToEnum(Of MsoShapeType)(description("Type").ToString()) = MsoShapeType.msoGroup Then
                Me._type = PointTypes.Symbols
                'TODO : Add functionality to detect other list types 
            End If
            Me._type = Type
            Me._content = content
            Me._description = description

            Me.Rerender = True
        End Sub

        Public Sub Render(slide As Slide)

            If (TextToEnum(Of MsoShapeType)(description("Type").ToString()) = MsoShapeType.msoGroup) Then
                Dim Orientation = MsoOrientation.msoOrientationHorizontal

                Dim Left As Integer = description("Left")
                Dim Top As Integer = description("Top")
                Dim Width As Integer = description("Width")
                Dim Height As Integer = description("Left")

                'Adding the height of the item to ensure that sufficient gap is maintained
                'Do a proper calculation and improve the implementation
                _gaps += Height

                Dim ItemShapes As New List(Of PowerPoint.Shape)
                Dim pointDescription = FindPointFrame(description("GroupItems"))
                Dim i As Integer = 0
                For Each point In Me.Content
                    Dim shape As PowerPoint.Shape = GeneratePointItem(point, pointDescription, slide)

                    'TODO : Use the layout properties to correct the top and the left position
                    shape.Top = shape.Top + i * _gaps
                    shape.Left = shape.Left

                    ItemShapes.Add(shape)
                    i += 1
                Next

                Dim shapeIds As String() = ItemShapes.Select(Function(s) s.Name).ToArray()
                Try
                    Dim shapesToGroup As PowerPoint.ShapeRange = slide.Shapes.Range(shapeIds)
                    Dim groupedPoints As PowerPoint.Shape = shapesToGroup.Group()

                    groupedPoints.Top = Top
                    groupedPoints.Left = Left

                Catch ex As Exception
                    Debug.WriteLine($"Error Creating the elements : {ex.Message}")
                End Try

                ' TODO : Add Correct Implementation of  Shadows, Glows and Reflection


                ' TODO : Add Correct Implementation of Fill, Line Style 


                ' TODO : Add Shape Specific Style to the Shape 
            End If

            Rerender = False
        End Sub

        Private Function FindPointFrame(groupItems As JArray)
            ' TODO : Find the element with the Point text 
            Dim PointItem As New JObject()
            Dim SideItems As New JArray()

            For Each item As JObject In groupItems
                Dim itemType As MsoTriState = TextToEnum(Of MsoShapeType)(item("Type").ToString())
                If itemType = MsoShapeType.msoTextBox Then
                    Dim hasText As MsoTriState = TextToEnum(Of MsoTriState)(item("HasText").ToString())
                    If hasText = MsoTriState.msoTrue Then
                        If item("Text").ToString().Contains("Points") Then
                            PointItem = item
                        End If
                    Else
                        SideItems.Add(item)
                        'TODO : Add functionality of complext list type that can store different structure
                    End If
                Else
                    SideItems.Add(item)
                End If
            Next

            Return New PointTemplateDescription(PointItem, SideItems)
        End Function


        Private Function GeneratePointItem(pointText As String, pointTemplateDescription As PointTemplateDescription, slide As PowerPoint.Slide)
            Dim pointTop = pointTemplateDescription.PointItem("RelativeTop")
            Dim pointLeft = pointTemplateDescription.PointItem("RelativeLeft")
            Dim pointHeight = pointTemplateDescription.PointItem("Height")
            Dim pointWidth = pointTemplateDescription.PointItem("Width")
            Dim pointShape = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal,
                                                     pointLeft, pointTop, pointWidth, pointHeight)
            pointShape.TextFrame.TextRange.Text = pointText
            Debug.WriteLine($"PointShape : {pointShape.Id}")

            Dim additionalShapes = New List(Of PowerPoint.Shape)
            additionalShapes.Add(pointShape)
            For Each item In pointTemplateDescription.SideItems
                Dim Type = TextToEnum(Of MsoShapeType)(item("Type").ToString())
                Dim Height = item("Height")
                Dim Width = item("Width")
                Dim RelativeLeft = item("RelativeLeft")
                Dim RelativeTop = item("RelativeTop")

                Dim shape = slide.Shapes.AddShape(Type, RelativeLeft, RelativeTop, Width, Height)
                ' TODO : Implement Remaining Property of the Shapes 
                Debug.WriteLine($"Additional Shape : {shape.Id}")
                additionalShapes.Add(shape)
            Next


            Dim shapeIds As String() = additionalShapes.Select(Function(s) s.Name).ToArray()
            Try
                Dim shapesToGroup As PowerPoint.ShapeRange = slide.Shapes.Range(shapeIds)
                Dim groupedShape As PowerPoint.Shape = shapesToGroup.Group()
                Return groupedShape

            Catch ex As Exception
                Debug.WriteLine($"Error Grouping shapes: {ex.Message}")
            End Try
        End Function
    End Class

    Friend Structure PointTemplateDescription
        Public PointItem As JObject
        Public SideItems As JArray

        Public Sub New(pointItem As JObject, sideItems As JArray)
            Me.PointItem = pointItem
            Me.SideItems = sideItems
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            If Not (TypeOf obj Is PointTemplateDescription) Then
                Return False
            End If

            Dim other = DirectCast(obj, PointTemplateDescription)
            Return EqualityComparer(Of JObject).Default.Equals(PointItem, other.PointItem) AndAlso
                   EqualityComparer(Of JArray).Default.Equals(SideItems, other.SideItems)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (PointItem, SideItems).GetHashCode()
        End Function

        Public Sub Deconstruct(ByRef pointItem As JObject, ByRef sideItems As JArray)
            pointItem = Me.PointItem
            sideItems = Me.SideItems
        End Sub

        Public Shared Widening Operator CType(value As PointTemplateDescription) As (PointItem As JObject, SideItems As JArray)
            Return (value.PointItem, value.SideItems)
        End Operator

        Public Shared Widening Operator CType(value As (PointItem As JObject, SideItems As JArray)) As PointTemplateDescription
            Return New PointTemplateDescription(value.PointItem, value.SideItems)
        End Operator
    End Structure
End Namespace
