Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports System.Drawing.Imaging

Namespace Data.Content
    Public Enum PointTypes
        Ordered
        Unordered
        Symbols
    End Enum

    Public Class Points
        Private _content As List(Of String)
        Private _type As PointTypes
        Private _description As Dictionary(Of String, Object)
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

        Public Property Content As List(Of String)
            Get
                Return _content
            End Get
            Set(value As List(Of String))
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

        Public Property description As Dictionary(Of String, Object)
            Get
                Return _description
            End Get
            Set(value As Dictionary(Of String, Object))
                _description = value
                Rerender = True
            End Set
        End Property

        Public Sub New(content As List(Of String), description As Dictionary(Of String, Object))
            If description("Type") = MsoShapeType.msoGroup Then
                Me._type = PointTypes.Symbols
                'TODO : Add functionality to detect other list types 
            End If
            Me._type = Type
            Me._content = content
            Me._description = description

            Me.Rerender = True
        End Sub

        Public Sub Render(slide As Slide)

            If (description("Type") = MsoShapeType.msoGroup.ToString()) Then
                Dim Orientation = MsoOrientation.msoOrientationHorizontal

                Dim Left = description("Left")
                Dim Top = description("Top")
                Dim Width = description("Width")
                Dim Height = description("Left")

                Dim ItemShapes As New List(Of PowerPoint.Shape)
                Dim pointDescription = FindPointFrame(description("GroupItems"))
                For Each point In Me.Content
                    Dim shape As PowerPoint.Shape = GeneratePointItem(point, pointDescription, slide)

                    'TODO : Use the layout properties to correct the top and the left position
                    shape.Top = shape.Top + _gaps
                    shape.Left = shape.Left

                    ItemShapes.Add(shape)
                Next

                Dim shapeIds As Integer() = ItemShapes.Select(Function(s) s.Id).ToArray()
                Dim shapesToGroup As PowerPoint.ShapeRange = slide.Shapes.Range(shapeIds)
                Dim groupedPoints As PowerPoint.Shape = shapesToGroup.Group()

                groupedPoints.Top = Top
                groupedPoints.Left = Left

                ' TODO : Add Correct Implementation of  Shadows, Glows and Reflection


                ' TODO : Add Correct Implementation of Fill, Line Style 


                ' TODO : Add Shape Specific Style to the Shape 
            End If

            Rerender = False
        End Sub

        Private Function FindPointFrame(groupItems As List(Of Dictionary(Of String, Object)))
            ' TODO : Find the element with the Point text 
            Dim PointItem As Dictionary(Of String, Object)
            Dim SideItems As List(Of Dictionary(Of String, Object))

            For Each item As Dictionary(Of String, Object) In groupItems
                If item("Text").ToString().Contains("Points") Then
                    PointItem = item
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

            Dim additionalShapes = New List(Of PowerPoint.Shape)
            For Each item In pointTemplateDescription.SideItems
                Dim Type = CType(item("Type"), MsoShapeType)
                Dim Height = item("Height")
                Dim Width = item("Width")
                Dim RelativeLeft = item("RelativeLeft")
                Dim RelativeTop = item("RelativeTop")

                Dim shape = slide.Shapes.AddShape(Type, RelativeLeft, RelativeTop, Width, Height)
                ' TODO : Implement Remaining Property of the Shapes 
                additionalShapes.Add(shape)
            Next

            additionalShapes.Add(pointShape)

            Dim shapeIds As Integer() = additionalShapes.Select(Function(s) s.Id).ToArray()
            Dim shapesToGroup As PowerPoint.ShapeRange = slide.Shpaes.Range(shapeIds)
            Dim groupedShape As PowerPoint.Shape = shapesToGroup.Group()

            Return groupedShape
        End Function
    End Class

    Friend Structure PointTemplateDescription
        Public PointItem As Dictionary(Of String, Object)
        Public SideItems As List(Of Dictionary(Of String, Object))

        Public Sub New(pointItem As Dictionary(Of String, Object), sideItems As List(Of Dictionary(Of String, Object)))
            Me.PointItem = pointItem
            Me.SideItems = sideItems
        End Sub

        Public Overrides Function Equals(obj As Object) As Boolean
            If Not (TypeOf obj Is PointTemplateDescription) Then
                Return False
            End If

            Dim other = DirectCast(obj, PointTemplateDescription)
            Return EqualityComparer(Of Dictionary(Of String, Object)).Default.Equals(PointItem, other.PointItem) AndAlso
                   EqualityComparer(Of List(Of Dictionary(Of String, Object))).Default.Equals(SideItems, other.SideItems)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (PointItem, SideItems).GetHashCode()
        End Function

        Public Sub Deconstruct(ByRef pointItem As Dictionary(Of String, Object), ByRef sideItems As List(Of Dictionary(Of String, Object)))
            pointItem = Me.PointItem
            sideItems = Me.SideItems
        End Sub

        Public Shared Widening Operator CType(value As PointTemplateDescription) As (PointItem As Dictionary(Of String, Object), SideItems As List(Of Dictionary(Of String, Object)))
            Return (value.PointItem, value.SideItems)
        End Operator

        Public Shared Widening Operator CType(value As (PointItem As Dictionary(Of String, Object), SideItems As List(Of Dictionary(Of String, Object)))) As PointTemplateDescription
            Return New PointTemplateDescription(value.PointItem, value.SideItems)
        End Operator
    End Structure
End Namespace
