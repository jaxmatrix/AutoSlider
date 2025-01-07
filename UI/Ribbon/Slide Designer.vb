Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports Shape = Microsoft.Office.Interop.PowerPoint.Shape
Imports System.Windows.Forms
Imports Application = Microsoft.Office.Interop.PowerPoint.Application
Imports AutoSlider.SlideTemplates
Imports System.Diagnostics
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json
Imports System.IO
Imports System.Collections.Specialized

Public Class Slide_Designer
    Private Sub Slide_Designer_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnGenerate.Click
        Dim pptApp As Application = Globals.ThisAddIn.Application

        If pptApp.Presentations.Count > 0 AndAlso
                pptApp.ActiveWindow IsNot Nothing AndAlso
                pptApp.ActiveWindow.View.Slide IsNot Nothing Then
            Dim activeSlide As Slide = pptApp.ActiveWindow.View.Slide

            Dim shapeList As New List(Of String)
            For Each shp As Shape In activeSlide.Shapes
                Dim shapeInfo As String = shp.Name

                If shp.HasTextFrame <> MsoTriState.msoTrue Then
                    Continue For
                End If

                If shp.TextFrame.HasText <> MsoTriState.msoTrue Then
                    Continue For

                End If

                Dim textContent As String = shp.TextFrame.TextRange.Text
                shapeInfo &= " | Text: " & textContent
                shapeList.Add(shapeInfo)
            Next

            If shapeList.Count > 0 Then
                MessageBox.Show(String.Join(Environment.NewLine, shapeList),
                                "Shapes in Active Slide",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                )
            Else
                MessageBox.Show("No shapes found on this slide",
                                "Shapes in active Slide",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                )
            End If
        Else
            MessageBox.Show("No active slide found.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)

        End If


    End Sub

    Private Sub btnAutoSlide_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAutoSlide.Click
        Dim pptApp = Globals.ThisAddIn.Application
        Dim activePresentation As Presentation = pptApp.ActivePresentation
        Dim activeWindow As DocumentWindow = pptApp.ActiveWindow

        Debug.WriteLine($"Entering the btnAutoSlide Function {pptApp} {activePresentation} {activeWindow}")
        Debug.WriteLine($"Current Slide Porps {activeWindow.ViewType} ")

        If activeWindow.ViewType = PpViewType.ppViewNormal Then
            Dim currentSlide As Slide = activeWindow.View.Slide
            Dim newSlideIndex As Integer = currentSlide.SlideIndex + 1

            Dim newSlide As Slide = activePresentation.Slides.Add(newSlideIndex, PpSlideLayout.ppLayoutBlank)

            'TODO: Write Steps to Extract the text from the slide 

            'TODO: Write Steps to Transform the text to the different contents 

            'TODO: Write Steps to Find the Suitable Layouts from the predefined layouts 

            'Generating New Slide based on the available content
            Dim heading As String = "Test Heading"
            Dim desc As String = "Test Description"
            Dim pointList As List(Of String) = New List(Of String) From {
                "Test Point 1",
                "Test Point 2",
                "Test Point 3",
                "Test Point 4"
            }

            Dim NewSlideGenerator As SlideTemplates.Test1 = New SlideTemplates.Test1(newSlide, heading, desc, pointList)
            NewSlideGenerator.Render()
            Debug.WriteLine("Generation Succesful")
        End If

    End Sub

    Private Sub btnCaptureLayout_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCaptureLayout.Click
        'Get the active power point presentation and extract the information from the layout 
        Dim pptApp As PowerPoint.Application = Globals.ThisAddIn.Application
        Dim presentation As PowerPoint.Presentation = pptApp.ActivePresentation
        Dim activeWindow As DocumentWindow = pptApp.ActiveWindow


        'Get the current slide 
        If activeWindow.ViewType = PpViewType.ppViewNormal Then
            Dim currentSlide As Slide = activeWindow.View.Slide
            Dim LayoutProperties As New Dictionary(Of String, Object)
            Dim CosmeticShapes = New List(Of Dictionary(Of String, Object))

            For Each shape As PowerPoint.Shape In currentSlide.Shapes
                Dim elementType As SlideComponents = SlideComponents.Cosmetic

                If shape.Type = MsoShapeType.msoGroup Then
                    For Each groupItem As Shape In shape.GroupItems
                        If groupItem.Type = MsoShapeType.msoTextBox Then
                            If groupItem.HasTextFrame Then
                                If groupItem.TextFrame.HasText Then
                                    Dim textHint As String = groupItem.TextFrame.TextRange.Text
                                    elementType = CheckTextHint(textHint)
                                End If
                            End If
                        End If
                    Next
                    If elementType = SlideComponents.Cosmetic Then
                        CosmeticShapes.Add(CaptureGroupProperties(shape))
                    Else
                        LayoutProperties(SlideComponentsName(elementType)) = CaptureGroupProperties(shape)
                    End If

                    Debug.WriteLine($"Detected Group {elementType}")
                Else
                    If shape.HasTextFrame Then
                        If shape.TextFrame.HasText Then
                            Dim textHint As String = shape.TextFrame.TextRange.Text
                            elementType = CheckTextHint(textHint)
                        End If
                    End If

                    If elementType = SlideComponents.Cosmetic Then
                        CosmeticShapes.Add(CaptureShapeProperties(shape))
                    Else
                        LayoutProperties(SlideComponentsName(elementType)) = CaptureShapeProperties(shape)
                    End If

                End If
            Next
            LayoutProperties(SlideComponentsName(SlideComponents.Cosmetic)) = CosmeticShapes

            Dim json As String = JsonConvert.SerializeObject(LayoutProperties, Formatting.Indented)
            Dim randomFileName As String = "DictionaryData_" & Guid.NewGuid().ToString() & ".json"

            ' Define the file path
            Dim filePath As String = Path.Combine("C:\Temp\", randomFileName)

            ' Ensure the directory exists
            Directory.CreateDirectory("C:\Temp\")

            Try
                File.WriteAllText(filePath, json)
                Debug.WriteLine("Dictionary saved as JSON at: " & filePath)
            Catch ex As Exception
                Debug.WriteLine("Error saving JSON: " & ex.Message)
            End Try
        End If

    End Sub


    Private Function CaptureShapeProperties(shape As PowerPoint.Shape)
        Dim shapeProperties As New Dictionary(Of String, Object)

        ' General properties
        shapeProperties("Name") = shape.Name
        shapeProperties("Type") = shape.Type.ToString()
        shapeProperties("Left") = shape.Left
        shapeProperties("Top") = shape.Top
        shapeProperties("Width") = shape.Width
        shapeProperties("Height") = shape.Height
        shapeProperties("Rotation") = shape.Rotation
        shapeProperties("ZOrderPosition") = shape.ZOrderPosition
        shapeProperties("Visible") = shape.Visible.ToString()

        ' Fill and line properties
        If shape.Fill.Visible = MsoTriState.msoTrue Then
            shapeProperties("FillColor") = shape.Fill.ForeColor.RGB
        End If
        If shape.Line.Visible = MsoTriState.msoTrue Then
            shapeProperties("LineColor") = shape.Line.ForeColor.RGB
            shapeProperties("LineWeight") = shape.Line.Weight
        End If

        ' Text properties
        If shape.HasTextFrame Then
            Dim textFrame = shape.TextFrame
            shapeProperties("HasText") = textFrame.HasText
            If textFrame.HasText = MsoTriState.msoTrue Then
                shapeProperties("Text") = textFrame.TextRange.Text
                shapeProperties("TextFontName") = textFrame.TextRange.Font.Name
                shapeProperties("TextFontSize") = textFrame.TextRange.Font.Size
            End If
        End If

        ' Shape-specific properties
        If shape.Type = Office.MsoShapeType.msoAutoShape OrElse
           shape.Type = Office.MsoShapeType.msoFreeform Then
            shapeProperties("AutoShapeType") = shape.AutoShapeType.ToString()
        End If

        ' Shadow and effects
        shapeProperties("HasShadow") = shape.Shadow.Visible.ToString()
        shapeProperties("HasGlow") = shape.Glow.Radius > 0
        shapeProperties("HasReflection") = shape.Reflection.Type.ToString()

        ' Tags and metadata
        If shape.Tags.Count > 0 Then
            Dim tags As New Dictionary(Of String, String)
            For i As Integer = 1 To shape.Tags.Count
                tags.Add(shape.Tags.Name(i), shape.Tags.Value(i))
            Next
            shapeProperties("Tags") = tags
        End If

        ' Serialize the properties to JSON
        Return shapeProperties
    End Function

    Private Function CaptureGroupProperties(groupShape As PowerPoint.Shape)
        ' Ensure the shape is a group
        If groupShape.Type <> Office.MsoShapeType.msoGroup Then
            Throw New ArgumentException("The provided shape is not a group.")
        End If

        ' Dictionary to store group properties
        Dim groupProperties As New Dictionary(Of String, Object)

        ' General properties
        groupProperties("Name") = groupShape.Name
        groupProperties("Type") = groupShape.Type.ToString()
        groupProperties("Left") = groupShape.Left
        groupProperties("Top") = groupShape.Top
        groupProperties("Width") = groupShape.Width
        groupProperties("Height") = groupShape.Height
        groupProperties("Rotation") = groupShape.Rotation
        groupProperties("ZOrderPosition") = groupShape.ZOrderPosition
        groupProperties("LockAspectRatio") = groupShape.LockAspectRatio.ToString()
        groupProperties("Visible") = groupShape.Visible.ToString()

        ' Fill and line properties
        If groupShape.Fill.Visible = MsoTriState.msoTrue Then
            groupProperties("FillColor") = groupShape.Fill.ForeColor.RGB
        End If
        If groupShape.Line.Visible = MsoTriState.msoTrue Then
            groupProperties("LineColor") = groupShape.Line.ForeColor.RGB
            groupProperties("LineWeight") = groupShape.Line.Weight
        End If

        ' Shadow and effects
        groupProperties("HasShadow") = groupShape.Shadow.Visible.ToString()
        groupProperties("ShadowColor") = If(groupShape.Shadow.Visible = MsoTriState.msoTrue, groupShape.Shadow.ForeColor.RGB, Nothing)
        groupProperties("HasGlow") = groupShape.Glow.Radius > 0
        groupProperties("HasReflection") = groupShape.Reflection.Type.ToString()

        ' Tags and metadata
        If groupShape.Tags.Count > 0 Then
            Dim tags As New Dictionary(Of String, String)
            For i As Integer = 1 To groupShape.Tags.Count
                tags.Add(groupShape.Tags.Name(i), groupShape.Tags.Value(i))
            Next
            groupProperties("Tags") = tags
        End If

        ' Group items
        Dim groupItems As New List(Of Object)
        For Each groupItem As PowerPoint.Shape In groupShape.GroupItems
            Dim itemProperties As New Dictionary(Of String, Object)
            itemProperties("Name") = groupItem.Name
            itemProperties("Type") = groupItem.Type.ToString()
            itemProperties("RelativeLeft") = groupItem.Left
            itemProperties("RelativeTop") = groupItem.Top
            itemProperties("Width") = groupItem.Width
            itemProperties("Height") = groupItem.Height
            itemProperties("Rotation") = groupItem.Rotation

            ' Text properties
            If groupItem.HasTextFrame = MsoTriState.msoTrue Then
                itemProperties("HasText") = True
                itemProperties("Text") = groupItem.TextFrame.TextRange.Text
            Else
                itemProperties("HasText") = False
            End If

            ' Fill and line properties
            If groupItem.Fill.Visible = MsoTriState.msoTrue Then
                itemProperties("FillColor") = groupItem.Fill.ForeColor.RGB
            End If
            If groupItem.Line.Visible = MsoTriState.msoTrue Then
                itemProperties("LineColor") = groupItem.Line.ForeColor.RGB
                itemProperties("LineWeight") = groupItem.Line.Weight
            End If

            groupItems.Add(itemProperties)
        Next
        groupProperties("GroupItems") = groupItems

        ' Serialize to JSON
        Return groupProperties
    End Function
    Private Function CheckTextHint(textHint As String) As SlideTemplates.SlideComponents
        Debug.WriteLine($"Text Hint : {textHint}")
        Select Case textHint
            Case "Title"
                Return SlideTemplates.SlideComponents.TextHeading
            Case "Text"
                Return SlideTemplates.SlideComponents.Text
            Case "List Item"
                Debug.WriteLine($"Text Hint - List Item Detected : {textHint}")
                Return SlideTemplates.SlideComponents.TextList
            Case "Image"
                Return SlideTemplates.SlideComponents.Image
            Case Else
                Debug.WriteLine($"Text Hint - Cosmetic Item Detected : {textHint}")
                Return SlideTemplates.SlideComponents.Cosmetic
        End Select
    End Function
End Class
