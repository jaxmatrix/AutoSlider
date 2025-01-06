Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports Shape = Microsoft.Office.Interop.PowerPoint.Shape
Imports System.Windows.Forms
Imports Application = Microsoft.Office.Interop.PowerPoint.Application
Imports AutoSlider.SlideTemplates
Imports System.Diagnostics

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
End Class
