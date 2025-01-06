Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports Shape = Microsoft.Office.Interop.PowerPoint.Shape
Imports System.Windows.Forms
Imports Application = Microsoft.Office.Interop.PowerPoint.Application

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
End Class
