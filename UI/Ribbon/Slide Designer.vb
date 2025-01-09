Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core
Imports Shape = Microsoft.Office.Interop.PowerPoint.Shape
Imports System.Windows.Forms
Imports Application = Microsoft.Office.Interop.PowerPoint.Application
Imports AutoSlider.SlideTemplates
Imports System.Diagnostics
Imports Newtonsoft.Json
Imports System.IO
Imports AutoSlider.SlideTemplates.Enums
Imports Newtonsoft.Json.Linq
Imports System.Net.Http
Imports System.Net.Http.Headers

Public Class Slide_Designer
    Private Sub Slide_Designer_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load

    End Sub

    Private Async Sub btnGenerate_Click(sender As Object, e As RibbonControlEventArgs) Handles btnGenerate.Click
        Dim pptApp As Application = Globals.ThisAddIn.Application

        If pptApp.Presentations.Count > 0 AndAlso
                pptApp.ActiveWindow IsNot Nothing AndAlso
                pptApp.ActiveWindow.View.Slide IsNot Nothing Then
            Dim activeSlide As Slide = pptApp.ActiveWindow.View.Slide

            Dim shapeList As New List(Of String)
            For Each shp As Shape In activeSlide.Shapes
                If shp.Type = MsoShapeType.msoGroup Then
                    For Each groupItem In shp.GroupItems
                        Dim shapeInfo2 As String = shp.Name

                        If groupItem.HasTextFrame <> MsoTriState.msoTrue Then
                            Continue For
                        End If

                        If groupItem.TextFrame.HasText <> MsoTriState.msoTrue Then
                            Continue For

                        End If

                        Dim textContent2 As String = groupItem.TextFrame.TextRange.Text
                        shapeInfo2 = textContent2
                        shapeList.Add(shapeInfo2)
                    Next
                End If
                Dim shapeInfo As String = shp.Name

                If shp.HasTextFrame <> MsoTriState.msoTrue Then
                    Continue For
                End If

                If shp.TextFrame.HasText <> MsoTriState.msoTrue Then
                    Continue For

                End If

                Dim textContent As String = shp.TextFrame.TextRange.Text
                shapeInfo = textContent
                shapeList.Add(shapeInfo)
            Next

            If shapeList.Count > 0 Then
                Try
                    Using client As New HttpClient()
                        Dim DummyClasses As New JArray()
                        DummyClasses.Add(New JObject(New JProperty("name", "Title"),
                                                     New JProperty("value", 1))
                                        )
                        DummyClasses.Add(New JObject(New JProperty("name", "Description"),
                                                     New JProperty("value", 1))
                                        )
                        DummyClasses.Add(New JObject(New JProperty("name", "Points"),
                                                     New JProperty("value", 5))
                                        )

                        Dim jsonDummyClasses = JsonConvert.SerializeObject(DummyClasses)


                        Dim ContentObject As New JObject()
                        Dim jsonLines = JsonConvert.SerializeObject(shapeList)

                        ContentObject.Add("lines", JArray.FromObject(shapeList))
                        ContentObject.Add("classes", DummyClasses)

                        Dim jsonContent As String = JsonConvert.SerializeObject(ContentObject)
                        Debug.WriteLine($"JSON CONTENT {ContentObject}")

                        Dim content As New StringContent(jsonContent, Encoding.UTF8, "application/json")
                        Dim api As String = "http://localhost:8000/classify/"
                        Dim response As HttpResponseMessage = Await client.PostAsync(api, content)

                        If response.IsSuccessStatusCode Then
                            Dim responseBody As String = Await response.Content.ReadAsStringAsync()
                            MsgBox($"AI Generated Content {responseBody}")
                        End If

                    End Using

                Catch ex As Exception
                    Debug.WriteLine($"Error occured while generating AI insigths {ex.Message}")

                End Try
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

    Private Async Sub btnCaptureLayout_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCaptureLayout.Click
        'Get the active power point presentation and extract the information from the layout 
        Dim pptApp As Application = Globals.ThisAddIn.Application
        Dim presentation As Presentation = pptApp.ActivePresentation
        Dim activeWindow As DocumentWindow = pptApp.ActiveWindow


        'Get the current slide 
        If activeWindow.ViewType = PpViewType.ppViewNormal Then


            Dim currentSlide As Slide = activeWindow.View.Slide
            Dim LayoutProperties As New Dictionary(Of String, Object)
            Dim CosmeticShapes = New List(Of Dictionary(Of String, Object))

            Dim previewImageId As String = SaveAndSendSlidePreview(currentSlide)
            Debug.WriteLine($"Preview Image Generated with Id {previewImageId}")


            For Each shape As Shape In currentSlide.Shapes
                Dim elementType As SlideComponents = LayoutComponents.Cosmetic

                If shape.Type = MsoShapeType.msoGroup Then
                    For Each groupItem As Shape In shape.GroupItems
                        If groupItem.Type = MsoShapeType.msoTextBox Then
                            If groupItem.HasTextFrame Then
                                If groupItem.TextFrame.HasText Then
                                    Dim textHint As String = groupItem.TextFrame.TextRange.Text
                                    elementType = GetLayoutComponentEnum(textHint)
                                End If
                            End If
                        End If
                    Next
                    If elementType = LayoutComponents.Cosmetic Then
                        CosmeticShapes.Add(CaptureGroupProperties(shape))
                    Else
                        LayoutProperties(GetLayoutComponentName(elementType)) = CaptureGroupProperties(shape)
                    End If

                    Debug.WriteLine($"Detected Group {elementType}")
                ElseIf shape.Type = MsoShapeType.msoPicture Or shape.Type = MsoShapeType.msoLinkedPicture Then

                    elementType = LayoutComponents.Image
                    LayoutProperties(GetLayoutComponentName(elementType)) = CapturePictureProperties(shape)

                Else
                    If shape.HasTextFrame Then
                        If shape.TextFrame.HasText Then
                            Dim textHint As String = shape.TextFrame.TextRange.Text
                            elementType = GetLayoutComponentEnum(textHint)
                        End If
                    End If

                    If elementType = LayoutComponents.Cosmetic Then
                        CosmeticShapes.Add(CaptureShapeProperties(shape))
                    Else
                        LayoutProperties(GetLayoutComponentName(elementType)) = CaptureShapeProperties(shape)
                    End If

                End If
            Next
            LayoutProperties(GetLayoutComponentName(LayoutComponents.Cosmetic)) = CosmeticShapes

            'TODO : Send the Data to the backend server to store the  data in MongoDB database

            ' Define the file path

            Dim layoutContentDesc As New JObject()
            For Each prop As String In LayoutProperties.Keys
                Dim keyCount As Integer = Processor.LayoutComponents.GetContentRequriement(prop).ToString()
                If keyCount = 0 Then
                    Continue For
                End If
                layoutContentDesc.Add(prop, keyCount)
            Next


            Dim randomFileName As String = "DictionaryData_" & Guid.NewGuid().ToString() & ".json"
            Dim filePath As String = Path.Combine("C:\Temp\", randomFileName)
            ' Ensure the directory exists
            Directory.CreateDirectory("C:\Temp\")
            Dim layoutApi As String = "http://localhost:8000/layouts/"
            LayoutProperties("PreviewImageId") = previewImageId

            Try
                Dim descPrompt As New GetDescriptionPrompt()

                Dim layoutDesc As String = ""

                If descPrompt.ShowDialog() = DialogResult.OK Then
                    layoutDesc = descPrompt.UserInput
                End If

                If layoutDesc = "" Then
                    Throw New Exception("Empty layout description")
                End If

                Dim jsonString As String = JsonConvert.SerializeObject(LayoutProperties, Formatting.Indented)
                Dim jsonContentDesc As String = JsonConvert.SerializeObject(layoutContentDesc)

                Dim RequestObject As New JObject From {
                    {"jsonString", jsonString},
                    {"layoutContentDesc", jsonContentDesc},
                    {"description", layoutDesc}
                }

                Dim requestString As New StringContent(JsonConvert.SerializeObject(RequestObject), Encoding.UTF8, "application/json")
                Using client As New HttpClient()
                    Dim response As HttpResponseMessage = Await client.PostAsync(layoutApi, requestString)

                    If response.IsSuccessStatusCode Then
                        Dim responseBody As String = Await response.Content.ReadAsStringAsync()
                        Debug.WriteLine($"Layout Update Response : {responseBody}")
                    Else
                        Throw New Exception("Error updating the layout")
                    End If
                End Using

                File.WriteAllText(filePath, jsonString)
                Debug.WriteLine("Dictionary saved as JSON at: " & filePath)
            Catch ex As Exception
                Debug.WriteLine("Error saving JSON: " & ex.Message)
            End Try
        End If

    End Sub

    Private Function CaptureShapeProperties(shape As Shape)
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

            Dim subProperties = CaptureShapeProperties(groupItem)
            Dim combinedDict = CombineDictionaries(subProperties, itemProperties)

            ' TODO : Add Additional Case for creating a property if the text contains
            ' points and what type of points it is 

            groupItems.Add(combinedDict)
        Next
        groupProperties("GroupItems") = groupItems

        ' Serialize to JSON
        Return groupProperties
    End Function

    Private Function CaptureGroupProperties(groupShape As PowerPoint.Shape, Hint As SlideComponents)
        Dim groupProperties = CaptureGroupProperties(groupShape)
        ' TODO : Add additional functionality based on the text parsing of content of the list 
    End Function
    Private Function GetSlideComponentEnum(textHint As String) As SlideComponents
        Try
            Return Processor.SlideComponents.StringToEnum(textHint)
        Catch ex As Exception
            Debug.WriteLine($"No Slide Component Description Detected : Fallback to cosmetic : {ex}")
            Return SlideComponents.Cosmetic
        End Try
    End Function
    Private Function GetSlideComponentName(Hint As SlideComponents) As String
        Try
            Return Processor.SlideComponents.EnumToString(Hint)
        Catch ex As Exception
            Debug.WriteLine($"No Slide Component Description Detected : Fallback to cosmetic : {ex}")
            Return Processor.SlideComponents.EnumToString(SlideComponents.Cosmetic)
        End Try
    End Function

    Private Function GetLayoutComponentEnum(textHint As String) As LayoutComponents
        Try
            Return Processor.LayoutComponents.StringToEnum(textHint)
        Catch ex As Exception
            Debug.WriteLine($"No Slide Component Description Detected : Fallback to cosmetic : {ex}")
            Return LayoutComponents.Cosmetic
        End Try
    End Function
    Private Function GetLayoutComponentName(Hint As LayoutComponents) As String
        Try
            Return Processor.LayoutComponents.EnumToString(Hint)
        Catch ex As Exception
            Debug.WriteLine($"No Slide Component Description Detected : Fallback to cosmetic : {ex}")
            Return Processor.LayoutComponents.EnumToString(LayoutComponents.Cosmetic)
        End Try
    End Function

    Public Function CapturePictureProperties(
            shape As Shape
        ) As Dictionary(Of String, Object)

        Dim tempFolderPath As String = "C:\Temp\ImageCollection"
        ' Ensure the temporary folder exists
        If Not Directory.Exists(tempFolderPath) Then
            Directory.CreateDirectory(tempFolderPath)
        End If

        ' Validate if the shape is a picture or linked picture
        If shape.Type <> MsoShapeType.msoPicture AndAlso shape.Type <> MsoShapeType.msoLinkedPicture Then
            Throw New ArgumentException("The shape is not a picture or linked picture.")
        End If

        ' Create a dictionary to store the properties
        Dim properties As New Dictionary(Of String, Object)()

        ' Extract properties
        properties("Name") = shape.Name
        properties("Width") = shape.Width
        properties("Height") = shape.Height
        properties("Left") = shape.Left
        properties("Top") = shape.Top
        properties("Rotation") = shape.Rotation
        properties("ZOrderPosition") = shape.ZOrderPosition
        properties("AlternativeText") = shape.AlternativeText

        ' Save the picture to the temporary folder
        Dim tempImagePath As String = Path.Combine(tempFolderPath, $"{Guid.NewGuid().ToString()}.png")
        shape.Export(tempImagePath, PpShapeFormat.ppShapeFormatPNG)
        properties("TempImagePath") = tempImagePath

        ' Add additional properties for linked pictures
        If shape.Type = MsoShapeType.msoLinkedPicture Then
            properties("LinkFormat.SourceFullName") = shape.LinkFormat.SourceFullName
            properties("LinkFormat.AutoUpdate") = shape.LinkFormat.AutoUpdate
        End If

        ' Return the properties dictionary
        Return properties
    End Function

    Private Sub btnTestGenerator_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTestGenerator.Click
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
            Dim layoutPath As String = "C:\Temp\LayoutTest.json"

            Dim layoutContent As String = File.ReadAllText(layoutPath)

            Dim layoutData As JObject = JsonConvert.DeserializeObject(Of JObject)(layoutContent)

            Debug.WriteLine($"{layoutData}")

            Dim contentJObject = New JObject()
            contentJObject("Title") = "Test Title"
            contentJObject("Image") = "C:\Temp\ImageCollection\Picture 13.png"
            contentJObject("Description") = "Test Title"
            contentJObject("Points") = New JArray("Item 1", "Item 2", "Item 3")
            contentJObject("Cosmetic") = "Nothing"

            Dim nextSlide As SlideTemplates.Layouts = New SlideTemplates.Layouts(contentJObject, layoutData)
            nextSlide.Render(newSlide)
        End If

    End Sub

    Public Function CombineDictionaries(Of TKey, TValue)(dict1 As Dictionary(Of TKey, TValue), dict2 As Dictionary(Of TKey, TValue)) As Dictionary(Of TKey, TValue)
        ' Create a new dictionary to store the result
        Dim result As New Dictionary(Of TKey, TValue)()

        ' Add all elements from dict2
        For Each kvp In dict2
            result(kvp.Key) = kvp.Value
        Next

        ' Add all elements from dict1 (overwriting dict2 values if keys overlap)
        For Each kvp In dict1
            result(kvp.Key) = kvp.Value
        Next

        Return result
    End Function

    Public Function SaveAndSendSlidePreview(slide As Slide) As String
        ' Path to save the slide preview temporarily
        Dim tempFolder As String = Path.GetTempPath()
        Dim tempImagePath As String = Path.Combine(tempFolder, "SlidePreview.png")
        Dim JsonResponse As JObject

        ' PowerPoint Application

        ' Export the current slide as an image
        slide.Export(tempImagePath, "PNG", 1920, 1080)

        ' Send the image to the FastAPI server
        Dim apiUrl As String = "http://localhost:8000/upload-slide-preview"
        Using client As New HttpClient()
            Using formData As New MultipartFormDataContent()
                Dim fileContent As New ByteArrayContent(File.ReadAllBytes(tempImagePath))
                fileContent.Headers.ContentType = MediaTypeHeaderValue.Parse("image/png")
                formData.Add(fileContent, "file", "SlidePreview.png")

                ' Post the data
                Dim response As HttpResponseMessage = client.PostAsync(apiUrl, formData).Result
                If response.IsSuccessStatusCode Then
                    Debug.WriteLine("Slide preview sent successfully!")
                    Dim responseJSON As String = response.Content.ReadAsStringAsync().Result
                    Debug.WriteLine($"{response.Content.ReadAsStringAsync().Result}")
                    JsonResponse = JObject.Parse(responseJSON)
                Else
                    Throw New Exception($"Failed to send slide preview. Status Code : {response.StatusCode}")
                End If
            End Using
        End Using

        ' Clean up the temporary image
        If File.Exists(tempImagePath) Then
            File.Delete(tempImagePath)
        End If
        Return JsonResponse("file_id")
    End Function
End Class
