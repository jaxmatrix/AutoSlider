Imports Microsoft.Office.Interop.PowerPoint
Imports Microsoft.Office.Core

Namespace SlideTemplates
    ''' <summary>
    ''' Represents a test layout in the presentation. This slide will hold a title, description about the 
    ''' title and a list of information. All this data will be used to descripbe a image that will be seen at 2:4 of
    ''' of the width of the space
    ''' </summary>
    Public Class Test1

        Private _title As Data.Content.Text
        Private _description As Data.Content.Text
        Private _list As Data.Content.TextList
        Private _imageSrc As String
        Private _slide As Slide

        Public Sub New(NewSlide As Slide, title As String, desc As String, points As List(Of String))
            _slide = NewSlide
            _title = New Data.Content.Text(Data.Content.TextTypes.Header, title)
            _description = New Data.Content.Text(Data.Content.TextTypes.Text, desc)
            _list = New Data.Content.TextList(Data.Content.TextListTypes.Ordered, points)
        End Sub

        Public Sub Render()
            _title.Render(_slide, MsoTextOrientation.msoTextOrientationHorizontal, 10, 10, 200, 300)
            _description.Render(_slide, MsoTextOrientation.msoTextOrientationHorizontal, 10, 40, 200, 200)
            _list.Render(_slide, MsoTextOrientation.msoTextOrientationHorizontal, 10, 50, 300, 400)
        End Sub
    End Class
End Namespace

