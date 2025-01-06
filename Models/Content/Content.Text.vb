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
        Public Property Rerender As Boolean

        Public Property content As String
            Get
                Return _content
            End Get
            Set(value As String)
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

        Public Sub New(type As TextTypes, content As String)
        End Sub

        Public Sub Render()
            Rerender = False
        End Sub
    End Class

End Namespace
