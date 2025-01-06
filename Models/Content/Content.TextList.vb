Namespace Data.Content
    Public Enum TextListTypes
        Ordered
        Unordered
        Symbols
    End Enum

    Public Class TextList
        Private _content As List(Of String)
        Private _type As TextListTypes
        Public Property Rerender As Boolean

        Public Property content As List(Of String)
            Get
                Return _content
            End Get
            Set(value As List(Of String))
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

        Public Sub New(type As TextListTypes, content As List(Of String))
        End Sub

        Public Sub Render()
            Rerender = False
        End Sub
    End Class

End Namespace
