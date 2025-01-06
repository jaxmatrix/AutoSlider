Imports System.ComponentModel

Namespace Layouts
    Public Class LayoutData
        Implements INotifyPropertyChanged


        Private _data As String
        Private _layout As String

        Public Property Layout As String
            Get
                Return _layout
            End Get
            Set(value As String)
                _layout = value
                OnPropertyChanged(NameOf(Layout))
            End Set
        End Property
        Public Property Data As String
            Get
                Return _data
            End Get
            Set(value As String)
                _data = value
                OnPropertyChanged(NameOf(Data))
            End Set
        End Property

        Public Sub New(newData As String, newLayout As String)
            Layout = newLayout
            Data = newData
        End Sub

        Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

        Protected Sub OnPropertyChanged(propertyName As String)
            RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(propertyName))
        End Sub

    End Class

End Namespace
