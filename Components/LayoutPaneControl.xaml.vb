Imports System.Collections.ObjectModel
Imports System.Diagnostics
Imports AutoSlider.Layouts

Public Class LayoutPaneControl

    Private Property LayoutList As New ObservableCollection(Of LayoutSnapControl)
    Private _wv2Layouts As List(Of LayoutSnapControl)
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.


    End Sub

    Public Sub GenerateLayouts(LayoutIds As List(Of String), Data As String)
        Dim stackPanel = stkLayoutPanel
        stackPanel.Children.Clear()

        LayoutList.Clear()
        For Each layoutId As String In LayoutIds
            LayoutList.Add(New LayoutSnapControl(Data, layoutId))
            Debug.WriteLine($"Created a New LayoutList with {layoutId} {Data}")
        Next

        For Each layoutSnapControl As LayoutSnapControl In LayoutList
            layoutSnapControl.Width = 400
            layoutSnapControl.Height = layoutSnapControl.Width * 9 / 16 * 0.7
            stackPanel.Children.Add(layoutSnapControl)
        Next
    End Sub
End Class
