Imports Microsoft.Office.Tools

Public Class ThisAddIn
    Private layoutPane As CustomTaskPane


    Private Sub ThisAddIn_Startup() Handles Me.Startup
        Dim layoutPaneControl As New LayoutPaneHostControl()

        layoutPane = Me.CustomTaskPanes.Add(layoutPaneControl, "Layout Action Pane")

        layoutPane.Visible = True
        layoutPane.Width = 300

        Dim LayoutIds As New List(Of String) From {
            "1",
            "2",
            "3",
            "4"
        }

        Dim Data As String = "Random Data"

        layoutPaneControl.GenerateLayout(LayoutIds, Data)

    End Sub

    Private Sub ThisAddIn_Shutdown() Handles Me.Shutdown

    End Sub

End Class
