Public Class GetDescriptionPrompt
    Public Property UserInput As String
    Private Sub btnSubmitDesc_Click(sender As Object, e As EventArgs) Handles btnSubmitDesc.Click
        UserInput = txtBxLayoutDescription.Text

        If UserInput = "" Then
            Me.DialogResult = DialogResult.Abort
            Me.Close()
            Throw New Exception("Empty Description Detection")
        End If

        Me.DialogResult = DialogResult.OK
        Me.Close()
    End Sub
End Class