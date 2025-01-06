<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class LayoutPaneHostControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.hostLayoutPane = New System.Windows.Forms.Integration.ElementHost()
        Me.LayoutPaneControl1 = New AutoSlider.LayoutPaneControl()
        Me.SuspendLayout()
        '
        'hostLayoutPane
        '
        Me.hostLayoutPane.Dock = System.Windows.Forms.DockStyle.Fill
        Me.hostLayoutPane.Location = New System.Drawing.Point(0, 0)
        Me.hostLayoutPane.Name = "hostLayoutPane"
        Me.hostLayoutPane.Size = New System.Drawing.Size(155, 391)
        Me.hostLayoutPane.TabIndex = 0
        Me.hostLayoutPane.Text = "ElementHost1"
        Me.hostLayoutPane.Child = Me.LayoutPaneControl1
        '
        'LayoutPaneHostControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.hostLayoutPane)
        Me.Name = "LayoutPaneHostControl"
        Me.Size = New System.Drawing.Size(155, 391)
        Me.ResumeLayout(False)

    End Sub

    Public Sub GenerateLayout(LayoutIds As List(Of String), Data As String)
        Me.LayoutPaneControl1.GenerateLayouts(LayoutIds, Data)
    End Sub

    Friend WithEvents hostLayoutPane As Windows.Forms.Integration.ElementHost
    Friend LayoutPaneControl1 As LayoutPaneControl
End Class
