<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class GetDescriptionPrompt
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
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
        Me.txtBxLayoutDescription = New System.Windows.Forms.TextBox()
        Me.btnSubmitDesc = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'txtBxLayoutDescription
        '
        Me.txtBxLayoutDescription.Location = New System.Drawing.Point(17, 18)
        Me.txtBxLayoutDescription.Multiline = True
        Me.txtBxLayoutDescription.Name = "txtBxLayoutDescription"
        Me.txtBxLayoutDescription.Size = New System.Drawing.Size(341, 106)
        Me.txtBxLayoutDescription.TabIndex = 0
        '
        'btnSubmitDesc
        '
        Me.btnSubmitDesc.Location = New System.Drawing.Point(283, 130)
        Me.btnSubmitDesc.Name = "btnSubmitDesc"
        Me.btnSubmitDesc.Size = New System.Drawing.Size(75, 23)
        Me.btnSubmitDesc.TabIndex = 1
        Me.btnSubmitDesc.Text = "OK"
        Me.btnSubmitDesc.UseVisualStyleBackColor = True
        '
        'GetDescription
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(370, 170)
        Me.Controls.Add(Me.btnSubmitDesc)
        Me.Controls.Add(Me.txtBxLayoutDescription)
        Me.Name = "GetDescription"
        Me.Text = "Form1"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtBxLayoutDescription As Windows.Forms.TextBox
    Friend WithEvents btnSubmitDesc As Windows.Forms.Button
End Class
