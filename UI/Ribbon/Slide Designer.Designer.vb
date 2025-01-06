﻿Partial Class Slide_Designer
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
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

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Slide_Designer))
        Me.tbSlideGenerator = Me.Factory.CreateRibbonTab
        Me.Group1 = Me.Factory.CreateRibbonGroup
        Me.btnGenerate = Me.Factory.CreateRibbonButton
        Me.tbSlideGenerator.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.SuspendLayout()
        '
        'tbSlideGenerator
        '
        Me.tbSlideGenerator.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office
        Me.tbSlideGenerator.Groups.Add(Me.Group1)
        Me.tbSlideGenerator.Label = "Slide Generator"
        Me.tbSlideGenerator.Name = "tbSlideGenerator"
        '
        'Group1
        '
        Me.Group1.Items.Add(Me.btnGenerate)
        Me.Group1.Label = "Group1"
        Me.Group1.Name = "Group1"
        '
        'btnGenerate
        '
        Me.btnGenerate.Image = CType(resources.GetObject("btnGenerate.Image"), System.Drawing.Image)
        Me.btnGenerate.Label = "Generate"
        Me.btnGenerate.Name = "btnGenerate"
        Me.btnGenerate.ShowImage = True
        '
        'Slide_Designer
        '
        Me.Name = "Slide_Designer"
        Me.RibbonType = "Microsoft.PowerPoint.Presentation"
        Me.Tabs.Add(Me.tbSlideGenerator)
        Me.tbSlideGenerator.ResumeLayout(False)
        Me.tbSlideGenerator.PerformLayout()
        Me.Group1.ResumeLayout(False)
        Me.Group1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tbSlideGenerator As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents Group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnGenerate As Microsoft.Office.Tools.Ribbon.RibbonButton
End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property Slide_Designer() As Slide_Designer
        Get
            Return Me.GetRibbon(Of Slide_Designer)()
        End Get
    End Property
End Class
