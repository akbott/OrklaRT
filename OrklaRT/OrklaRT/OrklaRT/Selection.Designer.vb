<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Selection
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
        Me.selectionElementHost = New System.Windows.Forms.Integration.ElementHost()
        Me.SuspendLayout()
        '
        'selectionElementHost
        '
        Me.selectionElementHost.Location = New System.Drawing.Point(-1, -1)
        Me.selectionElementHost.Name = "selectionElementHost"
        Me.selectionElementHost.Size = New System.Drawing.Size(455, 600)
        Me.selectionElementHost.TabIndex = 0
        Me.selectionElementHost.Text = "Selection"
        Me.selectionElementHost.Child = Nothing
        '
        'Selection
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.selectionElementHost)
        Me.Name = "Selection"
        Me.Size = New System.Drawing.Size(455, 600)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents selectionElementHost As System.Windows.Forms.Integration.ElementHost

End Class
