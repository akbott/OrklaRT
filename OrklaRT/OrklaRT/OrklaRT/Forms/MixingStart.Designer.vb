<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class MixingStart
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MixingStart))
        Me.lblMixingStartDate = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnSave = New System.Windows.Forms.Button()
        Me.dtpDate = New System.Windows.Forms.DateTimePicker()
        Me.lblNewMixingStartTime = New System.Windows.Forms.Label()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.txtStartTime = New System.Windows.Forms.TextBox()
        Me.lblMaterialName = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lblMixingStartDate
        '
        resources.ApplyResources(Me.lblMixingStartDate, "lblMixingStartDate")
        Me.lblMixingStartDate.Name = "lblMixingStartDate"
        '
        'btnCancel
        '
        resources.ApplyResources(Me.btnCancel, "btnCancel")
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnSave
        '
        resources.ApplyResources(Me.btnSave, "btnSave")
        Me.btnSave.Name = "btnSave"
        Me.btnSave.UseVisualStyleBackColor = True
        '
        'dtpDate
        '
        resources.ApplyResources(Me.dtpDate, "dtpDate")
        Me.dtpDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpDate.MinDate = New Date(1900, 1, 1, 0, 0, 0, 0)
        Me.dtpDate.Name = "dtpDate"
        Me.dtpDate.Value = New Date(2015, 4, 12, 19, 19, 32, 0)
        '
        'lblNewMixingStartTime
        '
        resources.ApplyResources(Me.lblNewMixingStartTime, "lblNewMixingStartTime")
        Me.lblNewMixingStartTime.Name = "lblNewMixingStartTime"
        '
        'btnDelete
        '
        resources.ApplyResources(Me.btnDelete, "btnDelete")
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.UseVisualStyleBackColor = True
        '
        'txtStartTime
        '
        resources.ApplyResources(Me.txtStartTime, "txtStartTime")
        Me.txtStartTime.Name = "txtStartTime"
        '
        'lblMaterialName
        '
        resources.ApplyResources(Me.lblMaterialName, "lblMaterialName")
        Me.lblMaterialName.Name = "lblMaterialName"
        '
        'MixingStart
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lblMaterialName)
        Me.Controls.Add(Me.txtStartTime)
        Me.Controls.Add(Me.btnDelete)
        Me.Controls.Add(Me.lblNewMixingStartTime)
        Me.Controls.Add(Me.dtpDate)
        Me.Controls.Add(Me.lblMixingStartDate)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnSave)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "MixingStart"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.TopMost = True
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblMixingStartDate As System.Windows.Forms.Label
    Friend WithEvents btnCancel As System.Windows.Forms.Button
    Friend WithEvents btnSave As System.Windows.Forms.Button
    Friend WithEvents dtpDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblNewMixingStartTime As System.Windows.Forms.Label
    Friend WithEvents btnDelete As System.Windows.Forms.Button
    Friend WithEvents txtStartTime As System.Windows.Forms.TextBox
    Friend WithEvents lblMaterialName As System.Windows.Forms.Label
End Class
