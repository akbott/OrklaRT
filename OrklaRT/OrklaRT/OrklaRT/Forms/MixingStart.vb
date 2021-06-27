Public Class MixingStart
    Public order As Long
    Public Sub New(ByVal order As Long, ByVal material As String, mixingDate As Date, time As String)
        InitializeComponent()
        dtpDate.Value = mixingDate.Date
        lblMaterialName.Text = material
        txtStartTime.Text = time
        order = order        
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click        
        Me.Close()
    End Sub


    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        strEditComm = btnCancel.Text
        Me.Close()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        strEditComm = btnDelete.Text
        Me.Close()
    End Sub
End Class