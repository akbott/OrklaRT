Public Class PivotComments
    Public commentSeq As String
    Public Sub New(ByVal commentID As String)
        InitializeComponent()
        If commentID <> 0 Then
            commentSeq = commentID
            Using entities = New DAL.SAPExlEntities()
                If entities.ReportComments.Where(Function(p) p.CommentID = commentID And p.ModifiedBy = Environment.UserName).Count() > 0 Then
                    txtComment.Text = entities.ReportComments.SingleOrDefault(Function(p) p.CommentID = commentID And p.ModifiedBy = Environment.UserName).Comment
                    dtpDate.Value = entities.ReportComments.SingleOrDefault(Function(p) p.CommentID = commentID And p.ModifiedBy = Environment.UserName).Date1
                    btnUpdate.Visible = True
                Else
                    txtComment.Text = String.Empty
                    btnSave.Visible = True
                End If
            End Using
        End If
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        OrklaRTBPL.PivotFacade.SavePivotComment(gReportID, commentSeq, txtComment.Text, dtpDate.Value.Date, Environment.UserName)
        Me.Close()
        UpdatePivotComments()
    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        OrklaRTBPL.PivotFacade.SavePivotComment(gReportID, commentSeq, txtComment.Text, dtpDate.Value.Date, Environment.UserName)
        Me.Close()
        UpdatePivotComments()
    End Sub
    Public Sub UpdatePivotComments()
        Dim sFirstSheet As String
        Dim sFirstTable As String

        Application.EnableEvents = False
        Application.ScreenUpdating = False

        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Application.ActiveWorkbook.Sheets("Comments").ListObjects
            If listObject.Name.Equals("Comments") Then
                Try
                    Dim reportComments = OrklaRTBPL.CommonFacade.GetReportComments(gReportID)
                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(reportComments.Tables(0))
                    data.MoveFirst()
                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, reportComments.Tables(0).Rows.Count, reportComments.Tables(0).Columns.Count)
                Catch
                End Try
            End If
        Next
        
        sFirstSheet = Application.Sheets("Version").Range("FirstSheet").Value
        sFirstTable = Application.Sheets("Version").Range("FirstTable").Value
        
        Dim pvt As Excel.PivotTable = Application.Sheets(sFirstSheet).PivotTables(sFirstTable)
        pvt.PivotCache.Refresh()



        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End Sub
End Class