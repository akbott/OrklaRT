Public Class PivotLayoutVariant
    Public user As Integer, report As Integer, variantSeq As Integer, layoutDefinition As String
    Public Sub New(ByVal userId As Integer, ByVal reportId As Integer, ByVal variantId As Integer)
        InitializeComponent()
        user = userId
        report = reportId
        variantSeq = variantId
        If variantId <> 0 Then
            Using entities = New DAL.SAPExlEntities()
                Dim pivotLayoutVariants = entities.PivotLayoutVariants.SingleOrDefault(Function(p) p.ID = variantId)
                txtName.Text = pivotLayoutVariants.VariantName
                txtDescription.Text = pivotLayoutVariants.VariantDescription
            End Using
            btnUpdate.Visible = True
            btnSave.Visible = False
        Else
            txtName.Text = String.Empty
            txtDescription.Text = String.Empty
            btnCreateNew.Visible = False
            btnUpdate.Visible = False
            btnSave.Visible = True
        End If
    End Sub
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Try
            OrklaRTBPL.PivotFacade.SavePivotLayoutVariant(report, user, txtName.Text, txtDescription.Text)
            OrklaRTBPL.PivotFacade.SavePivotLayout(report, user, OrklaRTBPL.PivotFacade.GetPivotLayoutVariantID(report, user, txtName.Text, txtDescription.Text), ReturnPivotLayout)
            OrklaRTBPL.PivotFacade.UpdateCurrentUserReportPivotLayoutVariant(gUserId, gReportID, OrklaRTBPL.PivotFacade.GetPivotLayoutVariantID(report, user, txtName.Text, txtDescription.Text))
            LoadPivotLayouts()
            Dim pivotLayoutID = OrklaRTBPL.PivotFacade.GetPivotLayoutVariantID(report, user, txtName.Text, txtDescription.Text)
            Using entities = New DAL.SAPExlEntities()
                Dim reportPivotLayout = entities.PivotLayoutVariants.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = gUserId And rp.ID = pivotLayoutID)
                For Each pivotLayoitItem In Globals.Ribbons.OrklaRT.ddlPivotLayout.Items
                    If pivotLayoitItem.Tag = pivotLayoutID Then
                        Globals.Ribbons.OrklaRT.ddlPivotLayout.SelectedItem = pivotLayoitItem
                    End If
                Next
            End Using
            btnUpdate.Visible = True
            Me.Close()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.Name, gUserId, gReportID)
        End Try

    End Sub

    Private Sub btnUpdate_Click(sender As Object, e As EventArgs) Handles btnUpdate.Click
        OrklaRTBPL.PivotFacade.SavePivotLayout(report, user, variantSeq, ReturnPivotLayout)
        Me.Close()
    End Sub

    Private Sub btnCreateNew_Click(sender As Object, e As EventArgs) Handles btnCreateNew.Click
        txtName.Text = String.Empty
        txtDescription.Text = String.Empty
        btnUpdate.Visible = False
        btnSave.Visible = True
    End Sub
    Private Sub LoadPivotLayouts()
        Using entities = New DAL.SAPExlEntities()
            Globals.Ribbons.OrklaRT.ddlPivotLayout.Items.Clear()
            Dim newLayout As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
            newLayout.Tag = 0
            newLayout.Label = "Standard"
            Globals.Ribbons.OrklaRT.ddlPivotLayout.Items.Add(newLayout)
            Dim reportPivotLayouts = entities.PivotLayoutVariants.Where(Function(rp) rp.ReportID = gReportID And rp.UserID = gUserId)
            For Each reportPivotLayout In reportPivotLayouts
                Dim rdi As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                rdi.Tag = reportPivotLayout.ID
                rdi.Label = reportPivotLayout.VariantName
                Globals.Ribbons.OrklaRT.ddlPivotLayout.Items.Add(rdi)
            Next
        End Using
    End Sub
End Class