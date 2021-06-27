Module AllReports
    Public Sub StockDeviationShowDeviation()
        Dim pf As Excel.PivotField

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        On Error Resume Next
        For Each pf In Globals.ThisAddIn.Application.ActiveSheet.PivotTables(1).DataFields
            If pf.SourceName = "Unres" Or pf.SourceName = "Blocked" Then
                pf.Calculation = Excel.XlPivotFieldCalculation.xlDifferenceFrom
                pf.BaseField = "Sloc"
                pf.BaseItem = "(forrige)"
            End If
        Next
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End Sub


    Public Sub StockDeviationShowTotal()
        Dim pf As Excel.PivotField
        Const xlnormal = 1

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        For Each pf In Globals.ThisAddIn.Application.ActiveSheet.PivotTables(1).DataFields
            If pf.SourceName = "Unres" Or pf.SourceName = "Blocked" Then
                pf.Calculation = xlnormal
            End If
        Next
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End Sub


    Sub StockShelfLifeLocalUpdate()
        Dim c As Excel.Range
        Dim d As Excel.Range
        Dim sFirstAddress As String
        Dim lOrderRest As Long
        Dim lBatchRest As Long
        Dim lMaterial As Long
        Dim lNextRow As Long

        Application.Sheets("BatchRest").Cells(1, 1).CurrentRegion.Offset(1, 0).ClearContents()

        lMaterial = 0
        For Each c In Application.Range("OrklaRTData").Columns(8).Cells
            If c.Row = Application.Range("OrklaRTData").Cells(1, 1).Row Then GoTo ResumeHere
            If c.Offset(0, 18).Value <> 1 Then GoTo ResumeHere
            If c.Offset(0, 22).Value <> "0 - 30 days" Then GoTo ResumeHere
            If c.Offset(0, 23).Value = "Blocked" Then GoTo ResumeHere

            If c.Value <> lMaterial Then
                lMaterial = c.Value
                lOrderRest = 0
            End If

            'Gjenværende beholdning for batch
            lBatchRest = c.Offset(0, 20).Value
            d = Application.Sheets("Orders").Range("OrdersNotBilledNew").Columns(7).Find(c.Value, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            If Not IsNothing(d) And Not String.IsNullOrWhiteSpace(d.Address) Then
                sFirstAddress = d.Address
                Do    'Ordredato                 Batchdato
                    lOrderRest = d.Offset(0, 1).Value
                    If CDate(d.Offset(0, -5).Value).ToOADate() <= c.Offset(0, 17).Value And lOrderRest > 0 Then
                        If lOrderRest >= lBatchRest Then
                            lOrderRest = lOrderRest - lBatchRest
                            lBatchRest = 0
                            If lOrderRest >= 0 Then
                                d.Offset(0, 1).Value = lOrderRest
                            Else
                                d.Offset(0, 1).Value = 0
                                lOrderRest = 0
                            End If
                            Exit Do
                        Else
                            lBatchRest = lBatchRest - d.Offset(0, 1).Value
                            lOrderRest = 0
                            d.Offset(0, 1).Value = lOrderRest
                        End If
                    End If
                    d = Application.Sheets("Orders").Range("OrdersNotBilledNew").Columns(7).FindNext(d)
                Loop While Not d Is Nothing And d.Address <> sFirstAddress
                lNextRow = Application.Sheets("BatchRest").Cells(1, 1).CurrentRegion.Rows.Count
                Application.Sheets("BatchRest").Cells(lNextRow + 1, 1).Value = c.Offset(0, -7).Value
                Application.Sheets("BatchRest").Cells(lNextRow + 1, 2).Value = lBatchRest
            End If
ResumeHere:
        Next c
        Application.Sheets("Details").PivotTables(1).PivotCache.Refresh()

    End Sub

    Sub StockValuesAndCoverageRefreshProdData()

        Dim resultTable As New System.Data.DataTable

        Application.Application.DisplayAlerts = False

        resultTable = OrklaRTBPL.ReportSpecific.GetStockValuesAndCoverageProdPlanData(OrklaRTBPL.SelectionFacade.StockValuesAndCoverageProdPlanSelectionPlant, DateTime.Now.Date.ToString()).Tables("ProductionPlanData")

        Call Common.LoadListObjectData("ProductionPlanData", "ProdPlan", "tProdPlanAll", resultTable)

        Application.Application.DisplayAlerts = True

    End Sub
End Module
