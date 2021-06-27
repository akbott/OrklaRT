Module StockTransfer
   
    Sub LocalUpdate()
        Dim C As Excel.Range
        Dim x As Integer


        Call RefreshTransportGroups()
        Call RefreshBinTest()
        Call RefreshExcludedTypes()

CleanUp:
        Exit Sub
    End Sub

    Sub RefreshTransportGroups()
        Dim resultTable As New System.Data.DataTable

        Application.Application.DisplayAlerts = False


        resultTable = OrklaRTBPL.ReportSpecific.GetSTTransportGroups(OrklaRTBPL.SelectionFacade.StockTransferSelectionWarehouse).Tables("STTransportGroups")

        If resultTable.Rows.Count > 0 Then
            Dim rs = OrklaRTBPL.CommonFacade.ConvertToRecordset(resultTable)

            Application.Sheets("Grupper").Range("tTransportGroups").Clear()
            rs.MoveFirst()
            Application.Sheets("Grupper").Range("tTransportGroups").CopyFromRecordset(rs)
        End If

        Application.Application.DisplayAlerts = True

    End Sub
    Sub RefreshBinTest()
        Dim resultTable As New System.Data.DataTable

        Application.Application.DisplayAlerts = False


        resultTable = OrklaRTBPL.ReportSpecific.GetSTBinTest(OrklaRTBPL.SelectionFacade.StockTransferSelectionWarehouse).Tables("STBinTest")

        If resultTable.Rows.Count > 0 Then
            Dim rs = OrklaRTBPL.CommonFacade.ConvertToRecordset(resultTable)

            Application.Sheets("Grupper").Range("tBinTest").Clear()
            rs.MoveFirst()
            Application.Sheets("Grupper").Range("tBinTest").CopyFromRecordset(rs)
        End If


        Application.Application.DisplayAlerts = True


    End Sub
    Sub RefreshExcludedTypes()
        Dim resultTable As New System.Data.DataTable

        Application.Application.DisplayAlerts = False


        resultTable = OrklaRTBPL.ReportSpecific.GetSTExcludedTypes(OrklaRTBPL.SelectionFacade.StockTransferSelectionWarehouse).Tables("STExcludedTypes")

        If resultTable.Rows.Count > 0 Then
            Dim rs = OrklaRTBPL.CommonFacade.ConvertToRecordset(resultTable)

            Application.Sheets("Grupper").Range("tExclTypes").Clear()
            rs.MoveFirst()
            Application.Sheets("Grupper").Range("tExclTypes").CopyFromRecordset(rs)
        End If


        Application.Application.DisplayAlerts = True


    End Sub

    Sub WriteGroups()

        Dim x As Integer
        Dim r As Excel.Range

        r = Application.Sheets("Groups").Range("tTransportGroups").Cells(1, 1)
        For x = 0 To 1000
            If r.Offset(x, 1).Value = "" Then Exit For
            OrklaRTBPL.ReportSpecific.InsertSTTransportGroups(OrklaRTBPL.SelectionFacade.StockTransferSelectionWarehouse, String.Format("0:000", r.Offset(x, 2).Value) & "-" & String.Format("0:000", r.Offset(x, 3).Value) & "-" & r.Offset(x, 4).Value, r.Offset(x, 1).Value, r.Offset(x, 2).Value, r.Offset(x, 3).Value, r.Offset(x, 4).Value, r.Offset(x, 5).Value, r.Offset(x, 6).Value)
        Next x

        Call RefreshTransportGroups()

CleanUp:
        Exit Sub

    End Sub


    Sub WriteBinTest()
     
        Dim x As Integer
        Dim r As Excel.Range

        r = Application.Sheets("Groups").Range("tBinTest").Cells(1, 1)
        For x = 0 To 1000
            If r.Offset(x, 0).Value = "" Then Exit For
            OrklaRTBPL.ReportSpecific.InsertSTBinTest(OrklaRTBPL.SelectionFacade.StockTransferSelectionWarehouse, r.Offset(x, 0).Value, r.Offset(x, 1).Value)
        Next x

        Call RefreshBinTest()

CleanUp:
        Exit Sub

    End Sub

    Sub WriteExclTypes()

        Dim x As Integer
        Dim r As Excel.Range

        r = Application.Sheets("Groups").Range("tExclTypes").Cells(1, 1)
        For x = 0 To 1000
            If r.Offset(x, 0).Value = "" Then Exit For
            OrklaRTBPL.ReportSpecific.InsertSTExcludedTypes(OrklaRTBPL.SelectionFacade.StockTransferSelectionWarehouse, r.Offset(x, 0).Value)
        Next x
       
        Call RefreshExcludedTypes()

CleanUp:
        Exit Sub

    End Sub

End Module
