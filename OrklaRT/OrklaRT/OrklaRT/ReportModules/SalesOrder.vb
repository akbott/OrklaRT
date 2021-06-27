Module SalesOrder
    Sub LocalUpdate()

        Dim pi As Excel.PivotItem

        Application.Sheets("MankoKunder").PivotTables(1).PivotCache.Refresh()
        Try
            Application.Sheets("MankoKunder").PivotTables(1).PivotFields("Ønsk.levDa").CurrentPage = CDate(DateTime.Now.ToShortDateString())
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Sales Order Status", gUserId, gReportID)
        End Try
        Application.Sheets("Verdi").PivotTables(1).PivotCache.Refresh()
    End Sub
    'Dim pi As Excel.PivotItem
    'Dim pi1 As Excel.PivotItem
    'Dim Pvt As Excel.PivotTable
    'Dim pf As Excel.PivotField
    'Dim c As Excel.Range
    'Dim i As String

    'Try

    '    Call SalesOrderRefreshProdData()
    '    Application.Calculation = Excel.XlCalculation.xlCalculationManual
    '    'Application.Sheets("Value").PivotTables(1).PivotCache.Refresh()

    '    Select Case OrklaRTBPL.SelectionFacade.SalesOrderSelectionSalesOrg
    '        Case "NOA"
    '            Application.Sheets("Value").PivotTables(1).CalculatedFields("Deliv_SL"). _
    '               StandardFormula = "=(Deliv_Qty/Order_Qty)"
    '            Application.Sheets("Value").PivotTables(1).CalculatedFields("Conf_SL"). _
    '               StandardFormula = "=(Conf_Qty/Order_Qty)"

    '            With Application.Sheets("Descriptions").Range("Descriptions")
    '                c = .Columns(1).Find("Conf_SL", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '                If Not c Is Nothing Then
    '                    c.Offset(0, 1).Value = "Confirmed Service Level - calculation based on sales units"
    '                End If
    '                c = .Columns(1).Find("Deliv_SL", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '                If Not c Is Nothing Then
    '                    c.Offset(0, 1).Value = "Delivered Service Level - calculation based on sales units"
    '                End If
    '            End With

    '        Case Else
    '            Application.Sheets("Value").PivotTables(1).CalculatedFields("Deliv_SL"). _
    '               StandardFormula = "=(Deliv_Value/Order_Value)"
    '            Application.Sheets("Value").PivotTables(1).CalculatedFields("Conf_SL"). _
    '               StandardFormula = "=(Conf_Value/Order_Value)"

    '            With Application.Sheets("Descriptions").Range("Descriptions")
    '                c = .Columns(1).Find("Conf_SL", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '                If Not c Is Nothing Then
    '                    c.Offset(0, 1).Value = "Confirmed Service Level - calculation based on values"
    '                End If
    '                c = .Columns(1).Find("Deliv_SL", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '                If Not c Is Nothing Then
    '                    c.Offset(0, 1).Value = "Delivered Service Level - calculation based on values"
    '                End If
    '            End With
    '    End Select

    '    Application.Calculate()
    '    Application.Sheets("Value").PivotTables(1).PivotCache.Refresh()

    '    Application.Sheets("Undelivered_List").Activate()
    '    Application.Sheets("Undelivered_List").PivotTables(1).RowFields("Comments").LabelRange.ColumnWidth = 35
    '    Application.ActiveSheet.Shapes("OrderValue").Select()
    '    Application.Selection.Delete()
    '    Application.Sheets("Groups").Range("OrderValues").Copy()
    '    Application.Sheets("Undelivered_List").Range("E9").Activate()
    '    Application.ActiveSheet.Pictures.Paste(Link:=True).Select()
    '    Application.Selection.Name = "OrderValue"
    '    '   Selection.ShapeRange.IncrementLeft -33.75
    '    Application.CutCopyMode = False
    '    Application.Range("C17").Select()

    'Catch ex As Exception
    '    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "SalesOrder", gUserId, gReportID)
    'End Try

    'Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic

    'c = Nothing
    'pi = Nothing
    'pi1 = Nothing
    'Pvt = Nothing
    'pf = Nothing

    Sub SalesOrderRefreshProdData()

        Dim resultTable As New System.Data.DataTable

        Application.Application.DisplayAlerts = False


        resultTable = OrklaRTBPL.ReportSpecific.GetStockValuesAndCoverageProdPlanData(OrklaRTBPL.SelectionFacade.SalesOrderSelectionSalesOrg, DateTime.Now.Date.ToString()).Tables("ProductionPlanData")

        Dim rs = OrklaRTBPL.CommonFacade.ConvertToRecordset(resultTable)

        If Not rs.EOF Then
            Application.Sheets("ProdPlan").Range("tProdPlanAll").Clear()
            rs.MoveFirst()
            Application.Sheets("ProdPlan").Range("tProdPlanAll").CopyFromRecordset(rs)
        End If

        Application.Application.DisplayAlerts = True

    End Sub

    Sub MakeSheetCopyWeek()

        Dim w As Excel.Workbook
        Dim n As Integer
        Dim x As Integer
        Dim y As Integer
        Dim dblSG As Double
        Dim c As Excel.Range
        Dim xCounter As Integer

        On Error GoTo CleanUp

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual


        Dim ThisWorkbook = Globals.Factory.GetVstoObject(Application.ActiveWorkbook)

        x = 0
        n = Application.ActiveSheet.PivotTables(1).TableRange1.Columns.Count
        For Each c In Application.ActiveSheet.PivotTables(1).TableRange1.Columns(n).Cells
            If c.Value.ToString() = "1" Then Exit For
            x = x + 1
        Next c
        If x = 0 Then GoTo CleanUp

        y = Application.ActiveSheet.PivotTables(1).TableRange1.Cells(1, 1).Row + 1
        Application.Range("E1").FormulaR1C1 = "=GETPIVOTDATA(""Lever_SL"",R12C1)"
        Application.Range(Application.Cells(y, 1), Application.Cells(y + x - 2, n)).Copy()

        w = Application.Workbooks.Add()
        Application.Cells(5, 2).Select()
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.Constants.xlNone, SkipBlanks _
            :=False, Transpose:=False)
        Application.CutCopyMode = False
        Application.ActiveWindow.DisplayGridlines = False

        Application.Range("A1:B2").Interior.Color = 12900829
        Application.Range("A3:B3").Interior.Color = 255
        Application.Range("A4:E4").Interior.Color = 12900829
        Application.Range("A4:E4").Font.Bold = True
        Application.Range("A1:A3").Font.Bold = True
        Application.Range("A1").Value = "Uke:"
        Application.Range("A2").Value = "Leverandør:"
        Application.Range("B2").Value = "Orkla Foods Norge"
        Application.Range("A3").Value = "Total servicegrad:"
        Application.Range("B3").Value = ThisWorkbook.Sheets("MankoUke").Range("E1").Value
        Application.Range("B1").NumberFormat = "@"
        Application.Range("B3").NumberFormat = "0.0 %"
        Application.Range("C5:C100").NumberFormat = "0.0 %"
        Application.Range("A4").Value = "EPD"
        Application.Range("B4").Value = "Varetekst"
        Application.Range("C4").Value = "Servicegrad i %"
        Application.Range("D4").Value = "Leveringsdyktig dag"
        Application.Range("E4").Value = "Kommentar på produkter man ikke leverer"
        Application.Cells(1, 1).ColumnWidth = 20
        Application.Cells(1, 2).ColumnWidth = 50
        Application.Cells(1, 3).ColumnWidth = 15
        Application.Cells(1, 4).ColumnWidth = 20
        Application.Cells(1, 5).ColumnWidth = 40

        xCounter = 0
        For Each c In Application.Range("B5:B400")
            If c.Value.ToString() = "0" Then Exit For
            xCounter = xCounter + 1
            Call CLAF_CLASSIFICATION_Week(c, c.Value.ToString(), xCounter)
        Next
        Application.Range("B1").Select()

        Application.StatusBar = "Definere utskrifts innstillingene ..."

        '   Define printing settings
        With Application.ActiveSheet.PageSetup
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            '       .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlPortrait
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
        End With

CleanUp:
        w = Nothing
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End Sub
    Sub hentfarge()
        Application.GenerateGetPivotData = True
    End Sub
    Sub MakeSheetCopyLocal()

        Dim w As Excel.Workbook
        Dim n As Integer

        On Error GoTo CleanUp

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        Dim ThisWorkbook = Globals.Factory.GetVstoObject(Application.ActiveWorkbook)

        Application.StatusBar = "Definere utskrifts innstillingene ..."
        '   n = ActiveSheet.PivotTables(1).TableRange1.Cells(1, 1).Row
        '   n = n + ActiveSheet.PivotTables(1).TableRange1.Rows.Count
        '   ActiveSheet.Range(Cells(1, 1), Cells(n, ActiveSheet.PivotTables(1).TableRange1.Columns.Count)).Copy
        Application.Cells.Copy()
        w = Application.Workbooks.Add()
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.Constants.xlNone, SkipBlanks _
            :=False, Transpose:=False)
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteFormats, Operation:=Excel.Constants.xlNone, _
            SkipBlanks:=False, Transpose:=False)
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteColumnWidths, Operation:=Excel.Constants.xlNone, _
            SkipBlanks:=False, Transpose:=False)
        Application.CutCopyMode = False
        'ThisWorkbook.ActiveSheet.Shapes("Orkla").Copy()
        w.Activate()
        Application.Range("A1").Select()
        Application.ActiveSheet.Paste()
        Application.Range("A1").Select()
        Application.Selection.ColumnWidth = 10.5
        ThisWorkbook.ActiveSheet.Shapes("Avrundet rektangel 1").Copy()
        w.Activate()
        Application.Range("A4").Select()
        Application.ActiveSheet.Paste()
        Application.Range("A4").Select()

        Application.ActiveWindow.DisplayGridlines = False
        Application.ActiveWindow.Zoom = 85
        '   Range("D:D").Delete
        Application.Range("E20").Value = "EPD-nr"
        Application.Range("F20").Value = "NKL-nr"

        Call Loop_EPD()

        '   Formatting EPD-nr
        Application.Range("C20:C200").Select()
        Application.Selection.Copy()
        Application.Range("E20:F200").Select()
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteFormats, Operation:=Excel.Constants.xlNone, _
            SkipBlanks:=False, Transpose:=False)
        Application.CutCopyMode = False
        Application.Columns("E:F").ColumnWidth = 9
        Application.Columns("D").ColumnWidth = 10
        Application.Range("D20:F20").Select()
        With Application.Selection
            .HorizontalAlignment = Excel.Constants.xlRight
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
        End With
        Application.Columns("G:J").Delete()
        Application.Range("C15").Select()

        '   Define printing settings
        With Application.ActiveSheet.PageSetup
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = Excel.XlPrintLocation.xlPrintNoComments
            '       .PrintQuality = 600
            .CenterHorizontally = False
            .CenterVertically = False
            .Orientation = Excel.XlPageOrientation.xlPortrait
            .Draft = False
            .PaperSize = Excel.XlPaperSize.xlPaperA4
            .FirstPageNumber = Excel.Constants.xlAutomatic
            .Order = Excel.XlOrder.xlDownThenOver
            .BlackAndWhite = False
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .PrintErrors = Excel.XlPrintErrors.xlPrintErrorsDisplayed
        End With

CleanUp:
        w = Nothing
        Application.ActiveWindow.DisplayZeros = False

        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True

    End Sub

    Sub MakeSheetCopyManko()

        Dim w As Excel.Workbook
        Dim n As Integer

        On Error GoTo CleanUp

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        Dim ThisWorkbook = Globals.Factory.GetVstoObject(Application.ActiveWorkbook)

        On Error Resume Next
        Application.StatusBar = "Definere utskrifts innstillingene ..."
        n = Application.ActiveSheet.PivotTables(1).TableRange1.Cells(1, 1).Row
        n = n + Application.ActiveSheet.PivotTables(1).TableRange1.Rows.Count
        Application.ActiveSheet.Range(Application.Cells(1, 1), Application.Cells(n, Application.ActiveSheet.PivotTables(1).TableRange1.Columns.Count)).Copy()
        '   Cells.Copy
        w = Application.Workbooks.Add()
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.Constants.xlNone, SkipBlanks _
            :=False, Transpose:=False)
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteFormats, Operation:=Excel.Constants.xlNone, _
            SkipBlanks:=False, Transpose:=False)
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteColumnWidths, Operation:=Excel.Constants.xlNone, _
            SkipBlanks:=False, Transpose:=False)
        Application.CutCopyMode = False

        ThisWorkbook.ActiveSheet.Shapes("Avrundet rektangel 1").Copy()
        w.Activate()
        Application.Range("A1").Select()
        Application.ActiveSheet.Paste()
        Application.Range("A1").Select()

        ThisWorkbook.ActiveSheet.Shapes("OrderValue").Copy()
        w.Activate()
        Application.Range("E9").Select()
        Application.ActiveSheet.Paste()
        Application.Range("A1").Select()
        'Application.ActiveWorkbook.BreakLink(Name:="C:\Tmp\OrklaRT\Order_Status_Norway_M.xls", Type:=Excel.XlLinkType.xlLinkTypeExcelLinks)
        Application.ActiveSheet.PageSetup.PrintArea = Application.ActiveSheet.UsedRange.Address
        Application.ActiveSheet.PageSetup.PrintTitleRows = "$16:$16"

        With Application.ActiveSheet
            With .PageSetup
                .Orientation = Excel.XlPageOrientation.xlLandscape
                .LeftHeader = "Orkla RTR Rapporter"
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .LeftFooter = "Utskrift dato: &D"
                .CenterFooter = "Side &P / &N"
                .RightFooter = "&F - &A "
            End With
        End With

        Application.ActiveWindow.DisplayGridlines = False
        Application.ActiveWindow.Zoom = 85

CleanUp:
        w = Nothing
        Application.ActiveWindow.DisplayZeros = False

        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
    End Sub

    Sub Loop_EPD()
        Dim c As Excel.Range
        Dim xCounter As Integer

        Try
            xCounter = 0
            For Each c In Application.Range("C21:C200")
                If Not c.Value Is Nothing Then
                    If c.Value.ToString() = "0" Then Exit For
                    xCounter = xCounter + 1
                    Call CLAF_CLASSIFICATION_OF_OBJECTS_All(c, c.Value.ToString(), xCounter)
                End If
            Next
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "SalesOrder", gUserId, gReportID)
        End Try
    End Sub

    Sub CLAF_CLASSIFICATION_OF_OBJECTS_All(Target As Excel.Range, strMaterial As String, counter As Integer)
        Dim rfcTable As SAP.Middleware.Connector.IRfcTable
        Try
            Application.StatusBar = "Henter EPD-nr " & counter.ToString()

            rfcTable = BPL.RfcFunctions.GetCLAFCLASSIFICATIONOFOBJECTSAll(strMaterial.Split(" ").GetValue(0).ToString())

            Target.Offset(0, 2).Value = rfcTable(25).GetValue("AUSP1")
            Target.Offset(0, 3).Value = rfcTable(26).GetValue("AUSP1")
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "SalesOrder", gUserId, gReportID)
        End Try
    End Sub


    Sub CLAF_CLASSIFICATION_Week(Target As Excel.Range, strMaterial As String, counter As Integer)
        Dim rfcTable As SAP.Middleware.Connector.IRfcTable

        Try

            Application.StatusBar = "Henter EPD-nr " & counter.ToString()

            rfcTable = BPL.RfcFunctions.GetCLAFCLASSIFICATIONOFOBJECTSAll(strMaterial.Split(" ").GetValue(0).ToString())

            Target.Offset(0, -1).Value = rfcTable(25).GetValue("AUSP1")
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "SalesOrder", gUserId, gReportID)
        End Try

    End Sub

    'Sub BAPI_MATERIAL_AVAILABILITY(strMaterial As String, plant As String)
    '    Dim rfcTable As SAP.Middleware.Connector.IRfcTable

    '    rfcTable = BPL.RfcFunctions.GetBAPIMATERIALAVAILABILITY(strMaterial, plant)

    '    'strATP = rfcTable(0)("AV_QTY_PLT").ToString()
    '    'strATP1 = rfcTable(0)("Endleadtme").ToString()
    'End Sub
End Module
