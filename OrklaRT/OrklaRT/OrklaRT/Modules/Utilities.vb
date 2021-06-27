Module Utilities
    Public Sub MakeSheetCopy()
        Dim w As Excel.Workbook
        Dim iOrientation As Integer
        Dim sPrintRows As String
        Dim wbOpen As Excel.Workbook

        If gwbReport Is Nothing Then
            Exit Sub
        End If

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        wbOpen = Application.ActiveWorkbook

        Application.StatusBar = "Defining Printing Settings..."
        If fnIsPivotTable() = True Then
            Application.ActiveSheet.PivotTables(1).TableRange1.Columns.AutoFit()
        End If

        If Application.ActiveSheet.UsedRange.Width > 700 Then
            iOrientation = Excel.XlPageOrientation.xlLandscape
        Else
            iOrientation = Excel.XlPageOrientation.xlPortrait
        End If

        If fnIsPivotTable() = True Then
            If Application.ActiveSheet.PivotTables(1).DataFields.Count > 1 Then
                sPrintRows = Application.ActiveSheet.PivotTables(1).ColumnRange.EntireRow.Address
            Else
                sPrintRows = Application.ActiveSheet.PivotTables(1).RowRange.Rows(1).EntireRow.Address
            End If
        Else
            sPrintRows = ""
        End If

        Application.Cells.Copy()
        w = Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
        w.Sheets(1).Name = "Kopi_Fra_Rapport"
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
        Application.Selection.PasteSpecial(Paste:=Excel.XlPasteType.xlPasteFormats, Operation:=Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, SkipBlanks:=False, Transpose:=False)
        Application.CutCopyMode = False
        w.Activate()
        Application.Calculate()
        Application.Range("A1").Select()
        Application.ActiveWindow.DisplayGridlines = False
        Application.ActiveWindow.Zoom = 85

        Try
            With Application.ActiveSheet
                With .PageSetup
                    .LeftHeaderPicture.Filename = Left(Right(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase, Len(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) - 8), Len(Right(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase, Len(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase) - 8)) - 11) + "Resources/OFN.jpg"
                    .LeftHeader = "&G"
                    .PrintTitleRows = sPrintRows
                    .Orientation = iOrientation
                    .RightHeader = "OrklaRT rapporter - strengt konfidensiell" & Chr(10) & "Kan ikke bli distribuert til hvem det ikke er bekymring"
                    .Zoom = False
                    .FitToPagesWide = 1
                    .FitToPagesTall = False
                    .LeftFooter = "Utskrift dato: &D"
                    .CenterFooter = "side &P / &N"
                    .RightFooter = "&F - &A "
                End With
            End With
        Catch ex As Exception

        End Try        

        w.Activate()
        Application.Range("A1").Select()
        Application.Calculate()

CleanUp:
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        w = Nothing
        wbOpen = Nothing
        Exit Sub
    End Sub


    Public Sub TransformPivotToTable()
        Dim w As Excel.Workbook
        Dim shNew As Excel.Worksheet
        Dim shPivot As Excel.Worksheet
        Dim shNewTemp As Excel.Worksheet
        Dim pf As Excel.PivotField
        Dim pvtRow As Excel.Range

        If fnIsPivotTable() = False Then
            Exit Sub
        End If

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        '  Define and fill variables.
        shPivot = Application.ActiveSheet
        w = Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
        w.Sheets(1).Name = "Table_Rapport"
        shNewTemp = w.Sheets.Add(, Application.Sheets("Table_Rapport"))
        shNew = w.Sheets(1)

        '  Copy actual Pivot Ranges to the new worksheet.
        shPivot.PivotTables(1).TableRange2.Copy()
        w.Activate()
        shNewTemp.Activate()
        Application.ActiveSheet.Paste()
        shNewTemp.PivotTables(1).Format(Excel.XlPivotFormatType.xlTable4)

        '  Transform the Pivot table to standard flat Pivot view.
        For Each pf In shNewTemp.PivotTables(1).ColumnFields
            If pf.Name <> "Data" Then
                pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField
                pf.Position = 1
            End If
        Next

        '  Transform the Pivot table to standard flat Pivot view.
        On Error Resume Next
        With shNewTemp.PivotTables(1).DataPivotField
            .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        End With
        On Error GoTo 0

        For Each pf In shNewTemp.PivotTables(1).RowFields
            pf.Subtotals = New Boolean() {False, False, False, False, False, False, False, False, False, False, False, False}
            pf.LayoutBlankLine = False
        Next pf
        shNewTemp.PivotTables(1).RepeatAllLabels(Excel.XlPivotFieldRepeatLabels.xlRepeatLabels)
        With shNewTemp.PivotTables(1)
            .ColumnGrand = False
            .RowGrand = False
        End With

        '  Copy actual Pivot Ranges to the new worksheet.
        shNewTemp.PivotTables(1).TableRange1.Copy()
        shNew.Activate()
        shNew.Cells(1, 1).PasteSpecial(Excel.XlPasteType.xlPasteValues)

        shNew.Cells(2, 1).Activate()
        Application.ActiveWindow.FreezePanes = True
        Application.ActiveWindow.Zoom = 85
        Application.ActiveWindow.DisplayGridlines = False
        Application.Selection.Rows(1).Font.Bold = True
        Application.Selection.Columns.AutoFit()
        shNew.Cells(1, 1).Select()
        Application.DisplayAlerts = False
        shNewTemp.Delete()

CleanUp:
        '  Turn on normal screen refreshing and reset statusbar.
        Application.EnableEvents = True
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Exit Sub

errHandler:
        MsgBox("Active sheet doesn't contain a Pivot Table.", , gSysTitle)

        Application.EnableEvents = True
        Application.StatusBar = False
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.ScreenUpdating = True
    End Sub


    Public Sub TransformPivotToList()
        Dim w As Excel.Workbook
        Dim shNew As Excel.Worksheet
        Dim shPivot As Excel.Worksheet
        Dim shNewTemp As Excel.Worksheet
        Dim pf As Excel.PivotField

        If fnIsPivotTable() = False Then
            Exit Sub
        End If

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        '  Define and fill variables.
        shPivot = Application.ActiveSheet
        w = Application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet)
        w.Sheets(1).Name = "Table_Rapport"
        shNewTemp = w.Sheets.Add(, Application.Sheets("Table_Rapport"))
        shNew = w.Sheets(1)

        '  Copy actual Pivot Ranges to the new worksheet.
        shPivot.PivotTables(1).TableRange2.Copy()
        w.Activate()
        shNewTemp.Activate()
        Application.ActiveSheet.Paste()

        GoTo ResumeHere

        '  Transform the Pivot table to standard flat Pivot view.
        On Error Resume Next
        With shNewTemp.PivotTables(1).DataPivotField
            .Orientation = Excel.XlPivotFieldOrientation.xlColumnField
        End With
        On Error GoTo 0
        For Each pf In shNewTemp.PivotTables(1).ColumnFields
            If pf.Name <> "Data" Then
                pf.Orientation = Excel.XlPivotFieldOrientation.xlRowField
                pf.Position = 1
            End If
        Next

ResumeHere:
        For Each pf In shNewTemp.PivotTables(1).RowFields
            pf.Subtotals = New Boolean() {False, False, False, False, False, False, False, False, False, False, False, False}
            pf.LayoutBlankLine = False
        Next pf
        shNewTemp.PivotTables(1).RepeatAllLabels(Excel.XlPivotFieldRepeatLabels.xlRepeatLabels)
        With shNewTemp.PivotTables(1)
            .ColumnGrand = False
            .RowGrand = False
        End With

        '  Copy actual Pivot Ranges to the new worksheet.
        shNewTemp.PivotTables(1).TableRange1.Copy()
        shNew.Activate()
        shNew.Cells(1, 1).PasteSpecial(Excel.XlPasteType.xlPasteValues)

        shNew.Activate()
        shNew.UsedRange.Borders.LineStyle = Excel.XlLineStyle.xlLineStyleNone
        shNew.UsedRange.Interior.ColorIndex = Excel.XlColorIndex.xlColorIndexNone
        shNew.UsedRange.Font.Bold = True
        shNew.UsedRange.Offset(shNewTemp.PivotTables(1).ColumnFields.Count + 1, 0).Font.Bold = False
        shNew.Cells(shNewTemp.PivotTables(1).ColumnFields.Count + 2, 1).Activate()
        Application.Selection.Columns.AutoFit()
        Application.ActiveWindow.FreezePanes = True
        Application.ActiveWindow.Zoom = 85
        Application.ActiveWindow.DisplayGridlines = False
        shNew.Cells(1, 1).Select()
        Application.DisplayAlerts = False
        shNewTemp.Delete()

CleanUp:
        '  Turn on normal screen refreshing and reset statusbar.
        Application.EnableEvents = True
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Exit Sub

errHandler:
        MsgBox("Active sheet doesn't contain a Pivot Table.", , gSysTitle)

        '  Turn on normal screen refreshing and reset statusbar.
        Application.EnableEvents = True
        Application.StatusBar = False
        Application.ScreenUpdating = True
        Application.DisplayAlerts = True
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
    End Sub


    Public Sub HideAllSheets()
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Call ProtectWorkbook()
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End Sub


    Public Sub ProtectWorkbook() 'Ctrl w
        Dim shActive As Excel.Worksheet
        Dim sh As Excel.Worksheet

        Try
            Application.ActiveWorkbook.Unprotect("next")

            shActive = Application.ActiveSheet
            For Each sh In Application.ActiveWorkbook.Sheets
                If sh.Visible = Excel.XlSheetVisibility.xlSheetVisible Then
                    If sh.Tab.ColorIndex = 22 Then
                        sh.Select()
                        Application.ActiveWindow.SelectedSheets.Visible = False
                    End If
                End If
            Next
            Application.ActiveWorkbook.Protect(Password:="next", Structure:=True, Windows:=False)
            shActive.Activate()

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Utilities", gUserId, gReportID)
        End Try

    End Sub

    Public Sub ShowAllSheets()
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Call UnProtectWorkbook()
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End Sub


    Public Sub UnProtectWorkbook() 'Ctrl d
        Dim shActive As Excel.Worksheet

        shActive = Application.ActiveSheet
        Application.ActiveWorkbook.Unprotect("next")
        Call UnHideSheets()
        shActive.Activate()

    End Sub


    'Sub Endre_Tilgang_S() 'Ctrl s       
    '    Application.ActiveWorkbook.ChangeFileAccess(Mode:=Excel.XlFileAccess.xlReadWrite, WritePassword:="next")
    'End Sub


    Public Sub UnHideSheets()
        Dim sh As Excel.Worksheet

        Try
            '   If ActiveWorkbook.Protect = True Then
            Application.ActiveWorkbook.Unprotect("next")
            '   End If
            For Each sh In Application.ActiveWorkbook.Sheets
                sh.Visible = True
            Next
            Application.ActiveWindow.DisplayWorkbookTabs = True
            Application.ActiveWindow.ScrollWorkbookTabs(Position:=Excel.XlTabPosition.xlTabPositionFirst)
            Application.ActiveWindow.ScrollWorkbookTabs(Position:=Excel.XlTabPosition.xlTabPositionLast)
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Utilities", gUserId, gReportID)
        End Try       

    End Sub


    Public Sub SheetUnprotect()

        Try
            Application.ActiveSheet.Unprotect("next")
            If Application.ActiveSheet.Name = "Rapport info" Then
                Application.ActiveSheet.Range("Hide").Columns.Hidden = False
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Utilities", gUserId, gReportID)
        End Try

    End Sub


    Public Sub SheetProtect()

        Try
            Application.ActiveSheet.EnableSelection = Excel.XlEnableSelection.xlNoRestrictions
            Application.ActiveSheet.Protect(Password:="next", AllowUsingPivotTables:=True)
            If Application.ActiveSheet.Name = "Rapport info" Then
                Application.ActiveSheet.Range("Hide").Columns.Hidden = True
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Utilities", gUserId, gReportID)
        End Try

    End Sub

End Module
