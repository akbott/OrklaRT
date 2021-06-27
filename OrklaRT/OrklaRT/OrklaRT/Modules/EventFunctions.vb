Module EventFunctions
    'This module content common functions and subs used by addin events module

    Function IsOrklaRTReport(wb As Excel.Workbook) As String
        Try
            gsErrTest = wb.Sheets("Version").Range("SapExlReportName").Value
            gwbReport = Application.ActiveWorkbook
        Catch
            gsErrTest = String.Empty
        End Try
        Return gsErrTest
    End Function
    'Sub IsOrklaRTReport(wb As Excel.Workbook)
    '    Try
    '        gsErrTest = wb.Sheets("Version").Range("SapExlReportName").value
    '    Catch
    '        gwbReport = Nothing
    '        Exit Sub
    '    End Try
    '    gwbReport = wb
    'End Sub

    Function fnIsPivotTable() As Boolean
        Try
            gsErrTest = Application.ActiveCell.PivotTable.Name
        Catch e As Exception
            fnIsPivotTable = False
            Exit Function
        End Try
        fnIsPivotTable = True
    End Function

    Function fnIsSheetDetails(sh As Excel.Worksheet) As Boolean
        fnIsSheetDetails = False
        On Error Resume Next
        gsErrTest = sh.Range("ShowDetails").Cells.Count
        If Err.Number <> 0 Then Exit Function
        fnIsSheetDetails = True
    End Function

    Sub ResetExcelFind()
        Dim c As Excel.Range
        'Reset the Excel Search Dialog Box to Excel default
        c = Application.Cells.Find("", LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
    End Sub

    Sub ResetExcelDefaults() 'Probably not relevant for use in the .NET solution.
        On Error Resume Next
        Application.Caption = ""
        Application.OnKey("{F8}")
        Application.OnKey("{F1}")
        Call ResetExcelFind()
        Application.StatusBar = False
    End Sub

    Sub SetOrklaRTSettings() 'Probably not relevant for use in the .NET solution.
        Application.Caption = gSYSTITLE
        Application.OnKey("{F8}", "RunUpdatePivot")
        Application.OnKey("{F1}") 'OrklaRT Help
        Application.StatusBar = False
    End Sub

    Sub RefreshStandardPivot()
        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Dim sFirstSheet As String = gwbReport.Sheets("Version").range("FirstSheet").value
        Dim sFirstTable As String = gwbReport.Sheets("Version").range("FirstTable").value
        Dim pvt As Excel.PivotTable = gwbReport.Sheets(sFirstSheet).Pivottables(sFirstTable)
        pvt.PivotCache.Refresh()

        Application.EnableEvents = True
        Application.ScreenUpdating = True

    End Sub

    Sub StopEvents(Optional bScreen As Boolean = False, _
      Optional bEvents As Boolean = False, _
      Optional bCalc As Boolean = False, _
      Optional bAlerts As Boolean = False, _
      Optional bCKey As Boolean = False)

        If bScreen Then Application.ScreenUpdating = False
        If bEvents Then Application.EnableEvents = False
        If bCalc Then Application.DisplayAlerts = False
        If bAlerts Then Application.Calculation = Excel.XlCalculation.xlCalculationManual
        If bCKey Then Application.EnableCancelKey = Excel.XlEnableCancelKey.xlDisabled
    End Sub

    Sub ResetEvents(Optional bScreen As Boolean = True, _
       Optional bEvents As Boolean = True, _
       Optional bAlerts As Boolean = True, _
       Optional bCalc As Boolean = True, _
       Optional bCKey As Boolean = True, _
       Optional bStatusBar As Boolean = True)

        If bScreen = False Then Application.ScreenUpdating = True
        If bEvents = False Then Application.EnableEvents = True
        If bAlerts = False Then Application.DisplayAlerts = True
        If bCalc Then Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        If bCKey Then Application.EnableCancelKey = Excel.XlEnableCancelKey.xlInterrupt

    End Sub
    Sub ResetAllEvents()
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.EnableCancelKey = Excel.XlEnableCancelKey.xlInterrupt
        Application.StatusBar = String.Empty
    End Sub


    Sub ShowReportSheets()
        Dim shActive As Excel.Worksheet
        Dim sh As Excel.Worksheet

        shActive = Application.ActiveSheet
        If gwbReport.ProtectStructure = True Then
            gwbReport.Unprotect(Password:="next")
            For Each sh In gwbReport.Sheets
                If sh.Tab.Color = 192 Then
                    sh.Visible = True
                End If
            Next
        End If
        Application.ActiveWindow.DisplayWorkbookTabs = True
        gwbReport.Protect(Password:="next", Structure:=True, Windows:=False)
        shActive.Activate()
    End Sub

    Sub PivotTablePlacement(pt As Excel.PivotTable)
        'Dim lStartRow As Long
        'Dim lFixedRow As Long = 6
        'Dim lStartRowOffset As Long
        'Dim oSelection As Object

        'oSelection = Application.Selection

        'Adjust Pivot Table vertical placement on the sheet.
        'lStartRow = pt.TableRange2.Rows(1).Row
        'lStartRowOffset = lStartRow - lFixedRow
        'If lStartRowOffset > 0 Then
        '    Application.Rows(lFixedRow & ":" & lFixedRow + lStartRowOffset - 1).delete(Shift:=Excel.XlDeleteShiftDirection.xlShiftUp)
        'ElseIf lStartRowOffset < 0 Then
        '    Application.Rows(lStartRow & ":" & lStartRow - lStartRowOffset - 1).Insert(Excel.XlDeleteShiftDirection.xlShiftUp)
        'End If

        'oSelection = Nothing
    End Sub

    Sub ReadDescriptions(rTarget As Excel.Range)
        Dim sFieldName As String
        Dim sDescription As String
        Dim sCaption As String
        Dim c As Excel.Range
        Dim bolInfo As Boolean
        Dim bolDetails As Boolean

        On Error GoTo CleanUp
        bolDetails = False
        sDescription = vbEmpty
        If gwbReport.ActiveSheet.Names.Count > 0 Then
            If gwbReport.ActiveSheet.Names(1).Name = gwbReport.ActiveSheet.Name & "!ShowDetails" Then
                sFieldName = gwbReport.Sheets(rTarget.Parent.name).Cells(1, rTarget.Column)
                sCaption = sFieldName
                bolDetails = True
            Else
                sFieldName = rTarget.PivotField.SourceName
                sCaption = rTarget.PivotField.Caption
            End If
        Else
            sFieldName = rTarget.PivotField.SourceName
            sCaption = rTarget.PivotField.Caption
        End If

        c = gwbReport.Sheets("Descriptions").Range("Descriptions").Columns(1).Find(sFieldName, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
        If Not c Is Nothing Then
            sDescription = sCaption & ": " & c.Offset(0, 1).Value
            If bolDetails = False Then
                If gwbReport.ActiveSheet.PivotTables(1).PivotFields(sFieldName).IsCalculated Then
                    sDescription = sDescription & "       Formula: " & gwbReport.ActiveSheet.PivotTables(1).PivotFields(sFieldName).Formula
                End If
            End If
        Else
            If IsNumeric(Right(sFieldName, 1)) Then
                sFieldName = Mid(sFieldName, 1, Len(sFieldName) - 1)
                c = gwbReport.Sheets("Descriptions").Range("Descriptions").Columns(1).Find(sFieldName, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
                If Not c Is Nothing Then
                    sDescription = sCaption & ": " & c.Offset(0, 1).Value
                    If bolDetails = False Then
                        If gwbReport.ActiveSheet.PivotTables(1).PivotFields(sFieldName).IsCalculated Then
                            sDescription = sDescription & "       Formula: " & gwbReport.ActiveSheet.PivotTables(1).PivotFields(sFieldName).Formula
                        End If
                    End If
                End If
            End If
        End If
        Call ResetExcelFind()
        c = Nothing
        If sDescription <> "" Then
            Application.StatusBar = sDescription
        Else
            Application.StatusBar = False
        End If
        Exit Sub

CleanUp:
        Call ResetExcelFind()
        c = Nothing
        Application.StatusBar = False
    End Sub
    Public Sub LocalLockUnlockOrder()
        'Dim intOrderCol As Integer
        'Dim c As Excel.Range

        'If (New DAL.SAPExlEntities()).vwCurrentUser.SingleOrDefault().ProductionPlant <> OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant Then GoTo CleanUp 'Only specified users are allowed.
        '        Application.EnableEvents = False
        '        Application.ScreenUpdating = False
        '        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        '        Try
        '            If Application.ActiveCell.PivotTable.Name <> "" Then

        '                If Application.ActiveCell.PivotField.SourceName = "Locked" Then

        '                    For Each c In Application.Selection
        '                        If Not IsError(Application.ActiveCell.PivotTable.PivotFields("Order").LabelRange.Column) Then
        '                            intOrderCol = Application.ActiveCell.PivotTable.PivotFields("Order").LabelRange.Column
        '                            If c.Row > Application.ActiveCell.PivotTable.PivotFields("Order").LabelRange.Row Then
        '                                If Application.Cells(c.Row, intOrderCol).Value > 0 Then
        '                                    Call WriteLockedOrders(Application.Cells(c.Row, intOrderCol).Value)
        '                                End If
        '                            End If                      
        '                        End If
        '                    Next

        '                Else
        '                    MsgBox("Active cell must be a Pivot table cell in the 'Locked' column to use the Lock / Unlock function.", , "Orkla SAP Intergation")
        '                    Exit Sub
        '                End If
        '            End If
        '        Catch ex As Exception
        '            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Event Functions - LocalLockUnlockOrder", gUserId, gReportID)
        '        End Try

        'CleanUp:
        '        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic        
        '        Application.ScreenUpdating = True
        '        Application.EnableEvents = True
        'Exit Sub

    End Sub
    Public Sub PreviewPrinting()
        Dim iOrientation As Integer

        Application.ScreenUpdating = False

        If fnIsPivotTable() = True Then
            Application.ActiveCell.PivotTable.TableRange1.Columns.AutoFit()
        End If

        If Application.ActiveSheet.UsedRange.Width > 700 Then
            iOrientation = Excel.XlPageOrientation.xlLandscape
        Else
            iOrientation = Excel.XlPageOrientation.xlPortrait
        End If

        With Application.ActiveSheet
            With .PageSetup
                If fnIsPivotTable() = True Then
                    .PrintTitleRows = Application.ActiveSheet.PivotTables(1).ColumnRange.EntireRow.Address
                End If
                .Orientation = iOrientation
                If Not gwbReport Is Nothing Then
                    .LeftHeader = "OrklaRT Reports"
                End If
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = False
                .LeftFooter = "Utskrift dato: &D"
                .CenterFooter = "Side &P / &N"
                .RightFooter = "&F - &A "
            End With
        End With

ResumeHere:
        Application.ActiveWindow.SelectedSheets.PrintPreview()

CleanUp:
        Application.ScreenUpdating = True
        Exit Sub
    End Sub
End Module
