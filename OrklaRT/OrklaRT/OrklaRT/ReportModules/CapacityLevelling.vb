Module CapacityLevelling
    Sub LocalUpdate()

        Call GetStartEndWC()
        gwbReport.Sheets("ReportOptions").Range("WorkCenterGroup").Value = OrklaRTBPL.SelectionFacade.CapacityLevellingWorkGroupCenter
        Application.ActiveWorkbook.Sheets("Ordrer").PivotTables(1).PivotCache.Refresh()
        Application.ActiveWorkbook.Sheets("Uke_Load").PivotTables(1).PivotCache.Refresh()

    End Sub
    Function WCHelpDataTable() As System.Data.DataTable
        Dim table As New System.Data.DataTable
        table.Columns.Add("Work Center", GetType(String))
        table.Columns.Add("Date", GetType(Date))
        table.Columns.Add("Start", GetType(Double))
        table.Columns.Add("Stop", GetType(Double))
        table.Columns.Add("Capacity", GetType(Double))
        Return table
    End Function
    Sub GetStartEndWC()
        Dim x As Integer
        Dim z As Integer
        Dim n As Long
        Dim intHolyday As Integer
        Dim strWC As String
        Dim d As Excel.Range
        Dim w As Integer

        Dim wcHelpTable = WCHelpDataTable()
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        Application.Range("WC_Cap1").ClearContents()
        Application.Sheets("Capacity").PivotTables(1).PivotCache.Refresh()

        Application.Calculate()

        For w = 1 To Application.Sheets("Capacity").PivotTables(1).RowRange.Cells.Count - 1
            strWC = Application.Sheets("Capacity").PivotTables(1).RowRange.Cells(1, 1).Offset(w, 0).Value
            d = Application.Range("CapacityWC").Columns(13).Find(strWC, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)

            For n = DateTime.Now.Date.AddDays(-30).ToOADate() To CDate(Application.Range("MaxDate").Value).ToOADate()
                intHolyday = fnWorkDaysBetween(Application.Sheets("Database").Range("AA1").Offset(1, 0).Value, Date.FromOADate(n - 1), Date.FromOADate(n))
                If intHolyday = 0 Then
                    wcHelpTable.Rows.Add(strWC, Date.FromOADate(n).Date, 0, 0, 0)
                    GoTo NextDate
                End If
                If Not d Is Nothing Then
                    Do While CDate(d.Offset(0, -2).Value).Date < Date.FromOADate(n).Date   'If valid to date is lower than cap. date...
                        If d.Offset(0, -2).Value = Nothing Then Exit Do
                        d = d.Offset(1, 0)  'move to the next row.
                    Loop

                    z = 0 'Initialize capacity date counter.
                    For x = 0 To 6 'Loop through all week days for actual capacity valid to date.
                        Try
                            If Not d.Offset(x, -7).Value Is Nothing Then
                                If d.Offset(x, -7).Value.ToString <> "Ikke tilordnet" Then
                                    If d.Offset(x, -7).Value = Weekday(Date.FromOADate(n).Date, 2) Then
                                        If d.Offset(x, 1).Value = 0 Then
                                            wcHelpTable.Rows.Add(strWC, Date.FromOADate(n).Date, 0, 0, 0)
                                        Else
                                            wcHelpTable.Rows.Add(strWC, Date.FromOADate(n).Date, d.Offset(x, 8).Value, d.Offset(x, 6).Value, d.Offset(x, 1).Value)
                                        End If
                                        z = 1
                                        Exit For
                                    End If
                                End If
                            End If
                        Catch ex As Exception
                            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "CapacityLevelling - GetStartEndWC", gUserId, gReportID)
                        End Try
                    Next x
                End If

                If z = 0 Then
                    wcHelpTable.Rows.Add(strWC, Date.FromOADate(n).Date, 0, 0, 0)
                End If
NextDate:
            Next n
        Next w

        Common.LoadListObjectData("WC_Cap", "WC_Help", "tWC_Help", wcHelpTable)

        Application.Sheets("WC_Pivot").Activate()
        Application.Sheets("WC_Pivot").PivotTables(1).PivotCache.Refresh()

        'On Error Resume Next
        'For Each pi In Application.Sheets("WC_Pivot").PivotTables(1).RowFields
        '    If pi.Name = "Dato" Then
        '        With pi
        '            .ClearAllFilters()
        '            .PivotFilters.Add(Excel.XlPivotFilterType.xlValueEquals, , "(blank)") 'for exact matching
        '        End With               
        '    End If
        'Next
        'On Error GoTo 0

        Application.Sheets("WC_Pivot").PivotTables(1).TableRange1.Copy()

        Application.Sheets("WC_Cap").Activate()
        Application.Sheets("WC_Cap").Range("WC_Paste").PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Operation:=Excel.Constants.xlNone, SkipBlanks:=False, Transpose:=False)
        Application.CutCopyMode = False

        'Application.Range("WC_All").Sort(Key1:=Range("B2"), Order1:=xlDescending, Header:= _
        'xlGuess, OrderCustom:=1, MatchCase:=False, Orientation:=xlTopToBottom, _
        'DataOption1:=xlSortNormal)

        Application.Calculate()
        Call CreateGantt()
        Application.Calculate()
        'Call FormatGraph1()

        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic

    End Sub

    Sub CreateGantt()
        Dim x As Integer
        Dim c As Excel.Range
        Dim dblLength As Double
        Dim dblHeigth As Double
        Dim dblStart As Double
        Dim dblTop As Double
        Dim intTopOffset As Integer
        Dim dblFirst As Double
        Dim intOrder As Integer
        Dim dblConst As Double
        Dim y As Integer
        Dim z As Integer
        Dim intMaterial As Integer
        Dim intNewCap As Integer
        Dim intSapCap As Integer
        Dim intCapacity As Integer
        Dim intStartDate As Integer
        Dim intFirmed As Integer
        Dim LastDate As Date
        Dim rngActive As Excel.Range
        Dim strName As String
        Dim intName As Integer
        Dim strFirmed As String
        Dim intOrder1 As Integer

        '   Application.ScreenUpdating = False
        'Application.Calculation = Excel.XlCalculation.xlCalculationManual
        '   Application.EnableEvents = False

        Application.StatusBar = "Creating Gantt Graph..."
        x = 0
        y = Application.Range("DateCap").Columns(1).Cells.Count
        c = Application.Range("DateCap").Columns(1).Cells(1, 1)
        rngActive = Application.Selection
        Application.Range("GanttRange").ClearContents()

        For z = y To 0 Step -1
            If c.Offset(z, 3).Value > 0 Then
                If x > 200 Then Exit For
                x = x + 1
                Application.Range("StartHere").Offset(0, 1 + x).Value = c.Offset(z, 0).Value
                Application.Range("StartHere").Offset(0, 1 + x).ColumnWidth = c.Offset(z, 3).Value / 86400 * 6
                'Application.Range("StartHere").Offset(0, 1 + x).ColumnWidth = 8
            End If
        Next z

        If z = -1 Then
            LastDate = c.Offset(z + 2, 0).Value
        Else
            LastDate = c.Offset(z - 1, 0).Value
        End If


        x = 1
        intTopOffset = 2
        dblFirst = Application.Range("MaxCap").Value
        Application.Calculate()
        dblConst = Application.Range("GanttCap").Value / Application.Range("CapRange").Width '* 0.985
        '   dblConst = 1020600 / 708
        Application.Range("Const").Value = dblConst

        '  Find column number of db fields.
        intMaterial = fnFieldNo("Material Navn")
        intNewCap = fnFieldNo("Start_D")
        intSapCap = fnFieldNo("Sap_Start")
        intCapacity = fnFieldNo("Kapasitet")
        intStartDate = fnFieldNo("Ord.start")
        intName = fnFieldNo("Material")
        intFirmed = fnFieldNo("Ind.")
        intOrder1 = fnFieldNo("Plnd ordre")

        '  Delete old shapes.
        Application.Sheets("Gantt").Activate()
        Application.Sheets("Gantt").DrawingObjects.Delete()
        Application.Range("StartHere").Offset(0, 0).Value = "Ordre"
        Application.Range("StartHere").Offset(0, 1).Value = "Material Navn"

        For Each c In Application.Range("SapExlData").Columns(1).Cells
            If c.Row > 1 Then
                If c.Value = "" Then Exit For
                If c.Offset(0, intStartDate).Value > LastDate Then Exit For
                Application.Range("StartHere").Offset(x, 0).Value = c.Offset(0, intOrder1).Value 'Order number
                Application.Range("StartHere").Offset(x, 1).Value = c.Offset(0, intMaterial).Value 'Material Name

                dblHeigth = Application.Range("StartHere").Offset(x, 1).Height - 4
                dblStart = Application.Range("StartHere").Offset(x, 2).Left + ((dblFirst - c.Offset(0, intNewCap).Value) / dblConst)
                dblLength = c.Offset(0, intCapacity).Value / dblConst
                dblTop = Application.Range("StartHere").Offset(x, 1).Top + intTopOffset
                intOrder = 1
                strName = c.Offset(0, intName).Value
                strFirmed = c.Offset(0, intFirmed).Value

                Call CreateOrder(dblStart, dblTop, dblLength, dblHeigth, intOrder, strName, strFirmed)

                dblHeigth = Application.Range("StartHere").Offset(x, 1).Height - 4
                dblStart = Application.Range("StartHere").Offset(x, 2).Left + ((dblFirst - c.Offset(0, intSapCap).Value) / dblConst)
                dblLength = c.Offset(0, intCapacity).Value / dblConst
                dblTop = Application.Range("StartHere").Offset(x, 1).Top + intTopOffset
                intOrder = 2
                strName = c.Offset(0, intName).Value
                strFirmed = c.Offset(0, intFirmed).Value

                Call CreateOrder(dblStart, dblTop, dblLength, dblHeigth, intOrder, strName, strFirmed)

                x = x + 1
            End If
        Next c

        For Each c In Application.Range("CapRange").Cells
            If c.Value < Application.Range("MinDate").Value Then
                c.EntireColumn.Hidden = True
            Else
                Application.Range("StartHere").Activate()
                Exit For
            End If
        Next c

        '   rngActive.Activate
        Application.StatusBar = String.Empty

        '   Application.EnableEvents = True
        '   Application.ScreenUpdating = True
        '   Application.Calculation = xlCalculationAutomatic
    End Sub


    Sub CreateOrder(dblStart As Double, dblTop As Double, dblLength As Double, dblHeigth As Double, intOrder As Integer, strName As String, strFirmed As String)
        Dim s As Excel.Shape

        s = Application.Sheets("Gantt").Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, dblStart, dblTop, dblLength, dblHeigth)
        s.DrawingObject.ShapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
        s.DrawingObject.ShapeRange.Fill.Solid()
        If intOrder = 1 Then
            s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 50
            s.DrawingObject.ShapeRange.Fill.Transparency = 0
        Else
            If strFirmed = "X" Then
                s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 10
                s.DrawingObject.ShapeRange.Fill.Transparency = 0.5
            Else
                s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 29
                s.DrawingObject.ShapeRange.Fill.Transparency = 0.75
            End If
        End If
        s.DrawingObject.ShapeRange.Line.Weight = 0.5
        s.DrawingObject.Name = "Material " & strName
        s.DrawingObject.OnAction = "LaunchMD04"
    End Sub
    Function fnFieldNo(strFieldName As String)
        Dim c As Excel.Range
        c = Application.Range("SapExlData").ListObject.HeaderRowRange.Find(strFieldName, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
        fnFieldNo = c.Column - 1
    End Function
    Sub LocalPrepareSaving()
        Application.Sheets("Gantt").Activate()
        Application.Sheets("Gantt").DrawingObjects.Delete()
        Application.Range("GanttRange").ClearContents()
        Application.Range("StartHere").Activate()
        Application.Sheets("Selections").Activate()
    End Sub
    Sub FormatGraph()

        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Application.Sheets("Uke_Load").Activate()
        Application.Sheets("Uke_Load").Range("A10").Activate()

        Application.ActiveSheet.ChartObjects("Chart 154").Activate()
        Application.ActiveChart.SeriesCollection(2).Select()
        With Application.Selection.Border
            .Weight = Excel.XlBorderWeight.xlThin
            .LineStyle = Excel.Constants.xlAutomatic
        End With
        Application.Selection.Shadow = False
        Application.Selection.InvertIfNegative = False
        With Application.Selection.Interior
            .ColorIndex = 19
            .Pattern = Excel.XlPattern.xlPatternSolid
        End With
        Application.ActiveChart.SeriesCollection(3).Select()
        Application.ActiveChart.SeriesCollection(3).ChartType = Excel.XlChartType.xlLineMarkers
        Application.ActiveChart.SeriesCollection(3).Select()
        With Application.Selection.Border
            .ColorIndex = 3
            .Weight = Excel.XlBorderWeight.xlMedium
            .LineStyle = Excel.XlLineStyle.xlContinuous
        End With
        With Application.Selection
            .MarkerBackgroundColorIndex = Excel.Constants.xlAutomatic
            .MarkerForegroundColorIndex = Excel.Constants.xlAutomatic
            .MarkerStyle = Excel.Constants.xlAutomatic
            .Smooth = True
            .MarkerSize = 7
            .Shadow = False
        End With
        Application.ActiveChart.SeriesCollection(3).AxisGroup = 2
        With Application.Selection.Border
            .ColorIndex = 3
            .Weight = Excel.XlBorderWeight.xlMedium
            .LineStyle = Excel.XlLineStyle.xlContinuous
        End With
        With Application.Selection
            .MarkerBackgroundColorIndex = Excel.Constants.xlAutomatic
            .MarkerForegroundColorIndex = Excel.Constants.xlAutomatic
            .MarkerStyle = Excel.Constants.xlNone
            .Smooth = True
            .MarkerSize = 7
            .Shadow = False
        End With
        Application.ActiveChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary).Select()
        With Application.ActiveChart.Axes(Excel.XlAxisType.xlValue, Excel.XlAxisGroup.xlSecondary)
            .MinimumScale = 0
            .MaximumScaleIsAuto = True
            .MinorUnitIsAuto = True
            .MajorUnitIsAuto = True
            .Crosses = Excel.Constants.xlAutomatic
            .ReversePlotOrder = False
            .ScaleType = Excel.XlScaleType.xlScaleLinear
            .DisplayUnit = Excel.Constants.xlNone
        End With
        Application.ActiveChart.ChartArea.Select()
        Application.ActiveWindow.Visible = False

        Application.Range("A10").Activate()

        Application.EnableEvents = True
        Application.ScreenUpdating = True

    End Sub
End Module
