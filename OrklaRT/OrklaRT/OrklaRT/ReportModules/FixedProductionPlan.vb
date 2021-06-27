Imports Microsoft.Office.Interop.Excel
Imports System.Reflection
Imports System.Data.SqlClient
Imports System.Runtime.InteropServices


Module FixedProductionPlan
    Public bolNoPlan As Boolean
    Public bolOnlyRefresh As Boolean
    Public boolErr As Boolean
    Public Sub LocalUpdate()

        Try
            Application.StatusBar = "Retrieving Saved Plan Version..."
            Application.Calculation = Excel.XlCalculation.xlCalculationManual

            Call GetExistingPlan()

            Call Globals.Ribbons.OrklaRT.GetLockedOrders()

            If bolNoPlan = True Then GoTo CleanUp
            Application.Calculate()
            Call GetStartEndWC()

            Application.StatusBar = "Refreshing Production Plan Status..."
            Application.Calculate()
            Call CreatePlanShapes()

            Application.Calculate()
            Call RefreshStatus()

            Application.Calculate()
            Call PlanStatusUpdate()

            Application.Calculate()
            Call GetTime()

            Call GetPriorities()

            Call RefreshIntTables()

CleanUp:
            Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
            Application.StatusBar = String.Empty

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Fixed Production Plan - LocalUpdate", gUserId, gReportID)
        End Try

        Application.ActiveWindow.Zoom = 75
        Application.ActiveWindow.Zoom = 85

    End Sub
    Public Sub GetStartEndWC()

        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim d As Range
        Dim r As Range
        Dim ActDate As Date
        Dim CapacityDate As Date
        Dim intHolyday As Integer
        Dim strWC As String
        '  ———————————————————————————————————————————————————————————————————————————

        '   Application.EnableEvents = False
        Application.Range("WC_Cap").Clear()
        y = 0

        On Error Resume Next

        For Each c In Application.Range("SapExlData").Columns(1).Cells
            '   For Each c In Range("PlanData").Columns(1).Cells
            If c.Value.ToString() = "" Then Exit For
            '      Debug.Print "Rad: " & c.Row

            If c.Row = 1 Then GoTo ResumeHere

            '     Check if we start at a new Work center number.
            If c.Offset(0, 42).Value <> c.Offset(-1, 42).Value Then
                ActDate = Application.Range("CapDate").Value 'Start at the first capacity date.                
                '        Find the first line of the new work venter in sheet Capacity.
                d = Application.Sheets("Capacity").Columns(13).Find(c.Offset(0, 42).Value, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                strWC = c.Offset(0, 75).Value
                If Not d Is Nothing Then
                    For Each r In Application.Sheets("WC_Cap").Range("CapDates").Cells 'Loop through all capacity dates.
                        'If r is a holyday goto next.
                        CapacityDate = r.Value
                        intHolyday = fnWorkDaysBetween(Application.Sheets("Database").Range("AP2").Value, Microsoft.VisualBasic.DateValue(r.Value).AddDays(-1), Microsoft.VisualBasic.DateValue(r.Value))
                        r.Offset(0, 1).Value = intHolyday
                        If intHolyday = 0 Then
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 0).Value = strWC
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 1).Value = r.Value
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 2).Value = 0
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 3).Value = 0
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 4).Value = 0
                            y = y + 1
                            GoTo NextDate
                        End If

                        ' Do While CDate(d.Offset(0, -2).Value).Date < CapacityDate  'If valid to date is lower than cap. date...
                        'If d.Offset(0, -2).Value = Nothing Then Exit Do
                        'd = d.Offset(1, 0)  'move to the next row.
                        'Loop
                        'ContHere:

                        z = 0 'Initialize capacity date counter.

                        For x = 0 To 6 'Loop through all week days for actual capacity valid to date.
                            'If x > 0 And d.Offset(x, -2).Value > d.Offset(x - 1, -2).Value Then Exit For
                            '                  Debug.Print d.Offset(x, 5).Value
                            '                  Debug.Print d.Row
                            '                  Debug.Print Weekday(r.Value, 2)
                            If CInt(d.Offset(x, -7).Value) = Weekday(r.Value, 2) Then
                                If d.Offset(x, 1).Value = 0 Then
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 0).Value = strWC
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 1).Value = r.Value
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 2).Value = 0
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 3).Value = 0
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 4).Value = 0
                                Else
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 0).Value = strWC
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 1).Value = r.Value
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 2).Value = d.Offset(x, 8).Value
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 3).Value = d.Offset(x, 6).Value
                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 4).Value = d.Offset(x, 1).Value
                                End If
                                '                     Debug.Print d.Offset(x, 7).Row
                                y = y + 1
                                z = 1
                                Exit For
                            End If
                        Next x

                        If z = 0 Then
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 0).Value = strWC
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 1).Value = r.Value
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 2).Value = 0
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 3).Value = 0
                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 4).Value = 0
                            y = y + 1
                        End If
NextDate:
                    Next r
                End If
            End If
ResumeHere:
        Next c

        Call CapacityStart()
        '   Application.EnableEvents = True

    End Sub
    '    Public Sub GetStartEndWC()

    '        Dim x As Integer
    '        Dim y As Integer
    '        Dim z As Integer
    '        Dim d As Range
    '        Dim r As Range
    '        Dim ActDate As Date
    '        Dim CapacityDate As Date
    '        Dim intHolyday As Integer
    '        Dim strWC As String
    '        '  ———————————————————————————————————————————————————————————————————————————

    '        '   Application.EnableEvents = False
    '        Application.Range("WC_Cap").Clear()
    '        y = 0

    '        On Error Resume Next

    '        For Each c In Application.Range("SapExlData").Columns(1).Cells
    '            '   For Each c In Range("PlanData").Columns(1).Cells
    '            If c.Value.ToString() = "" Then Exit For
    '            '      Debug.Print "Rad: " & c.Row

    '            If c.Row = 1 Then GoTo ResumeHere

    '            '     Check if we start at a new Work center number.
    '            If c.Offset(0, 42).Value <> c.Offset(-1, 42).Value Then
    '                ActDate = Application.Range("CapDate").Value 'Start at the first capacity date.                
    '                '        Find the first line of the new work venter in sheet Capacity.
    '                d = Application.Sheets("Capacity").Columns(13).Find(c.Offset(0, 42).Value, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '                strWC = c.Offset(0, 75).Value
    '                If Not d Is Nothing Then
    '                    For Each r In Application.Sheets("WC_Cap").Range("CapDates").Cells 'Loop through all capacity dates.
    '                        'If r is a holyday goto next.
    '                        CapacityDate = r.Value
    '                        intHolyday = fnWorkDaysBetween(Application.Sheets("Database").Range("AP2").Value, Microsoft.VisualBasic.DateValue(r.Value).AddDays(-1), Microsoft.VisualBasic.DateValue(r.Value))
    '                        r.Offset(0, 1).Value = intHolyday
    '                        If intHolyday = 0 Then
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 0).Value = strWC
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 1).Value = r.Value
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 2).Value = 0
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 3).Value = 0
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 4).Value = 0
    '                            y = y + 1
    '                            GoTo NextDate
    '                        End If

    '                        ' Do While CDate(d.Offset(0, -2).Value).Date < CapacityDate  'If valid to date is lower than cap. date...
    '                        'If d.Offset(0, -2).Value = Nothing Then Exit Do
    '                        'd = d.Offset(1, 0)  'move to the next row.
    '                        'Loop
    '                        'ContHere:

    '                        z = 0 'Initialize capacity date counter.

    '                        For x = 0 To 6 'Loop through all week days for actual capacity valid to date.
    '                            'If x > 0 And d.Offset(x, -2).Value > d.Offset(x - 1, -2).Value Then Exit For
    '                            '                  Debug.Print d.Offset(x, 5).Value
    '                            '                  Debug.Print d.Row
    '                            '                  Debug.Print Weekday(r.Value, 2)
    '                            If CInt(d.Offset(x, -7).Value) = Weekday(r.Value, 2) Then
    '                                If d.Offset(x, 1).Value = 0 Then
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 0).Value = strWC
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 1).Value = r.Value
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 2).Value = 0
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 3).Value = 0
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 4).Value = 0
    '                                Else
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 0).Value = strWC
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 1).Value = r.Value
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 2).Value = d.Offset(x, 8).Value
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 3).Value = d.Offset(x, 6).Value
    '                                    Application.Range("WC_Cap").Cells(1, 1).Offset(y, 4).Value = d.Offset(x, 1).Value
    '                                End If
    '                                '                     Debug.Print d.Offset(x, 7).Row
    '                                y = y + 1
    '                                z = 1
    '                                Exit For
    '                            End If
    '                        Next x

    '                        If z = 0 Then
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 0).Value = strWC
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 1).Value = r.Value
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 2).Value = 0
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 3).Value = 0
    '                            Application.Range("WC_Cap").Cells(1, 1).Offset(y, 4).Value = 0
    '                            y = y + 1
    '                        End If
    'NextDate:
    '                    Next r
    '                End If
    '            End If
    'ResumeHere:
    '        Next c

    '        Call CapacityStart()
    '        '   Application.EnableEvents = True

    '    End Sub

    Public Sub CapacityStart()
        Dim c As Range
        Dim d As Range
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer

        '   Application.EnableEvents = False
        x = 2
        Application.Sheets("Cap_Start").Cells.ClearContents()

        Application.Calculate()
        For Each c In Application.Sheets("WC_Cap").Columns(1).Cells
            If c.Row > 1 Then

                '         If Left(c.Value, 4) = "3107" Then Stop

                If c.Value <> c.Offset(-1, 0).Value Then
                    If c.Value.ToString() = "" Then Exit For
                    x = x + 1
                    y = 2
                    '            Debug.Print c.Row
                    Application.Sheets("Cap_Start").Cells(x, 1).Value = c.Value
                    Application.Sheets("Cap_Start").Cells(x, y).Value = c.Offset(0, 7).Value
                    '            If c.Row = 257 Then Stop
                    '            Debug.Print "Rad " & c.Row
                Else
                    y = y + 1
                    Application.Sheets("Cap_Start").Cells(x, y).Value = c.Offset(0, 7).Value
                    '            Debug.Print "Rad " & c.Row
                End If
            Else
                z = 1
                Application.Sheets("Cap_Start").Cells(1, 1).Value = "Work Center"
                For Each d In Application.Range("CapDates").Columns(1).Cells
                    z = z + 1
                    Application.Sheets("Cap_Start").Cells(1, z).Value = d.Value
                    Application.Sheets("Cap_Start").Cells(2, z).Value = z
                Next d
            End If
        Next c
        '   Application.EnableEvents = True

    End Sub

    Public Sub GetExistingPlan()

        Dim c As Range
        Dim sh As Worksheet
        Dim s As Range
        Dim PlanTime As Date
        Dim resultTable As New System.Data.DataTable

        Try
            If OrklaRTBPL.ReportSpecific.CheckPlanExists(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate).Tables("ProductionPlanData").Rows.Count > 0 Then
                bolNoPlan = False
                resultTable = OrklaRTBPL.ReportSpecific.CheckPlanExists(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate).Tables("ProductionPlanData")
                PlanTime = OrklaRTBPL.ReportSpecific.GetPlanTime(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate)

                Application.Sheets("PlanData").Range("PlanTime").Value = PlanTime
                Application.Sheets("PlanData").Range("PlanTime1").Value = "'" & PlanTime

                Common.LoadListObjectData("PlanData", "PlanData", "PlanAll", resultTable)                

            Else
                MsgBox("No Plan version exists for your selections," & Chr(10) &
                      "Please check your input selections.", , "Orkla SAP Integration")
                bolNoPlan = True
                GoTo CleanUp
            End If

            Exit Sub

CleanUp:
            sh = Application.Sheets("ProdPlan")
            c = Application.ActiveCell
            sh.Activate()
            Application.Range("Night").EntireColumn.Hidden = False

            'Delete all shapes.
            If sh.Shapes.Count > 4 Then
                For Each s In sh.Shapes
                    s.Select()
                    If (TypeName(Application.Selection) = "TextBox" Or TypeName(Application.Selection) = "Rectangle") And s.Name <> "TimeShape" And s.Name <> "TimeStart" Then s.Delete()
                Next s
                sh.Rows("7:100").Clear()
                sh.Rows("7:100").RowHeight = 12.75
                c.Activate()
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Fixed Production Plan", gUserId, gReportID)
        End Try
    End Sub

    Public Sub CreatePlanShapes()
        Dim c As Range
        Dim s As Shape
        Dim sh As Worksheet
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim dblLength As Double
        Dim dblHeight As Double
        Dim dblStart As Double
        Dim dblSetup As Double
        Dim dblEnd As Double
        Dim dblSecValue As Double
        Dim dblTop As Double
        Dim Day1 As Date
        Dim Day2 As Date
        Dim strMaterial As String
        Dim strQty As String
        Dim strQty1 As String
        Dim strPeriod As String
        Dim ActCell As Range
        Dim intTopOffset As Integer
        Dim strMixed As String

        '   Application.ScreenUpdating = False
        '   Application.EnableEvents = False

        'On Error GoTo ResumeHere
        Application.Sheets("ProdPlan").Activate()
        ActCell = Application.Selection

        '   GoTo CleanUp

        sh = Application.Sheets("ProdPlan")
        sh.Activate()
        Application.Range("Night").EntireColumn.Hidden = False
        'Delete all shapes.
        For Each s In sh.Shapes
            If s.Name = "TimeShape" Or s.Name = "TimeStart" Or s.Name = "TimeLine" Then
                'Debug.Print s.Name
            Else
                '         Debug.Print TypeName(s.DrawingObject)
                If TypeName(s.DrawingObject) = "TextBox" Or TypeName(s.DrawingObject) = "Rectangle" Then s.Delete()
            End If
        Next s
        sh.Rows("7:100").Clear()
        sh.Rows("7:100").RowHeight = 12.75
        '   GoTo CleanUp

        Day1 = Application.Range("CapDate").Value
        Day2 = fnCalendarDay(Application.Sheets("Database").Range("AP2").Value, Day1, 1)
        dblSecValue = 1602
        dblHeight = 37
        intTopOffset = 3

        x = 0
        For Each c In Application.Sheets("PlanData").Range("PlanData").Columns(1).Cells
            If c.Value Is Nothing Then Exit For
            '      Debug.Print "Rad: " & c.Row

            If c.Row > 1 Then

                '         If c.Row = 9 Then Stop

                'Fill in Work center name and format row heights.

                If c.Offset(0, 42).Value.ToString() <> c.Offset(-1, 42).Value.ToString() And c.Offset(0, 42).Value.ToString() <> "0" Then
                    x = x + 1
                    sh.Range("StartWC").Offset(x, 0).RowHeight = 43
                    sh.Range("StartWC").Offset(x, 0).Value = c.Offset(0, 75).Value
                    sh.Range("StartWC").Offset(x, 0).VerticalAlignment = Excel.Constants.xlCenter
                    '         sh.Range("StartWC").Offset(x, 0).Font.Bold = True
                End If

                'If start day and time (66) > (day 1 + shift start time) in the plan than calculate setup length.        

                If CDate(c.Offset(0, 72).Value) > CDate(CStr(Day1) + " " + c.Offset(0, 74).Text) Then
                    dblSetup = CDbl(c.Offset(0, 64).Value) / 24 * dblSecValue
                Else
                    dblSetup = 0
                End If

                'If active tag not 0 and start time <= day2 then create shapes...
                '      If c.Offset(0, 80).Value > 0 And Int(c.Offset(0, 66).Value) <= Day2 Then
                If Int(c.Offset(0, 69).Value) > 0 And CDate(c.Offset(0, 72).Value).Date <= Day2 Then
                    dblStart = Application.Range("_Day1").Left + (CDate(c.Offset(0, 72).Value) - Day1).TotalDays * dblSecValue + dblSetup
                    dblEnd = Application.Range("_Day1").Left + (CDate(c.Offset(0, 73).Value) - Day1).TotalDays * dblSecValue
                    dblLength = dblEnd - dblStart
                    If dblLength < 0 Then dblLength = 0
                    strMaterial = c.Offset(0, 61).Value
                    '            strQty = c.Offset(0, 56).Value
                    strQty = c.Offset(0, 113).Value
                    strQty1 = c.Offset(0, 114).Value
                    If IsNumeric(c.Offset(0, 111).Value) Then
                        If Int(c.Offset(0, 111).Value) > 1 Then
                            strMixed = ". Mixed " & String.Format("{0:0%}", 1)
                        Else
                            strMixed = ". Mixed " & String.Format("{0:0%}", c.Offset(0, 111).Value)
                        End If
                    End If

                    strPeriod = CDate(c.Offset(0, 72).Value).ToString("dd.MM") & " " & CDate(c.Offset(0, 72).Value).ToString("HH.mm") & " - " & CDate(c.Offset(0, 73).Value).ToString("dd.MM") & " " & CDate(c.Offset(0, 73).Value).ToString("HH.mm") & strMixed


                    dblTop = CInt(sh.Range("StartWC").Offset(x, 0).Top) + intTopOffset

                    s = Application.ActiveSheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, dblStart, dblTop, dblLength, dblHeight)
                    s.DrawingObject.Characters.Text = strMaterial & Chr(10) & "Order " & c.Value & ", Rem. Time " & String.Format("{0:#,##0.0}", CDbl(strQty)) & ", Rem. Qty " & String.Format("{0:#,##0}", CDbl(strQty1)) & Chr(10) & strPeriod
                    s.DrawingObject.Font.Name = "Arial"
                    s.DrawingObject.Font.Size = 10
                    s.DrawingObject.Font.ColorIndex = 1
                    s.DrawingObject.HorizontalAlignment = Excel.Constants.xlLeft
                    s.DrawingObject.ShapeRange.Line.Weight = 0.5
                    s.DrawingObject.ShapeRange.TextFrame.MarginLeft = 1.5
                    s.DrawingObject.ShapeRange.TextFrame.MarginRight = 0.5
                    s.DrawingObject.ShapeRange.TextFrame.MarginTop = 2
                    s.DrawingObject.ShapeRange.TextFrame.MarginBottom = 0.5
                    s.DrawingObject.ShapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                    s.DrawingObject.ShapeRange.Fill.Solid()
                    '         If c.Offset(0, 80).Value = 1 Then
                    If Int(c.Offset(0, 69).Value) = 1 Then
                        s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 29
                    Else
                        s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 52  'Organge
                    End If
                    s.DrawingObject.Name = "Order" & c.Value
                    s.DrawingObject.OnAction = "LaunchSAPOrder"
                    s.DrawingObject.ShapeRange.Fill.Transparency = 0.6
                    s.DrawingObject.ShapeRange.Line.ForeColor.SchemeColor = 55
                    s.DrawingObject.ShapeRange.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack)

                    If CDate(c.Offset(0, 72).Value) > CDate(CStr(Day1) + " " + c.Offset(0, 74).Text) Then
                        dblEnd = dblStart
                        dblStart = dblEnd - CDbl(c.Offset(0, 64).Value) / 24 * dblSecValue
                        dblLength = dblEnd - dblStart
                        If dblLength < 0 Then dblLength = 0
                        dblTop = CInt(sh.Range("StartWC").Offset(x, 0).Top) + intTopOffset
                        strQty = c.Offset(0, 64).Value

                        s = Application.ActiveSheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeActionButtonCustom, dblStart, dblTop, dblLength, dblHeight)
                        s.DrawingObject.Characters.Text = String.Format("{0:#,##0.0}", CDbl(strQty)) & " h" & Chr(10) & "Setup"
                        s.DrawingObject.Font.Name = "Arial"
                        s.DrawingObject.Font.Size = 10
                        s.DrawingObject.Font.ColorIndex = 1
                        s.DrawingObject.HorizontalAlignment = Excel.Constants.xlLeft
                        s.DrawingObject.ShapeRange.Line.Weight = 0.5
                        s.DrawingObject.ShapeRange.TextFrame.MarginLeft = 1.5
                        s.DrawingObject.ShapeRange.TextFrame.MarginRight = 0.5
                        s.DrawingObject.ShapeRange.TextFrame.MarginTop = 2
                        s.DrawingObject.ShapeRange.TextFrame.MarginBottom = 0.5
                        s.DrawingObject.ShapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        s.DrawingObject.ShapeRange.Fill.Solid()
                        s.DrawingObject.ShapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        s.DrawingObject.ShapeRange.Fill.Solid()
                        s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 44
                        s.DrawingObject.Name = "Setup" & c.Value
                        s.DrawingObject.OnAction = "LaunchSAPOrder"
                        s.DrawingObject.ShapeRange.Fill.Transparency = 0.6
                        s.DrawingObject.ShapeRange.Line.ForeColor.SchemeColor = 55
                        '"#N/A"
                    End If

                End If
            End If
            '   On Error Resume Next
        Next c

ResumeHere:
        Call CreateBreakes()

CleanUp:
        Application.Range("Night").EntireColumn.Hidden = True
        ActCell.Activate()
    End Sub

    Sub CreateBreakes()
        Dim c As Range
        Dim d As Range
        Dim x As Integer
        Dim y As Integer
        Dim intTopOffset As Integer
        Dim sh As Worksheet
        Dim dblLength As Double
        Dim dblHeight As Double
        Dim dblStart As Double
        Dim dblEnd As Double
        Dim dblSecValue As Double
        Dim dblTop As Double
        Dim Day1 As Date
        Dim Day2 As Date

        '   Application.EnableEvents = False

        sh = Application.Sheets("ProdPlan")
        sh.Activate()

        Day1 = DateTime.Now.Date
        Day2 = fnCalendarDay(Application.Sheets("Database").Range("AP2").Value, Day1, 1)
        dblSecValue = sh.Range("_Day1").Width / CLng(86400)
        '   Debug.Print dblSecValue
        dblHeight = 37
        intTopOffset = 3
        y = 0
        x = 0

        For Each c In Application.Range("WC_Plan").Columns(1).Cells
            If c.Value Is Nothing Then Exit For
            '      Debug.Print c.Value

            '      If c.Value = "3201 Grunwald 3,Mesanin A" Then Stop

            d = Application.Sheets("WC_Cap").Columns(1).Find(c.Value, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            '      Debug.Print d.Value
            If Not d Is Nothing Then
                For y = 0 To 2
                    If d.Offset(y, 2).Value <> d.Offset(y, 3).Value Then
                        '               Debug.Print d.Offset(y, 2).Value
                        '               Debug.Print d.Offset(y, 2).Value * dblSecValue
                        dblStart = Application.Range("_Day1").Left + Application.Range("_Day1").Width * y + (d.Offset(y, 3).Value * dblSecValue)
                        '               dblEnd = Range("_Day1").Left + Range("_Day1").Width * (y + 1) + (d.Offset(y, 2).Value * dblSecValue)
                        dblEnd = Application.Range("_Day1").Left + Application.Range("_Day1").Width * (y + 1) + (d.Offset(y + 1, 2).Value * dblSecValue)
                    ElseIf d.Offset(y, 2).Value = 0 Then
                        dblStart = Application.Range("_Day1").Left + Application.Range("_Day1").Width * y + (d.Offset(y, 2).Value * dblSecValue)
                        dblEnd = Application.Range("_Day1").Left + Application.Range("_Day1").Width * y + Application.Range("_Day1").Width
                        If d.Offset(y + 1, 2).Value <> d.Offset(y + 1, 3).Value Then
                            dblEnd = Application.Range("_Day1").Left + Application.Range("_Day1").Width * y + Application.Range("_Day1").Width + (d.Offset(y + 1, 2).Value * dblSecValue)
                        End If
                    End If
                    dblLength = dblEnd - dblStart
                    If dblLength < 0 Then dblLength = 0
                    dblTop = c.Top + intTopOffset

                    If (d.Offset(y, 2).Value <> d.Offset(y, 3).Value) Or d.Offset(y, 2).Value = 0 Then
                        Application.ActiveSheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, dblStart, dblTop, dblLength, dblHeight).Select()
                        Application.Selection.ShapeRange.Line.Weight = 0.5
                        Application.Selection.ShapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                        Application.Selection.ShapeRange.Fill.Solid()
                        Application.Selection.ShapeRange.Fill.ForeColor.SchemeColor = 55 '22 light grey, 55 medium grey, 23 dark grey
                        Application.Selection.ShapeRange.Fill.Transparency = 0
                        Application.Selection.ShapeRange.Line.Visible = Microsoft.Office.Core.MsoTriState.msoFalse
                        Application.Selection.ShapeRange.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoSendToBack)
                    End If
                Next y
                If c.Value.ToString() <> c.Offset(0, 1).Value Then
                    x = x + 1
                End If
            End If
        Next c
CleanUp:

    End Sub
    Public Sub RefreshStatus()
        Dim c As Range
        Dim d As Range
        Dim dblEndCap As Double
        Dim dblRemCap As Double
        Dim EndDate As Date
        Dim dblEndTime As Double

        '   Application.EnableEvents = False
        '   Application.Calculation = xlCalculationManual

        For Each c In Application.Range("SapExlData").Columns(1).Cells

            If c.Value.ToString() = "" Then Exit For
            '      Debug.Print "Rad: " & c.Row

            If c.Row > 1 Then

                'Find current work center in the capacity sheet.
                d = Application.Sheets("WC_Cap").Columns(1).Find(c.Offset(0, 75).Value, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)

                '         If c.Row = 9 Then Stop

                'Find shift start and stop times.
                If Not d Is Nothing Then
                    If d.Offset(0, 2).Value = d.Offset(0, 3).Value Then
                        c.Offset(0, 74).Value = d.Offset(0, 2).Value / 3600 / 24
                        c.Offset(0, 76).Value = d.Offset(0, 3).Value / 3600 / 24
                        '               c.Offset(0, 68).Value = 0
                        '               c.Offset(0, 70).Value = 0
                    Else
                        c.Offset(0, 74).Value = d.Offset(0, 2).Value / 3600 / 24
                        c.Offset(0, 76).Value = d.Offset(0, 3).Value / 3600 / 24
                    End If

                    Application.Calculate()
                    If Not IsError(c.Offset(0, 77).Value) Then
                        dblEndCap = c.Offset(0, 77).Value 'Her fikk jeg feil...
                    End If
                    '            Debug.Print d.Row                  
                End If

                If Not d Is Nothing Then
                    If dblEndCap <= d.Offset(0, 7).Value Then
                        '               Debug.Print d.Row
                    Else
                        Do While dblEndCap > d.Offset(0, 7).Value
                            If d.Value <> c.Offset(0, 75).Value Then Exit Do
                            d = d.Offset(1, 0)
                            '                  Debug.Print d.Offset(0, 7).Value
                            '                  Debug.Print d.Row
                        Loop
                        '               Debug.Print d.Row
                        '               Debug.Print d.Offset(0, 7).Value
                    End If

                    '         If c.Row = 2 Then Stop

                    dblRemCap = d.Offset(0, 7).Value - dblEndCap
                    If d.Offset(0, 2).Value <> d.Offset(0, 3).Value Then
                        dblEndTime = d.Offset(0, 3).Value - dblRemCap
                        EndDate = d.Offset(0, 1).Value
                    Else
                        dblEndTime = 86400 + 25200 - dblRemCap
                        If dblEndTime < 86400 Then
                            EndDate = d.Offset(0, 1).Value
                        Else
                            '                  Debug.Print d.Offset(0, 1).Value
                            '                  Debug.Print d.Row

                            Do While d.Offset(1, 2).Value = 0
                                If d.Value <> c.Offset(0, 75).Value Then Exit Do
                                d = d.Offset(1, 0)
                                '                     Debug.Print d.Offset(0, 1).Value
                                '                     Debug.Print d.Row
                            Loop
                            EndDate = d.Offset(1, 1).Value

                        End If
                    End If

                    If dblEndTime > 86400 Then dblEndTime = dblEndTime - 86400

                    c.Offset(0, 70).Value = EndDate
                    c.Offset(0, 71).Value = dblEndTime / 3600 / 24
                    '            Debug.Print c.Row                   
                End If
            End If
        Next c

        Try
            Application.ActiveWorkbook.Sheets("Database").ListObjects("SapExlData").ListColumns("Work Center").DataBodyRange.FormulaR1C1 = "=RIGHT([Work ctr],LEN([Work ctr])-5)&"" ""&[Work Center Description]"
        Catch ex As Exception
        End Try
        '   Call PlanStatusUpdate

        '   Application.EnableEvents = True
        '   Calculate     

    End Sub

    Public Sub PlanStatusUpdate()
        Dim c As Range
        Dim s As Shape
        Dim sh As Worksheet
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim dblLength As Double
        Dim dblHeight As Double
        Dim dblStart As Double
        Dim dblEnd As Double
        Dim dblSecValue As Double
        Dim dblTop As Double
        Dim Day1 As Date
        Dim Day2 As Date
        Dim strMaterial As String
        Dim strQty As String
        Dim strPeriod As String
        Dim ActCell As Range
        Dim intTopOffset As Integer
        Dim intLength As Integer
        Dim intPlanColumn As Integer

        '   Application.ScreenUpdating = False
        '   Application.EnableEvents = False

        Application.Sheets("ProdPlan").Activate()
        ActCell = Application.Selection
        dblHeight = 12
        intLength = 29
        intPlanColumn = Application.Range("SapExlData").Columns.Count - 1

        On Error Resume Next

        For Each s In Application.ActiveSheet.Shapes
            Select Case s.Name
                Case "Time+"
                    s.Delete()
                Case "Time-"
                    s.Delete()
            End Select
        Next s

        x = 0
        For Each c In Application.Sheets("PlanData").Range("PlanData").Columns(1).Cells
            If c.Value.ToString() = "" Then Exit For
            '      Debug.Print "Rad: " & c.Row

            If c.Row = 1 Then GoTo ResumeHere

            If c.Offset(0, 42).Value <> c.Offset(-1, 42).Value And c.Offset(0, 42).Value <> "0" Then
                x = x + 1
            End If

            If c.Offset(0, intPlanColumn + 4).Value = 1 Then
                If c.Offset(0, intPlanColumn + 3).Value = "Pos" Then
                    dblStart = Application.ActiveSheet.Shapes("TimeLine").Left

                Else
                    dblStart = Application.ActiveSheet.Shapes("TimeLine").Left - intLength

                End If

                dblEnd = dblStart + intLength
                dblLength = dblEnd - dblStart
                strQty = Date.FromOADate(c.Offset(0, intPlanColumn + 2).Value).ToString("HH:mm")
                dblTop = Application.ActiveSheet.Range("StartWC").Offset(x, 0).Top + 15

                s = Application.ActiveSheet.Shapes.AddShape(Microsoft.Office.Core.MsoAutoShapeType.msoShapeRectangle, dblStart, dblTop, dblLength, dblHeight)
                s.DrawingObject.Characters.Text = strQty
                s.DrawingObject.Font.Name = "Arial"
                s.DrawingObject.Font.Size = 10
                s.DrawingObject.Font.ColorIndex = 1
                s.DrawingObject.HorizontalAlignment = Excel.Constants.xlLeft
                s.DrawingObject.ShapeRange.Line.Weight = 0.5
                s.DrawingObject.ShapeRange.TextFrame.MarginLeft = 1.5
                s.DrawingObject.ShapeRange.TextFrame.MarginRight = 0.5
                s.DrawingObject.ShapeRange.TextFrame.MarginTop = 2
                s.DrawingObject.ShapeRange.TextFrame.MarginBottom = 0.5
                s.DrawingObject.ShapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                s.DrawingObject.ShapeRange.Fill.Solid()
                s.DrawingObject.ShapeRange.Fill.Visible = Microsoft.Office.Core.MsoTriState.msoTrue
                s.DrawingObject.ShapeRange.Fill.Solid()
                If c.Offset(0, intPlanColumn + 3).Value = "Pos" Then
                    s.DrawingObject.Name = "Time+"
                    s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 50
                    s.DrawingObject.Font.ColorIndex = 0
                Else
                    s.DrawingObject.Name = "Time-"
                    s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 53
                    s.DrawingObject.Font.ColorIndex = 2
                End If
                s.DrawingObject.ShapeRange.Fill.Transparency = 0
                s.DrawingObject.ShapeRange.Line.ForeColor.SchemeColor = 55
            End If

            ActCell.Activate()

            On Error GoTo 0
            On Error Resume Next
            s = Application.ActiveSheet.Shapes("Order" & c.Value)
            If Err.Number = 0 Then
                '      Select Case c.Offset(0, intPlanColumn + 12).Value
                Select Case c.Offset(0, intPlanColumn + 4).Value
                    Case 0
                        s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 22
                    Case 1
                        If c.Offset(0, intPlanColumn + 13).Value = 1 Then
                            s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 29
                        Else
                            s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 52
                        End If
                    Case Else
                        s.DrawingObject.ShapeRange.Fill.ForeColor.SchemeColor = 52
                End Select
                ActCell.Activate()
            End If

ResumeHere:
        Next c

CleanUp:
        ActCell.Activate()

        '   Application.ScreenUpdating = True
        '   Application.EnableEvents = True

    End Sub
    Sub GetTime()
        Dim StartTime As Double
        Dim EndTime As Double
        Dim PlanTime As Double
        Dim ActTime As Double
        Dim ActCell
        Dim x As Integer
        Dim s As Shape
        Dim lngDate As Long

        On Error GoTo CleanUp
        ActCell = Application.Selection
        lngDate = Fix(DateTime.Now.Date.ToOADate())
        x = 0

        Select Case lngDate
            Case CDate(Application.Range("CapDate").Value).Date.ToOADate()
                x = 0
            Case CDate(Application.Range("CapDate").Offset(1, 0).Value).Date.ToOADate()
                x = 1
            Case CDate(Application.Range("CapDate").Offset(2, 0).Value).Date.ToOADate()
                x = 2
        End Select

        x = lngDate - CDate(Application.Range("CapDate").Value).Date.ToOADate()
        StartTime = Application.Range("StartTime").Offset(0, 24 * x + 7).Left
        EndTime = Application.Range("StartTime").Offset(0, 24 * x + 24 + 7).Left

        ActTime = StartTime + (EndTime - StartTime) * (CDate(DateTime.Now.TimeOfDay.ToString()).ToOADate() - 7 / 24)

        Application.ActiveSheet.Shapes("TimeLine").Left = ActTime
        Application.ActiveSheet.Shapes("TimeShape").Left = ActTime
        Application.ActiveSheet.Shapes("TimeShape").Select()
        Application.Selection.Characters.Text = DateTime.Now.ToString("HH:mm:ss")

        For Each s In Application.ActiveSheet.Shapes
            Select Case s.Name
                Case "Time+"
                    s.Left = ActTime
                Case "Time-"
                    s.Left = ActTime - s.Width
            End Select
        Next s

        StartTime = Application.Range("StartTime").Offset(0, 24 * 0 + 7).Left
        EndTime = Application.Range("StartTime").Offset(0, 24 * 0 + 24 + 7).Left
        PlanTime = Application.Range("PlanStart").Value

        ActTime = StartTime + (EndTime - StartTime) * (PlanTime - 7 / 24)

        Application.ActiveSheet.Shapes("TimePlan").Left = ActTime
        Application.ActiveSheet.Shapes("TimeStart").Left = ActTime
        Application.ActiveSheet.Shapes("TimeStart").Select()
        Application.Selection.Characters.Text = Date.FromOADate(PlanTime).ToString("HH:mm:ss")


        ActCell.Activate()
CleanUp:

    End Sub
    Public Sub GetPriorities()

        Dim c As Range
        Dim resultTable As System.Data.DataTable

        resultTable = OrklaRTBPL.ReportSpecific.GetPriorities(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate).Tables("PPPriorities")
        For Each dataRow In resultTable.Rows

            For Each c In Application.Sheets("ProdPlan").Range("WC_Plan").Cells
                If c.Value = dataRow("WorkCenter").ToString() Then
                    c.Interior.ColorIndex = Excel.Constants.xlNone
                    c.Interior.ColorIndex = dataRow("ColorIndex")
                End If
            Next
        Next

        Application.ActiveWorkbook.Activate()

CleanUp:

    End Sub
    Public Sub SavePriorities()

        Dim answ As Integer
        'Dim userId As Integer
        Dim c As Range

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False

        If OrklaRTBPL.ReportSpecific.GetPriorities(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate).Tables("PPPriorities").Rows.Count > 0 Then
            'userId = Convert.ToInt32(OrklaRTBPL.ReportSpecific.GetPriorities(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate).Tables("PPPriorities").Rows(0)("UserId"))
            answ = MsgBox("Work Center Priorities for active Plan already exists," & Chr(10) & _
                        "Do you want to overwrite?", vbYesNo, "Orkla SAP Integration")
            If answ = vbYes Then
                For Each c In Application.Sheets("ProdPlan").Range("WC_Plan").Cells
                    If Not c.Value Is Nothing Then
                        OrklaRTBPL.ReportSpecific.UpdateProductionPlanPriorities(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, c.Value, c.Interior.ColorIndex, DateTime.Now.ToString(), 0)
                    Else
                        Exit For
                        GoTo CleanUp
                    End If
                Next
            Else
                GoTo CleanUp
            End If
            'Else
            '    For Each c In Application.Sheets("ProdPlan").Range("WC_Plan").Cells
            '        If c.Interior.ColorIndex > 0 Then
            '            OrklaRTBPL.ReportSpecific.UpdateProductionPlanPriorities(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, c.Value, c.Interior.ColorIndex, DateTime.Now.ToString(), 0)
            '        End If
            '    Next
        End If

CleanUp:
        Application.DisplayAlerts = True
        Application.ScreenUpdating = True
        Application.EnableEvents = True



        Exit Sub
    End Sub

    Public Sub WriteLockedOrders(lngOrder As Long)
        Dim c As Range

        Try
            c = Application.Range("LockedTable").Columns(1).Find(lngOrder.ToString(), LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                OrklaRTBPL.ReportSpecific.DeleteLockedOrder(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, lngOrder.ToString(), gUserId)
            Else
                OrklaRTBPL.ReportSpecific.InsertLockedOrder(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, lngOrder.ToString(), gUserId)
            End If

            Globals.Ribbons.OrklaRT.GetLockedOrders()

            Call LockedUpdate()

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Fixed Production Plan - WriteLockedOrders", gUserId, gReportID)
        End Try

    End Sub

    Public Sub LockedUpdate()

        Application.StatusBar = "Retrieving Saved Plan Version..."
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        Call GetStartEndWC()
        Application.Calculate()

        Call RefreshStatus()
        Application.Calculate()

        Try
            If IsNothing(gwbReport) Then Exit Sub
            GC.Collect()
            GC.WaitForPendingFinalizers()
            gwbReport.Activate()
            Application.Sheets("Sequence").PivotTables(1).PivotCache.Refresh()
            Application.Sheets("Sequence").Activate()
        Catch ex As Exception
            boolErr = True
            GC.Collect()
            GC.WaitForPendingFinalizers()
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Fixed Production Plan - LockedUpdate", gUserId, gReportID)
        Finally
            If boolErr = True Then                
                LocalUpdate()                
                boolErr = False
            End If
        End Try
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.StatusBar = String.Empty

    End Sub

    Public Sub RefreshIntTables()

        Dim intRows As Integer
        Dim dblScale As Double

        Try

            Application.Sheets("Sequence").PivotTables(1).PivotCache.Refresh()
            Application.Sheets("Låst plan").PivotTables(1).PivotCache.Refresh()
            Application.Sheets("Deviations").PivotTables(1).PivotCache.Refresh()
            Application.Sheets("Mixing Status").PivotTables(1).PivotCache.Refresh()

            GoTo ContHere

            Application.Sheets("Deviations").Activate()
            Application.ActiveSheet.ChartObjects(1).Activate()
            Application.ActiveChart.SeriesCollection(1).Select()
            Application.ActiveChart.SeriesCollection(1).ApplyDataLabels(AutoText:=True, LegendKey:=False, ShowSeriesName:=False, ShowCategoryName:=False, ShowValue:=True, ShowPercentage:=False, ShowBubbleSize:=False)

            intRows = Application.Sheets("Deviations").PivotTables(1).RowRange.Rows.Count - 1
            dblScale = intRows / 21 * 267 + 10

            Application.ActiveSheet.Shapes(1).Select()
            Application.ActiveSheet.Shapes(1).Top = 126.3
            Application.ActiveChart.ChartArea.Select()
            Application.ActiveSheet.Shapes(1).Height = dblScale
            Application.ActiveChart.ChartArea.Select()
            Application.ActiveWindow.Visible = False

            'On Error Resume Next
            Application.ActiveSheet.Shapes(2).Select()
            Application.Selection.Top = Application.ActiveSheet.PivotTables(1).PageRange.Top + Application.ActiveSheet.PivotTables(1).PageRange.Height + 0 - Application.Selection.Height
            Application.Selection.Left = Application.ActiveSheet.PivotTables(1).PageRange.Left + Application.ActiveSheet.PivotTables(1).PageRange.Width + 20
            'On Error GoTo 0

ContHere:
            Application.Sheets("Mixing Status").Activate()
            Application.Sheets("Mixing Status").PivotTables(1).PivotSelect("Mix_Status[All]", XlPTSelectionMode.xlLabelOnly, True)
            Application.Selection.NumberFormat = "#,##0 %;[Red]-#,##0 %"
            Application.Sheets("Mixing Status").Cells(12, 1).Select()
            'bolNoPlan = False
            Application.Sheets("Sequence").Activate()
            Application.Sheets("Sequence").PivotTables(1).PivotSelect("Mix_Status[All]", XlPTSelectionMode.xlLabelOnly, True)
            Application.Selection.NumberFormat = "#,##0 %;[Red]-#,##0 %"
            Application.Sheets("Sequence").Cells(12, 1).Select()
            Application.Sheets("Låst plan").Activate()
            Application.Sheets("Låst plan").PivotTables(1).PivotSelect("Mixed[All]", XlPTSelectionMode.xlLabelOnly, True)
            Application.Selection.NumberFormat = "#,##0 %;[Red]-#,##0 %"
            Application.Sheets("Låst plan").Cells(12, 1).Select()
            Application.Sheets("ProdPlan").Activate()


        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Fixed Production Plan - LockedUpdate", gUserId, gReportID)
        End Try

    End Sub
    Sub CreateNewPlan()
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        Try
            Application.OnKey("{F8}", String.Empty)
            Application.StatusBar = "Refreshing Production Plan Status..."
            Application.Calculate()
            Call GetStartEndWC()   'Get capacity data.

            Application.Calculate()
            Call RefreshStatus()

            Application.StatusBar = "Checking user Credibilities..."
            Application.Calculate()
            BuildNewPlan()

            If bolOnlyRefresh = True Then GoTo CleanUp

            Call GetExistingPlan()

            Application.StatusBar = "Creating Production Plan..."
            Application.Calculate()
            Call CreatePlanShapes()   'Including brakes.

            Application.Calculate()
            Call PlanStatusUpdate()   'Updating time line deviations.

            Application.Calculate()
            Call GetTime()

            Call GetPriorities()
CleanUp:
            Call RefreshIntTables()
            bolOnlyRefresh = False
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Fixed Production Plan - CreateNewPlan", gUserId, gReportID)
        End Try

        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.EnableEvents = True
        Application.StatusBar = String.Empty
        Application.ScreenUpdating = True

        Exit Sub

    End Sub

    Public Sub BuildNewPlan()
        Dim answ As Integer
        Dim planTable As New System.Data.DataTable
        Dim dataTable As New System.Data.DataTable

        Application.DisplayAlerts = False

        Try

            If OrklaRTBPL.ReportSpecific.CheckPlanExists(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate).Tables("ProductionPlanData").Rows.Count > 0 Then
                answ = MsgBox("A Production Plan for your selections already exists," & Chr(10) & _
                              "created by " & OrklaRTBPL.CommonFacade.GetUserName(OrklaRTBPL.ReportSpecific.GetPlannerId(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate)) & "." & Chr(10) & _
                              "Do you want to overwrite?", vbYesNo, "Orkla SAP Integration")
                If answ = vbYes Then
                    OrklaRTBPL.ReportSpecific.DeleteProductionPlanData(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant, OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate)
                    'Do Nothing
                Else
                    bolOnlyRefresh = True
                    GoTo CleanUp
                End If
            End If

            Application.StatusBar = "Saving Plan Version..."
            planTable = OrklaRTBPL.ReportSpecific.ClonePlanData().Tables(0).Clone()
            dataTable = mc_ExcelTableToDataTable("Database", "SapExlData")

            Try
                For i = 0 To dataTable.Rows.Count - 1
                    planTable.Rows.Add(i)
                    For j = 0 To dataTable.Columns.Count - 1
                        planTable.Rows(i)(1) = gUserId
                        planTable.Rows(i)(2) = DateTime.Now
                        If Not dataTable.Rows(i)(3).Equals(String.Empty) Then
                            If dataTable.Columns(j).DataType.FullName <> planTable.Columns(j + 3).DataType.FullName Then
                                If Not IsDBNull(dataTable.Rows(i)(j)) Then
                                    'If dataTable.Columns(j).DataType.Name.Equals("DateTime") Then
                                    '    If dataTable.Rows(i)(j) = Convert.ToDateTime("31.12.9999").Date Then
                                    '        planTable.Rows(i)(j + 3) = System.Convert.ChangeType(0, Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                    '    Else
                                    '        planTable.Rows(i)(j + 3) = System.Convert.ChangeType(dataTable.Rows(i)(j), Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                    '    End If
                                    'Else
                                    If planTable.Columns(j + 3).DataType.Name.Equals("DateTime") Then
                                        If IsNumeric(dataTable.Rows(i)(j)) Then
                                            If Not dataTable.Rows(i)(j).Equals("0") Then
                                                planTable.Rows(i)(j + 3) = System.Convert.ChangeType(DateTime.FromOADate(dataTable.Rows(i)(j)), Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                            Else
                                                planTable.Rows(i)(j + 3) = DBNull.Value
                                            End If
                                        Else
                                            If Not dataTable.Rows(i)(j).ToString().Equals("Ikke tilordnet") And Not dataTable.Rows(i)(j).ToString().Equals(String.Empty) Then
                                                planTable.Rows(i)(j + 3) = System.Convert.ChangeType(Convert.ToDateTime(dataTable.Rows(i)(j)).Date, Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                            Else
                                                planTable.Rows(i)(j + 3) = DBNull.Value
                                            End If
                                        End If
                                    ElseIf planTable.Columns(j + 3).DataType.Name.Equals("TimeSpan") Then
                                        If IsNumeric(dataTable.Rows(i)(j)) Then
                                            planTable.Rows(i)(j + 3) = System.Convert.ChangeType(New TimeSpan(TimeSpan.FromDays(dataTable.Rows(i)(j)).Hours, TimeSpan.FromDays(dataTable.Rows(i)(j)).Minutes, TimeSpan.FromDays(dataTable.Rows(i)(j)).Seconds), Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                        Else
                                            If Not dataTable.Rows(i)(j).ToString().Equals("Ikke tilordnet") Then
                                                If Convert.ToDateTime(dataTable.Rows(i)(j)).Date = Convert.ToDateTime("30.12.1899").Date Then
                                                    planTable.Rows(i)(j + 3) = System.Convert.ChangeType(New TimeSpan(CInt(dataTable.Rows(i)(j).ToString().Substring(11, 2)), CInt(dataTable.Rows(i)(j).ToString().Substring(14, 2)), CInt(dataTable.Rows(i)(j).ToString().Substring(17, 2))), Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                                Else
                                                    planTable.Rows(i)(j + 3) = System.Convert.ChangeType(New TimeSpan(CInt(dataTable.Rows(i)(j).ToString().Substring(11, 2)), CInt(dataTable.Rows(i)(j).ToString().Substring(14, 2)), CInt(dataTable.Rows(i)(j).ToString().Substring(17, 2))), Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                                End If
                                                'Else
                                                '    planTable.Rows(i)(j + 3) = DBNull.Value
                                            End If
                                        End If
                                    ElseIf planTable.Columns(j + 3).DataType.Name.Contains("Int") Then
                                        If IsNumeric(dataTable.Rows(i)(j)) Then
                                            planTable.Rows(i)(j + 3) = System.Convert.ChangeType(Convert.ToInt32(dataTable.Rows(i)(j)), Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                        Else
                                            planTable.Rows(i)(j + 3) = DBNull.Value
                                        End If
                                    Else
                                        If Not dataTable.Rows(i)(j).ToString().Equals("30.12.1899 00:00:00") Then
                                            planTable.Rows(i)(j + 3) = System.Convert.ChangeType(dataTable.Rows(i)(j), Type.GetType(planTable.Columns(j + 3).DataType.FullName))
                                        Else
                                            planTable.Rows(i)(j + 3) = DBNull.Value
                                        End If
                                    End If
                                    'End If
                                Else
                                    planTable.Rows(i)(j + 3) = DBNull.Value
                                End If
                            Else
                                planTable.Rows(i)(j + 3) = dataTable.Rows(i)(j)
                            End If
                        End If
                    Next
                Next
            Catch ex As Exception
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "FixedProductionPlan", gUserId, gReportID)
            End Try

            Using bulkCopy As SqlBulkCopy = New SqlBulkCopy(System.Configuration.ConfigurationManager.AppSettings("ConnectionString").ToString())
                bulkCopy.DestinationTableName = "dbo.ProductionPlanData"

                Try
                    bulkCopy.WriteToServer(planTable)
                Catch ex As Exception
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "FixedProductionPlan", gUserId, gReportID)
                End Try
            End Using

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "FixedProductionPlan", gUserId, gReportID)
        End Try


CleanUp:
        Application.DisplayAlerts = True
    End Sub



    Public Function mc_ExcelTableToDataTable(ByVal SheetName As String, _
                                          ByVal TableName As String, _
                                          Optional ByVal FilePath As String = "", _
                                          Optional ByVal SQLsentence As String = "") As System.Data.DataTable


        Dim vRange As String = Application.Sheets(SheetName).ListObjects(TableName).Range.AddressLocal
        vRange = vRange.Replace("$", "")

        Dim vCNNstring As String = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                    "Data Source= " & Application.ActiveWorkbook.FullName & ";" & _
                                     "Extended Properties=""Excel 12.0 Macro;HDR=YES;IMEX=1"""

        Dim ExcelCNN As System.Data.OleDb.OleDbConnection = New System.Data.OleDb.OleDbConnection(vCNNstring)

        Try
            ExcelCNN.Open()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "FixedProductionPlan", gUserId, gReportID)
        End Try


        Dim vSQL As String = IIf(SQLsentence = "", _
                                 "SELECT * FROM [" + SheetName + "$" & vRange & "]", _
                                 SQLsentence)
        Dim ExcelCMD As System.Data.OleDb.OleDbDataAdapter = New System.Data.OleDb.OleDbDataAdapter(vSQL, ExcelCNN)
        Dim ExcelDS As System.Data.DataSet = New System.Data.DataSet()
        ExcelCMD.TableMappings.Add("Table", "SapExlData")

        Try
            ExcelCMD.Fill(ExcelDS)
            mc_ExcelTableToDataTable = ExcelDS.Tables(0).Copy()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "FixedProductionPlan", gUserId, gReportID)
        End Try

        'Application.DefaultFilePath = excelDefaultPath

        ExcelCMD = Nothing
        ExcelCNN.Close()
        ExcelCNN.Dispose()


        Return mc_ExcelTableToDataTable
    End Function
End Module
