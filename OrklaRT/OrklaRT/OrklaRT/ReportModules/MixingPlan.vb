Module MixingPlan
    Public strEditComm As String
    Sub LocalUpdate()
        'Dim c As Excel.Range
        'Dim x As Integer

        Call RefreshProdData()
        Call RefreshMixWC()
        Call RefreshMixPlan()
        Call RefreshPriPlan()
        Call RefreshRSTest()
        Call RefreshNewStart()
        Call RefreshProdStatus()

        'Call RefreshComments(ThisWorkbook, "MixingPlanComments.txt")

        Application.Sheets("MaterialPlan").Activate()
        Application.Sheets("MaterialPlan").PivotTables(1).PivotSelect("Stock[All]", Excel.XlPTSelectionMode.xlLabelOnly, True)
        Application.Selection.NumberFormat = "#,##0;[Red]-#,##0"
        'Application.Sheets("MaterialPlan").PivotTables(1).PivotSelect("Rem_Stock[All]", Excel.XlPTSelectionMode.xlLabelOnly, True)
        'Application.Selection.NumberFormat = "#,##0;[Red]-#,##0"
        Try
            Application.Range("E17").Select()
            Application.Sheets("BlandePlan").Activate()
            Application.Sheets("BlandePlan").PivotTables(1).PivotSelect("Buffer[All]", Excel.XlPTSelectionMode.xlLabelOnly, True)
            Application.Selection.NumberFormat = "#,##0;[Red]-#,##0"
            Application.Sheets("BlandePlan").PivotTables(1).PivotSelect("Tid[All]", Excel.XlPTSelectionMode.xlLabelOnly, True)
            Application.Selection.NumberFormat = "hh:mm"
            Application.Range("A14").Select()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "MixingPlan - LocalUpdate", gUserId, gReportID)
        End Try

        If Application.Version <> "12.0" Then
            '   Sheets("MixingPlan").PivotTables(1).PivotFields("Mach").ShowAllItems = True
            '   Sheets("MaterialPlan").PivotTables(1).PivotFields("Shift").ShowAllItems = True
        Else
            Application.Sheets("BlandePlan").PivotTables(1).showdrillindicators = False
        End If

CleanUp:
        Exit Sub

    End Sub

    Sub RefreshProdData()

        Dim resultTable As New System.Data.DataTable

        Application.Application.DisplayAlerts = False

        resultTable = OrklaRTBPL.ReportSpecific.GetProdPlanData(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, OrklaRTBPL.SelectionFacade.MixingPlanProdPlanSelectionDate).Tables("ProductionPlanData")

        Call Common.LoadListObjectData("ProductionPlanData", "ProdPlan", "tProdPlanAll", resultTable)

        'Dim rs = OrklaRTBPL.CommonFacade.ConvertToRecordset(resultTable)

        'Application.Sheets("ProdPLan").Range("tProdPlanAll").Clear()
        'rs.MoveFirst()
        'Application.Sheets("ProdPLan").Range("tProdPlanAll").CopyFromRecordset(rs)

        Application.Application.DisplayAlerts = True

    End Sub

    Sub RefreshProdStatus()

        Dim resultTable As New System.Data.DataTable

        Application.Application.DisplayAlerts = False


        resultTable = OrklaRTBPL.ReportSpecific.GetMPProdStatus(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant).Tables("MPProdStatus")

        Call Common.LoadListObjectData("ProductionPlanData", "ProdStatus", "tProdStatus", resultTable)
        'Dim rs = OrklaRTBPL.CommonFacade.ConvertToRecordset(resultTable)

        'Application.Sheets("ProdStatus").Range("tProdStatus").Clear()
        'rs.MoveFirst()
        'Application.Sheets("ProdStatus").Range("tProdStatus").CopyFromRecordset(rs)

        Application.Application.DisplayAlerts = True

    End Sub

    Sub WriteNewStart()
        Dim x As Integer
        Dim r As Excel.Range

        r = Application.Sheets("OrderStart").Cells(1, 1)
        For x = 1 To 5000
            If r.Offset(x, 0).Value <> Nothing Then
                OrklaRTBPL.ReportSpecific.InsertMPOrderStart(r.Offset(x, 0).Value.ToString(), r.Offset(x, 1).Value.ToString(), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
                'Globals.Ribbons.OrklaRT.GetLockedOrders()
            Else
                Exit For
            End If
        Next x
        RefreshNewStart()
    End Sub

    Sub WriteMixWC()

        Dim x As Integer
        Dim r As Excel.Range
        Dim y As Integer

        OrklaRTBPL.ReportSpecific.DeleteMPMixWC(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)

        y = 0
        r = Application.Sheets("Oppsett").Range("MixWC1").Cells(1, 1)
        For x = 1 To 5000
            If r.Offset(x, 0).Value <> "" Then
                OrklaRTBPL.ReportSpecific.InsertMPMixWC(r.Offset(x, 0).Value, OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
            Else
                y = y + 1
                If y = 10 Then Exit For
            End If
        Next x

    End Sub

    Sub WriteMixPlan()

        Dim x As Integer
        Dim r As Excel.Range

        r = Application.Sheets("MixPlan").Cells(1, 1)
        For x = 1 To 5000
            If r.Offset(x, 0).Value <> Nothing Then
                OrklaRTBPL.ReportSpecific.InsertMPMixPlan(r.Offset(x, 0).Value.ToString(), Convert.ToInt32(r.Offset(x, 1).Value), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
            Else
                Exit For
            End If
        Next x

    End Sub

    Sub WritePriPlan()

        Dim x As Integer
        Dim r As Excel.Range

        r = Application.Sheets("PriPlan").Cells(1, 1)
        For x = 1 To 5000
            If r.Offset(x, 0).Value <> Nothing Then
                OrklaRTBPL.ReportSpecific.InsertMPPriPlan(r.Offset(x, 0).Value.ToString(), Convert.ToInt32(r.Offset(x, 1).Value), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
            Else
                Exit For
            End If
        Next x

    End Sub

    Sub WriteRSTest()

        Dim x As Integer
        Dim r As Excel.Range

        r = Application.Sheets("RSTest").Cells(1, 1)
        For x = 1 To 5000
            If r.Offset(x, 0).Value <> "" Then
                OrklaRTBPL.ReportSpecific.InsertMPRSTest(r.Offset(x, 0).Value.ToString(), r.Offset(x, 1).Value.ToString(), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
            Else
                Exit For
            End If
        Next x
    End Sub

    Sub RefreshNewStart()
        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("OrderStart").ListObjects
            If listObject.Name.Equals("tOrderStart") Then
                Try
                    If Not listObject.DataBodyRange Is Nothing Then
                        listObject.DataBodyRange.Delete()
                    End If
                    Dim orderStart = OrklaRTBPL.ReportSpecific.GetMPOrderStart(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(orderStart.Tables(0))
                    data.MoveFirst()
                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, orderStart.Tables(0).Rows.Count, orderStart.Tables(0).Columns.Count)
                Catch
                End Try
            End If
        Next
    End Sub
    Sub RefreshMixWC()

        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("Oppsett").ListObjects
            If listObject.Name.Equals("tMixWC") Then
                Try
                    If Not listObject.DataBodyRange Is Nothing Then
                        listObject.DataBodyRange.Delete()
                    End If
                    Dim mixWC = OrklaRTBPL.ReportSpecific.GetMPMixWC(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(mixWC.Tables(0))
                    data.MoveFirst()
                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, mixWC.Tables(0).Rows.Count, mixWC.Tables(0).Columns.Count)
                Catch
                End Try
            End If
        Next

    End Sub


    Sub RefreshMixPlan()

        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("MixPlan").ListObjects
            If listObject.Name.Equals("tMixPlan") Then
                Try
                    If Not listObject.DataBodyRange Is Nothing Then
                        listObject.DataBodyRange.Delete()
                    End If
                    Dim mixPlan = OrklaRTBPL.ReportSpecific.GetMPMixPlan(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(mixPlan.Tables(0))
                    data.MoveFirst()
                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, mixPlan.Tables(0).Rows.Count, mixPlan.Tables(0).Columns.Count)
                Catch
                End Try
            End If
        Next

    End Sub

    Sub RefreshPriPlan()

        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("PriPlan").ListObjects
            If listObject.Name.Equals("tPriPlan1") Then
                Try
                    If Not listObject.DataBodyRange Is Nothing Then
                        listObject.DataBodyRange.Delete()
                    End If
                    Dim priPlan = OrklaRTBPL.ReportSpecific.GetMPPriPlan(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
                    If priPlan.Tables(0).Rows.Count > 0 Then
                        Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(priPlan.Tables(0))
                        data.MoveFirst()
                        Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, priPlan.Tables(0).Rows.Count, priPlan.Tables(0).Columns.Count)
                    End If
                Catch
                End Try
            End If
        Next

    End Sub

    Sub RefreshRSTest()

        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("RSTest").ListObjects
            If listObject.Name.Equals("tRSTest") Then
                Try
                    If Not listObject.DataBodyRange Is Nothing Then
                        listObject.DataBodyRange.Delete()
                    End If
                    Dim rsTest = OrklaRTBPL.ReportSpecific.GetMPMixPlan(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(rsTest.Tables(0))
                    data.MoveFirst()
                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, rsTest.Tables(0).Rows.Count, rsTest.Tables(0).Columns.Count)
                Catch
                End Try
            End If
        Next

    End Sub

    Public Sub FindOrder(lngOrder As Long, intMach As Integer)
        Dim c As Excel.Range

        Call RefreshMixPlan()

        With Application.Sheets("MixPlan")
            c = .UsedRange.Columns(1).Find(lngOrder, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                c.Offset(0, 1).Value = intMach
            Else
                .Cells(2, 1).EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown)
                .Cells(2, 1).Value = lngOrder
                .Cells(2, 2).Value = intMach
            End If
        End With

        Call WriteMixPlan()
        Call RefreshMixPlan()
        'Application.Sheets(Application.Sheets("Version").Range("FirstSheet").Value).PivotTables(1).PivotCache.Refresh()

CleanUp:

    End Sub


    Sub FindPri(lngOrder As Long, intPri As Integer)

        Dim c As Excel.Range

        Call RefreshPriPlan()

        With Application.Sheets("PriPlan")
            c = .UsedRange.Columns(1).Find(lngOrder, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                c.Offset(0, 1).Value = intPri
            Else
                .Cells(2, 1).EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown)
                .Cells(2, 1).Value = lngOrder
                .Cells(2, 2).Value = intPri
                'Application.Sheets("PriPlan").Cells(Application.Sheets("PriPlan").Range("PriPlan1").Rows.Count, 1).Value = lngOrder
                'Application.Sheets("PriPlan").Cells(Application.Sheets("PriPlan").Range("PriPlan1").Rows.Count, 2).Value = intPri
            End If
        End With

        Call WritePriPlan()
        Call RefreshPriPlan()
        '   ThisWorkbook.Sheets(fnfirstsheet(ThisWorkbook)).PivotTables(1).PivotCache.Refresh

CleanUp:

    End Sub


    Sub FindRS(lngOrder As Long, strRS As Object)
        Dim c As Excel.Range

        Call RefreshRSTest()

        c = Application.Sheets("RSTest").Columns(1).Find(lngOrder, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
        If Not c Is Nothing Then
            c.Offset(0, 1).Value = strRS
        Else
            Application.Sheets("RSTest").Cells(Application.Sheets("RSTest").Range("RSTest").Rows.Count + 1, 1).Value = lngOrder
            Application.Sheets("RSTest").Cells(Application.Sheets("RSTest").Range("RSTest").Rows.Count + 1, 2).Value = strRS
        End If

        Call WriteRSTest()
        Call RefreshRSTest()

CleanUp:

    End Sub


    Sub FindNewStart(lngOrder As Long, lngDate As Date)
        Dim c As Excel.Range

        '  Call RefreshNewStart

        With Application.Sheets("OrderStart")
            c = .UsedRange.Columns(1).Find(lngOrder, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                If strEditComm = "Delete" Then
                    c.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
                    OrklaRTBPL.ReportSpecific.DeleteMPOrderStart(lngOrder, lngDate.ToString(), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
                Else
                    c.Offset(0, 1).Value = CDate(lngDate)
                End If
                strEditComm = String.Empty
            Else
                .Cells(2, 1).EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown)
                .Cells(2, 1).Value = CLng(lngOrder)
                .Cells(2, 2).Value = CDate(lngDate)
            End If
        End With

        'Call WriteNewStart()
        'Call RefreshNewStart()
        '   ThisWorkbook.Sheets(fnfirstsheet(ThisWorkbook)).PivotTables(1).PivotCache.Refresh

CleanUp:

    End Sub

    Sub MixingPlanSheetChangeStart(Target As Excel.Range)
        Dim d As Integer
        Dim e As Integer
        Dim f As Integer
        Dim g As Integer
        Dim x As Integer
        Dim lngOrder As Long
        Dim lngDate As Date
        Dim intHour As Integer
        Dim c As Excel.Range
        Dim sngTime As Date
        Dim mixingStartForm As MixingStart

        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Try
            If Target.Text <> String.Empty Then
                sngTime = CDate(Target.Text)
            Else
                sngTime = CDate("00:00")
            End If            

            d = Target.PivotTable.PivotFields("Navn blanding").LabelRange.Column
            e = Target.PivotTable.PivotFields("Start_Date").LabelRange.Column
            f = Target.PivotTable.PivotFields("Tid").LabelRange.Column
            g = Target.PivotTable.PivotFields("Ordre").LabelRange.Column

            lngOrder = Application.Cells(Target.Row, g).Value
            Dim z = Application.Cells(Target.Row, e).PivotCell.PivotItem.Value.ToString().Split("/")
            lngDate = New Date(CInt(z(2)), CInt(z(0)), CInt(z(1)))
            sngTime = Application.Cells(Target.Row, f).PivotCell.PivotItem.Value

            If Not Application.Cells(Target.Row, d).Value Is Nothing Then
                mixingStartForm = New MixingStart(lngOrder, Application.Cells(Target.Row, d).Value.ToString(), lngDate, Format("hh:mm", sngTime))
            Else                
                mixingStartForm = New MixingStart(lngOrder, String.Empty, lngDate, Format("hh:mm", sngTime))
            End If

            Call mixingStartForm.ShowDialog()

            lngDate = DateValue(mixingStartForm.dtpDate.Value)
            sngTime = mixingStartForm.txtStartTime.Text
            lngDate = CDate(lngDate + " " + sngTime)

            If strEditComm = "Cancel" Then GoTo CleanUp

            Call FindNewStart(lngOrder, lngDate)

            Call WriteNewStart()
            Call RefreshNewStart()

            strEditComm = String.Empty
            'Call FindNewStart(lngOrder, lngDate)
            '   ThisWorkbook.Sheets("Comments").Cells.ClearContents
            '   Call RefreshCommentsNew            
            'Debug.Print lngOrder; lngDate; intHour
CleanUp:
            d = Nothing
            e = Nothing
            f = Nothing
            g = Nothing
            Application.EnableEvents = True
            Application.ScreenUpdating = True


        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "MixingPlan", gUserId, gReportID)
        End Try

      
    End Sub


    Sub MixingPlanSheetCalculateTime(intMach1 As Integer, lngDate1 As Date)
        Dim c As Excel.Range
        Dim d As Excel.Range
        Dim e As Excel.Range
        Dim f As Excel.Range
        Dim accTime As Double
        Dim lngOrder As Long
        Dim g As Excel.Range
        Dim lngDate As Date
        Dim intMach As Integer
        Dim x As Integer

      
        Try

            '   Set c = ActiveSheet.PivotTables(1).PivotFields("Start").LabelRange
            d = Application.ActiveSheet.PivotTables(1).PivotFields("Pri").LabelRange
            e = Application.ActiveSheet.PivotTables(1).PivotFields("Mach").LabelRange
            f = Application.Cells(e.Row, Application.ActiveSheet.PivotTables(1).PivotFields("Blandetid(t)").LabelRange.Column)
            g = Application.ActiveSheet.PivotTables(1).PivotFields("Ordre").LabelRange


            x = 0
            accTime = Nothing
            For Each c In Application.ActiveSheet.PivotTables(1).PivotFields("Start_Date").DataRange.Cells
                x = x + 1
                'Debug.Print c.Address
                If c.PivotCell.PivotCellType < 2 Then
                    '         If c.Row = 22 Then Stop
                    If e.Offset(x, 0).Value > 0 Then
                        Dim dato = c.PivotItem.Value.ToString().Split("/")
                        lngDate = New Date(CInt(dato(2)), CInt(dato(0)), CInt(dato(1)))
                        If e.Offset(x, 0).PivotItem.Value = intMach1 And lngDate = lngDate1 Then
                            intMach = e.Offset(x, 0).PivotItem.Value
                            lngOrder = g.Offset(x, 0).PivotItem.Value
                            'lngDate = DateValue(c.PivotItem.Value)
                            If lngDate = DateTime.Now.Date Then
                                lngDate = Date.FromOADate(lngDate.Date.ToOADate() + Application.Range("StartHour").Value)
                            Else
                                lngDate = Date.FromOADate(lngDate.Date.ToOADate() + Application.Range("WorkStart").Value)
                            End If
                            Call FindNewStart(lngOrder, lngDate)
                            accTime = Application.Range("StartHour").Value
                        End If
                    Else
                        Dim dato = c.PivotItem.Value.ToString().Split("/")
                        lngDate = New Date(CInt(dato(2)), CInt(dato(0)), CInt(dato(1)))
                        If e.Offset(x, 0).PivotItem.Value = intMach1 And lngDate = lngDate1 Then
                            If e.Offset(x, 0).PivotItem.Value = intMach And d.Offset(x, 0).PivotItem.Value < 999 Then
                                accTime = accTime + f.Offset(x - 1, 0).Value / 24
                                lngOrder = g.Offset(x, 0).PivotItem.Value
                                lngDate = Date.FromOADate(lngDate.Date.ToOADate() + accTime)
                                Call FindNewStart(lngOrder, lngDate)
                            End If
                        End If
                    End If
                    'Debug.Print x & " - " & Format(accTime, "hh:mm")
                    '         d.Offset(x, 0).Select
                    '         Debug.Print "X = " & x
                End If
            Next c

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "MixingPlan", gUserId, gReportID)
        End Try

        d = Nothing
        e = Nothing
        f = Nothing
        g = Nothing
        Application.EnableEvents = True
        Application.ScreenUpdating = True

    End Sub
    Sub MixingPlanAllWorkCentersSheetChangeStart(Target As Excel.Range)
        Dim d As Integer
        Dim e As Integer
        Dim f As Integer
        Dim g As Integer
        Dim x As Integer
        Dim lngOrder As Long
        Dim lngDate As Date
        Dim intHour As Integer
        Dim c As Excel.Range

        Application.EnableEvents = False
        Application.ScreenUpdating = False

        Try

            intHour = Target.Value

            d = Target.PivotTable.PivotFields("Material Name Mix").LabelRange.Column
            e = Target.PivotTable.PivotFields("Start_Mix").LabelRange.Column
            f = Target.PivotTable.PivotFields("Hour").LabelRange.Column
            g = Target.PivotTable.PivotFields("Ordre").LabelRange.Column

            lngOrder = Application.Cells(Target.Row, g).Value
            lngDate = Application.Cells(Target.Row, e).PivotCell.PivotItem.Value
            intHour = Application.Cells(Target.Row, f).PivotCell.PivotItem.Value

            'frmMixingStart.labMaterialName.Caption = Application.Cells(Target.Row, d).Value
            'frmMixingStart.tbxDate.Text = String.Format("dd.mm.yyyy", lngDate)
            'frmMixingStart.tbxHour.Text = String.Format("00", intHour)

            Dim misingStartForm As New MixingStart(lngOrder, Application.Cells(Target.Row, d).ValueToString(), lngDate, String.Format("00", intHour))
            Call misingStartForm.Show()

            'lngDate = DateValue(frmMixingStart.tbxDate.Text)
            'intHour = frmMixingStart.tbxHour.Text

            'If strEditComm = "cancel" Then GoTo CleanUp

            'Call MixingPlan.FindNewStart(lngOrder, lngDate)
            'Call MixingPlan.FindNewStart(lngOrder, lngDate, intHour)
            '   ThisWorkbook.Sheets("Comments").Cells.ClearContents
            '   Call RefreshCommentsNew

            'Debug.Print lngOrder; lngDate; intHour

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "MixingPlan", gUserId, gReportID)
        End Try

        d = Nothing
        e = Nothing
        f = Nothing
        g = Nothing
        Application.EnableEvents = True
        Application.ScreenUpdating = True

    End Sub

    Sub resall()
        '   MsgBox ActiveCell.PivotCell.PivotCellType

        'ActiveSheet.PivotTables(1).PivotFields("Start_Date").DataRange.Select
        'ActiveSheet.PivotTables(1).PivotFields("Mach").LabelRange.Select
        Application.ActiveSheet.PivotTables(1).PivotFields("Blandetid(t)").LabelRange.Select()
        On Error GoTo 0
        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End Sub

    '    Sub LocalAutoOpen()
    '        Dim intDays As Integer   'Number of days for backward order evaluaton.

    '        intDays = 2

    '        Application.Range("F_Date").Value = DateTime.Now.Date.AddDays(-30).Date
    '        Application.Range("T_Date").Value = DateTime.Now.Date.AddDays(14).Date

    '        If DateTime.Now.TimeOfDay.TotalHours < (10 / 24) Then
    '        Range("PlanDate").Value = Date - 1
    '            Select Case Weekday(Of Date, vbMonday)()
    '                Case 1
    '              Range("PlanDate").Value = Date - 3
    '                Case 2, 3, 4, 5
    '              Range("PlanDate").Value = Date - 1
    '                Case 6
    '              Range("PlanDate").Value = Date - 1
    '                Case 7
    '              Range("PlanDate").Value = Date - 2
    '            End Select
    '        Else
    '        Range("PlanDate").Value = Date
    '        End If

    'CleanUp:

    '        Exit Sub

    '    End Sub
    '    '—————————————————————————————————————————————————————————————————————————————
    '    Sub SaveMultiSelect()
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        '  Scope   :
    '        '  Author  : Bjørn Tømmerbakk.
    '        '  Date    : 28.01.2009.
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        '  Local constants and variabel declarations:
    '        Dim strFileName As String
    '        Dim strField As String
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        strFileName = Sheets("Selections").Range("Home").Value & "MSMixingPlan"
    '        strField = Sheets("MultiSelect").Range("Multiselect").Cells(1, 1).Offset(-1, 0).Value

    '        Call WriteMultiSelect(ThisWorkbook, strFileName, strField)
    '        '  ———————————————————————————————————————————————————————————————————————————
    '    End Sub
    '    '—————————————————————————————————————————————————————————————————————————————


    '    '—————————————————————————————————————————————————————————————————————————————
    '    Sub RefreshMultiSelect()
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        '  Scope   :
    '        '  Author  : Bjørn Tømmerbakk.
    '        '  Date    : 28.01.2009.
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        '  Local constants and variabel declarations:
    '        Dim strPath As String
    '        '  ———————————————————————————————————————————————————————————————————————————

    '        On Error GoTo CleanUp
    '        strPath = strSystemPath & "SavedReports\" & Sheets("Selections").Range("Home").Value

    '        With ThisWorkbook.Sheets("MultiSelect").QueryTables(1)
    '            .Connection = "TEXT;" & strPath & "MSMixingPlan.txt"
    '            .TextFilePlatform = 65000
    '            .TextFileStartRow = 1
    '            .TextFileParseType = xlDelimited
    '            .TextFileTextQualifier = xlTextQualifierDoubleQuote
    '            .TextFileConsecutiveDelimiter = False
    '            .TextFileTabDelimiter = True
    '            .TextFileSemicolonDelimiter = False
    '            .TextFileCommaDelimiter = True
    '            .TextFileSpaceDelimiter = False
    '            .TextFileColumnDataTypes = Array(1, 1, 1)
    '            .TextFileTrailingMinusNumbers = True
    '            '      .Refresh BackgroundQuery:=False
    '        End With
    '        Call RefreshQueryTable("MultiSelect", ThisWorkbook)

    'CleanUp:
    '        '  ———————————————————————————————————————————————————————————————————————————
    '    End Sub
    '    '—————————————————————————————————————————————————————————————————————————————


    '    Sub reset12()
    '        Application.EnableEvents = True
    '    End Sub


    '    Sub testpivot()

    '        Debug.Print ActiveSheet.PivotTables(1).PivotFields("Open_Qty_Mix").Value

    '    End Sub


End Module
