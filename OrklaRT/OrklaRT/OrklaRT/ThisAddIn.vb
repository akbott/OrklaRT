Imports System.Data
Imports MouseKeyboardActivityMonitor
Imports MouseKeyboardActivityMonitor.WinApi
Imports Microsoft.VisualStudio.Tools.Applications.Runtime
Imports System.Runtime.InteropServices
Imports System.IO
Imports System.IO.Packaging
Imports System.Reflection
Imports BPL
Imports SAP.Middleware.Connector
Imports System.Windows.Threading
Imports System.Windows.Forms
Imports System.ComponentModel
Imports System.Xml

Public Class ThisAddIn
    Dim keyBoardListener As KeyboardHookListener
    Private comCalls As COMCalls
    Private Sub ThisAddIn_Startup() Handles Me.Startup
        keyBoardListener = New KeyboardHookListener(New AppHooker())
        keyBoardListener.Enabled = True
        AddHandler keyBoardListener.KeyDown, AddressOf keyListener_KeyDown
    End Sub
    Protected Overrides Function RequestComAddInAutomationService() As Object
        If comCalls Is Nothing Then
            comCalls = New COMCalls()
        End If
        Return comCalls
    End Function

    Protected Sub keyListener_KeyDown(ByVal sender As Object, ByVal e As KeyEventArgs)
        Dim firstResultTable As New DataTable
        Dim secondResultTable As New DataTable
        Dim mixingPlanResultTable As New DataTable
        Dim startTick, endTick As Long
        Dim sFirstSheet As String
        Dim sFirstTable As String
        Dim sSecondSheet As String
        Dim sSecondTable As String

        Try
            If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then Exit Sub 'If not a OrklaRT report do nothing.            
            If e.KeyCode = Keys.F8 Then
                For Each taskpane In CustomTaskPanes
                    Dim selection = DirectCast(DirectCast(taskpane.Control, UserControl).Controls.Owner, Selection)
                    If Not DirectCast(selection.selectionElementHost.Child, SelectionPane.Selection).CheckReportMandatoryInput(gReportID, gUserId) Then
                        MessageBox.Show("Vennligst fyll ut alle obligatoriske felter (rødt felt)")
                        Application.OnKey("{F8}", String.Empty)
                        Exit Sub
                    End If
                    If Not DirectCast(selection.selectionElementHost.Child, SelectionPane.Selection).CheckReportRequiredInput(gReportID, gUserId) Then
                        MessageBox.Show("Vennligst fyll ut ett av de nødvendige feltene (gult felt)")
                        Application.OnKey("{F8}", String.Empty)
                        Exit Sub
                    End If
                Next

                Try
                    OrklaRTBPL.CommonFacade.InsertReportLog(gUserId, gReportID)
                Catch ex As Exception
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                End Try

                Application.EnableEvents = False
                Application.ScreenUpdating = False

                Application.Cursor = Excel.XlMousePointer.xlWait
                Application.Calculation = Excel.XlCalculation.xlCalculationManual

                startTick = DateTime.Now.Ticks
                Using entities = New DAL.SAPExlEntities()

                    If gReportID.Equals(33) Then
                        secondResultTable = OrklaRTBPL.ReportSpecific.GetT157EData(OrklaRTBPL.SelectionFacade.ScrappingOverviewSelectionLanguage).Tables(0)
                        Call Common.LoadListObjectData("T157E", "RFC_T157E_Data", "RFC_T157E", secondResultTable)
                        secondResultTable = New DataTable()
                    End If

                    System.Windows.Forms.Application.DoEvents()
                    System.Windows.Forms.Application.EnableVisualStyles()
                    Dim report = entities.Reports.SingleOrDefault(Function(p) p.ReportID = gReportID)
                    Application.StatusBar = String.Format("Henter data fra BW Query {0}, vennligst vent .....", report.QueryName)
                    Try
                        firstResultTable = New DataTable()
                        firstResultTable = RfcFunctions.BWFunctionCall(report.QueryName, report.ReportID, entities.vwCurrentUser.SingleOrDefault().ID, OrklaRTBPL.CommonFacade.GetVariantID(gUserId, gReportID), String.Empty, OrklaRTBPL.SelectionFacade.ReportSelectionLanguage)
                        If gReportID.Equals(7) Or gReportID.Equals(63) Then
                            ProductionPlanTable = firstResultTable
                            Globals.Ribbons.OrklaRT.GetLockedOrders()
                            If OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate <> String.Empty Then
                                gwbReport.Sheets("ReportOptions").Range("CapDate").Value = Convert.ToDateTime(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate).Date
                            End If
                        End If
                        If firstResultTable.Rows.Count > 0 Then
                            If gReportID.Equals(7) Or gReportID.Equals(63) Then
                                Dim dt As New DataTable
                                dt = firstResultTable.Clone
                                dt.Columns("0ACT_START").DataType = GetType(Date)
                                dt.Columns("ZIGSUZI").DataType = GetType(TimeSpan)
                                For i = 0 To firstResultTable.Rows.Count - 1
                                    dt.Rows.Add(i)
                                    For j = 0 To firstResultTable.Columns.Count - 1
                                        If firstResultTable.Columns(j).ToString().Equals("0ACT_START") Then
                                            If Not firstResultTable.Rows(i)(j).Equals(String.Empty) Then dt.Rows(i)(j) = CDate(firstResultTable.Rows(i)(j)).Date
                                        ElseIf firstResultTable.Columns(j).ToString().Equals("ZIGSUZI") Then
                                            If Not firstResultTable.Rows(i)(j).Equals(String.Empty) Then dt.Rows(i)(j) = CDate(firstResultTable.Rows(i)(j)).TimeOfDay
                                        ElseIf firstResultTable.Columns(j).ToString().Equals("20COORDER") Then
                                            dt.Rows(i)(10) = firstResultTable.Rows(i)(j)
                                        Else
                                            dt.Rows(i)(j) = firstResultTable.Rows(i)(j)
                                        End If
                                    Next
                                Next
                                dt.Columns.Remove("20COORDER")
                                dt.DefaultView.Sort = "[20WORKCENTER] ASC,[0ACT_START] DESC,[ZIGSUZI] DESC"
                                Call Common.LoadListObjectData(report.QueryName, "DataBase", "SapExlData", dt)
                                Call Common.LoadListObjectData(report.QueryName, "AllergenData", "tblAllergenData", dt)
                                Call Common.LoadListObjectData("Allergen Data", "Allergen", "tblAllergen", OrklaRTBPL.ReportSpecific.GetAllergenData().Tables(0))
                                Call Common.LoadListObjectData("AllergenType Data", "AllergenType", "tblAllergenType", OrklaRTBPL.ReportSpecific.GetAllergenType(False).Tables(0))
                            ElseIf gReportID.Equals(8) Or gReportID.Equals(12) Or gReportID.Equals(13) Or gReportID.Equals(20) Or gReportID.Equals(62) Then
                                Dim dt As New DataTable
                                dt = firstResultTable.Clone
                                For i = 0 To firstResultTable.Rows.Count - 1
                                    dt.Rows.Add(i)
                                    For j = 0 To firstResultTable.Columns.Count - 1
                                        If firstResultTable.Columns(j).ToString().Equals("20COORDER") Then
                                            If gReportID.Equals(8) Then
                                                dt.Rows(i)(4) = firstResultTable.Rows(i)(j)
                                            ElseIf gReportID.Equals(12) Then
                                                dt.Rows(i)(3) = firstResultTable.Rows(i)(j)
                                            ElseIf gReportID.Equals(13) Or gReportID.Equals(20) Then
                                                dt.Rows(i)(0) = firstResultTable.Rows(i)(j)
                                            ElseIf gReportID.Equals(62) Then
                                                dt.Rows(i)(11) = firstResultTable.Rows(i)(j)
                                            End If
                                        Else
                                            dt.Rows(i)(j) = firstResultTable.Rows(i)(j)
                                        End If
                                    Next
                                Next
                                dt.Columns.Remove("20COORDER")
                                Call Common.LoadListObjectData(report.QueryName, "DataBase", "SapExlData", dt)
                            ElseIf gReportID.Equals(11) Then
                                Dim dt As New DataTable
                                dt = firstResultTable.Clone
                                dt.Columns("ZIPSTTR").DataType = GetType(Date)
                                dt.Columns("ZIPEDTR").DataType = GetType(Date)
                                For i = 0 To firstResultTable.Rows.Count - 1
                                    dt.Rows.Add(i)
                                    For j = 0 To firstResultTable.Columns.Count - 1
                                        If firstResultTable.Columns(j).ToString().Equals("ZIPSTTR") Then
                                            If Not firstResultTable.Rows(i)(j).Equals(String.Empty) Then dt.Rows(i)(j) = CDate(firstResultTable.Rows(i)(j)).Date
                                        ElseIf firstResultTable.Columns(j).ToString().Equals("ZIPEDTR") Then
                                            If Not firstResultTable.Rows(i)(j).Equals(String.Empty) Then dt.Rows(i)(j) = CDate(firstResultTable.Rows(i)(j)).Date
                                        Else
                                            dt.Rows(i)(j) = firstResultTable.Rows(i)(j)
                                        End If
                                    Next
                                Next
                                dt.DefaultView.Sort = "ZIPSTTR ASC,Measures_00O2TGU64XTN259QEGXJIVPTB DESC"
                                Call Common.LoadListObjectData(report.QueryName, "DataBase", "SapExlData", dt)
                            Else
                                Call Common.LoadListObjectData(report.QueryName, "DataBase", "SapExlData", firstResultTable)
                            End If

                            If gReportID.Equals(34) Then
                                secondResultTable = BPL.RfcFunctions.GetBAPIMATERIALGETAll(OrklaRTBPL.SelectionFacade.StockHistorySelectionMaterial, OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant)
                                Call Common.LoadListObjectData(report.QueryName, "Y084_MatData", "MatData", secondResultTable)
                                secondResultTable = New DataTable()
                                secondResultTable = BPL.RfcFunctions.GetBAPIMATERIALSTOCKREQLIST(OrklaRTBPL.SelectionFacade.StockHistorySelectionMaterial, OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant)
                                Call Common.LoadListObjectData(report.QueryName, "Y084_StockData", "tStockData", secondResultTable)
                                secondResultTable = New DataTable()
                            ElseIf gReportID.Equals(63) Then
                                Globals.Ribbons.OrklaRT.ppTimer.Enabled = True
                            End If

                            If (entities.ReportsLinkedQuery.Where(Function(rsp) rsp.ReportID = gReportID).Count > 0) Then
                                For Each row In entities.ReportsLinkedQuery.Where(Function(rsp) rsp.ReportID = gReportID)
                                    Try
                                        If gReportID.Equals(35) Then
                                            secondResultTable = BPL.RfcFunctions.GetBAPIMATERIALGETAll(OrklaRTBPL.SelectionFacade.StockSimulationSelectionMaterial, OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant)
                                            Call Common.LoadListObjectData(report.QueryName, "Y084_MatData", "MatData", secondResultTable)
                                            secondResultTable = New DataTable()
                                            secondResultTable = BPL.RfcFunctions.GetBAPIMATERIALSTOCKREQLIST(OrklaRTBPL.SelectionFacade.StockSimulationSelectionMaterial, OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant)
                                            Call Common.LoadListObjectData(report.QueryName, "Y084_StockData", "tStockData", secondResultTable)
                                            secondResultTable = New DataTable()
                                        End If


                                        If row.SubReports.Equals(52) Then
                                            Dim dt As New DataTable
                                            If gReportID.Equals(7) Then
                                                If OrklaRTBPL.ReportSpecific.GetWorkCentersCapacityData().Tables(0).Rows.Count > 0 Then
                                                    dt = OrklaRTBPL.ReportSpecific.GetWorkCentersCapacityData().Tables(0)
                                                Else
                                                    secondResultTable = RfcFunctions.BWFunctionCall(row.QueryName, row.SubReports, entities.vwCurrentUser.SingleOrDefault().ID, 0, String.Empty, String.Empty)
                                                    dt = secondResultTable.Clone
                                                    dt.Columns("0CALDAY").DataType = GetType(Date)
                                                    For i = 0 To secondResultTable.Rows.Count - 1
                                                        dt.Rows.Add(i)
                                                        For j = 0 To secondResultTable.Columns.Count - 1
                                                            If secondResultTable.Columns(j).ToString().Equals("0CALDAY") Then
                                                                dt.Rows(i)(j) = CDate(secondResultTable.Rows(i)(j)).Date
                                                            Else
                                                                dt.Rows(i)(j) = secondResultTable.Rows(i)(j)
                                                            End If
                                                        Next
                                                    Next
                                                    dt.DefaultView.Sort = "20WORKCENTER ASC,0CALDAY ASC"
                                                End If
                                            Else
                                                secondResultTable = RfcFunctions.BWFunctionCall(row.QueryName, row.SubReports, entities.vwCurrentUser.SingleOrDefault().ID, 0, String.Empty, String.Empty)
                                                dt = secondResultTable.Clone
                                                dt.Columns("0CALDAY").DataType = GetType(Date)
                                                For i = 0 To secondResultTable.Rows.Count - 1
                                                    dt.Rows.Add(i)
                                                    For j = 0 To secondResultTable.Columns.Count - 1
                                                        If secondResultTable.Columns(j).ToString().Equals("0CALDAY") Then
                                                            dt.Rows(i)(j) = CDate(secondResultTable.Rows(i)(j)).Date
                                                        Else
                                                            dt.Rows(i)(j) = secondResultTable.Rows(i)(j)
                                                        End If
                                                    Next
                                                Next
                                                dt.DefaultView.Sort = "20WORKCENTER ASC,0CALDAY ASC"
                                            End If
                                            Call Common.LoadListObjectData(row.QueryName, row.SheetName, row.ListObjectName, dt)
                                        Else
                                            secondResultTable = RfcFunctions.BWFunctionCall(row.QueryName, row.SubReports, entities.vwCurrentUser.SingleOrDefault().ID, 0, String.Empty, String.Empty)
                                            If secondResultTable.Rows.Count > 0 Then
                                                Call Common.LoadListObjectData(row.QueryName, row.SheetName, row.ListObjectName, secondResultTable)
                                            End If
                                        End If
                                    Catch ex As Exception
                                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                                    End Try
                                Next
                            End If
                            Try
                                Select Case gReportID
                                    Case 7, 63
                                        Call FixedProductionPlan.LocalUpdate()                                        
                                        sSecondSheet = Application.Sheets("Version").Range("SecondSheet").Value
                                        sSecondTable = Application.Sheets("Version").Range("SecondTable").Value
                                        Dim pvt As Excel.PivotTable = Application.Sheets(sSecondSheet).PivotTables(sSecondTable)
                                        pvt.PivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone
                                        pvt.PivotCache.Refresh()
                                    Case 8
                                        Call MixingPlan.LocalUpdate()
                                        Call Common.LoadListObjectData("Allergen Data", "Allergen", "tblAllergen", OrklaRTBPL.ReportSpecific.GetAllergenData().Tables(0))
                                        Call Common.LoadListObjectData("AllergenType Data", "AllergenType", "tblAllergenType", OrklaRTBPL.ReportSpecific.GetAllergenType(True).Tables(0))
                                    Case 11
                                        Call CapacityLevelling.LocalUpdate()
                                    Case 14
                                        Call StockTransfer.LocalUpdate()
                                    Case 15
                                        Call OptimizedLotSize.LocalUpdate()
                                    Case 16
                                        Call DeliveryAgent.LocalUpdate()
                                    Case 24
                                        Call SalesOrder.LocalUpdate()
                                    Case 34
                                        Call StockHistory.LocalUpdate()
                                    Case 35
                                        Call StockSimulation.LocalUpdate()
                                    Case 38
                                        Call AllReports.StockValuesAndCoverageRefreshProdData()
                                    Case 39
                                        Call Common.LoadListObjectData("MARM Data", "MARM", "tblMARM", OrklaRTBPL.ReportSpecific.GetTPKData().Tables(0))
                                    Case 62
                                        Call Common.LoadListObjectData("MARM Data", "MARMDPK", "tblMARMDPK", OrklaRTBPL.ReportSpecific.GetDPKData().Tables(0))
                                        Call Common.LoadListObjectData("MARM Data", "MARMFPK", "tblMARMFPK", OrklaRTBPL.ReportSpecific.GetFPKData().Tables(0))
                                End Select
                            Catch
                            End Try
                        Else
                            System.Windows.MessageBox.Show("Ingen data hentet fra spørringen!")
                            Application.StatusBar = "Ingen data hentet fra spørringen!"
                            Application.ActiveWorkbook.Sheets("Rapport info").Activate()
                        End If

                    Catch ex As Exception
                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                    End Try

                End Using

                endTick = DateTime.Now.Ticks
                Application.ActiveWorkbook.Sheets("Rapport info").Unprotect("next")
                Application.ActiveWorkbook.Sheets("Rapport info").Range("RecNum").Value = "Nummer av rader hentet: " & firstResultTable.Rows.Count.ToString()
                Application.ActiveWorkbook.Sheets("Rapport info").Range("Updated").Value = "Sist oppdatert: " & DateTime.Now.ToString()
                Application.ActiveWorkbook.Sheets("Rapport info").Range("Time_Used").Value = "Tid brukt: " & CInt(TimeSpan.FromTicks(endTick - startTick).TotalSeconds).ToString() + " Sec"
                Application.ActiveWorkbook.Sheets("Rapport info").protect("next")

                If gReportID.Equals(10) Then Application.EnableEvents = True

                sFirstSheet = Application.Sheets("Version").Range("FirstSheet").Value

                If Not Application.Sheets("Version").Range("FirstTable").Value Is Nothing Then
                    sFirstTable = Application.Sheets("Version").Range("FirstTable").Value
                    Dim pvt As Excel.PivotTable = Application.Sheets(sFirstSheet).PivotTables(sFirstTable)
                    pvt.PivotCache.MissingItemsLimit = Excel.XlPivotTableMissingItems.xlMissingItemsNone
                    pvt.PivotCache.Refresh()
                End If

                If Not (gReportID.Equals(7) Or gReportID.Equals(63)) And firstResultTable.Rows.Count > 0 Then
                    Application.Sheets(sFirstSheet).Activate()
                ElseIf (gReportID.Equals(10) And firstResultTable.Rows.Count.Equals(0)) Then
                    Application.ActiveWorkbook.Sheets("Rapport info").Activate()
                End If


                Using entities = New DAL.SAPExlEntities()
                    Dim xml = New XmlDocument()
                    Dim pivotLayoutID = DirectCast(OrklaRTBPL.PivotFacade.GetCurrentUserReportPivotLayoutVariant(gUserId, gReportID).Rows(0)("PivotLayoutVariantID"), Integer)
                    If Not pivotLayoutID.Equals(0) Then
                        xml.Load(New XmlTextReader(New StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = gUserId And rp.VariantID = pivotLayoutID).PivotLayout)))
                        Application.DisplayAlerts = False
                        Application.ActiveWorkbook.XmlMaps.Add(xml.InnerXml, "XtraSerializer")
                        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("PvtTableDef").ListObjects
                            If listObject.Name.Equals("PvtTableDef") Then
                                Try
                                    If (pivotLayoutID.Equals(0)) Then
                                        listObject.XmlMap.ImportXml(xml.InnerXml, True)
                                    Else
                                        listObject.XmlMap.ImportXml(ReturnDiffPivotLayout(pivotLayoutID), True)
                                    End If
                                Catch
                                End Try
                            End If
                        Next
                        Application.DisplayAlerts = True
                        Call PivotFunctions.LoadPivotLayout()

                        For Each pivotLayoitItem In Globals.Ribbons.OrklaRT.ddlPivotLayout.Items
                            If pivotLayoitItem.Tag = pivotLayoutID Then
                                Globals.Ribbons.OrklaRT.ddlPivotLayout.SelectedItem = pivotLayoitItem
                            End If
                        Next
                    End If
                End Using

                Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic

                If Not gReportID.Equals(10) Then
                    Application.EnableEvents = True
                End If

                Application.ScreenUpdating = True
                Application.Cursor = Excel.XlMousePointer.xlDefault

                GC.Collect()
                GC.WaitForPendingFinalizers()

                Application.OnKey("{F8}", String.Empty)

            ElseIf e.KeyCode = Keys.F4 Then
                'Launch Sap field help
            ElseIf e.KeyCode = Keys.F1 Then
                'Launch OrklaRT help
            ElseIf e.KeyCode = Keys.F10 Then
                For Each customTaskPane In Globals.ThisAddIn.CustomTaskPanes
                    If customTaskPane.Title.Equals(Application.ActiveWorkbook.Name.Remove(Application.ActiveWorkbook.Name.Length - 5, 5)) Then
                        If customTaskPane.Visible Then
                            customTaskPane.Visible = False
                        Else
                            customTaskPane.Visible = True
                        End If
                        Exit For
                    End If
                Next
                Application.OnKey("{F10}", String.Empty)
            Else
                'myCustomTaskPane.Control.Hide();
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub


    Private Sub Application_SheetActivate(Sh As Object) Handles Application.SheetActivate
        If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then Exit Sub

        Dim c As Excel.Range
        Dim e As Excel.Range
        Dim d As Integer
        Dim f As Excel.Range
        Dim d1 As Integer
        Dim g As Excel.Range
        Dim y As Integer
        Dim x As Integer

        If gReportID.Equals(8) Then

            If Sh.Name.Equals("MaterialPlan") Then
                Application.ScreenUpdating = False
                Application.EnableEvents = False
                Try
                    Application.ActiveSheet.PivotTables(1).PivotFields("Alerts").CurrentPage = "All"
                    Application.ActiveSheet.PivotTables(1).PivotCache.Refresh()
                    d = Application.ActiveSheet.PivotTables(1).DataBodyRange.Columns.Count
                    d1 = d + Application.ActiveSheet.PivotTables(1).DataBodyRange.Cells(1, 1).Column - 1
                    e = Application.ActiveSheet.PivotTables(1).PivotFields("Stock").LabelRange
                    f = Application.ActiveSheet.PivotTables(1).PivotFields("Material Navn").LabelRange
                    g = Application.ActiveSheet.PivotTables(1).PivotFields("Forsyningsområde").LabelRange
                    Application.Sheets("Alerts").Range("Alerts").ClearContents()

                    x = 0
                    y = 0
                    For Each c In Application.ActiveSheet.PivotTables(1).DataBodyRange.Columns(d).Cells
                        x = x + 1
                        If c.Offset(0, -d1 + 1).PivotCell.PivotCellType < 2 Then
                            If e.Offset(x, 0).Value < c.Value Then
                                y = y + 1
                                e.Offset(x, 0).Interior.Color = 5263615
                                Application.Sheets("Alerts").Cells(y, 1).Value = g.Offset(x, 0).PivotItem.Value & f.Offset(x, 0).Value
                            ElseIf e.Offset(x, 0).Value * 0.8 < c.Value Then
                                y = y + 1
                                e.Offset(x, 0).Interior.Color = 123391
                                Application.Sheets("Alerts").Cells(y, 1).Value = g.Offset(x, 0).PivotItem.Value & f.Offset(x, 0).Value
                            Else
                                e.Offset(x, 0).Interior.ColorIndex = Excel.Constants.xlNone
                            End If
                        End If
                    Next c

                    e = Nothing
                    f = Nothing
                    Application.ActiveSheet.PivotTables(1).PivotCache.Refresh()
                    Application.ActiveSheet.PivotTables(1).PivotFields("Alerts").CurrentPage = "1"
                Catch
                End Try
                Application.ScreenUpdating = True
                Application.EnableEvents = True
            End If
        End If

    End Sub


    Private Sub Application_SheetBeforeDoubleClick(Sh As Object, Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles Application.SheetBeforeDoubleClick

        If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then Exit Sub

        Dim pi As Excel.PivotItem
        Dim lngOrder As Long
        Dim intMach As Integer
        Dim intPri As Integer
        Dim sv As String
        Dim g As Excel.Range
        Dim e As Excel.Range
        Dim f As Excel.Range
        Dim lngDate As Date
        Dim m As Integer
        Dim d As Integer
        Dim y As Integer

        Try
            Try
                If Target.PivotField.Name.Equals("Kommentarer") Then
                    If gReportID.Equals(24) Then
                        For Each col In Target.EntireRow.Columns
                            If col.PivotField.Name.Equals("Material Navn") Then
                                Dim pivotCommentsForm As New PivotComments(col.Value.ToString().Split(" ").GetValue(0))
                                Call pivotCommentsForm.Show()
                                Cancel = True
                                Exit For
                            End If
                        Next
                    ElseIf gReportID.Equals(39) Then
                        For Each col In Target.EntireRow.Columns
                            If col.PivotField.Name.Equals("Batch") Then
                                If Not IsNothing(col.Value) Then
                                    Dim pivotCommentsForm As New PivotComments(col.Value.ToString().Split(" ").GetValue(0))
                                    Call pivotCommentsForm.Show()
                                    Cancel = True
                                    Exit For
                                Else
                                    Cancel = True
                                    Exit For
                                End If
                            End If
                        Next
                    ElseIf gReportID.Equals(33) Then
                        For Each col In Target.EntireRow.Columns
                            If col.PivotField.Name.Equals("Material Navn") Then
                                Dim pivotCommentsForm As New PivotComments(col.Value.ToString().Split(" ").GetValue(0))
                                Call pivotCommentsForm.Show()
                                Cancel = True
                                Exit For
                            End If
                        Next
                    ElseIf gReportID.Equals(8) Then
                        For Each col In Target.EntireRow.Columns
                            If col.PivotField.Name.Equals("Ordre") Or col.PivotField.Name.Equals("Order Mix") Then
                                Dim pivotCommentsForm As New PivotComments(col.Value.ToString())
                                Call pivotCommentsForm.Show()
                                Cancel = True
                                Exit For
                            End If
                        Next
                    End If
                ElseIf Target.PivotField.Name.Equals("Locked") Then
                    If Sh.Name.Equals("Sequence") Then
                        If OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(2) Or OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(3) Then
                            If Target.Count > 1 Then
                                MsgBox("Active cell must be a Pivot table cell in the 'Locked' column to use the Lock / Unlock function.", , "Orkla SAP Intergation")
                                Exit Sub
                            Else
                                If gReportID.Equals(7) Or gReportID.Equals(63) Then
                                    For Each col In Target.EntireRow.Columns
                                        If col.PivotField.Name.Equals("Order") Then
                                            Application.EnableEvents = False
                                            Application.ScreenUpdating = False
                                            Application.Calculation = Excel.XlCalculation.xlCalculationManual
                                            Call FixedProductionPlan.WriteLockedOrders(col.Value)
                                            Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
                                            Application.ScreenUpdating = True
                                            Application.EnableEvents = True
                                            Cancel = True
                                            Exit For
                                        End If
                                    Next
                                End If
                            End If
                        Else
                            MsgBox("You are not allowed to lock an order,please contact produksjonplanlegger.", , "Orkla SAP Intergation")
                            Cancel = True
                            Exit Sub
                        End If
                    End If
                End If
            Catch
            End Try


            If gReportID.Equals(8) Then

                If Sh.Name.Equals("BlandePlan") Then
                    Application.ScreenUpdating = False
                    Application.EnableEvents = False

                    intMach = 0
                    lngOrder = 0

                    Try
                        If Target.PivotCell.PivotField.Name = "Tid" Then
                            Call MixingPlanSheetChangeStart(Target)
                            'Application.ActiveSheet.PivotTables(1).PivotCache.Refresh()                          
                        End If
                    Catch
                    End Try

                    Try
                        If Target.PivotCell.PivotField.Name = "Mach" Then
                            sv = InputBox("Please input machine number to use.", "Orkla SAP Integration")
                            g = Target.PivotTable.PivotFields("Ordre").LabelRange

                            If Not String.IsNullOrWhiteSpace(sv) Then
                                If Convert.ToInt32(sv) > 0 And Convert.ToInt32(sv) < 4 Then
                                    intMach = Convert.ToInt32(sv)
                                Else
                                    intMach = 1
                                End If
                            End If

                            lngOrder = Application.Cells(Target.Row, g.Column).PivotItem.Value

                            If intMach > 0 And lngOrder > 0 Then
                                Call MixingPlan.FindOrder(lngOrder, intMach)
                            End If
                        End If
                    Catch
                    End Try

                    Try
                        If Target.PivotCell.PivotField.Name = "RS" Then
                            sv = InputBox("Please input type of container to use.", "Orkla SAP Integration")
                            g = Target.PivotTable.PivotFields("Order").LabelRange
                            lngOrder = Application.Cells(Target.Row, g.Column).PivotItem.Value

                            If Not String.IsNullOrWhiteSpace(sv) Then
                                If UCase(sv) = "R" Or UCase(sv) = "S" Then
                                    Call MixingPlan.FindRS(lngOrder, UCase(sv))
                                    Application.ActiveSheet.PivotTables(1).PivotCache.Refresh()
                                Else
                                    MsgBox("Wrong input. Has to be R or S", , "Orkla SAP Integration")
                                End If
                            End If
                        End If
                    Catch
                    End Try

                    Try
                        If Target.PivotCell.PivotField.Name = "Pri" Then
                            sv = InputBox("Please input order priority.", "Orkla SAP Integration")
                            g = Target.PivotTable.PivotFields("Ordre").LabelRange
                            e = Target.PivotTable.PivotFields("Start_Date").LabelRange
                            f = Target.PivotTable.PivotFields("Mach").LabelRange

                            If Not String.IsNullOrWhiteSpace(sv) Then
                                intPri = Convert.ToInt32(sv)
                            End If

                            lngOrder = Application.Cells(Target.Row, g.Column).PivotItem.Value
                            Dim c = Application.Cells(Target.Row, e.Column).PivotItem.Value.ToString().Split("/")
                            lngDate = New Date(CInt(c(2)), CInt(c(0)), CInt(c(1)))
                            intMach = Application.Cells(Target.Row, f.Column).PivotItem.Value

                            If intPri > 0 And lngOrder > 0 Then
                                Call MixingPlan.FindPri(lngOrder, intPri)
                                Application.ActiveSheet.PivotTables(1).PivotCache.Refresh()
                                Call MixingPlanSheetCalculateTime(intMach, lngDate)
                                Call MixingPlan.WriteNewStart()
                                Call MixingPlan.RefreshNewStart()
                                If Not Application.Sheets("Version").Range("FirstTable").Value Is Nothing Then
                                    Dim pvt As Excel.PivotTable = Application.Sheets(Application.Sheets("Version").Range("FirstSheet").Value).PivotTables(Application.Sheets("Version").Range("FirstTable").Value)
                                    pvt.PivotCache.Refresh()
                                End If
                            End If
                        End If
                    Catch
                    End Try

                    Cancel = True
                    Application.EnableEvents = True
                    Application.ScreenUpdating = True

                ElseIf Sh.Name.Equals("Alle_Arbeidsstasjoner") Then

                    Application.ScreenUpdating = False
                    Application.EnableEvents = False

                    intMach = 0
                    lngOrder = 0

                    Try
                        If Target.PivotCell.PivotField.Name = "Hour" Then
                            Call MixingPlanAllWorkCentersSheetChangeStart(Target)
                        End If
                    Catch
                    End Try


                    Try
                        For Each pi In Target.PivotCell.ColumnItems
                            If pi.Parent.SourceName = "Mach" Then
                                If Application.ActiveSheet.PivotTables(1).PivotFields("Mach").PivotItems.Count < 3 Then
                                    sv = InputBox("Please input machine number to use.", "Orkla SAP Integration")
                                    If Not String.IsNullOrWhiteSpace(sv) Then
                                        If Convert.ToInt32(sv) > 0 And Convert.ToInt32(sv) < 4 Then
                                            intMach = Convert.ToInt32(sv)
                                        Else
                                            intMach = pi.SourceName
                                        End If
                                    End If
                                Else
                                    intMach = pi.SourceName
                                    Exit For
                                End If
                            End If
                        Next
                    Catch
                    End Try

                    Try
                        For Each pi In Target.PivotCell.RowItems
                            If pi.Parent.SourceName = "Ordre" Then
                                lngOrder = pi.SourceName
                                Exit For
                            End If
                        Next
                    Catch
                    End Try

                    If intMach > 0 And lngOrder > 0 Then
                        Call FindOrder(lngOrder, intMach)
                    End If

                    Cancel = True
                    Application.EnableEvents = True
                    Application.ScreenUpdating = True

                End If
                If Not Application.Sheets("Version").Range("FirstTable").Value Is Nothing Then
                    Dim pvt As Excel.PivotTable = Application.Sheets(Application.Sheets("Version").Range("FirstSheet").Value).PivotTables(Application.Sheets("Version").Range("FirstTable").Value)
                    pvt.PivotCache.Refresh()
                End If
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try

    End Sub



    Private Sub Application_SheetBeforeRightClick(Sh As Object, Target As Microsoft.Office.Interop.Excel.Range, ByRef Cancel As Boolean) Handles Application.SheetBeforeRightClick
        If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then
            Try
                Application.CommandBars("Cell").ShowPopup()
            Catch
            End Try
            Cancel = True
        Else
            DefineShortcutMenu()
            Cancel = True
        End If
    End Sub

    Private Sub Application_SheetCalculate(Sh As Object) Handles Application.SheetCalculate
        If gReportID.Equals(7) Or gReportID.Equals(63) Then
            If Sh.Name <> "ProdPlan" Then
                Call FixedProductionPlan.GetTime()
            End If
        End If
    End Sub

    Private Sub Application_SheetChange(Sh As Object, Target As Microsoft.Office.Interop.Excel.Range) Handles Application.SheetChange
        If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then Exit Sub

        Application.ScreenUpdating = False
        Application.EnableEvents = False

        Try

            If gReportID.Equals(10) Then

                If Sh.name.Equals("Lager") Then
                    If Target.Address = Application.Range("RollReq").Address Then
                        Sh.ChartObjects(1).Activate()
                        Application.ActiveChart.SeriesCollection(1).XValues = "=Lager!x_axis"
                        Application.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).TickLabels.NumberFormat = "dd.mm.yyyy"
                        Sh.Pivottables(1).PivotCache.Refresh()
                        Sh.ChartObjects(1).Chart.Refresh()
                    End If
                End If

            ElseIf gReportID.Equals(39) Then

                If Sh.Name.Equals("Oversikt") Then
                    If Left(Target.Name.Name, 5) = "Limit" Then
                        Application.ActiveSheet.PivotTables(1).PivotCache.Refresh()
                    End If
                End If

            ElseIf gReportID.Equals(35) Then

                Call StopEvents(True, True, True)
                If Target.Address = Application.Range("crLead_Time").Address Or Target.Address = Application.Range("crSaf_Time").Address Then
                    Call StockSimulation.UpdateStockInFrom()
                End If
                Call ResetAllEvents()

            ElseIf gReportID.Equals(16) Then

                If Sh.Name.Equals("Sikkerhetsdager ved levering") Then
                    Application.Calculation = Excel.XlCalculation.xlCalculationManual
                    If Target.Address = Application.Range("GreenLimit").Address Then
                        Application.Calculate()
                        Application.ActiveSheet.PivotTables(1).PivotCache.Refresh()
                        Call DeliveryAgent.FormatPivotTable()
                    End If
                    Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
                End If
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try

        Application.ScreenUpdating = True
        Application.EnableEvents = True

    End Sub

    Private Sub Application_SheetDeactivate(Sh As Object) Handles Application.SheetDeactivate
        If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then Exit Sub
    End Sub


    Private Sub Application_SheetPivotTableUpdate(Sh As Object, Target As Microsoft.Office.Interop.Excel.PivotTable) Handles Application.SheetPivotTableUpdate
        If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then Exit Sub
        Dim x1 As Integer, x2 As Integer, x3 As Integer
        Dim c1 As Integer, c2 As Integer, c3 As Integer, c4 As Integer, c5 As Integer

        Application.ScreenUpdating = False
        Application.EnableEvents = False

        If gReportID.Equals(10) Then

            If Sh.Name.Equals("Detaljer") Then
                Try

                    x1 = Target.RowRange.Cells(2, 1).Row
                    x2 = Target.RowRange.Cells(2, 1).Row + Target.RowRange.Cells.Count - 3
                    x3 = Target.RowRange.Cells(1, 1).Row + 1
                    c1 = Target.RowRange.Cells(1, 1).Column
                    c2 = Target.PivotFields("Beholding").LabelRange.Column
                    c3 = Target.PivotFields("Behov ").LabelRange.Column
                    c4 = Target.PivotFields("Snitt dekn.").LabelRange.Column
                    c5 = Target.PivotFields("Rullerende_beh ").LabelRange.Column

                    Application.ActiveWorkbook.Names.Add(Name:="Detaljer!x_axis", RefersToR1C1:="=R" & x1 & "C" & c1 & ":R" & x2 & "C" & c1)
                    Application.ActiveWorkbook.Names.Add(Name:="Detaljer!S_Stock", RefersToR1C1:="=R" & x1 & "C" & c2 & ":R" & x2 & "C" & c2)
                    Application.ActiveWorkbook.Names.Add(Name:="Detaljer!S_Req", RefersToR1C1:="=R" & x1 & "C" & c3 & ":R" & x2 & "C" & c3)
                    Application.ActiveWorkbook.Names.Add(Name:="Detaljer!S_Cov", RefersToR1C1:="=R" & x1 & "C" & c4 & ":R" & x2 & "C" & c4)
                    Application.ActiveWorkbook.Names.Add(Name:="Detaljer!S_Roll", RefersToR1C1:="=R" & x1 & "C" & c5 & ":R" & x2 & "C" & c5)

                    Sh.ChartObjects(1).Activate()
                    Application.ActiveChart.SeriesCollection(1).XValues = "=Detaljer!x_axis"
                    Application.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).TickLabels.NumberFormat = "dd.mm.yyyy"
                    Sh.Pivottables(1).PivotCache.Refresh()
                    Sh.ChartObjects(1).Chart.Refresh()

                Catch ex As Exception
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                End Try

            ElseIf Sh.Name.Equals("MD04") Then

                Try

                    x1 = Target.RowRange.Cells(2, 1).Row
                    x2 = Target.RowRange.Cells(2, 1).Row + Target.RowRange.Cells.Count - 3
                    x3 = Target.RowRange.Cells(1, 1).Row + 1
                    c1 = Target.RowRange.Cells(1, 1).Column
                    c5 = Target.PivotFields("Snitt lagerdekn. ").LabelRange.Column

                    Application.ActiveWorkbook.Names.Add(Name:="MD04!x_axis", RefersToR1C1:="=R" & x1 & "C" & c1 & ":R" & x2 & "C" & c1)
                    Application.ActiveWorkbook.Names.Add(Name:="MD04!S_Cov", RefersToR1C1:="=R" & x1 & "C" & c5 & ":R" & x2 & "C" & c5)

                    Sh.ChartObjects(1).Activate()
                    Application.ActiveChart.SeriesCollection(1).XValues = "=MD04!x_axis"
                    Application.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).TickLabels.NumberFormat = "dd.mm.yyyy"
                    Sh.Pivottables(1).PivotCache.Refresh()
                    Sh.ChartObjects(1).Chart.Refresh()

                Catch ex As Exception
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                End Try

            ElseIf Sh.Name.Equals("Lager") Then

                Try

                    x1 = Target.RowRange.Cells(2, 1).Row
                    x2 = Target.RowRange.Cells(2, 1).Row + Target.RowRange.Cells.Count - 2
                    x3 = Target.RowRange.Cells(1, 1).Row + 1
                    c1 = Target.RowRange.Cells(1, 1).Column
                    c2 = Target.PivotFields("Beholding ").LabelRange.Column
                    c3 = Target.PivotFields("Behov ").LabelRange.Column
                    c4 = Target.PivotFields("Rullerende_beh ").LabelRange.Column
                    c5 = Target.PivotFields("Snitt dekn.  ").LabelRange.Column

                    Application.ActiveWorkbook.Names.Add(Name:="Lager!x_axis", RefersToR1C1:="=R" & x1 & "C" & c1 & ":R" & x2 & "C" & c1)
                    Application.ActiveWorkbook.Names.Add(Name:="Lager!S_Stock", RefersToR1C1:="=R" & x1 & "C" & c2 & ":R" & x2 & "C" & c2)
                    Application.ActiveWorkbook.Names.Add(Name:="Lager!S_Req", RefersToR1C1:="=R" & x1 & "C" & c3 & ":R" & x2 & "C" & c3)
                    Application.ActiveWorkbook.Names.Add(Name:="Lager!S_Roll", RefersToR1C1:="=R" & x1 & "C" & c4 & ":R" & x2 & "C" & c4)
                    Application.ActiveWorkbook.Names.Add(Name:="Lager!S_Cov", RefersToR1C1:="=R" & x1 & "C" & c5 & ":R" & x2 & "C" & c5)

                    Sh.ChartObjects(1).Activate()
                    Application.ActiveChart.SeriesCollection(1).XValues = "=Lager!x_axis"
                    Application.ActiveChart.Axes(Excel.XlAxisType.xlCategory, Excel.XlAxisGroup.xlPrimary).TickLabels.NumberFormat = "dd.mm.yyyy"
                    Sh.Pivottables(1).PivotCache.Refresh()
                    Sh.ChartObjects(1).Chart.Refresh()
                Catch ex As Exception
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                End Try
            End If

        ElseIf gReportID.Equals(16) Then

            If Sh.Name.Equals("Sikkerhetsdager ved levering") Then
                Call DeliveryAgent.FormatPivotTable()
            End If

        End If

        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End Sub

    Private Sub Application_SheetSelectionChange(Sh As Object, Target As Microsoft.Office.Interop.Excel.Range) Handles Application.SheetSelectionChange
        Dim r As Integer

        If gReportID.Equals(11) Then
            Application.EnableEvents = False
            If Sh.Name.Equals("Gantt") Then
                If Target.Row > 6 And Target.Column > 2 Then
                    Application.StatusBar = Target.Cells(7, Target.Column).Value & " " & Target.Cells(Target.Row, 2).value
                Else
                    Application.StatusBar = String.Empty
                End If
            End If
            Application.EnableEvents = True
        End If

    End Sub

    Private Sub Application_WorkbookActivate(Wb As Microsoft.Office.Interop.Excel.Workbook) Handles Application.WorkbookActivate
        If String.IsNullOrWhiteSpace(IsOrklaRTReport(Wb)) Then
            Globals.Ribbons.OrklaRT.grpOptions.Visible = False
            Globals.Ribbons.OrklaRT.grpPivotLayout.Visible = False
            Exit Sub
        Else
            Using entities = New DAL.SAPExlEntities()
                Dim reportDefinition = entities.Reports.Where(Function(r) r.ReportName = Wb.Name.Remove(Wb.Name.Length - 5, 5)).SingleOrDefault()
                gReportID = reportDefinition.ReportID
            End Using

            Call Common.ShowReportOptions()
            Call Globals.Ribbons.OrklaRT.LoadPivotLayouts()
            Dim pivotLayoutID = DirectCast(OrklaRTBPL.PivotFacade.GetCurrentUserReportPivotLayoutVariant(gUserId, gReportID).Rows(0)("PivotLayoutVariantID"), Integer)
            If Not pivotLayoutID.Equals(0) Then
                For Each pivotLayoitItem In Globals.Ribbons.OrklaRT.ddlPivotLayout.Items
                    If pivotLayoitItem.Tag = pivotLayoutID Then
                        Globals.Ribbons.OrklaRT.ddlPivotLayout.SelectedItem = pivotLayoitItem
                    End If
                Next
            End If

            Globals.Ribbons.OrklaRT.grpOptions.Visible = True
            Globals.Ribbons.OrklaRT.grpPivotLayout.Visible = True
        End If

        Try
            For Each customTaskPane In Globals.ThisAddIn.CustomTaskPanes
                If customTaskPane.Title.Equals(Wb.Name.Substring(0, Wb.Name.Length - 5)) Then
                    customTaskPane.Visible = True
                    Exit For
                End If
            Next
        Catch
        End Try

        Application.Caption = gSysTitle
    End Sub

    Private Sub Application_WorkbookBeforeClose(Wb As Microsoft.Office.Interop.Excel.Workbook, ByRef Cancel As Boolean) Handles Application.WorkbookBeforeClose

        Try
            If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then
                Globals.Ribbons.OrklaRT.grpOptions.Visible = False
                Globals.Ribbons.OrklaRT.grpPivotLayout.Visible = False
                Exit Sub
            Else
                For Each customTaskPane In Globals.ThisAddIn.CustomTaskPanes
                    If customTaskPane.Title.Equals(Wb.Name.Substring(0, Wb.Name.Length - 5)) Then
                        Globals.ThisAddIn.CustomTaskPanes.Remove(customTaskPane)
                        Globals.Ribbons.OrklaRT.grpOptions.Visible = False
                        Exit For
                    End If
                Next
                If gReportID.Equals(63) Then
                    Globals.Ribbons.OrklaRT.ppTimer.Enabled = False
                End If
                gwbReport.Saved = True
            End If
        Catch
            gwbReport.Saved = True
        End Try

    End Sub

    Private Sub Application_WorkbookDeactivate(Wb As Microsoft.Office.Interop.Excel.Workbook) Handles Application.WorkbookDeactivate
        If String.IsNullOrWhiteSpace(IsOrklaRTReport(Application.ActiveWorkbook)) Then Exit Sub
        Try
            For Each customTaskPane In Globals.ThisAddIn.CustomTaskPanes
                If customTaskPane.Title.Equals(Wb.Name.Substring(0, Wb.Name.Length - 5)) Then
                    customTaskPane.Visible = False
                    Exit For
                End If
            Next
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub DefineShortcutMenu()
        If fnIsPivotTable() Then
            Try
                For Each pf In Application.ActiveCell.PivotTable.PageFields
                    If pf.CurrentPage.SourceName <> "(All)" Then
                        If Not String.IsNullOrWhiteSpace(pf.SourceName()) Then
                            Select Case pf.SourceName()
                                Case "Fabrikk"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Plant", pf.CurrentPage.SourceName(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Plant", pf.CurrentPage.SourceName(), gUserId, 2)
                                Case "Material"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pf.SourceName(), pf.CurrentPage.SourceName(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pf.SourceName(), pf.CurrentPage.SourceName(), gUserId, 2)
                                Case "Material Navn"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", pf.CurrentPage.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", pf.CurrentPage.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, 2)
                                Case "Material Navn - Inngår i"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", pf.CurrentPage.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", pf.CurrentPage.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, 2)
                                Case "Navn blanding"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", pf.CurrentPage.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", pf.CurrentPage.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, 2)
                                Case "Ordre"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("OrderNumber", pf.CurrentPage.SourceName(), gUserId, gReportID)
                                Case "Order"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("OrderNumber", pf.CurrentPage.SourceName(), gUserId, gReportID)
                                Case "Salgs dok."
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("SalesOrder", pf.CurrentPage.SourceName(), gUserId, gReportID)
                                Case "Levering"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pf.SourceName(), pf.CurrentPage.SourceName(), gUserId, gReportID)
                                Case "Purch.Dok."
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("PurchaseOrder", pf.CurrentPage.SourceName(), gUserId, gReportID)
                                Case "PO"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("PurchaseOrder", pf.CurrentPage.SourceName(), gUserId, gReportID)
                            End Select
                        End If
                    End If
                Next pf
            Catch
            End Try

            Try
                For Each rf In Application.ActiveCell.PivotTable.RowFields
                    If Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceNameStandard <> "(blank)" Then
                        If Not String.IsNullOrWhiteSpace(rf.SourceName()) Then
                            Select Case rf.SourceName()
                                Case "Fabrikk"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Plant", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Plant", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, 2)
                                Case "Material"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(rf.SourceName(), Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(rf.SourceName(), Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, 2)
                                Case "Material Navn"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, 2)
                                Case "Material Navn - Inngår i"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, 2)
                                Case "Navn blanding"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, gReportID)
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Material", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName().ToString().Split(" ").GetValue(0).ToString(), gUserId, 2)
                                Case "Ordre"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("OrderNumber", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, gReportID)
                                Case "Order"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("OrderNumber", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, gReportID)
                                Case "Salgs dok."
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("SalesOrder", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, gReportID)
                                Case "Levering"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(rf.SourceName(), Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, gReportID)
                                Case "Purch.Dok."
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("PurchaseOrder", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, gReportID)
                                Case "PO"
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("PurchaseOrder", Application.ActiveCell.EntireRow.Cells(1, rf.LabelRange.Column).PivotItem.SourceName(), gUserId, gReportID)
                            End Select
                        End If
                    End If
                Next rf
            Catch
            End Try

            Try
                For Each pi In Application.ActiveCell.PivotCell.ColumnItems
                    If Not String.IsNullOrWhiteSpace(pi.Parent.SourceName()) Then
                        Select Case pi.Parent.SourceName()
                            Case "Fabrikk"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Plant", pi.SourceName(), gUserId, gReportID)
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Plant", pi.SourceName(), gUserId, 2)
                            Case "Material"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName(), gUserId, gReportID)
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName(), gUserId, 2)
                            Case "Material Navn"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName().ToString().Split(" ").GetValue(0), gUserId, gReportID)
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName().ToString().Split(" ").GetValue(0), gUserId, 2)
                            Case "Material Navn - Inngår i"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName().ToString().Split(" ").GetValue(0), gUserId, gReportID)
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName().ToString().Split(" ").GetValue(0), gUserId, 2)
                            Case "Navn blanding"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName().ToString().Split(" ").GetValue(0), gUserId, gReportID)
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName().ToString().Split(" ").GetValue(0), gUserId, 2)
                            Case "Ordre"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("OrderNumber", pi.SourceName(), gUserId, gReportID)
                            Case "Order"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("OrderNumber", pi.SourceName(), gUserId, gReportID)
                            Case "Salgs dok."
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("SalesOrder", pi.SourceName(), gUserId, gReportID)
                            Case "Levering"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields(pi.Parent.SourceName(), pi.SourceName(), gUserId, gReportID)
                            Case "Purch.Dok."
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("PurchaseOrder", pi.SourceName(), gUserId, gReportID)
                            Case "PO"
                                OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("PurchaseOrder", pi.SourceName(), gUserId, gReportID)
                        End Select
                    End If
                Next pi
            Catch
            End Try

            Try
                Dim commandBar As Office.CommandBar
                Dim subMenuCommandBarButton As Office.CommandBarButton
                Dim commandBarButton As Office.CommandBarButton
                Dim rightClickMenuID As Integer

                Try
                    Application.CommandBars("SheetRightClick").Delete()
                Catch
                End Try
                commandBar = Application.CommandBars.Add("SheetRightClick", Office.MsoBarPosition.msoBarPopup)
                Using entities = New DAL.SAPExlEntities()
                    For Each rightClickMenu As DataRow In OrklaRTBPL.CommonFacade.GetReportRightClickMenu(gReportID).Tables(0).Rows
                        If OrklaRTBPL.CommonFacade.GetRightClickSubMenuCount(rightClickMenu("MenuID")) <> 0 Then
                            Dim rightClickPopupMenu = commandBar.Controls.Add(Office.MsoControlType.msoControlPopup)
                            rightClickPopupMenu.Caption = rightClickMenu("Caption")
                            rightClickPopupMenu.BeginGroup = rightClickMenu("BeginGroup")
                            If IsNumeric(Application.ActiveCell.Value) Then
                                rightClickPopupMenu.Enabled = True
                            Else
                                rightClickPopupMenu.Enabled = False
                            End If
                            Try
                                rightClickMenuID = rightClickMenu("MenuID")
                                For Each rightClickSubMenu In entities.RightClickSubMenu.Where(Function(rcsm) rcsm.RightClickMenuID = rightClickMenuID)
                                    subMenuCommandBarButton = DirectCast(rightClickPopupMenu.Controls.Add(Office.MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
                                    subMenuCommandBarButton.Style = Office.MsoButtonStyle.msoButtonCaption
                                    subMenuCommandBarButton.Caption = rightClickSubMenu.Caption
                                    subMenuCommandBarButton.BeginGroup = rightClickSubMenu.BeginGroup
                                    subMenuCommandBarButton.Tag = rightClickSubMenu.FunctionName
                                    subMenuCommandBarButton.Enabled = True
                                    AddHandler subMenuCommandBarButton.Click, AddressOf SubMenuCommandBarButton_Click
                                Next
                            Catch
                            End Try
                        Else
                            commandBarButton = DirectCast(commandBar.Controls.Add(Office.MsoControlType.msoControlButton), Microsoft.Office.Core.CommandBarButton)
                            commandBarButton.Style = Office.MsoButtonStyle.msoButtonCaption
                            commandBarButton.Caption = rightClickMenu("Caption")
                            commandBarButton.BeginGroup = rightClickMenu("BeginGroup")
                            commandBarButton.Tag = rightClickMenu("FunctionName")
                            commandBarButton.Enabled = rightClickMenu("Enable")
                            AddHandler commandBarButton.Click, AddressOf CommandBarButton_Click
                        End If
                    Next
                End Using
                commandBar.ShowPopup()

                commandBarButton = Nothing
                subMenuCommandBarButton = Nothing
                commandBar = Nothing
            Catch
            End Try
        Else
            Try
                Application.CommandBars("Cell").ShowPopup()
            Catch
            End Try
        End If
    End Sub
    Private Sub CommandBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        If Not String.IsNullOrWhiteSpace(Ctrl.Tag) And Ctrl.Tag <> "No Action" Then
            Using entities = New DAL.SAPExlEntities()
                If Ctrl.Caption.ToString().StartsWith("SAP") Then
                    Try
                        Dim rightClickMenuFeilds = OrklaRTBPL.CommonFacade.GetRightClickReportMenuFields(Ctrl.Tag)
                        For Each rightClickMenuField In rightClickMenuFeilds.Tables(0).Rows
                            rightClickMenuField("FieldValue") = OrklaRTBPL.CommonFacade.GetCurrentUserReportFields(gUserId, gReportID).Tables(0).Rows(0)(rightClickMenuField("ValueFieldName"))
                        Next

                        BPL.RfcFunctions.RfcTransactionCallUsing(rightClickMenuFeilds.Tables(0), Ctrl.Tag)
                    Catch ex As Exception
                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                    End Try
                    Application.CommandBars("SheetRightClick").Delete()
                ElseIf IsNumeric(Ctrl.Tag) Then
                    Try
                        If Convert.ToInt32(Ctrl.Tag).Equals(2) Or Convert.ToInt32(Ctrl.Tag).Equals(3) Then
                            OrklaRTBPL.SelectionFacade.DeleteCurrentUserReportSelections(2, gUserId)
                            OrklaRTBPL.SelectionFacade.InsertReportSelectionToCurrentUserReportSelections(2, gUserId)
                            Dim dt = OrklaRTBPL.CommonFacade.GetCurrentUserReportFields(gUserId, 2).Tables(0)
                            If Convert.ToInt32(Ctrl.Tag).Equals(2) Then
                                OrklaRTBPL.SelectionFacade.UpdateCurrentUserReportSelectionLowValue(2, gUserId, "ZVR013", "S", "I", "EQ", dt.Rows(0)("Material"))
                                OrklaRTBPL.SelectionFacade.UpdateCurrentUserReportSelectionLowValue(2, gUserId, "0P_PLANT", "S", "I", "EQ", dt.Rows(0)("Plant"))
                            Else
                                OrklaRTBPL.SelectionFacade.UpdateCurrentUserReportSelectionLowValue(2, gUserId, "ZVR028", "S", "I", "EQ", dt.Rows(0)("Material"))
                                OrklaRTBPL.SelectionFacade.UpdateCurrentUserReportSelectionLowValue(2, gUserId, "0P_PLANT", "S", "I", "EQ", dt.Rows(0)("Plant"))
                            End If
                            For Each workbook As Excel.Workbook In Application.Workbooks
                                If workbook.Name = "Inngår i,består av.xlsm" Or workbook.FullName = "Inngår i,består av.xlsm" Then
                                    For Each customTaskPane In Globals.ThisAddIn.CustomTaskPanes
                                        If customTaskPane.Title.Equals("Inngår i,består av") Then
                                            Globals.ThisAddIn.CustomTaskPanes.Remove(customTaskPane)
                                            Globals.Ribbons.OrklaRT.grpOptions.Visible = False
                                            Exit For
                                        End If
                                    Next
                                    workbook.Saved = True
                                    workbook.Close()
                                End If
                            Next
                            Call Globals.Ribbons.OrklaRT.OpenReport(2, True)
                            System.Windows.Forms.SendKeys.Send("{F8}")
                        End If
                    Catch ex As Exception
                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                    End Try
                Else
                    Try
                        Dim type As Type = GetType(EventFunctions)
                        Dim method As MethodInfo = type.GetMethod(Ctrl.Tag.ToString())
                        If method IsNot Nothing Then
                            method.Invoke(Me, Nothing)
                        End If
                    Catch ex As Exception
                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
                    End Try
                End If
            End Using
            CancelDefault = True
        Else
            Try
                Application.CommandBars("SheetRightClick").Delete()
                Application.CommandBars("PivotTable Context Menu").ShowPopup()
            Catch ex As Exception
            End Try
            CancelDefault = True
        End If
    End Sub
    Private Sub SubMenuCommandBarButton_Click(Ctrl As Microsoft.Office.Core.CommandBarButton, ByRef CancelDefault As Boolean)
        Try
            Dim pivotFunctions As Type = GetType(PivotFunctions)
            Dim method As MethodInfo = pivotFunctions.GetMethod(Ctrl.Tag)
            method.Invoke(Me, Nothing)
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub

End Class

