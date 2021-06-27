Imports System.Runtime.InteropServices

<ComVisible(True)>
<ClassInterface(ClassInterfaceType.None)>
Public Class COMCalls
    Implements ICOMCalls
    Public Sub DeliveryAgentFormatPivotTable()
        Application.ScreenUpdating = False
        Application.EnableEvents = False

        Call DeliveryAgent.FormatPivotTable()

        Application.ScreenUpdating = True
        Application.EnableEvents = True
    End Sub
    Public Sub LaunchSAPOrder(orderNumber As String) Implements ICOMCalls.LaunchSAPOrder
        Dim cor3Table As System.Data.DataTable
        cor3Table = RfcTransactionTable()
        cor3Table.Rows.Add("OrderNumber", orderNumber)
        BPL.RfcFunctions.RfcTransactionCallUsing(cor3Table, "COR3")
    End Sub

    Function RfcTransactionTable() As System.Data.DataTable
        Dim dataTable As New System.Data.DataTable
        dataTable.Columns.Add("ValueFieldName", GetType(String))
        dataTable.Columns.Add("FieldValue", GetType(String))
        Return dataTable
    End Function


    Public Sub LaunchMD04(materialNumber As String) Implements ICOMCalls.LaunchMD04
        Dim md04Table As System.Data.DataTable
        md04Table = RfcTransactionTable()
        md04Table.Rows.Add("Material", materialNumber.Trim())
        md04Table.Rows.Add("Plant", OrklaRTBPL.SelectionFacade.CapacityLevellingSelectionPlant)
        BPL.RfcFunctions.RfcTransactionCallUsing(md04Table, "MD04")

    End Sub
    Public Sub LoopStockInLevels() Implements ICOMCalls.LoopStockInLevels
        Call StockSimulation.LoopStockInLevels()
    End Sub
    Public Sub LoopForecasts() Implements ICOMCalls.LoopForecasts       
        Call StockSimulation.LoopForecasts()
    End Sub
    Public Sub LoopAllLevels() Implements ICOMCalls.LoopAllLevels
        Call StockSimulation.LoopAllLevels()
    End Sub
    Public Sub CreateForecast() Implements ICOMCalls.CreateForecast
        Call StockSimulation.CreateForecast()
    End Sub
    Public Sub CreatePlainForecast() Implements ICOMCalls.CreatePlainForecast
        Call StockSimulation.CreatePlainForecast()
    End Sub
    Public Sub CreateActualForecast() Implements ICOMCalls.CreateActualForecast
        Call StockSimulation.CreateActualForecast()
    End Sub
    Public Sub ShowStockIn() Implements ICOMCalls.ShowStockIn
        Call StockSimulation.ShowStockIn()
    End Sub
    Public Sub ShowForecast1() Implements ICOMCalls.ShowForecast1
        Call StockSimulation.ShowForecast1()
    End Sub
    Public Sub ShowCurrentSim() Implements ICOMCalls.ShowCurrentSim
        Call StockSimulation.ShowCurrentSim()
    End Sub
    Public Sub ShowActualFC() Implements ICOMCalls.ShowActualFC
        Call StockSimulation.ShowActualFC()
    End Sub
    Public Sub ShowMRP() Implements ICOMCalls.ShowMRP
        Call StockSimulation.ShowMRP()
    End Sub
    Public Sub ShowDailyUsage() Implements ICOMCalls.ShowDailyUsage
        Call StockSimulation.ShowDailyUsage()
    End Sub

    Public Sub PurchasingCockpitGet_Requsitions() Implements ICOMCalls.PurchasingCockpitGet_Requsitions

        Dim c As Excel.Range
        Dim p As String
        Dim r As Excel.Range
        Dim f As Object
        Dim x As Integer
        Dim pi As Excel.PivotItem
        Dim bolPItem As Boolean


        Application.Sheets("Cockpit").Range("RequsitionRange").ClearContents()
        If Application.Range("cbRequsitions").Value = 0 Then Exit Sub

        x = Application.Range("RequsitionRange").Row - Application.Range("Periods").Row

        On Error GoTo CleanUp
        For Each pi In Application.Sheets("Pvt_Requsitions").PivotTables(1).PivotFields("MRP elmnt ind.").PivotItems
            If pi.SourceName = "BA" Then
                bolPItem = True
                Exit For
            End If
        Next pi

        If bolPItem = True Then
            Application.Sheets("Pvt_Requsitions").PivotTables(1).PivotFields("MRP elmnt ind.").CurrentPage = "BA"
        Else
            GoTo CleanUp
        End If

        If Application.Range("Plant").Value <> "All" Then
            p = Application.Range("Plant").Value
        Else
            p = "Total"
        End If

        If p <> "Total" Then
            With Application.Sheets("Pvt_Requsitions").PivotTables(1).RowRange
                c = .Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    r = c
                End If
            End With
        Else
            r = Application.Range("A1")
        End If

        For Each f In Application.Range("Periods")
            With Application.Sheets("Pvt_Requsitions").PivotTables(1).PivotFields("Period")
                c = .DataRange.Find(f, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)
                If Not c Is Nothing Then
                    f.Offset(x, 0) = Application.Sheets("Pvt_Requsitions").Cells(r.Row, c.Column)
                End If
            End With
        Next f

        On Error Resume Next
        c = Application.Range("Plant_List").Columns(1).Find("x", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)

CleanUp:
        c = Nothing
        r = Nothing
        pi = Nothing

        Exit Sub

    End Sub


    Public Sub PurchasingCockpitGet_OpenQty() Implements ICOMCalls.PurchasingCockpitGet_OpenQty

        Dim c As Excel.Range
        Dim p As String
        Dim r As Excel.Range
        Dim f As Excel.Range
        Dim pi As Excel.PivotItem
        Dim bolPItem As Boolean

        Application.Range("Contracts").ClearContents()

        On Error GoTo CleanUp
        For Each pi In Application.Sheets("Pvt_OpenQty").PivotTables(1).PivotFields("MRP elmnt ind.").PivotItems
            If pi.SourceName = "K" Then
                bolPItem = True
                Exit For
            End If
        Next pi

        If bolPItem = True Then
            Application.Sheets("Pvt_OpenQty").PivotTables(1).PivotFields("MRP elmnt ind.").CurrentPage = "K"
        Else
            GoTo CleanUp
        End If

        If Application.Range("Plant").Value <> "All" Then
            p = Application.Range("Plant").Value
        Else
            p = "Total"
        End If

        On Error Resume Next
        If p <> "Total" Then
            With Application.Sheets("Pvt_OpenQty").PivotTables(1).RowRange
                c = .Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    r = c
                End If
            End With
        Else
            r = Application.Range("A1")
        End If
        Application.Range("Contracts").Value = Application.Sheets("Pvt_OpenQty").Cells(r.Row, 2).Value

        On Error Resume Next
        c = Application.Range("Plant_List").Columns(1).Find("xyz", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)



CleanUp:
        c = Nothing
        f = Nothing
        r = Nothing
        pi = Nothing

        Exit Sub

    End Sub

    Public Sub PurchasingCockpitGet_SafetyTime() Implements ICOMCalls.PurchasingCockpitGet_SafetyTime

        Dim c As Excel.Range
        Dim t As String
        Dim p As Excel.Range


        Application.Range("Saf_Time").ClearContents()
        If Application.Sheets("Lists").Range("cbSafetyTime").Value = False Then Exit Sub

        For Each p In Application.Range("Plant_List")
            With Application.Sheets("Material")
                c = Application.Range("ZPC_STOCK").Columns(2).Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    p.Offset(10, 0).Value = c.Offset(0, 24).Value
                End If
            End With
        Next p

        c = Application.Sheets("Material").Range("ZPC_STOCK").Columns(1).Find("x", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)

CleanUp:
        c = Nothing
        p = Nothing

        Exit Sub

    End Sub
    Public Sub PurchasingCockpitGet_Requirements() Implements ICOMCalls.PurchasingCockpitGet_Requirements

        Dim c As Excel.Range
        Dim p As String
        Dim r As Excel.Range
        Dim f As Excel.Range
        Dim x As Integer

        Application.Range("Require_Range").ClearContents()
        x = Application.Range("Require_Range").Row - Application.Range("Periods").Row
        If Application.Range("Plant").Value <> "All" Then
            p = Application.Range("Plant").Value
        Else
            p = "Total"
        End If

        If p <> "Total" Then
            With Application.Sheets("Pvt_Req").PivotTables(1).RowRange
                c = .Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    r = c
                End If
            End With
        Else
            r = Application.Range("A1")
        End If

        For Each f In Application.Range("Periods")
            With Application.Sheets("Pvt_Req").PivotTables(1).PivotFields("Periode")
                c = .DataRange.Find(f, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)
                If Not c Is Nothing Then
                    f.Offset(x, 0).Value = -Application.Sheets("Pvt_Req").Cells(r.Row, c.Column).Value
                End If
            End With
        Next f

CleanUp:
        c = Nothing
        f = Nothing
        r = Nothing

        Exit Sub

    End Sub
    Public Sub PurchasingCockpitGet_GRTime() Implements ICOMCalls.PurchasingCockpitGet_GRTime

        Dim c As Excel.Range
        Dim t As String
        Dim p As Excel.Range

        Application.Sheets("Cockpit").Range("ProcTime").ClearContents()

        For Each p In Application.Range("Plant_List")
            With Application.Sheets("Material")
                c = .Range("ZPC_STOCK").Columns(2).Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    p.Offset(11, 0).Value = c.Offset(0, 12).Value
                End If
            End With
        Next p

        On Error Resume Next
        c = Application.Sheets("Material").Range("ZPC_STOCK").Columns(1).Find("x", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)


CleanUp:
        c = Nothing
        p = Nothing

        Exit Sub

    End Sub
    Public Sub PurchasingCockpitGet_SafStock() Implements ICOMCalls.PurchasingCockpitGet_SafStock
        Dim c As Excel.Range
        Dim t As String
        Dim p As Excel.Range

        For Each p In Application.Range("Plant_List")
            With Application.Sheets("Material")
                c = .Range("ZPC_STOCK").Columns(2).Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    p.Offset(9, 0).Value = c.Offset(0, 18).Value
                End If
            End With
        Next p

        On Error Resume Next
        c = Application.Sheets("Material").Range("ZPC_STOCK").Columns(1).Find("x", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)


CleanUp:
        c = Nothing
        p = Nothing

        Exit Sub
    End Sub

    Public Sub PurchasingCockpitSet_Formulas() Implements ICOMCalls.PurchasingCockpitSet_Formulas

        Dim c As Excel.Range
        Dim x As Integer
        Dim z As Integer
        Dim r1 As Integer
        Dim r2 As Integer
        Dim y As Double
        Dim y1 As Double
        Dim strFormula As String


        r1 = Application.Range("Unres_Stock").Row
        r2 = Application.Range("Saf_Stock").Row

        If Application.Range("Plant").Value = "All" Then
            z = Application.Range("Plant").Column
            Application.ActiveWorkbook.Names.Add(Name:="UnresStockValue", RefersToR1C1:="=Cockpit!R" & r1 & "C" & z)
            Application.ActiveWorkbook.Names.Add(Name:="SafStockValue", RefersToR1C1:="=Cockpit!R" & r2 & "C" & z)
        Else
            c = Application.Range("Plant_List").Find(Application.Range("Plant"), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
            z = c.Column
            Application.ActiveWorkbook.Names.Add(Name:="UnresStockValue", RefersToR1C1:="=Cockpit!R" & r1 & "C" & z)
            Application.ActiveWorkbook.Names.Add(Name:="SafStockValue", RefersToR1C1:="=Cockpit!R" & r2 & "C" & z)
        End If

        Application.Range("Stock_In").FormulaR1C1 = "=UnresStockValue-SafStockValue*cbSafStock"

CleanUp:
        c = Nothing

        Exit Sub

    End Sub

    Sub PurchasingCockpitSet_PlantRefresh() Implements ICOMCalls.PurchasingCockpitSet_PlantRefresh

        Dim strField As String
        Dim strDescr As String
        Dim c As Excel.Range
        Dim fs As Object
        Dim f As Object
        Dim p As Excel.Range


        Application.EnableEvents = False

        Application.Sheets("Cockpit").Unprotect("next")

        Application.StatusBar = "Checking Input Values..."
        If OrklaRTBPL.SelectionFacade.PurchaseCockpitSelectionMaterial = "" Then
            MsgBox("Material number has to be filled out.", , "Orkla Purchasing Cockpit")
            Exit Sub
        End If

        If Application.Range("Plant").Value <> "All" Then
            c = Application.Range("Plant_List").Find(Application.Range("Plant"), LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
            Else
                MsgBox("Not a valid plant. Choose valid plant or 'All'.", , "Orkla Purchasing Cockpit")
            End If
        End If
        On Error Resume Next
        c = Application.Range("Plant").Columns(1).Find("x", LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
        On Error GoTo 0

        If OrklaRTBPL.SelectionFacade.PurchaseCockpitSelectionYear = "" Then
            'Application.Range("Year").Value = Year(DateTime.Now) need to be fixed
        End If

        Call PurchasingCockpit.Get_Requsitions()
        Call PurchasingCockpit.Get_POs()
        Call PurchasingCockpit.Get_Requirements()
        Call PurchasingCockpit.Get_Budget()
        Call PurchasingCockpit.Get_Consumption()

        If Application.Range("cbSafStock").Value = True Then
            Call PurchasingCockpit.Get_SafStock()
        Else
            Application.Range("Saf_Stock").ClearContents()
        End If

        If Application.Range("cbRequsitions").Value = True Then
            Call PurchasingCockpit.Get_Requsitions()
        Else
            Application.Range("RequsitionRange").ClearContents()
        End If

        If Application.Range("cbContracts").Value = True Then
            Call PurchasingCockpit.Get_OpenQty()
        Else
            Application.Range("Contracts").ClearContents()
        End If
        Call PurchasingCockpit.Set_Formulas()

        Application.Sheets("Cockpit").Protect("next")

CleanUp:
        Application.EnableEvents = True
        c = Nothing
        fs = Nothing
        f = Nothing
        p = Nothing

        Exit Sub

    End Sub
















    '    Sub MixingPlanWriteNewStart()
    '        Dim x As Integer
    '        Dim A As String, B As String
    '        Dim r As Excel.Range

    '        r = Application.Sheets("OrderStart").Cells(1, 1)
    '        For x = 1 To 5000
    '            If r.Offset(x, 0).Value <> "" Then
    '                OrklaRTBPL.ReportSpecific.InsertMPOrderStart(r.Offset(x, 0).Value.ToString(), r.Offset(x, 1).Value.ToString(), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
    '                'Globals.Ribbons.OrklaRT.GetLockedOrders()
    '            Else
    '                Exit For
    '            End If
    '        Next x
    '        MixingPlanRefreshNewStart()
    '    End Sub

    '    Sub MixingPlanWriteMixWC()

    '        Dim x As Integer
    '        Dim r As Excel.Range
    '        Dim y As Integer

    '        y = 0
    '        r = Application.Sheets("Settings").Range("MixWC1").Cells(1, 1)
    '        For x = 1 To 5000
    '            If r.Offset(x, 0).Value <> "" Then
    '                OrklaRTBPL.ReportSpecific.InsertMPMixWC(Convert.ToInt32(r.Offset(x, 0).Value), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
    '            Else
    '                y = y + 1
    '                If y = 10 Then Exit For
    '            End If
    '        Next x

    '    End Sub

    '    Sub MixingPlanWriteMixPlan()

    '        Dim x As Integer
    '        Dim r As Excel.Range

    '        r = Application.Sheets("MixPlan").Cells(1, 1)
    '        For x = 1 To 5000
    '            If r.Offset(x, 0).Value <> "" Then
    '                OrklaRTBPL.ReportSpecific.InsertMPMixPlan(r.Offset(x, 0).Value.ToString(), Convert.ToInt32(r.Offset(x, 1).Value), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
    '            Else
    '                Exit For
    '            End If
    '        Next x

    '    End Sub

    '    Sub MixingPlanWritePriPlan()

    '        Dim x As Integer
    '        Dim r As Excel.Range

    '        r = Application.Sheets("PriPlan").Cells(1, 1)
    '        For x = 1 To 5000
    '            If r.Offset(x, 0).Value <> "" Then
    '                OrklaRTBPL.ReportSpecific.InsertMPPriPlan(r.Offset(x, 0).Value.ToString(), Convert.ToInt32(r.Offset(x, 1).Value), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
    '            Else
    '                Exit For
    '            End If
    '        Next x

    '    End Sub

    '    Sub MixingPlanWriteRSTest()

    '        Dim x As Integer
    '        Dim r As Excel.Range

    '        r = Application.Sheets("RSTest").Cells(1, 1)
    '        For x = 1 To 5000
    '            If r.Offset(x, 0).Value <> "" Then
    '                OrklaRTBPL.ReportSpecific.InsertMPRSTest(r.Offset(x, 0).Value.ToString(), r.Offset(x, 1).Value.ToString(), OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant, gUserId)
    '            Else
    '                Exit For
    '            End If
    '        Next x
    '    End Sub

    '    Sub MixingPlanRefreshNewStart()
    '        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("OrderStart").ListObjects
    '            If listObject.Name.Equals("tOrderStart") Then
    '                Try
    '                    If Not listObject.DataBodyRange Is Nothing Then
    '                        listObject.DataBodyRange.Delete()
    '                    End If
    '                    Dim orderStart = OrklaRTBPL.ReportSpecific.GetMPOrderStart(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
    '                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(orderStart.Tables(0))
    '                    data.MoveFirst()
    '                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, orderStart.Tables(0).Rows.Count, orderStart.Tables(0).Columns.Count)
    '                Catch
    '                End Try
    '            End If
    '        Next
    '    End Sub
    '    Sub MixingPlanRefreshMixWC()

    '        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("Settings").ListObjects
    '            If listObject.Name.Equals("tMixWC") Then
    '                Try
    '                    If Not listObject.DataBodyRange Is Nothing Then
    '                        listObject.DataBodyRange.Delete()
    '                    End If
    '                    Dim mixWC = OrklaRTBPL.ReportSpecific.GetMPMixWC(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
    '                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(mixWC.Tables(0))
    '                    data.MoveFirst()
    '                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, mixWC.Tables(0).Rows.Count, mixWC.Tables(0).Columns.Count)
    '                Catch
    '                End Try
    '            End If
    '        Next

    '    End Sub


    '    Sub MixingPlanRefreshMixPlan()

    '        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("MixPlan").ListObjects
    '            If listObject.Name.Equals("tMixPlan") Then
    '                Try
    '                    If Not listObject.DataBodyRange Is Nothing Then
    '                        listObject.DataBodyRange.Delete()
    '                    End If
    '                    Dim mixPlan = OrklaRTBPL.ReportSpecific.GetMPMixPlan(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
    '                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(mixPlan.Tables(0))
    '                    data.MoveFirst()
    '                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, mixPlan.Tables(0).Rows.Count, mixPlan.Tables(0).Columns.Count)
    '                Catch
    '                End Try
    '            End If
    '        Next

    '    End Sub

    '    Sub MixingPlanRefreshPriPlan()

    '        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("PriPlan").ListObjects
    '            If listObject.Name.Equals("tPriPlan1") Then
    '                Try
    '                    If Not listObject.DataBodyRange Is Nothing Then
    '                        listObject.DataBodyRange.Delete()
    '                    End If
    '                    Dim priPlan = OrklaRTBPL.ReportSpecific.GetMPMixPlan(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
    '                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(priPlan.Tables(0))
    '                    data.MoveFirst()
    '                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, priPlan.Tables(0).Rows.Count, priPlan.Tables(0).Columns.Count)
    '                Catch
    '                End Try
    '            End If
    '        Next

    '    End Sub

    '    Sub MixingPlanRefreshRSTest()

    '        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("RSTest").ListObjects
    '            If listObject.Name.Equals("tRSTest") Then
    '                Try
    '                    If Not listObject.DataBodyRange Is Nothing Then
    '                        listObject.DataBodyRange.Delete()
    '                    End If
    '                    Dim rsTest = OrklaRTBPL.ReportSpecific.GetMPMixPlan(OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant)
    '                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(rsTest.Tables(0))
    '                    data.MoveFirst()
    '                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, rsTest.Tables(0).Rows.Count, rsTest.Tables(0).Columns.Count)
    '                Catch
    '                End Try
    '            End If
    '        Next

    '    End Sub

    '    Public Sub MixingPlanFindOrder(lngOrder As String, intMach As String)
    '        Dim c As Excel.Range

    '        Call MixingPlanRefreshMixPlan()

    '        c = Application.Sheets("MixPlan").Columns(1).Find(lngOrder, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '        If Not c Is Nothing Then
    '            c.Offset(0, 1).Value = intMach
    '        Else
    '            Application.Sheets("MixPlan").Cells(Application.Sheets("MixPlan").Range("MixPlan").Rows.Count + 1, 1).Value = lngOrder
    '            Application.Sheets("MixPlan").Cells(Application.Sheets("MixPlan").Range("MixPlan").Rows.Count + 1, 2).Value = intMach
    '        End If

    '        Call MixingPlanWriteMixPlan()
    '        Call MixingPlanRefreshMixPlan()
    '        Application.Sheets(Application.Sheets("Version").Range("FirstSheet").Value).PivotTables(1).PivotCache.Refresh()

    'CleanUp:

    '    End Sub


    '    Sub MixingPlanFindPri(lngOrder As Long, intPri As Integer)

    '        Dim c As Excel.Range

    '        Call MixingPlanRefreshPriPlan()

    '        c = Application.Sheets("PriPlan").Columns(1).Find(lngOrder, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '        If Not c Is Nothing Then
    '            c.Offset(0, 1).Value = intPri
    '        Else
    '            Application.Sheets("PriPlan").Cells(Application.Sheets("PriPlan").Range("PriPlan1").Rows.Count + 1, 1).Value = lngOrder
    '            Application.Sheets("PriPlan").Cells(Application.Sheets("PriPlan").Range("PriPlan1").Rows.Count + 1, 2).Value = intPri
    '        End If

    '        Call MixingPlanWritePriPlan()
    '        Call MixingPlanRefreshPriPlan()
    '        '   ThisWorkbook.Sheets(fnfirstsheet(ThisWorkbook)).PivotTables(1).PivotCache.Refresh

    'CleanUp:

    '    End Sub


    '    Sub MixingPlanFindRS(lngOrder As Long, strRS As Object)
    '        Dim c As Excel.Range

    '        Call MixingPlanRefreshRSTest()

    '        c = Application.Sheets("RSTest").Columns(1).Find(lngOrder, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '        If Not c Is Nothing Then
    '            c.Offset(0, 1).Value = strRS
    '        Else
    '            Application.Sheets("RSTest").Cells(Application.Sheets("RSTest").Range("RSTest").Rows.Count + 1, 1).Value = lngOrder
    '            Application.Sheets("RSTest").Cells(Application.Sheets("RSTest").Range("RSTest").Rows.Count + 1, 2).Value = strRS
    '        End If

    '        Call MixingPlanWriteRSTest()
    '        Call MixingPlanRefreshRSTest()

    'CleanUp:

    '    End Sub


    '    Sub MixingPlanFindNewStart(lngOrder As Long, lngDate As Date)
    '        Dim c As Excel.Range

    '        '   Call RefreshNewStart

    '        With Application.Sheets("OrderStart")
    '            c = .UsedRange.Columns(1).Find(lngOrder, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
    '            If Not c Is Nothing Then
    '                If strEditComm = "delete" Then
    '                    c.EntireRow.Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
    '                Else
    '                    c.Offset(0, 1).Value = lngDate
    '                End If
    '            Else
    '                .Cells(2, 1).EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown)
    '                .Cells(2, 1).Value = lngOrder
    '                .Cells(2, 2).Value = lngDate
    '            End If
    '        End With

    '        '   Call WriteNewStart
    '        '   Call RefreshNewStart
    '        '   ThisWorkbook.Sheets(fnfirstsheet(ThisWorkbook)).PivotTables(1).PivotCache.Refresh

    'CleanUp:

    '    End Sub

End Class
