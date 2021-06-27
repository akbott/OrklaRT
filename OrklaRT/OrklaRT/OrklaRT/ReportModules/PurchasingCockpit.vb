
Imports System.IO
Module PurchasingCockpit
   
    Sub LocalUpdate(path As String)
        Dim c As Excel.Range
        Dim fs As Object
        Dim f As Object
        Dim x As Integer
        Dim sh As Excel.Worksheet
        'Dim aiArray As Object
        'Dim iColumn As Integer
        Dim qryActive As Excel.QueryTable


        Application.Sheets("Cockpit").Unprotect("next")

        Application.ActiveWorkbook.Activate()
        Application.Sheets("Cockpit").Range("Plant") = "All"

        Application.StatusBar = "Refreshing Purchasing Cockpit..."

        For Each sh In Application.Sheets
            If sh.QueryTables.Count > 0 Then
                qryActive = sh.QueryTables(1)
                qryActive.Connection = "TEXT;" + path + qryActive.Connection.ToString().Split("\").GetValue((qryActive.Connection.ToString().Split("\").Length) - 1)
            End If
        Next sh

        Call RefreshQueryTable("Stock")
        Call RefreshQueryTable("Orders")
        Call RefreshQueryTable("Database")
        Call RefreshQueryTable("Material")
        Call RefreshQueryTable("Consumption")
        Call RefreshQueryTable("Budget")
        
        Application.ActiveWorkbook.RefreshAll()

        Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Call Present_Plants()
        Application.Calculate()
        Call Get_Stocks()
        Application.Calculate()
        Call Get_Controller()
        Application.Calculate()
        Call Get_POs()
        Application.Calculate()
        Call Get_Budget()
        Application.Calculate()
        Call Get_Consumption()
        Application.Calculate()

        x = 0
        For Each c In Application.Sheets("Database").Range("Requirements").Columns(18).Cells
            x = x + 1
            If x > 1 Then
                c.Value = fnCalendarDay(c.Offset(0, -16).Value, c.Offset(0, -10).Value, -c.Offset(0, 7).Value)
            End If
        Next c
        x = 0
        For Each c In Application.Sheets("Orders").Range("AllOrders").Columns(2).Cells
            x = x + 1
            If c.Offset(0, -1).Value = "BA" And x > 1 Then
                c.Value = fnCalendarDay(c.Offset(0, 8).Value, c.Offset(0, 19).Value, -c.Offset(0, 23).Value)
            End If
        Next c


        Application.Calculate()
        Application.Sheets("Pvt_Req").PivotTables(1).PivotCache.Refresh()
        Application.Calculate()
        Call Get_Requirements()
        Application.Calculate()

        If Application.Range("cbSafStock").Value = True Then
            Call Get_SafStock()
        Else
            Application.Range("Saf_Stock").ClearContents()
        End If

        If Application.Range("cbSafetyTime").Value = True Then
            Call Get_SafetyTime()
        Else
            Application.Range("Saf_Time").ClearContents()
        End If

        If Application.Range("cbGRTime").Value = True Then
            Call Get_GRTime()
        Else
            Application.Range("ProcTime").ClearContents()
        End If

        Application.Calculate()
        If Application.Range("cbRequsitions").Value = True Then
            Application.Sheets("Pvt_Requsitions").PivotTables(1).PivotCache.Refresh()
            Application.Calculate()
            Call Get_Requsitions()
        Else
            Application.Range("RequsitionRange").ClearContents()
        End If

        Application.Calculate()
        If Application.Range("cbContracts").Value = True Then
            Call Get_OpenQty()
        Else
            Application.Range("Contracts").ClearContents()
        End If

        Application.Calculate()
        Call Set_Formulas()
        Call Get_Contracts()


        'Application.Sheets("Cockpit").Shapes("ShapeQuery").Delete()

        'Call Protection_On()
        Application.Sheets("Cockpit").protect("next")
        Application.Sheets("Cockpit").Activate()

CleanUp:
        c = Nothing
        fs = Nothing
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        '   Application.EnableEvents = True

        Exit Sub
    End Sub

    Sub Present_Plants()
        Dim strPlants As String
        Dim x As Integer
        Dim c As Excel.Range

        Application.Range("Plant_List").ClearContents()
        Application.Range("Plant_List").ClearComments()

        With Application.Sheets("Stock").Range("ZPC_STOCK").Columns(2)
            x = 2
            Do While Not .Cells(x, 1).Value = ""
                strPlants = strPlants & .Cells(x, 1).Value & " - "
                Application.Range("Start_Plant").Offset(0, x - 2).Value = .Cells(x, 1).Value
                x = x + 1
            Loop
        End With

        If Application.Sheets("Material").Range("AJ1").Offset(1, 0).Value <> "" Then
            Application.Range("Matname").Value = String.Format("{0:000000}", Convert.ToInt32(OrklaRTBPL.SelectionFacade.PurchaseCockpitSelectionMaterial)) & " " & Application.Sheets("Material").Range("AJ1").Offset(1, 0).Value
        Else
            Application.Range("Matname").Value = ""
        End If
        Application.Range("MatGrp").Value = Application.Sheets("Material").Range("AE1").Offset(1, 0).Value & " " & Application.Sheets("Material").Range("AG1").Offset(1, 0).Value
        Application.Range("LeadBuyer").Value = Application.Sheets("Material").Range("AL1").Offset(1, 0).Value
        Application.Range("LeadDeveloper").Value = Application.Sheets("Material").Range("AN1").Offset(1, 0).Value

        Application.ActiveWorkbook.Names.Add(Name:="Plant_List", RefersTo:=Application.Range(Application.Range("Start_Plant"), Application.Range("Start_Plant").Offset(0, x - 3)))
        If x > 3 Then
            Application.Range("Plant_List").Sort(Key1:=Application.Range("Start_Plant"), Order1:=Excel.XlSortOrder.xlAscending, Header:=Excel.XlYesNoGuess.xlNo, OrderCustom:=1, MatchCase:=False, Orientation:=Excel.Constants.xlLeftToRight, DataOption1:=Excel.XlSortDataOption.xlSortNormal)
        End If

        Application.Range("Plant_List_Valid").ClearContents()
        Application.Range("Plant_List_All").Value = "All"

        x = 1
        For Each c In Application.Range("Plant_List")
            Application.Range("Plant_List_All").Offset(x, 0).Value = c
            x = x + 1
        Next c
        Application.ActiveWorkbook.Names.Add(Name:="Plant_List_Valid", RefersTo:=Application.Range(Application.Range("Plant_List_All"), Application.Range("Plant_List_All").Offset(x - 1, 0)))

CleanUp:
        c = Nothing

        Exit Sub
    End Sub


    Sub LocalPrepareSaving()

        Dim strPlants As String
        Dim x As Integer
        Dim sh As Excel.Worksheet
     

        Application.Sheets("Cockpit").Unprotect("next")

        Application.ActiveWorkbook.Sheets("Cockpit").Activate()
        'Application.Range("MatNo").ClearContents()
        'Application.Range("Plant").Value = "All"
        'Application.Range("Year").Value = Year(DateTime.Now)

        Application.Range("Plant_List").ClearContents()
        Application.Range("Plant_List").ClearComments()
        Application.Range("Stocks").ClearContents()
        Application.Range("Saf_Stock").ClearContents()
        Application.Range("Saf_Time").ClearContents()
        Application.Range("ProcTime").ClearContents()
        Application.Range("DelivTime").ClearContents()
        Application.Range("Consump_Range").ClearContents()
        Application.Range("Bud_Range").ClearContents()
        Application.Range("Matname").ClearContents()
        Application.Range("MatGrp").ClearContents()
        Application.Range("LeadBuyer").ClearContents()
        Application.Range("LeadDeveloper").ClearContents()
        Application.Range("Last_Update").ClearContents()
        Application.Range("Stock_In").ClearContents()
        Application.Range("ListContracts").ClearContents()
        Application.Range("Require_Range").ClearContents()
        Application.Range("PO_Range").ClearContents()
        Application.Range("RequsitionRange").ClearContents()
        Application.Sheets("Cockpit").Activate()
        Application.Range("MatNo").Select()

        For Each sh In Application.ActiveWorkbook.Sheets
            If sh.PivotTables.Count > 0 Then
                sh.PivotTables(1).PivotCache.Refresh()
            End If
        Next sh

CleanUp:
        Application.Sheets("Cockpit").Protect("next")
        sh = Nothing

        Exit Sub

    End Sub
   
    Sub Field_Descr()

        Dim strField As String
        Dim strDescr As String
        Dim c As Excel.Range


        strDescr = String.Empty
        strField = Application.ActiveCell.Name
        With Application.Sheets("Descriptions")
            c = .Range("Descriptions").Columns(1).Find(strField, LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                strDescr = c.Value & ": " & c.Offset(0, 1).Value
                Application.StatusBar = strDescr
            Else
                Application.StatusBar = False
            End If

            On Error Resume Next
            c = .Range("Descriptions").Columns(1).Find("xyz", LookIn:=Excel.XlFindLookIn.xlValues, LookAt:=Excel.XlLookAt.xlPart)
        End With

CleanUp:
        c = Nothing

        Exit Sub
    End Sub
  
    Sub Plant_Refresh()
       
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

        Call Get_Requsitions()
        Call Get_POs()
        Call Get_Requirements()
        Call Get_Budget()
        Call Get_Consumption()

        If Application.Range("cbSafStock").Value = True Then
            Call Get_SafStock()
        Else
            Application.Range("Saf_Stock").ClearContents()
        End If

        If Application.Range("cbRequsitions").Value = True Then
            Call Get_Requsitions()
        Else
            Application.Range("RequsitionRange").ClearContents()
        End If

        If Application.Range("cbContracts").Value = True Then
            Call Get_OpenQty()
        Else
            Application.Range("Contracts").ClearContents()
        End If
        Call Set_Formulas()

        Application.Sheets("Cockpit").Protect("next")

CleanUp:
        Application.EnableEvents = True
        c = Nothing
        fs = Nothing
        f = Nothing
        p = Nothing

        Exit Sub

    End Sub

    Sub Set_Formulas()
        
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

    Sub Get_Stocks()

        Dim c As Excel.Range
        Dim p As Excel.Range

        Application.Range("Stocks").ClearContents()

        On Error Resume Next

        For Each p In Application.Range("Plant_List")
            With Application.Sheets("Stock")
                c = .Range("ZPC_STOCK").Columns(2).Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    p.Offset(1, 0).Value = c.Offset(0, 1).Value
                    p.Offset(2, 0).Value = c.Offset(0, 3).Value
                    p.Offset(3, 0).Value = c.Offset(0, 2).Value
                    p.Offset(5, 0).Value = c.Offset(0, 5).Value
                    p.Offset(6, 0).Value = 0
                Else
                    p.Offset(1, 0).Value = 0
                    p.Offset(2, 0).Value = 0
                    p.Offset(3, 0).Value = 0
                    p.Offset(5, 0).Value = 0
                    p.Offset(6, 0).Value = 0
                End If
            End With
        Next p

        c = Application.Sheets("Stock").Range("Descriptions").Columns(1).Find("x", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)
        On Error GoTo 0

CleanUp:
        c = Nothing
        p = Nothing

        Exit Sub

    End Sub

    Sub Get_Contracts()

        Dim r As Excel.Range
        Dim x As Integer
        Dim c As Excel.Range
        Dim t As String
        Dim p As Excel.Range
      

        Application.Range("ListContracts").ClearContents()
        Application.Range("ListContracts").ClearComments()

        c = Application.Range("ListContracts").Cells(1, 1)

        x = 0
        For Each r In Application.Range("AllOrders").Rows
            If r.Cells(1, 1).Offset(0, 0).Value = "K" Then
                c.Offset(x, 0).Value = r.Cells(1, 1).Offset(0, 0).Value
                c.Offset(x, 1).Value = r.Cells(1, 1).Offset(0, 2).Value
                c.Offset(x, 2).Value = r.Cells(1, 1).Offset(0, 9).Value
                c.Offset(x, 3).Value = r.Cells(1, 1).Offset(0, 10).Value & " " & r.Cells(1, 1).Offset(0, 11).Value
                c.Offset(x, 4).Value = r.Cells(1, 1).Offset(0, 17).Value
                c.Offset(x, 5).Value = r.Cells(1, 1).Offset(0, 18).Value
                If c.Offset(x, 4).Value <> 0 Then c.Offset(x, 6).Value = CDbl(c.Offset(x, 5).Value) / CDbl(c.Offset(x, 4).Value)
                c.Offset(x, 7).Value = r.Cells(1, 1).Offset(0, 13).Value
                c.Offset(x, 8).Value = r.Cells(1, 1).Offset(0, 16).Value
                c.Offset(x, 9).Value = r.Cells(1, 1).Offset(0, 15).Value
                c.Offset(x, 10).Value = r.Cells(1, 1).Offset(0, 14).Value
                c.Offset(x, 11).Value = r.Cells(1, 1).Offset(0, 3).Value
                c.Offset(x, 12).Value = r.Cells(1, 1).Offset(0, 4).Value
                c.Offset(x, 13).Value = r.Cells(1, 1).Offset(0, 23).Value
                c.Offset(x, 14).Value = r.Cells(1, 1).Offset(0, 21).Value
                x = x + 1
            End If
        Next r

        For Each r In Application.Range("AllOrders").Rows
            If r.Cells(1, 1).Offset(0, 0).Value = "BE" Or r.Cells(1, 1).Offset(0, 0).Value = "LA" Then
                c.Offset(x, 0).Value = r.Cells(1, 1).Offset(0, 0).Value
                c.Offset(x, 1).Value = r.Cells(1, 1).Offset(0, 7).Value
                c.Offset(x, 2).Value = r.Cells(1, 1).Offset(0, 9).Value
                c.Offset(x, 3).Value = r.Cells(1, 1).Offset(0, 10).Value & " " & r.Cells(1, 1).Offset(0, 11).Value
                c.Offset(x, 4).Value = r.Cells(1, 1).Offset(0, 17).Value
                c.Offset(x, 5).Value = r.Cells(1, 1).Offset(0, 18).Value
                c.Offset(x, 6).Value = 0
                c.Offset(x, 7).Value = r.Cells(1, 1).Offset(0, 13).Value
                c.Offset(x, 8).Value = r.Cells(1, 1).Offset(0, 16).Value
                c.Offset(x, 9).Value = r.Cells(1, 1).Offset(0, 15).Value
                c.Offset(x, 10).Value = r.Cells(1, 1).Offset(0, 14).Value
                c.Offset(x, 11).Value = r.Cells(1, 1).Offset(0, 3).Value
                c.Offset(x, 12).Value = r.Cells(1, 1).Offset(0, 4).Value
                c.Offset(x, 13).Value = r.Cells(1, 1).Offset(0, 23).Value
                If Right(r.Cells(1, 1).Offset(0, 21).Value, 4) = "0000" Then r.Cells(1, 1).Offset(0, 21).ClearContents()
                If r.Cells(1, 1).Offset(0, 21).Value > 0 Then c.Offset(x, 14).Value = r.Cells(1, 1).Offset(0, 21).Value
                x = x + 1
            End If
        Next r

        GoTo CleanUp

        For Each r In Application.Range("AllOrders").Rows
            If r.Cells(1, 1).Offset(0, 0).Value = "BA" Then
                c.Offset(x, 0).Value = r.Cells(1, 1).Offset(0, 0).Value
                c.Offset(x, 1).Value = r.Cells(1, 1).Offset(0, 2).Value
                c.Offset(x, 2).Value = r.Cells(1, 1).Offset(0, 9).Value
                c.Offset(x, 3).Value = r.Cells(1, 1).Offset(0, 10).Value
                c.Offset(x, 4).Value = r.Cells(1, 1).Offset(0, 17).Value
                c.Offset(x, 5).Value = r.Cells(1, 1).Offset(0, 18).Value
                c.Offset(x, 6).Value = 0
                c.Offset(x, 7).Value = r.Cells(1, 1).Offset(0, 13).Value
                c.Offset(x, 8).Value = r.Cells(1, 1).Offset(0, 16).Value
                c.Offset(x, 9).Value = r.Cells(1, 1).Offset(0, 15).Value
                c.Offset(x, 10).Value = r.Cells(1, 1).Offset(0, 14).Value
                c.Offset(x, 11).Value = r.Cells(1, 1).Offset(0, 3).Value
                c.Offset(x, 12).Value = r.Cells(1, 1).Offset(0, 4).Value
                c.Offset(x, 13).Value = r.Cells(1, 1).Offset(0, 22).Value
                If Int(r.Cells(1, 1).Offset(0, 21).Value) > 0 Then c.Offset(x, 14).Value = r.Cells(1, 1).Offset(0, 21).Value
                x = x + 1
            End If
        Next r

CleanUp:
        c = Application.Sheets("Material").Range("ZPC_STOCK").Columns(1).Find("xyz", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)

        r = Nothing
        c = Nothing
        p = Nothing

        Exit Sub

    End Sub

    Sub Get_Controller()
        Dim c As Excel.Range
        Dim t As String
        Dim p As Excel.Range

        Application.Range("Plant_List").ClearComments()
        Application.Range("Saf_Stock").ClearContents()
        Application.Range("Saf_Time").ClearContents()
        Application.Range("ProcTime").ClearContents()
        Application.Range("DelivTime").ClearContents()

        On Error Resume Next

        For Each p In Application.Range("Plant_List")
            With Application.Sheets("Material")
                c = .Range("ZPC_STOCK").Columns(2).Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    p.Offset(9, 0).Value = c.Offset(0, 18).Value
                    p.Offset(10, 0).Value = c.Offset(0, 24).Value
                    p.Offset(11, 0).Value = c.Offset(0, 12).Value
                    '            p.Offset(12, 0) = c.Offset(0, 11)
                    t = c.Offset(0, 10).Value
                    p.AddComment(t)
                End If
            End With
        Next p

        c = Application.Sheets("Material").Range("ZPC_STOCK").Columns(1).Find("x", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)

CleanUp:

        Exit Sub

    End Sub


    Sub Get_SafStock()
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

    Sub Get_SafetyTime()

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
    Sub Get_GRTime()

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

    Sub Get_Requirements()
   
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
  
    Sub Get_POs()

        Dim c As Excel.Range
        Dim p As String
        Dim r As Excel.Range
        Dim f As Excel.Range
        Dim x As Integer
        Dim pi As Excel.PivotItem
        Dim bolPItem As Boolean

        Application.Range("PO_Range").ClearContents()

        bolPItem = False
        On Error Resume Next
        For Each pi In Application.Sheets("Pvt_Orders").PivotTables(1).PivotFields("MRP elmnt ind.").PivotItems
            If pi.SourceName = "BE" Or pi.SourceName = "LA" Then
                pi.Visible = True
                bolPItem = True
            Else
                pi.Visible = False
            End If
        Next pi

        If bolPItem = True Then
            Application.Sheets("Pvt_Orders").PivotTables(1).PivotFields("MRP elmnt ind.").CurrentPage = "(All)"
        Else
            GoTo CleanUp
        End If

        x = Application.Range("PO_Range").Row - Application.Range("Periods").Row

        If Application.Range("Plant").Value <> "All" Then
            p = Application.Range("Plant").Value
        Else
            p = "Total"
        End If

        If p <> "Total" Then
            With Application.Sheets("Pvt_Orders").PivotTables(1).RowRange
                c = .Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    r = c
                End If
            End With
        Else
            r = Application.Range("A1") ' Ask Bjørn
        End If

        For Each f In Application.Range("Periods")
            With Application.Sheets("Pvt_Orders").PivotTables(1).PivotFields("Period")
                c = .DataRange.Find(f, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)
                If Not c Is Nothing Then
                    f.Offset(x, 0).Value = Application.Sheets("Pvt_Orders").Cells(r.Row, c.Column).Value
                End If
            End With
        Next f

CleanUp:
        c = Nothing
        f = Nothing
        r = Nothing
        pi = Nothing

        Exit Sub

    End Sub

    Sub Get_OpenQty()
    
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
  
    Sub Get_Budget()

        Dim c As Excel.Range
        Dim p As String
        Dim r As Excel.Range
        Dim f As Excel.Range
        Dim x As Integer

        Application.Range("Bud_Range").ClearContents()
        x = Application.Range("Bud_Range").Row - Application.Range("Periods").Row

        If Application.Range("Plant").Value <> "All" Then
            p = Application.Range("Plant").Value
        Else
            p = "Total"
        End If

        If p <> "Total" Then
            With Application.Sheets("Pvt_Budget").PivotTables(1).RowRange
                c = .Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    r = c
                End If
            End With
        Else
            r = Application.Range("A1")
        End If

        For Each f In Application.Range("Periods")
            With Application.Sheets("Pvt_Budget").PivotTables(1).PivotFields("Periode")
                '         Debug.Print f
                c = .DataRange.Find(f, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)
                If Not c Is Nothing Then
                    f.Offset(x, 0).Value = Application.Sheets("Pvt_Budget").Cells(r.Row, c.Column).Value
                End If
            End With
        Next f

        On Error Resume Next
        c = Application.Range("Plant_List").Columns(1).Find("x", LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)

CleanUp:
        c = Nothing
        f = Nothing
        r = Nothing

        Exit Sub

    End Sub
  
    Sub Get_Consumption()
   
        Dim c As Excel.Range
        Dim p As String
        Dim r As Excel.Range
        Dim f As Integer
        Dim x As Integer
        Dim x1 As Integer
        Dim x2 As Integer


        Application.Range("Consump_Range").ClearContents()

        x2 = Application.Range("Consump_Range").Row - Application.Range("Periods").Row

        If Application.Range("Plant").Value <> "All" Then
            p = Application.Range("Plant").Value
        Else
            p = "Total"
        End If

        If p <> "Total" Then
            With Application.Sheets("Consumption")
                c = .Range("ZPC_STOCK").Columns(2).Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                If Not c Is Nothing Then
                    r = c
                End If
            End With
        Else
            r = Application.Range("A2")
        End If

        f = 0
        x = 0
        With Application.Sheets("Consumption")
            f = Int(Right(Application.Range("Periods").Cells(1, 1).Value, 2)) + 2
            For x = 1 To 15 - f
                Application.Range("Periods").Cells(1, 1).Offset(x2, x - 1).Value = Application.Sheets("Consumption").Cells(r.Row, f + x - 1).Value
            Next
            For x1 = 1 To 12
                If Application.Range("Periods").Cells(1, 1).Offset(0, x + x1 - 2).Value = 0 Then Exit Sub
                Application.Range("Periods").Cells(1, 1).Offset(x2, x + x1 - 2).Value = Application.Sheets("Consumption").Cells(r.Row + 1, x1 + 2).Value
            Next
        End With

        c = Application.Sheets("Consumption").Range("ZPC_STOCK").Columns(2).Find(p, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)

CleanUp:
        c = Nothing
        r = Nothing

        Exit Sub
    End Sub

    Sub Get_Requsitions()

        Dim c As Excel.Range
        Dim p As String
        Dim r As Excel.Range
        Dim f As Object
        Dim x As Integer
        Dim pi As Excel.PivotItem
        Dim bolPItem As Boolean


        Application.Sheets("Cockpit").Range("RequsitionRange").ClearContents()
        If Application.Range("cbRequsitions").Value = False Then Exit Sub

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

    Sub ShowMatDetail(strPlant As String)
      
        Dim c As Excel.Range
        Dim x As Integer
        Dim s As String
        Dim r As Integer
     
        With Application.Sheets("Material")
            c = .Range("ZPC_STOCK").Columns(2).Find(strPlant, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                r = c.Row
            End If

            On Error Resume Next
            c = .Range("ZPC_STOCK").Columns(2).Find(strPlant, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlPart)
        End With
       

        s = ""
        For Each c In Application.Sheets("Material").Range("ZPC_STOCK").Columns
            s = s & c.Cells(1, 1) & ": " & c.Cells(r, 1) & Chr(10)
        Next
        MsgBox(s, , "Orkla SAP Integration")

CleanUp:
        c = Nothing

        Exit Sub

    End Sub
  
    '    Sub CreateValList()

    '        Dim strValList As String
    '        Dim c As Excel.Range

    '        If Application.Range("Plant_List").Cells.Count > 1 Then
    '            strValList = "All"
    '            For Each c In Application.Range("Plant_List")
    '                strValList = strValList & ";" & c.Value
    '            Next c
    '        Else
    '            strValList = Application.Range("Plant_List").Value
    '        End If

    '        With Application.Range("Plant").Validation
    '            .Delete()
    '            .Add(Type:=Excel.XlDVType.xlValidateList, AlertStyle:=Excel.XlDVAlertStyle.xlValidAlertStop, Operator:=Excel.XlFormatConditionOperator.xlBetween, Formula1:=strValList)
    '            .IgnoreBlank = True
    '            .InCellDropdown = True
    '                   .InputTitle = "Choose which Plant to view Stock period data for."
    '            .ErrorTitle = ""
    '            .InputMessage = ""
    '            .ErrorMessage = ""
    '            .ShowInput = True
    '            .ShowError = True
    '        End With

    'CleanUp:
    '        c = Nothing

    '        Exit Sub

    '    End Sub
  
'    Sub LocalVisibleFields()
    
'        Dim sh As Excel.Worksheet
'        Dim target As Excel.Range
'        Dim isect As Excel.Range


'        Application.Sheets("Cockpit").Unprotect("next")
'        Application.Sheets("Fields").Range("Input").ClearContents()
'        Application.Sheets("Fields").Range("Input").ClearFormats()

'        With Application.Sheets("Fields").Range("Input")
'            isect = Application.Intersect(Application.Sheets("Cockpit").Range("ListContracts"), target)
'            If Not isect Is Nothing Then
'                .Cells(1, 1).Value = Application.Sheets("Cockpit").Range("ListContracts").Cells(1, 1).Offset(-1, target.Column - 1).Value
'                .Cells(1, 2).Value = target.EntireRow.Cells(1, 1).Offset(0, 0).Value
'            End If
'        End With

'        Application.Sheets("Cockpit").Protect("next")

'CleanUp:
'        sh = Nothing

'        Exit Sub
'    End Sub

    Sub Protection_On()
      
        Application.ActiveSheet.Protect(Password:="next", DrawingObjects:=True, _
                    Contents:=True, Scenarios:=True, _
                    AllowFiltering:=True, AllowUsingPivotTables:=True, _
                    AllowSorting:=True)
        
    End Sub

    '    '—————————————————————————————————————————————————————————————————————————————
    '    Sub Protection_Off()
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        '  Scope   : Turns protection off.
    '        '  Author  : Bjørn Tømmerbakk.
    '        '  Date    : 10.01.2006.
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        '  Local constants and variabel declarations:
    '        '  ———————————————————————————————————————————————————————————————————————————

    '        '   Exit Sub
    '        ActiveSheet.Unprotect Password:="next"
    '        '   thisworkbook.Unprotect Password:="next"
    '        Application.EnableCancelKey = xlDisabled
    '        '  ———————————————————————————————————————————————————————————————————————————
    '    End Sub
    '    '—————————————————————————————————————————————————————————————————————————————
    '    Private Sub Worksheet_Change(ByVal target As Range)
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        '  Scope    :
    '        '  Comments :
    '        '  Author   : Bjørn Tømmerbakk.
    '        '  Date     : 13.06.2008.
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        '  Declaration of local constants and variables:
    '        Dim u
    '        '  ———————————————————————————————————————————————————————————————————————————
    '        On Error GoTo CleanUp

    '        Application.ScreenUpdating = False
    '        Application.EnableEvents = False    'To avoid event loop.
    '        On Error GoTo CleanUp

    '        Select Case target.Name.Name
    '            Case "MatNo", "Year", "LangCode", "Scenario"
    '                Sheets("Selections").Unprotect Password:="next"
    '                Sheets("Selections").Range("Home").Value = Range("MatNo").Value
    '                Sheets("Selections").Range("Home").Offset(1, 0).Value = Range("Year").Value
    '                Sheets("Selections").Range("Home").Offset(2, 0).Value = Range("Scenario").Value
    '                Sheets("Selections").Range("Home").Offset(3, 0).Value = Range("LangCode").Value
    '                Sheets("Selections").Protect Password:="next"
    '            Case "Plant"
    '                Call Plant_Refresh()
    '        End Select

    'CleanUp:
    '        '   Application.StatusBar = "Refresh time: " & Now - u
    '        Application.StatusBar = False
    '        Application.EnableEvents = True
    '        Application.ScreenUpdating = True
    '        '  ———————————————————————————————————————————————————————————————————————————
    '    End Sub
    '—————————————————————————————————————————————————————————————————————————————


End Module
