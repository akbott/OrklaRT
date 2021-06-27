Option Explicit On

Module StockSimulation
    Public mdStockTable = MDSTOCKDataTable()
    Public shSimulation = Application.Sheets("Simulation")
    Public shLogForecast = Application.Sheets("Log_Forecast")
    Public shLogLLS = Application.Sheets("Log_LLS")
    Public shLogStockIn = Application.Sheets("Log_StockIn")
    Function MDSTOCKDataTable() As System.Data.DataTable

        Dim dataTable As New System.Data.DataTable
        dataTable.Columns.Add("PLAAB", GetType(Int32))
        dataTable.Columns.Add("SORT1", GetType(String))
        dataTable.Columns.Add("SORT2", GetType(String))
        dataTable.Columns.Add("DELKZ", GetType(String))
        dataTable.Columns.Add("PLUMI", GetType(String))
        dataTable.Columns.Add("DAT00", GetType(Date))
        dataTable.Columns.Add("DAT01", GetType(Date))
        dataTable.Columns.Add("DELB0", GetType(String))
        dataTable.Columns.Add("EXTRA", GetType(String))
        dataTable.Columns.Add("AUSSL", GetType(String))
        dataTable.Columns.Add("AUSKT", GetType(String))
        dataTable.Columns.Add("MNG01", GetType(Double))
        dataTable.Columns.Add("DAT02", GetType(Date))
        dataTable.Columns.Add("WRK01", GetType(String))
        dataTable.Columns.Add("WRK02", GetType(String))
        dataTable.Columns.Add("SLOC2", GetType(String))
        dataTable.Columns.Add("LGORT", GetType(String))
        dataTable.Columns.Add("LIFNR", GetType(String))
        dataTable.Columns.Add("KUNNR", GetType(String))
        dataTable.Columns.Add("MD4KD", GetType(String))
        dataTable.Columns.Add("MD4LI", GetType(String))

        Return dataTable
    End Function
    Sub LocalUpdate()
        Dim d As Excel.Range
        Dim c As Excel.Range
        Dim row As Object
        Dim x As Integer
        Dim y As Integer
        Dim s As String
        Dim lngDate As Long
        Dim lngToday As Long
        Dim pi As Excel.PivotItem
        Dim bolFirstStock As Boolean
        Dim TestDate As Date
        Dim dbMaxStock As Double
        'Dim shSimulation As Excel.Worksheet
        Dim mappedColumns As String()
        Dim TListObject As Microsoft.Office.Tools.Excel.ListObject
        Dim rfcTable As SAP.Middleware.Connector.IRfcTable
        Dim Date01 As Date
        Dim Date02 As Date
        Dim Date03 As Date

        Application.Calculation = Excel.XlCalculation.xlCalculationManual
        'Dim ThisWorkbook = Globals.Factory.GetVstoObject(Application.ActiveWorkbook)
        'ThisWorkbook.Activate()

        rfcTable = BPL.RfcFunctions.GetMDSTOCKREQUIREMENTSLISTAPI(OrklaRTBPL.SelectionFacade.StockSimulationSelectionMaterial, OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant)
        mdStockTable = MDSTOCKDataTable()

        If rfcTable.RowCount > 0 Then
            For j As Integer = 0 To rfcTable.RowCount - 1
                rfcTable.CurrentIndex = j

                If rfcTable(j).GetValue("DAT00").ToString() = "0000-00-00" Then
                    Date01 = New DateTime().Date
                Else
                    Date01 = rfcTable(j).GetValue("DAT00")
                End If
                If rfcTable(j).GetValue("DAT01").ToString() = "0000-00-00" Then
                    Date02 = New DateTime().Date
                Else
                    Date02 = rfcTable(j).GetValue("DAT01")
                End If
                If rfcTable(j).GetValue("DAT02").ToString() = "0000-00-00" Then
                    Date03 = New DateTime().Date
                Else
                    Date03 = rfcTable(j).GetValue("DAT02")
                End If


                mdStockTable.Rows.Add(rfcTable(j).GetValue("PLAAB"), rfcTable(j).GetValue("SORT1"), rfcTable(j).GetValue("SORT2"), rfcTable(j).GetValue("DELKZ"), rfcTable(j).GetValue("PLUMI"), Date01, Date02, rfcTable(j).GetValue("DELB0"),
                                   rfcTable(j).GetValue("EXTRA"), rfcTable(j).GetValue("AUSSL"), rfcTable(j).GetValue("AUSKT"), rfcTable(j).GetValue("MNG01"), Date03, rfcTable(j).GetValue("WRK01"), rfcTable(j).GetValue("WRK02"),
                                   rfcTable(j).GetValue("SLOC2"), rfcTable(j).GetValue("LGORT"), rfcTable(j).GetValue("LIFNR"), rfcTable(j).GetValue("KUNNR"), rfcTable(j).GetValue("MD4KD"), rfcTable(j).GetValue("MD4LI"))                
            Next j
        End If

        If mdStockTable.Rows.Count > 0 Then
            For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Application.ActiveWorkbook.Sheets("MD04_Data").ListObjects
                If listObject.Name.Equals("MD04Data") Then
                    Try
                        mappedColumns = New String() {}
                        ReDim mappedColumns(listObject.ListColumns.Count - 1)
                        For Each col In listObject.ListColumns
                            If col.Index - 1 < mdStockTable.Columns.Count Then
                                mdStockTable.Columns(col.Index - 1).ColumnName = col.Name
                                mappedColumns(col.Index - 1) = col.Name
                            Else
                                mappedColumns(col.Index - 1) = String.Empty
                            End If
                        Next
                        TListObject = Globals.Factory.GetVstoObject(listObject)
                        TListObject.SetDataBinding(mdStockTable, String.Empty, mappedColumns)
                        TListObject.RefreshDataRows()
                        TListObject.Disconnect()
                        Application.StatusBar = String.Format("Data Successfully Loaded {0}", mdStockTable.Rows.Count)
                    Catch
                    End Try
                End If
            Next           
        End If


        Application.StatusBar = "Calculating days of Coverage, please wait..."
        Application.EnableEvents = False
        Application.Sheets("SourceData").Unprotect("next")

        'Application.Sheets("Y084_MatData").QueryTables(1).Refresh(True)
        'Application.Sheets("Y084_StockData").QueryTables(1).Refresh(True)

        'Application.Sheets("Y084_StockData").Activate()
        'For Each r In Application.Sheets("Y084_StockData").Range("StockData").Rows
        '    If UCase(r.Cells(1, 2).Value) = UCase(OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant) Then
        '        r.Range(r.Cells(r.Row, 2), r.Cells(r.Row, 7)).Copy(Application.Sheets("Y084_StockData").Range("B2"))
        '        Exit For
        '    End If
        'Next

        '  Find Factory Calendar days.
        'Debug.Print Now; "start calendar"
        Application.Sheets("SourceData").Range("FactCal").ClearContents()
        '   On Error Resume Next
        '   lngDate = Int(Now) + 270 - Sheets("Selections").Range("Home").Offset(1, 0).Value
        lngDate = DateTime.Now.Date.ToOADate() + 400
        '   Call fnCalendarDay(Sheets("Selections").Range("Home").Value, Sheets("Selections").Range("Home").Offset(1, 0).Value, CInt(lngDate))
        Call fnCalendarDay(OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant, Date.Now.Date, 305)
        x = 0
        For y = 1 To 600
            If CInt(Mid(strFactoryCal, y, 1)) = 1 Then
                x = x + 1
                Application.Range("rngToDay").Offset(-x, 0).Value = CDate(Application.Range("rngToDay").Value).AddDays(y)
                Application.Range("rngToDay").Offset(-x, 10).Value = CInt(Mid(strFactoryCal, y, 1))
                If Application.Range("rngToDay").Offset(-x, 0).Row = 2 Then Exit For
            End If
        Next y

        lngDate = DateTime.Now.Date.ToOADate() - CDate(OrklaRTBPL.SelectionFacade.StockSimulationSelectionFromDate).Date.ToOADate()
        Call fnCalendarDay(OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant, Date.Now.Date, -305)
        '   Debug.Print strfactorycal
        Application.Range("rngToDay").Offset(0, 10).Value = 1
        x = 0
        For y = 1 To lngDate
            If CInt(Mid(strFactoryCal, y, 1)) = 1 Then
                x = x + 1
                Application.Range("rngToDay").Offset(x, 0).Value = CDate(Application.Range("rngToDay").Value).AddDays(-y)
                Application.Range("rngToDay").Offset(x, 10).Value = CInt(Mid(strFactoryCal, y, 1))
            End If
        Next y
        Application.Range("NumbDays").Value = x

        y = Application.Range("rngToDay").Row
        Application.ActiveWorkbook.Names.Add(Name:="Dates", RefersToR1C1:= _
            "=SourceData!R" & y & "C1:R" & y + x & "C1")

        'Application.Sheets("Database").Range(Application.Sheets("Database").Range("NewData").Offset(1, 0), Application.Sheets("Database").Range("NewData").Offset(Application.Range("SapExlData").Rows.Count - 1, 0)).FormulaR1C1 = "=DATEVALUE(RC[-29])"


        Dim lO As Microsoft.Office.Interop.Excel.ListObject = Application.Sheets("Database").ListObjects("SapEXlData")
        TListObject = Globals.Factory.GetVstoObject(lO)

        TListObject.Sort.SortFields.Add(Application.Sheets("Database").Range("AQ2", "AQ" & Application.Range("SapEXlData").Rows.Count - 1), Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending)
        TListObject.Sort.Apply()

        If OrklaRTBPL.SelectionFacade.StockSimulationSelectionMaterial.Equals("502323") Then
            row = lO.ListRows(lO.ListRows.Count - 1).Range
            Dim NewRow = TListObject.ListRows.AddEx(AlwaysInsert:=True)
            For ind = 1 To TListObject.ListColumns.Count - 7
                NewRow.Range(1, ind).Value = row.Cells(1, ind).Value
            Next ind
            TListObject.ListRows(TListObject.ListRows.Count - 2).Range.Delete()
        End If

        Application.Sheets("Database").Range(Application.Sheets("Database").Range("NewData").Offset(1, 0), Application.Sheets("Database").Range("NewData").Offset(Application.Sheets("Database").Range("SapEXlData").Rows.Count - 1, 0)).FormulaR1C1 = "=DATEVALUE(RC[-29])"

        For y = 1 To Application.Range("Database").Rows.Count - 1
            TestDate = Application.Sheets("Database").Range("PostDate").Offset(y, 0).Value
            Do While fnWorkDaysBetween(OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant, TestDate.AddDays(-1), TestDate) < 1
                TestDate = TestDate.AddDays(1)
            Loop
            If TestDate > Application.Sheets("Database").Range("PostDate").Offset(y, 0).Value Then
                '         Sheets("Database").Activate
                Application.Sheets("Database").Range("NewData").Offset(y, 0).Value = TestDate
            End If
        Next y

        Application.Sheets("SourceData").Calculate()


        '  Prepare the SourceData sheet.
        Application.Range("DataRange").ClearContents()
        Application.Sheets("Usage").PivotTables(1).PivotCache.Refresh()
        Application.Calculate()

        Application.Sheets("StockOut").PivotTables(1).PivotCache.Refresh()
        Application.Calculate()
        x = 0

        '  Get Stock values at the end of each day.
        Application.Sheets("SourceData").Activate()
        d = Application.Range("RngToDay")
        Application.Range("Req_Today").Value = 0
        bolFirstStock = False

        For x = Application.Range("Dates").Cells.Count - 1 To 0 Step -1

            If DateValue(d.Offset(x, 0).Value) >= DateValue(OrklaRTBPL.SelectionFacade.StockSimulationSelectionFromDate) Then
                If x > 0 Then
                    c = Application.Sheets("StockOut").PivotTables(1).RowRange.Columns(1).Find(CDate(d.Offset(x, 0).Value).Date.ToString("dd.MM.yyyy"), LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                    If Not c Is Nothing Then
                        bolFirstStock = True
                        d.Offset(x, 1).Value = c.Offset(0, 3).Value
                    Else
                        If bolFirstStock = True Then
                            d.Offset(x, 1).Value = d.Offset(x + 1, 1).Value
                        Else
                            '                     d.Offset(x, 1).FormulaR1C1 = "=R[-1]C+R[-1]C[1]"
                            If Not c Is Nothing Then
                                d.Offset(x, 1).Value = Application.Sheets("Database").Range("StockIn").Value
                            Else
                                d.Offset(x, 1).Value = 0
                            End If
                        End If
                    End If
                Else
                    d.Offset(x, 1).Value = Application.Range("TotStock").Offset(1, 0) 'Quantity of valuated stock.
                End If
            End If
        Next x

        '  Find Safety Stock.
        Application.Sheets("SourceData").Range("SafStock").ClearContents()
        c = Application.Sheets("Y084_MatData").Columns(2).Find(OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
        If Not c Is Nothing Then
            Application.Sheets("SourceData").Range("SafStock").Value = c.Offset(0, 18).Value
        End If

        '  Get usage per day.
        For Each d In Application.Range("Dates")
            If DateValue(d.Value) < DateValue(OrklaRTBPL.SelectionFacade.StockSimulationSelectionFromDate) Then Exit For
            If d.Offset(0, 1).Value.ToString() = "" Then
                d.Offset(0, 2).Value = ""
                Exit For
                'If d.Offset(0, 1).Value Is Nothing Then
                '    d.Offset(0, 2).Value = Nothing
                '    Exit For
            End If
            c = Application.Sheets("Usage").PivotTables(1).RowRange.Find(CDate(d.Value).Date.ToString("dd.MM.yyyy"), LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                d.Offset(0, 2).Value = c.Offset(0, 1).Value * -1
            Else
                d.Offset(0, 2).Value = 0
            End If
        Next d

        '  Get MD04 Requirements.
        x = 0
        Application.Sheets("Pvt_Req").PivotTables(1).PivotCache.Refresh()
        Application.Sheets("Pvt_Req").PivotTables(1).PivotFields("PLUMI").CurrentPage = "-"

        For Each d In Application.Range("Dates1")
            x = x + 1
            c = Application.Sheets("Pvt_Req").PivotTables(1).RowRange.Find(CDate(d.Value).Date.ToString("dd.MM.yyyy"), LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
            If Not c Is Nothing Then
                If x < Application.Range("Dates1").Cells.Count Then
                    d.Offset(0, 2).Value = -c.Offset(0, 1).Value
                Else
                    d.Offset(0, 2).Value = -c.Offset(0, 1).Value + d.Offset(0, 2).Value
                    Application.Range("Req_Today").Value = -c.Offset(0, 1).Value
                End If
            Else
                d.Offset(0, 2).Value = 0
            End If
        Next d

        Call ActualCoverage()

        Application.Sheets("SourceData").Range("UnresStock").Value = 0
        For x = 1 To 3
            Application.Sheets("SourceData").Range("UnresStock").Value = Application.Sheets("SourceData").Range("UnresStock").Value _
                              + Application.Sheets("Y084_StockData").Range("ChartRange").Cells(1, 1).Offset(0, x - 1).Value
        Next x

        Application.Sheets("Y084_StockData").Range("ChartRange").NumberFormat = "#,##0"

        Call EmptyStock()

        Application.Calculate()

        lngToday = Application.Range("RngToday").Row
        x = Application.Range("RngToday").Row - 40 'The chart range start row.
        Application.Sheets("Sourcedata").Activate()

        Application.ActiveWorkbook.Names.Add(Name:="X_Range", RefersToR1C1:= _
            "=SourceData!R" & x & "C1:R" & lngToday + Application.Range("NumbDays").Value & "C1")
        Application.ActiveWorkbook.Names.Add(Name:="Stock_Range", RefersToR1C1:= _
            "=SourceData!R" & x & "C2:R" & lngToday + Application.Range("NumbDays").Value & "C2")
        Application.ActiveWorkbook.Names.Add(Name:="Cov_Range", RefersToR1C1:= _
            "=SourceData!R" & x & "C4:R" & lngToday + Application.Range("NumbDays").Value & "C4")
        Application.ActiveWorkbook.Names.Add(Name:="Req_Range", RefersToR1C1:= _
            "=SourceData!R" & x & "C6:R" & lngToday + Application.Range("NumbDays").Value & "C6")
        Application.ActiveWorkbook.Names.Add(Name:="Future_Range", RefersToR1C1:= _
            "=SourceData!R" & x & "C5:R" & lngToday + Application.Range("NumbDays").Value & "C5")
        Application.ActiveWorkbook.Names.Add(Name:="SafStock_Range", RefersToR1C1:= _
            "=SourceData!R" & lngToday & "C13:R" & lngToday + Application.Range("NumbDays").Value & "C13")
        Application.ActiveWorkbook.Names.Add(Name:="SafDays_Range", RefersToR1C1:= _
            "=SourceData!R" & lngToday & "C14:R" & lngToday + Application.Range("NumbDays").Value & "C14")
        Application.ActiveWorkbook.Names.Add(Name:="X1_Range", RefersToR1C1:= _
            "=SourceData!R" & lngToday & "C1:R" & lngToday + Application.Range("NumbDays").Value & "C1")
        Application.ActiveWorkbook.Names.Add(Name:="SafDays1_Range", RefersToR1C1:= _
            "=SourceData!R" & x & "C14:R" & lngToday + Application.Range("NumbDays").Value & "C14")
        Application.ActiveWorkbook.Names.Add(Name:="Used_Range", RefersToR1C1:= _
            "=SourceData!R" & lngToday & "C3:R" & lngToday + Application.Range("NumbDays").Value & "C3")

        Application.ActiveWorkbook.Names.Add(Name:="DataPeriod", RefersToR1C1:= _
            "=SourceData!R1C1:R" & lngToday + Application.Range("NumbDays").Value & "C10")

        Application.Sheets("Material Usage Req").PivotTables(1).PivotCache.Refresh()

        Application.Sheets("SourceData").Range("rSourceData").Copy(Application.Sheets("Sim_Data").Range("rDataSource"))
        Application.Range("crFixedLot").Value = 0
        Application.Calculate()
        Application.Sheets("Simulation").Range("crStockInFrom").Value = Application.Sheets("Simulation").Range("crTotalUsage").Value / 255 _
           * (Application.Sheets("Simulation").Range("crLead_Time").Value + Application.Sheets("Simulation").Range("crSaf_Time").Value)
        dbMaxStock = Application.WorksheetFunction.Max(Application.Sheets("Sim_Data").Range("$B$236:$B$489"))
        shSimulation.Range("crStockInTo").Value = dbMaxStock
        shSimulation.Range("crIn_Stock").Value = Application.Sheets("Sim_Data").Range("crActStockIn").Value

        'shSimulation.Range("crBatchSize").Value = 1
        shSimulation.Range("rSapMasterData").ClearContents()
        Call ReadMARC()
        Application.Calculate()
        Call InitialForecast()
        Call DeleteLogs()
        shSimulation.Activate()
        shSimulation.Range("crSaf_Time").Select()

CleanUp:

        Application.Sheets("SourceData").Protect("next")
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic

        Exit Sub

    End Sub


    Sub ActualCoverage()

        Dim dblStock As Double
        Dim c As Excel.Range
        Dim x As Integer
        Dim y As Integer

        Application.EnableEvents = False
        Application.Sheets("SourceData").Unprotect("next")

        '  Calucalte actual Stock Coverage per day.
        y = 0
        For Each d In Application.Range("Dates")
            y = y + 1
            If DateValue(d.Value) < DateValue(OrklaRTBPL.SelectionFacade.StockSimulationSelectionFromDate) Then Exit For
            If y = 1 Then
                dblStock = d.Offset(0, 1).Value - Application.Range("Req_Today").Value
            Else
                dblStock = d.Offset(0, 1).Value
            End If
            x = 1
            For x = 1 To Application.Range("RngToDay").Row - 2
                If d.Offset(0, 1).Value.ToString() = "" Then
                    d.Offset(0, 3).Value = ""
                    Exit For
                    'If d.Offset(0, 1).Value Is Nothing Then
                    '    d.Offset(0, 3).Value = String.Empty
                    'Exit For
                ElseIf d.Offset(0, 1).Value = 0 Then
                    d.Offset(0, 3).Value = 0
                    Exit For
                End If

                If dblStock < 0 Then
                    If y = 1 Then
                        d.Offset(0, 3).Value = d.Offset(-x + 2, 11).Value - d.Offset(1, 11).Value '(x - 2) / 7 * 5
                        '               Debug.Print d.Offset(-x + 1, 11).Value; d.Offset(-1, 11).Value; d.Address
                    Else
                        If d.Offset(0, 1).Value.ToString() <> "" Then
                            d.Offset(0, 3).Value = d.Offset(-x + 2, 11).Value - d.Offset(0, 11).Value '(x - 1) / 7 * 5
                            '                  Debug.Print d.Offset(-x + 2, 11).Value; d.Offset(0, 11).Value; d.Address
                        End If
                    End If
                    Exit For
                Else
                    If x = Application.Range("RngToDay").Row - 2 Then
                        d.Offset(0, 3).Value = d.Offset(-x + 1, 11).Value - d.Offset(0, 11).Value '(x - 1) / 7 * 5
                        Exit For
                    Else
                        dblStock = dblStock - d.Offset(-x, 2).Value
                    End If
                End If
            Next x
        Next d

        '  Make range
        'Application.Sheets("Y084_StockData").Activate()
        'For Each c In Application.Sheets("Y084_StockData").Range("ChartRange")
        '    If c.Value = 0 Then c.ClearContents()
        'Next c

CleanUp:
        '   Application.EnableEvents = True
        '   Sheets("SourceData").Protect Password:="next"

        Exit Sub
    End Sub


    Sub EmptyStock()
        Dim dblStock As Double

        '   Application.EnableEvents = False
        dblStock = Application.Sheets("SourceData").Range("UnresStock").Value - Application.Range("RngToDay").Offset(0, 2).Value
        If dblStock < 0 Then
            Application.Range("RngToDay").Offset(-1, 4).Value = 0
        Else
            Application.Range("RngToDay").Offset(-1, 4).Value = dblStock
        End If

        For x = Application.Range("Dates1").Cells.Count - 1 To 1 Step -1

            If Application.Range("Dates1").Cells(1, 1).Offset(x - 1, 4).Row = 285 Then Exit For

            dblStock = Application.Range("Dates1").Cells(1, 1).Offset(x - 1, 4).Value - _
                       Application.Range("Dates1").Cells(1, 1).Offset(x - 1, 2).Value
            If dblStock < 0 Then
                dblStock = 0
                Application.Range("Dates1").Cells(1, 1).Offset(x - 2, 4).Value = dblStock
                Exit For
            Else
                Application.Range("Dates1").Cells(1, 1).Offset(x - 2, 4).Value = dblStock
            End If
        Next x
        '   Range("TopRange").Value = x
        '   Application.EnableEvents = False

    End Sub



    Sub LocalPrepareSaving()
        Application.Sheets("SourceData").Unprotect("next")
        Application.Sheets("Y084_StockData").Range("ChartRange").ClearContents()
        Application.Range("DataRange").ClearContents()
        Application.Sheets("Stock Dates").PivotTables(1).PivotCache.Refresh()
        Application.Sheets("Pvt_Req").PivotTables(1).PivotCache.Refresh()
        Application.Sheets("Material Usage Req").PivotTables(1).PivotCache.Refresh()

        Application.Sheets("SourceData").Protect("next")
    End Sub

    'SysCode Module
    Sub InitialForecast()

        '   On Error GoTo CleanUp
        'Copy actual last 12 months into forecast column.
        shSimulation.Range("rForecastManual").Columns(4).Copy()
        shSimulation.Range("rForecastManual").Columns(2).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=False)
        If shSimulation.Range("crSapLT").Value > 0 And IsNumeric(shSimulation.Range("crSapLT").Value) Then
            shSimulation.Range("crLead_Time").Value = shSimulation.Range("crSapLT").Value
        Else
            shSimulation.Range("crLead_Time").Value = 0
        End If

        If Not shSimulation.Range("crSapST").Value Is Nothing Then
            If shSimulation.Range("crSapST").Value > 0 And IsNumeric(shSimulation.Range("crSapST").Value) Then
                shSimulation.Range("crSaf_Time").Value = shSimulation.Range("crSapST").Value
            End If
        Else
            shSimulation.Range("crSaf_Time").Value = 5
        End If

        If IsNumeric(shSimulation.Range("crSapLS").Value) Then
            If shSimulation.Range("crSapLS").Value > 0 Then
                shSimulation.Range("crLot_Size").Value = shSimulation.Range("crSapLS").Value
            Else
                shSimulation.Range("crLot_Size").Value = 20
            End If
        Else
            shSimulation.Range("crLot_Size").Value = 20
        End If

        
        If shSimulation.Range("crSapRV").Value > 0 And IsNumeric(shSimulation.Range("crSapRV").Value) Then
            shSimulation.Range("crRoundingValue").Value = shSimulation.Range("crSapRV").Value
        Else
            shSimulation.Range("crRoundingValue").Value = 1
        End If
        shSimulation.Range("rSimLeadTime").Cells(1, 1).Value = shSimulation.Range("crLead_Time").Value
        shSimulation.Range("rSimLeadTime").Cells(1, 2).Value = shSimulation.Range("crLead_Time").Value

        shSimulation.Range("rSimSafStock").Cells(1, 1).Value = Int(shSimulation.Range("crSaf_Time").Value / 2) + 1
        shSimulation.Range("rSimSafStock").Cells(1, 2).Value = Int(shSimulation.Range("crSaf_Time").Value * 2)
        shSimulation.Range("rSimSafStock").Cells(1, 3).Value = Int((shSimulation.Range("rSimSafStock").Cells(1, 2).Value - shSimulation.Range("rSimSafStock").Cells(1, 1).Value) / 5)
        If shSimulation.Range("rSimSafStock").Cells(1, 3).Value = 0 Then shSimulation.Range("rSimSafStock").Cells(1, 3).Value = 1

        shSimulation.Range("rSimLotSize").Cells(1, 1).Value = Int(shSimulation.Range("crLot_Size").Value / 2) + 1
        shSimulation.Range("rSimLotSize").Cells(1, 2).Value = Int(shSimulation.Range("crLot_Size").Value * 1.5) + 1
        shSimulation.Range("rSimLotSize").Cells(1, 3).Value = Int((shSimulation.Range("rSimLotSize").Cells(1, 2).Value - shSimulation.Range("rSimLotSize").Cells(1, 1).Value) / 5)
        If shSimulation.Range("rSimLotSize").Cells(1, 3).Value = 0 Then shSimulation.Range("rSimLotSize").Cells(1, 3).Value = 1

        shSimulation.Range("crFCTotLY").Value = (shSimulation.Range("crTotalUsageLY").Value + shSimulation.Range("crTotalUsage").Value) / 2

CleanUp:
        Application.CutCopyMode = False
    End Sub

    Sub DeleteLogs()
        On Error GoTo CleanUp
        shLogForecast.Rows("2:100000").ClearContents()
        shLogLLS.Rows("2:100000").ClearContents()
        shLogStockIn.Rows("2:100000").ClearContents()
CleanUp:
    End Sub


    Sub ReadMARC()
        Dim options() As String
        Dim fields() As String
        Dim result() As String
        Dim rfcTable As SAP.Middleware.Connector.IRfcTable
        Dim i As Integer

        options = New String() {"MATNR EQ '" & String.Format("{0:000000000000000000}", Convert.ToInt64(OrklaRTBPL.SelectionFacade.StockSimulationSelectionMaterial)) & "' AND ", "WERKS EQ '" & OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant & "'"}
        fields = New String() {"PLIFZ", "RWPRO", "DISLS", "BSTRF"}

        rfcTable = RfcFunctions.GetRFCREADTABLE("MARC", options, fields)

        If rfcTable.RowCount > 0 Then
            For i = 0 To rfcTable.RowCount - 1
                rfcTable.CurrentIndex = i
                result = rfcTable(i)(0).ToString().Split(New Char() {"=", ";"})
                shSimulation.Range("crSapLT").Value = result(1).Trim()
                shSimulation.Range("crSapST").Value = result(2).Trim()
                shSimulation.Range("crSapLS").Value = result(3).Trim()
                shSimulation.Range("crSapRV").Value = result(4).Trim()
            Next i
        End If
    End Sub

    Sub LoopStockInLevels()
        Dim x As Long
        Dim y As Long
        Dim lStockIn As Long
        Dim lLoops As Long
        Dim rSim As Excel.Range
        Dim rLogValues As Excel.Range
        Dim lInitValue As Long        

        On Error GoTo CleanUp
        Call StopEvents(True, True)        
        shLogStockIn.Rows("2:100000").ClearContents()
        rSim = Application.Range("rSimStockIn")
        rLogValues = Application.Range("rLogValues").Columns(2)
        lInitValue = Application.Range("crIn_Stock").Value

        x = 0
        For lStockIn = rSim.Cells(1, 1).Value To rSim.Cells(1, 2).Value Step rSim.Cells(1, 3).Value
            x = x + 1
            rSim.Cells(1, 1).Offset(0, -2).Value = lStockIn
            rLogValues.Copy()
            shLogStockIn.Cells(1, 1).Offset(x, 0).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=True)
        Next lStockIn
        shSimulation.Range("crIn_Stock").Value = lInitValue
        shSimulation.Select()

CleanUp:
        Application.CutCopyMode = False
        Call ResetAllEvents()
    End Sub

    Sub CreateForecast(Optional bDontStopEvents As Boolean = False)
        Dim rForecast As Excel.Range
        Dim rRow As Excel.Range
        Dim dRandom As Double
        Dim lFcCol As Long
        Dim dTotDev As Double
        Dim lMonthsNotZero As Long
        Const lLastyear2 = 3
        Const lLastyear = 4
        Const lThisYear = 5

        On Error GoTo CleanUp
        If bDontStopEvents = False Then Call StopEvents(True, True, True)

        rForecast = Application.Range("rForecastManual")
        Select Case Application.Range("crFcBase").Value
            Case "LY"
                lFcCol = lLastyear
            Case "LY2"
                lFcCol = lLastyear2
        End Select
        For Each rRow In rForecast.Rows
            dRandom = Application.WorksheetFunction.RandBetween(Int(rRow.Cells(1, lFcCol).Value * 10 * (1 - Application.Range("cMonthErr").Value)), Int(rRow.Cells(1, lFcCol).Value * 10 * (1 + Application.Range("cMonthErr").Value))) / 10
            '      dRandom = dRandom * (1 + Range("cTotErr").Value)
            rRow.Cells(1, 2).Value = dRandom
        Next rRow

        '   GoTo CleanUp

        Application.Calculate()
        '   dRandom = Application.WorksheetFunction.RandBetween(Int(Range("crTotalUsage").Value * 10 * (1 - Range("cTotErr").Value)), Int(Range("crTotalUsage").Value * 10 * (1 + Range("cTotErr").Value))) / 10
        '   dTotDev = rForecast.Columns(lFcCol).Cells(1, 1).Offset(12, 0).Value - dRandom 'Total for selected FC base
        '   dTotDev = Range("crTotFC").Value - Range("crFCTotLY").Value 'Manual total in crFCTotLY
        dTotDev = Application.Range("crFCTotLY").Value / Application.Range("crTotFC").Value 'Manual total in crFCTotLY

        lMonthsNotZero = 0
        For Each rRow In rForecast.Rows
            '      If (rRow.Cells(1, 2).Value + (dTotDev / 12)) > 0 Then lMonthsNotZero = lMonthsNotZero + 1
            rRow.Cells(1, 2).Value = rRow.Cells(1, 2).Value * dTotDev
        Next rRow

        GoTo CleanUp

        For Each rRow In rForecast.Rows
            If (rRow.Cells(1, 2).Value + (dTotDev / 12)) > 0 Then
                rRow.Cells(1, 2).Value = rRow.Cells(1, 2).Value - (dTotDev / lMonthsNotZero)
            End If
            If rRow.Cells(1, 2).Value < 0 Then rRow.Cells(1, 2).Value = 0
        Next rRow
CleanUp:
        If bDontStopEvents = False Then Call ResetAllEvents()
    End Sub

    Sub LoopForecasts()
        Dim lLoop As Long
        Dim rLogValues As Excel.Range
        Dim oSelection As Object
        Dim lInitValue As Long

        '   On Error GoTo CleanUp
        Call StopEvents(True, True, True)
        oSelection = Application.Selection
        shLogForecast.Rows("2:100000").ClearContents()
        rLogValues = Application.Range("rFcastDevLog")
        lInitValue = Application.Range("crSaf_Time").Value

        If Application.Range("crFCSafTime").Value > 0 Then Application.Range("crSaf_Time").Value = Application.Range("crFCSafTime").Value
        For lLoop = 1 To Application.Range("crFcastLoops").Value
            Call CreateForecast(True)
            Application.Calculate()
            rLogValues.Copy()
            shLogForecast.Cells(1, 1).Offset(lLoop, 0).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=True)
        Next lLoop

        shLogForecast.Sort.SortFields.Clear()
        shLogForecast.Sort.SortFields.Add(Key:=Application.Range("M2"), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlAscending, DataOption:=Excel.XlSortDataOption.xlSortNormal)
        With shLogForecast.Sort
            '.SetRange shLogForecast.Cells(1, 1).CurrentRegion
            .Header = Excel.XlYesNoGuess.xlYes
            .MatchCase = False
            .Orientation = Excel.Constants.xlTopToBottom
            .SortMethod = Excel.XlSortMethod.xlPinYin
            .Apply()
        End With
        Application.Calculate()
        Application.Range("rFCLogAvg").Copy()
        shSimulation.Range("rForecastManual").Columns(2).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=False)
        Application.Range("crSaf_Time").Value = lInitValue
        oSelection.Select()

CleanUp:
        Application.CutCopyMode = False
        Call ResetAllEvents()
    End Sub

    '    Sub LoopSafetyStockLevels()
    '        Dim x As Long
    '        Dim y As Long
    '        Dim lStockIn As Long
    '        Dim lLoops As Long
    '        Dim rSim As Excel.Range
    '        Dim rLogValues As Excel.Range
    '        Dim lFirst As Long        

    '        On Error GoTo CleanUp
    '        Call StopEvents(True, True, True)        
    '        shLogSafStock.Rows("2:100000").ClearContents()
    '        rSim = Application.Range("rSimSafStock")
    '        rLogValues = Application.Range("rLogValues").Columns(2)
    '        lFirst = Application.Range("crSaf_Time").Value

    '        x = 0
    '        For lStockIn = rSim.Cells(1, 1).Value To rSim.Cells(1, 2).Value Step rSim.Cells(1, 3).Value
    '            x = x + 1
    '            rSim.Cells(1, 1).Offset(0, -2).Value = lStockIn
    '            rLogValues.Copy()
    '            shLogSafStock.Cells(1, 1).Offset(x, 0).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=True)
    '        Next lStockIn
    '        Application.Range("crSaf_Time").Value = lFirst

    'CleanUp:
    '        Application.CutCopyMode = False
    '        Call ResetAllEvents()
    '    End Sub

    '    Sub LoopLeadTimeLevels()
    '        Dim x As Long
    '        Dim y As Long
    '        Dim lStockIn As Long
    '        Dim lLoops As Long
    '        Dim rSim As Excel.Range
    '        Dim rLogValues As Excel.Range
    '        Dim lFirst As Long

    '        On Error GoTo CleanUp
    '        Call StopEvents(True, True, True)
    '        shLogLeadTime.Rows("2:100000").ClearContents()
    '        rSim = Application.Range("rSimLeadTime")
    '        rLogValues = Application.Range("rLeadTimeLog")
    '        lFirst = Application.Range("crLead_Time").Value

    '        x = 0
    '        For lStockIn = rSim.Cells(1, 1).Value To rSim.Cells(1, 2).Value Step rSim.Cells(1, 3).Value
    '            x = x + 1
    '            rSim.Cells(1, 1).Offset(0, -2).Value = lStockIn
    '            rLogValues.Copy()
    '            shLogLeadTime.Cells(1, 1).Offset(x, 0).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=True)
    '        Next lStockIn
    '        Application.Range("crLead_Time").Value = lFirst

    'CleanUp:
    '        Application.CutCopyMode = False
    '        Call ResetAllEvents()
    '    End Sub

    '    Sub LoopLotSizeLevels()
    '        Dim x As Long
    '        Dim y As Long
    '        Dim lStockIn As Long
    '        Dim lLoops As Long
    '        Dim rSim As Excel.Range
    '        Dim rLogValues As Excel.Range
    '        Dim lFirst As Long

    '        On Error GoTo CleanUp
    '        Call StopEvents(True, True, True)
    '        shLogLotSize.Rows("2:100000").ClearContents()
    '        rSim = Application.Range("rSimLotSize")
    '        rLogValues = Application.Range("rLotSizeLog")
    '        lFirst = Application.Range("crLot_Size").Value

    '        x = 0
    '        For lStockIn = rSim.Cells(1, 1).Value To rSim.Cells(1, 2).Value Step rSim.Cells(1, 3).Value
    '            x = x + 1
    '            rSim.Cells(1, 1).Offset(0, -2).Value = lStockIn
    '            rLogValues.Copy()
    '            shLogLotSize.Cells(1, 1).Offset(x, 0).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=True)
    '        Next lStockIn
    '        Application.Range("crLot_Size").Value = lFirst

    'CleanUp:
    '        Application.CutCopyMode = False
    '        Call ResetAllEvents()
    '    End Sub

    Sub LoopAllLevels()
        Dim x As Long
        Dim lStockIn As Long
        Dim lStockIn1 As Long
        Dim lStockIn2 As Long
        Dim rSim As Excel.Range
        Dim rSim1 As Excel.Range
        Dim rSim2 As Excel.Range
        Dim rLogValues As Excel.Range
        Dim lFirst As Long
        Dim lFirst1 As Long
        Dim lFirst2 As Long

        On Error GoTo CleanUp
        Call StopEvents(True, True)
        shLogLLS.Rows("2:100000").ClearContents()

        rLogValues = Application.Range("rLotSizeLog")
        rSim = Application.Range("rSimLeadTime")
        lFirst = Application.Range("crLead_Time").Value

        rSim1 = Application.Range("rSimSafStock")
        lFirst1 = Application.Range("crSaf_Time").Value

        rSim2 = Application.Range("rSimLotSize")
        lFirst2 = Application.Range("crLot_Size").Value

        x = 0
        For lStockIn = rSim.Cells(1, 1).Value To rSim.Cells(1, 2).Value Step rSim.Cells(1, 3).Value
            For lStockIn1 = rSim1.Cells(1, 1).Value To rSim1.Cells(1, 2).Value Step rSim1.Cells(1, 3).Value
                For lStockIn2 = rSim2.Cells(1, 1).Value To rSim2.Cells(1, 2).Value Step rSim2.Cells(1, 3).Value
                    x = x + 1
                    rSim.Cells(1, 1).Offset(0, -2).Value = lStockIn
                    rSim1.Cells(1, 1).Offset(0, -2).Value = lStockIn1
                    rSim2.Cells(1, 1).Offset(0, -2).Value = lStockIn2
                    rLogValues.Copy()
                    shLogLLS.Cells(1, 1).Offset(x, 0).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=True)
                Next lStockIn2
            Next lStockIn1
        Next lStockIn
        Application.Range("crLead_Time").Value = lFirst
        Application.Range("crSaf_Time").Value = lFirst1
        Application.Range("crLot_Size").Value = lFirst2

        shLogLLS.Sort.SortFields.Clear()
        shLogLLS.Sort.SortFields.Add(Key:=Application.Range("A2"), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlDescending, DataOption:=Excel.XlSortDataOption.xlSortNormal)
        shLogLLS.Sort.SortFields.Add(Key:=Application.Range("B2"), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlAscending, DataOption:=Excel.XlSortDataOption.xlSortNormal)
        shLogLLS.Sort.SortFields.Add(Key:=Application.Range("C2"), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlAscending, DataOption:=Excel.XlSortDataOption.xlSortNormal)
        With shLogLLS.Sort
            '.SetRange shLogLLS.Cells(1, 1).CurrentRegion
            .Header = Excel.XlYesNoGuess.xlYes
            .MatchCase = False
            .Orientation = Excel.Constants.xlTopToBottom
            .SortMethod = Excel.XlSortMethod.xlPinYin
            .Apply()
        End With

CleanUp:
        Application.CutCopyMode = False
        Call ResetAllEvents()
    End Sub

    Sub ShowTextInfo(sInfoShape As String)
        Dim x As Long
        Dim y As Long
        Dim lStockIn As Long
        Dim lLoops As Long
        Dim rSim As Excel.Range
        Dim rLogValues As Excel.Range
        Dim lFirst As Long

        On Error GoTo CleanUp
        If Application.ActiveSheet.Shapes(sInfoShape & "_t").Visible = True Then
            Application.ActiveSheet.Shapes(sInfoShape & "_t").Visible = False
        Else
            Application.ActiveSheet.Shapes(sInfoShape & "_t").Visible = True
        End If
CleanUp:
    End Sub

    Sub ShowStockIn()
        Call ShowTextInfo("sInfoStockIn")
    End Sub

    Sub ShowForecast1()
        Call ShowTextInfo("sInfoForecast1")
    End Sub

    Sub ShowCurrentSim()
        Call ShowTextInfo("sInfoCurrentSim")
    End Sub

    Sub ShowActualFC()
        Call ShowTextInfo("sInfoActualFC")
    End Sub

    Sub ShowMRP()
        Call ShowTextInfo("sInfoMRP")
    End Sub

    Sub ShowDailyUsage()
        Call ShowTextInfo("sInfoSafetyStock")
    End Sub

    Sub UpdateStockInFrom()
        On Error GoTo CleanUp
        shSimulation.Range("crStockInFrom").Value = shSimulation.Range("crTotalUsage").Value / 255 _
           * (shSimulation.Range("crLead_Time").Value + shSimulation.Range("crSaf_Time").Value)
        shSimulation.Range("crIn_Stock").Value = Application.Range("crActStockIn").Value
CleanUp:
    End Sub

    Sub CreatePlainForecast()
        Dim rForecast As Excel.Range
        Dim rRow As Excel.Range
        Dim dTotalDev As Double
        Dim dPrevMonth As Double
        Dim x As Long
        Dim oSelection As Object

        '   On Error GoTo CleanUp
        oSelection = Application.Selection
        Call StopEvents(True, True, True)
        rForecast = Application.Range("rForecastManual")
        dTotalDev = Application.Range("crTotalUsageLY").Value - Application.Range("crTotalUsage").Value
        dPrevMonth = Application.Range("crTotalUsageLY").Value / 12
        For Each rRow In rForecast.Rows
            x = x + 1
            If x > 1 Then
                dPrevMonth = dPrevMonth - (dTotalDev / 11 / 12)
            End If
            rRow.Cells(1, 2).Value = dPrevMonth
        Next rRow

        '   Ark25.Range("rFCTrend1").Copy
        Application.Range("rFCPlain").Copy()
        shSimulation.Range("rForecastManual").Columns(2).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=False)
        oSelection.Select()

CleanUp:
        Call ResetAllEvents()
    End Sub

    Sub CreateActualForecast()
        Dim oSelection As Object

        '   On Error GoTo CleanUp
        Call StopEvents(True, True, True)
        oSelection = Application.Selection
        shSimulation.Range("rForecastManual").Columns(4).Copy()
        shSimulation.Range("rForecastManual").Columns(2).PasteSpecial(Paste:=Excel.XlPasteType.xlPasteValues, Transpose:=False)
        oSelection.Select()

CleanUp:
        Call ResetAllEvents()
    End Sub



End Module
