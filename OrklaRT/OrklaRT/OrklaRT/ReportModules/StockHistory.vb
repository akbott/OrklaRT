Module StockHistory
    Public mdTable = MDDataTable()
    Function MDDataTable() As System.Data.DataTable

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
        Dim pi As Excel.PivotItem
        Dim bolFirstStock As Boolean
        Dim TestDate As Date
        Dim lngToday As Long
        Dim mappedColumns As String()
        Dim TListObject As Microsoft.Office.Tools.Excel.ListObject
        Dim rfcTable As SAP.Middleware.Connector.IRfcTable
        Dim Date01 As Date
        Dim Date02 As Date
        Dim Date03 As Date

        Application.Calculation = Excel.XlCalculation.xlCalculationManual
        'ThisWorkbook.Activate()

        'Call MD_STOCK_REQUIREMENTS_LIST_API(ThisWorkbook, _
        '      Sheets("Selections").Range("Home").Offset(2, 0).Value, _
        '      Sheets("Selections").Range("Home").Value, _
        '      "MD04_Data")

        rfcTable = RfcFunctions.GetMDSTOCKREQUIREMENTSLISTAPI(OrklaRTBPL.SelectionFacade.StockHistorySelectionMaterial, OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant)
        mdTable = MDDataTable()

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

                mdTable.Rows.Add(rfcTable(j).GetValue("PLAAB"), rfcTable(j).GetValue("SORT1"), rfcTable(j).GetValue("SORT2"), rfcTable(j).GetValue("DELKZ"), rfcTable(j).GetValue("PLUMI"), Date01, Date02, rfcTable(j).GetValue("DELB0"),
                                   rfcTable(j).GetValue("EXTRA"), rfcTable(j).GetValue("AUSSL"), rfcTable(j).GetValue("AUSKT"), rfcTable(j).GetValue("MNG01"), Date03, rfcTable(j).GetValue("WRK01"), rfcTable(j).GetValue("WRK02"),
                                   rfcTable(j).GetValue("SLOC2"), rfcTable(j).GetValue("LGORT"), rfcTable(j).GetValue("LIFNR"), rfcTable(j).GetValue("KUNNR"), rfcTable(j).GetValue("MD4KD"), rfcTable(j).GetValue("MD4LI"))
            Next j
        End If

        If mdTable.Rows.Count > 0 Then
            For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Application.ActiveWorkbook.Sheets("MD04_Data").ListObjects
                If listObject.Name.Equals("MD04Data") Then
                    Try
                        mappedColumns = New String() {}
                        ReDim mappedColumns(listObject.ListColumns.Count - 1)
                        For Each col In listObject.ListColumns
                            If col.Index - 1 < mdTable.Columns.Count Then
                                mdTable.Columns(col.Index - 1).ColumnName = col.Name
                                mappedColumns(col.Index - 1) = col.Name
                            Else
                                mappedColumns(col.Index - 1) = String.Empty
                            End If
                        Next
                        TListObject = Globals.Factory.GetVstoObject(listObject)
                        TListObject.SetDataBinding(mdTable, String.Empty, mappedColumns)
                        TListObject.RefreshDataRows()
                        TListObject.Disconnect()
                        Application.StatusBar = String.Format("Data Successfully Loaded {0}", mdTable.Rows.Count)
                    Catch
                    End Try
                End If
            Next
        End If

        Application.StatusBar = "Calculating days of Coverage, please wait..."
        Application.EnableEvents = False
        Application.Sheets("SourceData").Unprotect("next")

        'Sheets("Y084_MatData").QueryTables(1).Refresh BackgroundQuery:=True
        'Sheets("Y084_StockData").QueryTables(1).Refresh BackgroundQuery:=True

        'Application.Sheets("Y084_StockData").Activate()
        'For Each r In Application.Sheets("Y084_StockData").Range("StockData").Rows
        '    If UCase(r.Cells(1, 2).Value) = UCase(OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant) Then
        '        r.Range(r.Cells(r.Row, 2), r.Cells(r.Row, 7)).Copy(r.Range("B2"))
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
        Call fnCalendarDay(OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant, Date.Now.Date, 300)
        x = 0
        For y = 1 To 600
            If CInt(Mid(strfactorycal, y, 1)) = 1 Then
                x = x + 1
                Application.Range("rngToDay").Offset(-x, 0).Value = CDate(Application.Range("rngToDay").Value).AddDays(y)
                Application.Range("rngToDay").Offset(-x, 10).Value = CInt(Mid(strFactoryCal, y, 1))
                If Application.Range("rngToDay").Offset(-x, 0).Row = 2 Then Exit For
            End If
        Next y

        lngDate = DateTime.Now.Date.ToOADate() - CDate(OrklaRTBPL.SelectionFacade.StockHistorySelectionFromDate).Date.ToOADate()
        Call fnCalendarDay(OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant, Date.Now.Date, -300)
        '   Debug.Print strfactorycal
        Application.Range("rngToDay").Offset(0, 10).Value = 1
        x = 0
        For y = 1 To lngDate
            If CInt(Mid(strfactorycal, y, 1)) = 1 Then
                x = x + 1
                Application.Range("rngToDay").Offset(x, 0).Value = CDate(Application.Range("rngToDay").Value).AddDays(-y)
                Application.Range("rngToDay").Offset(x, 10).Value = CInt(Mid(strFactoryCal, y, 1))
            End If
        Next y
        Application.Range("NumbDays").Value = x

        y = Application.Range("rngToDay").Row
        Application.ActiveWorkbook.Names.Add(Name:="Dates", RefersToR1C1:= _
            "=SourceData!R" & y & "C1:R" & y + x & "C1")

        'Application.Sheets("Database").Sort.SortFields.Add(Key:=Application.Range("AQ2"), SortOn:=Excel.XlSortOn.xlSortOnValues, Order:=Excel.XlSortOrder.xlDescending, DataOption:=Excel.XlSortDataOption.xlSortNormal)

        Dim lO As Microsoft.Office.Interop.Excel.ListObject = Application.Sheets("Database").ListObjects("SapEXlData")
        TListObject = Globals.Factory.GetVstoObject(lO)

        TListObject.Sort.SortFields.Add(Application.Sheets("Database").Range("AR2", "AR" & Application.Range("SapEXlData").Rows.Count - 1), Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlDescending)
        TListObject.Sort.Apply()

        If OrklaRTBPL.SelectionFacade.StockHistorySelectionMaterial.Equals("502323") Then
            row = lO.ListRows(lO.ListRows.Count - 1).Range
            Dim NewRow = TListObject.ListRows.AddEx(AlwaysInsert:=True)
            For ind = 1 To TListObject.ListColumns.Count - 7
                NewRow.Range(1, ind).Value = row.Cells(1, ind).Value
            Next ind
            TListObject.ListRows(TListObject.ListRows.Count - 2).Range.Delete()
        End If


        Application.Sheets("Database").Range(Application.Range("NewData").Offset(1, 0), Application.Range("NewData").Offset(Application.Range("SapEXlData").Rows.Count - 1, 0)).FormulaR1C1 = "=DATEVALUE(RC[-31])"

        For y = 1 To Application.Range("Database").Rows.Count - 1
            TestDate = New Date(CInt(Application.Sheets("Database").Range("PostDate").Offset(y, 0).Value.ToString().Split(".")(2)), CInt(Application.Sheets("Database").Range("PostDate").Offset(y, 0).Value.ToString().Split(".")(1)), CInt(Application.Sheets("Database").Range("PostDate").Offset(y, 0).Value.ToString().Split(".")(0)))
            Do While fnWorkDaysBetween(OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant, TestDate.AddDays(-1), TestDate) < 1
                TestDate = TestDate.AddDays(1)
            Loop
            If TestDate > New Date(CInt(Application.Sheets("Database").Range("PostDate").Offset(y, 0).Value.ToString().Split(".")(2)), CInt(Application.Sheets("Database").Range("PostDate").Offset(y, 0).Value.ToString().Split(".")(1)), CInt(Application.Sheets("Database").Range("PostDate").Offset(y, 0).Value.ToString().Split(".")(0))) Then
                '         Sheets("Database").Activate
                Application.Sheets("Database").Range("NewData").Offset(y, 0).Value = TestDate
            End If
        Next y

        Application.Sheets("SourceData").Calculate()
        'Debug.Print Now; "slutt calendar"

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

            If DateValue(d.Offset(x, 0).Value) >= DateValue(OrklaRTBPL.SelectionFacade.StockHistorySelectionFromDate) Then
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
                    d.Offset(x, 1).Value = Application.Range("TotStock").Offset(1, 0).Value 'Quantity of valuated stock.
                End If
            End If
        Next x

        '  Find Safety Stock.
        Application.Sheets("SourceData").Range("SafStock").ClearContents()
        c = Application.Sheets("Y084_MatData").Columns(2).Find(OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
        If Not c Is Nothing Then
            Application.Sheets("SourceData").Range("SafStock").Value = c.Offset(0, 18).Value
        End If

        '  Get usage per day.
        For Each d In Application.Range("Dates")
            If DateValue(d.Value) < DateValue(OrklaRTBPL.SelectionFacade.StockHistorySelectionFromDate) Then Exit For
            If d.Offset(0, 1).Value.ToString() = "" Then
                d.Offset(0, 2).Value = ""
                Exit For
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

        Application.ActiveWorkbook.Names.Add(Name:="X_Range", RefersToR1C1:=
            "=SourceData!R" & x & "C1:R" & lngToday + Application.Range("NumbDays").Value & "C1")
        Application.ActiveWorkbook.Names.Add(Name:="Stock_Range", RefersToR1C1:=
            "=SourceData!R" & x & "C2:R" & lngToday + Application.Range("NumbDays").Value & "C2")
        Application.ActiveWorkbook.Names.Add(Name:="Cov_Range", RefersToR1C1:=
            "=SourceData!R" & x & "C4:R" & lngToday + Application.Range("NumbDays").Value & "C4")
        Application.ActiveWorkbook.Names.Add(Name:="Req_Range", RefersToR1C1:=
            "=SourceData!R" & x & "C6:R" & lngToday + Application.Range("NumbDays").Value & "C6")
        Application.ActiveWorkbook.Names.Add(Name:="Future_Range", RefersToR1C1:=
            "=SourceData!R" & x & "C5:R" & lngToday + Application.Range("NumbDays").Value & "C5")
        Application.ActiveWorkbook.Names.Add(Name:="SafStock_Range", RefersToR1C1:=
            "=SourceData!R" & lngToday & "C13:R" & lngToday + Application.Range("NumbDays").Value & "C13")
        Application.ActiveWorkbook.Names.Add(Name:="SafDays_Range", RefersToR1C1:=
            "=SourceData!R" & lngToday & "C14:R" & lngToday + Application.Range("NumbDays").Value & "C14")
        Application.ActiveWorkbook.Names.Add(Name:="X1_Range", RefersToR1C1:=
            "=SourceData!R" & lngToday & "C1:R" & lngToday + Application.Range("NumbDays").Value & "C1")
        Application.ActiveWorkbook.Names.Add(Name:="SafDays1_Range", RefersToR1C1:=
            "=SourceData!R" & x & "C14:R" & lngToday + Application.Range("NumbDays").Value & "C14")
        Application.ActiveWorkbook.Names.Add(Name:="Used_Range", RefersToR1C1:=
            "=SourceData!R" & lngToday & "C3:R" & lngToday + Application.Range("NumbDays").Value & "C3")

        Application.ActiveWorkbook.Names.Add(Name:="DataPeriod", RefersToR1C1:=
            "=SourceData!R1C1:R" & lngToday + Application.Range("NumbDays").Value & "C10")

        Application.Sheets("Material Usage Req").PivotTables(1).PivotCache.Refresh()

        Application.Sheets("Chart").ChartObjects("Chart 8").Activate()
        Application.Range("X_Range").NumberFormat = "dd.mm.åååå"
        Application.Range("X_Range").NumberFormatLocal = "dd.mm.åååå"
        Application.Sheets("Chart").ChartObjects("Chart 8").Chart.Refresh()

CleanUp:

        Application.Sheets("SourceData").Protect("next")      
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.StatusBar = String.Empty
        Exit Sub
    End Sub


    Sub ActualCoverage()

        Dim dblStock As Double
        Dim lngToday As Long
        Dim c As Excel.Range
        Dim x As Integer
        Dim y As Integer

        Application.EnableEvents = False
        Application.Sheets("SourceData").Unprotect("next")

        '  Calucalte actual Stock Coverage per day.
        y = 0
        For Each d In Application.Range("Dates")
            y = y + 1
            If DateValue(d.Value) < DateValue(OrklaRTBPL.SelectionFacade.StockHistorySelectionFromDate) Then Exit For
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

End Module
