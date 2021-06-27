Imports System.IO

Module Common
    Public wbIni As Excel.Workbook
    Public strIniName As String
    Public strFactoryCal As String
    'Option Explicit
    'Option Private Module

    Private Declare Function GetUserName Lib "advapi32.dll" Alias _
            "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

    'Sub ResetAllEvents()
    '    Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
    '    Application.ScreenUpdating = True
    '    Application.EnableEvents = True
    '    Application.StatusBar = False
    '    Application.DisplayAlerts = True
    '    Application.EnableCancelKey = Excel.XlEnableCancelKey.xlInterrupt
    'End Sub

    Sub ZoomAdjust(rViewRange As Excel.Range, Optional bHeightAlso As Boolean = False, Optional bButtonMode As Boolean = False, _
                   Optional sngX As Single = 0, Optional sngY As Single = 0)
        Dim x As Long
        Dim y As Long
        Dim z As Long

        Try
            If sngX = Microsoft.VisualBasic.vbEmpty Then sngX = 0.94
            If sngY = Microsoft.VisualBasic.vbEmpty Then sngY = 0.91
            z = Application.ActiveWindow.Zoom
            x = Int((Application.ActiveWindow.Width * sngX / rViewRange.Width) * 100)
            y = Int((Application.ActiveWindow.Height * sngY / rViewRange.Height) * 100)

            Application.ScreenUpdating = False
            Application.EnableEvents = False

            If bHeightAlso Then
                If x <= y Then
                    If x = z And bButtonMode Then
                        Application.ActiveWindow.Zoom = 100
                    Else
                        Application.ActiveWindow.Zoom = x
                    End If
                Else
                    If y = z And bButtonMode Then
                        Application.ActiveWindow.Zoom = 100
                    Else
                        Application.ActiveWindow.Zoom = y
                    End If
                End If
            Else
                If x = z And bButtonMode Then
                    Application.ActiveWindow.Zoom = 100
                Else
                    Application.ActiveWindow.Zoom = x
                End If
            End If

            Application.ScreenUpdating = True
            Application.EnableEvents = True
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Common", gUserId, gReportID)
        End Try

    End Sub
    Public Function fnFindUserName() As String
        Dim s1 As String
        s1 = Space(512)
        GetUserName(s1, Len(s1))
        fnFindUserName = Trim$(s1)
        fnFindUserName = UCase(Left(fnFindUserName, Len(fnFindUserName) - 1))
    End Function



    Public Function fnCalendarDay(strPlant As String, lngStartDate As Date, intWorkDays As Integer) As Date

        Dim c As Excel.Range
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim n As Integer
        Dim strPlantYear As String
        Dim intMonth As Integer
        Dim intDay As Integer
        Dim intWD As Integer
        Dim intWDCounter As Integer
        Dim intCDCounter As Date
        Dim s As String

        Try
            strPlantYear = strPlant & Year(lngStartDate)
            intCDCounter = lngStartDate
            intMonth = Month(lngStartDate)
            intDay = Microsoft.VisualBasic.DateAndTime.Day(lngStartDate) + 1
            intWD = intWorkDays
            intWDCounter = 0

            If intWorkDays = 0 Then GoTo ResumeHere

            If intWorkDays < 0 Then
                n = -1
                intDay = intDay - 2
            Else
                n = 1
            End If

            If wbIni Is Nothing Then
                Try

                    wbIni = Application.Workbooks.Open(systemIni, Password:="next", UpdateLinks:=False, ReadOnly:=True)
                    Application.Windows(wbIni.Name).Visible = False
                    strIniName = wbIni.Name
                Catch ex As Exception
                    strIniName = "SystemIni_07_M.xls"
                End Try
            End If

            c = Application.Workbooks(strIniName).Sheets("Calendar").Range("Index").Find(strPlantYear, Lookin:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)

            s = String.Empty
            Select Case n
                Case 1
                    For z = 0 To 9
                        For x = 0 To 12 - intMonth
                            For y = intDay To Microsoft.VisualBasic.Len(c.Offset(z, -17 + intMonth + x).Text)
                                intCDCounter = intCDCounter.AddDays(1)
                                s = s & c.Offset(z, -17 + intMonth + x).Characters(y, 1).Text
                                If c.Offset(z, -17 + intMonth + x).Characters(y, 1).Text = "1" Then
                                    intWDCounter = intWDCounter + 1
                                    If intWDCounter = intWD Then
                                        GoTo ResumeHere
                                    End If
                                End If
                            Next y
                            intDay = 1
                        Next x
                        intMonth = 1
                    Next z

                Case -1
                    For z = 0 To -9 Step -1
                        For x = 0 To intMonth - 1
                            For y = intDay To 1 Step -1
                                intCDCounter = intCDCounter.AddDays(-1)
                                s = s & c.Offset(z, -17 + intMonth - x).Characters(y, 1).Text
                                If c.Offset(z, -17 + intMonth - x).Characters(y, 1).Text = "1" Then
                                    intWDCounter = intWDCounter + 1
                                    If intWDCounter = System.Math.Abs(intWD) Then
                                        GoTo ResumeHere
                                    End If
                                End If
                            Next y
                            intDay = Microsoft.VisualBasic.Len(c.Offset(z, -17 + intMonth - x - 1).Text)
                        Next x
                        intMonth = 12
                        intDay = Microsoft.VisualBasic.Len(c.Offset(z - 1, -17 + intMonth).Text)
                    Next z
            End Select

            'wbIni.Close()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Common", gUserId, gReportID)
        End Try
ResumeHere:
        strFactoryCal = s
        fnCalendarDay = intCDCounter
        c = Nothing
    End Function

    Public Function fnWorkDaysBetween(strPlant As String, lngStartDate As Date, lngEndDate As Date) As Long

        Dim c As Excel.Range
        Dim x As Integer
        Dim y As Integer
        Dim z As Integer
        Dim n As Integer
        Dim strPlantYear As String
        Dim intMonth As Integer
        Dim intDay As Integer
        Dim intWD As Integer
        Dim intRestOfMonth As Integer
        Dim intWDCounter As Integer
        Dim intCDCounter
        Dim intCDDays As Integer
        Dim s As String

        Try

            strPlantYear = strPlant & Microsoft.VisualBasic.Year(lngStartDate)
            intCDCounter = Microsoft.VisualBasic.Int(lngStartDate.ToOADate)
            intMonth = Microsoft.VisualBasic.DateAndTime.Month(lngStartDate)
            intDay = Microsoft.VisualBasic.DateAndTime.Day(lngStartDate) + 1
            '   intWD = intWorkDays
            intWDCounter = 0
            intCDDays = Microsoft.VisualBasic.Int(lngEndDate.ToOADate) - Microsoft.VisualBasic.Int(lngStartDate.ToOADate)

            If intCDDays = 0 Then GoTo ResumeHere

            If intCDDays < 0 Then
                n = -1
                intDay = intDay - 2
            Else
                n = 1
            End If

            If wbIni Is Nothing Then
                Try
                    wbIni = Application.Workbooks.Open(systemIni, Password:="next", UpdateLinks:=False, ReadOnly:=True)
                    Application.Windows(wbIni.Name).Visible = False
                    strIniName = wbIni.Name
                Catch ex As Exception
                    strIniName = "SystemIni_07_M.xls"
                End Try
            End If

            c = Application.Workbooks(strIniName).Sheets("Calendar").Range("Index").Find(strPlantYear, Lookin:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)

            s = String.Empty
            Select Case n
                Case 1
                    For z = 0 To 9
                        For x = 0 To 12 - intMonth
                            For y = intDay To Microsoft.VisualBasic.Len(c.Offset(z, -17 + intMonth + x).Text)
                                intCDCounter = intCDCounter + 1
                                s = s & c.Offset(z, -17 + intMonth + x).Characters(y, 1).Text
                                If c.Offset(z, -17 + intMonth + x).Characters(y, 1).Text = "1" Then
                                    intWDCounter = intWDCounter + 1
                                End If
                                If intCDCounter = Microsoft.VisualBasic.Int(lngEndDate.ToOADate) Then
                                    GoTo ResumeHere
                                End If
                            Next y
                            intDay = 1
                        Next x
                        intMonth = 1
                    Next z

                Case -1
                    For z = 0 To -9 Step -1
                        For x = 0 To intMonth - 1
                            For y = intDay To 1 Step -1
                                intCDCounter = intCDCounter - 1
                                If c.Offset(z, -17 + intMonth - x).Characters(y, 1).Text = "1" Then
                                    intWDCounter = intWDCounter + 1
                                End If
                                If intCDCounter = Microsoft.VisualBasic.Int(lngEndDate.ToOADate) Then
                                    intWDCounter = -intWDCounter
                                    GoTo ResumeHere
                                End If
                            Next y
                            intDay = Microsoft.VisualBasic.Len(c.Offset(z, -17 + intMonth - x - 1).Text)
                        Next x
                        intMonth = 12
                        intDay = Microsoft.VisualBasic.Len(c.Offset(z - 1, -17 + intMonth - x).Text)
                    Next z

            End Select

            'wbIni.Close()            

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Common", gUserId, gReportID)
        End Try

ResumeHere:
        strFactoryCal = s
        fnWorkDaysBetween = intWDCounter
        c = Nothing
    End Function

    Public Sub LoadListObjectData(queryName As String, sheetName As String, listObjectName As String, dataTable As System.Data.DataTable)
        Dim TListObject As Microsoft.Office.Tools.Excel.ListObject
        Dim listObject As Microsoft.Office.Interop.Excel.ListObject
        Dim mappedColumns As String()

        'For Each listObject In Application.ActiveWorkbook.Sheets(sheetName).ListObjects
        'If listObject.Name.Equals(listObjectName) Then
        listObject = Application.ActiveWorkbook.Sheets(sheetName).ListObjects(listObjectName)        
        'dataTable = New Data.DataTable
        Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Try            
            If sheetName <> "Capacity" Then Application.StatusBar = String.Format("Laster data fra BW Query {0} med rader {1}, vennligst vent .....", queryName, dataTable.Rows.Count)
            mappedColumns = New String() {}
            ReDim mappedColumns(listObject.ListColumns.Count - 1)
            For Each col In listObject.ListColumns
                If (gReportID.Equals(7) Or gReportID.Equals(63)) And sheetName <> "Capacity" And sheetName <> "PlanData" And sheetName <> "Allergen" And sheetName <> "AllergenType" And sheetName <> "MixingPlan" Then
                    If col.Index - 2 < dataTable.Columns.Count Then
                        If col.Index > 1 Then
                            dataTable.Columns(col.Index - 2).ColumnName = col.Name
                            mappedColumns(col.Index - 1) = col.Name
                        Else
                            mappedColumns(col.Index - 1) = String.Empty
                        End If
                    Else
                        mappedColumns(col.Index - 1) = String.Empty
                    End If
                Else
                    If col.Index - 1 < dataTable.Columns.Count Then
                        dataTable.Columns(col.Index - 1).ColumnName = col.Name
                        mappedColumns(col.Index - 1) = col.Name
                    Else
                        mappedColumns(col.Index - 1) = String.Empty
                    End If
                End If
            Next

            TListObject = Globals.Factory.GetVstoObject(listObject)
            TListObject.SetDataBinding(dataTable, String.Empty, mappedColumns)
            'TListObject.RefreshDataRows()
            TListObject.Disconnect()
            TListObject.Dispose()

            Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic

            If (gReportID.Equals(7) Or gReportID.Equals(63)) And sheetName <> "Capacity" And sheetName <> "Allergen" And sheetName <> "AllergenType" And sheetName <> "MixingPlan" Then
                listObject.Range.Sort(
                Key1:=listObject.ListColumns(21).Range, Order1:=Excel.XlSortOrder.xlAscending,
                Orientation:=Excel.XlSortOrientation.xlSortColumns,
                Header:=Excel.XlYesNoGuess.xlYes)
            End If

            Application.Calculate()
            If sheetName <> "Capacity" Then Application.StatusBar = String.Format("Data lastet {0} rader", dataTable.Rows.Count)
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Common", gUserId, gReportID)
        End Try
        'End If
        'Next

    End Sub
    Public Function resort(dt As System.Data.DataTable, colName As String, direction As String) As System.Data.DataTable
        Dim dtOut As New System.Data.DataTable
        dt.DefaultView.Sort = colName & " " & direction
        dtOut = dt.DefaultView.ToTable()
        Return dtOut
    End Function
    Private Sub SortTable(sheet As Excel.Worksheet, _
    tableName As String, sortyBy As Excel.Range)

        sheet.ListObjects(tableName).Sort.SortFields.Clear()
        sheet.ListObjects(tableName).Sort.SortFields.Add(sortyBy, _
            Excel.XlSortOn.xlSortOnValues, Excel.XlSortOrder.xlAscending)
        With sheet.ListObjects(tableName).Sort
            .Header = Excel.XlYesNoGuess.xlYes
            .MatchCase = False
            .SortMethod = Excel.XlSortMethod.xlPinYin
            .Apply()
        End With
    End Sub

    Public Sub ShowReportOptions()
        Try
            Using entities = New DAL.SAPExlEntities()

                Dim report = entities.Reports.Where(Function(r) r.ReportID = gReportID).SingleOrDefault()
                Globals.Ribbons.OrklaRT.grpOptions.Visible = report.ReportOptions

                If report.ReportOptions <> False Then
                    Dim reportOptions = entities.ReportOptions.Where(Function(rs) rs.ReportID = gReportID).SingleOrDefault()
                    For Each optionControl In Globals.Ribbons.OrklaRT.grpOptions.Items
                        Select Case optionControl.Tag
                            'Case "MaterialPrice"
                            '    optionControl.Visible = reportOptions.MaterialPrice
                            'Case "ProductHierarchy"
                            '    optionControl.Visible = reportOptions.ProductHierarchy
                            'Case "CustomerHierarchy"
                            '    optionControl.Visible = reportOptions.CustomerHierarchy
                            'Case "SalesValue"
                            '    optionControl.Visible = reportOptions.SalesValue
                            Case "QuantityUnit"
                                optionControl.Visible = reportOptions.QuantityUnit
                                'Case "Currency"
                                '    optionControl.Visible = reportOptions.Currency
                                'Case "CurrencyYear"
                                '    optionControl.Visible = reportOptions.CurrencyYear
                            Case "CreateNewPlan"
                                If OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(2) Or OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(3) Then
                                    optionControl.Visible = reportOptions.CreateNewPlan
                                Else
                                    optionControl.Visible = False
                                End If
                            Case "SavePriorities"
                                If OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(2) Or OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(3) Then
                                    optionControl.Visible = reportOptions.SavePriorities
                                Else
                                    optionControl.Visible = False
                                End If
                                'Case "BudgetVersion"
                                '    optionControl.Visible = reportOptions.BudgetVersion
                            Case "FormatGraph"
                                optionControl.Visible = reportOptions.FormatGraph
                            Case "SaveGroup"
                                optionControl.Visible = reportOptions.SaveGroup
                            Case "SaveBinTest"
                                optionControl.Visible = reportOptions.SaveBinTest
                            Case "SaveExcludedTypes"
                                optionControl.Visible = reportOptions.SaveExcludedTypes
                                'Case "UpdateSAP"
                                '    optionControl.Visible = reportOptions.UpdateSAP
                                'Case "ShowOptions"
                                '    optionControl.Visible = reportOptions.ShowOptions
                                'Case "ShowStock"
                                '    optionControl.Visible = reportOptions.ShowStocks
                            Case "ShowMD04Data"
                                optionControl.Visible = reportOptions.ShowMD04Data
                                If reportOptions.ShowMD04Data = True Then Globals.Ribbons.OrklaRT.LoadShowMD04Data()
                            Case "ShelfLifeType"
                                optionControl.Visible = reportOptions.ShelfLifeTypes
                                If reportOptions.ShelfLifeTypes = True Then Globals.Ribbons.OrklaRT.LoadShelfLifeTypes()
                                'Case "MaterialsIncluded"
                                '    optionControl.Visible = reportOptions.MaterialsIncluded
                            Case "SaveList"
                                optionControl.Visible = reportOptions.SaveList
                            Case "SaveManko"
                                optionControl.Visible = reportOptions.SaveManko
                            Case "KopiArk"
                                optionControl.Visible = True
                            Case "TransformPivotToTable"
                                optionControl.Visible = True
                            Case "TransformPivotToList"
                                optionControl.Visible = True                                
                        End Select
                    Next
                End If
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Common", gUserId, gReportID)
        End Try
    End Sub
    Public Sub EnableReportSheetOptions()
        Try
            If OrklaRTBPL.CommonFacade.GetReportSheetCount(gReportID, Application.ActiveSheet.CodeName) <> 0 Then
                Dim sheetName As String = Application.ActiveSheet.CodeName
                Using entities = New DAL.SAPExlEntities()
                    Dim reportSheetOptions = entities.ReportSheetOptions.Where(Function(rsp) rsp.ReportID = gReportID And rsp.ReportSheet = sheetName).SingleOrDefault()
                    For Each optionControl In Globals.Ribbons.OrklaRT.grpOptions.Items
                        Select Case optionControl.Tag
                            'Case "MaterialPrice"
                            '    optionControl.Enabled = reportSheetOptions.MaterialPrice
                            'Case "ProductHierarchy"
                            '    optionControl.Enabled = reportSheetOptions.ProductHierarchy
                            'Case "CustomerHierarchy"
                            '    optionControl.Enabled = reportSheetOptions.CustomerHierarchy
                            'Case "SalesValue"
                            '    optionControl.Enabled = reportSheetOptions.SalesValue
                            'Case "QuantityUnit"
                            '    optionControl.Enabled = reportSheetOptions.QuantityUnit
                            'Case "Currency"
                            '    optionControl.Enabled = reportSheetOptions.Currency
                            'Case "CurrencyYear"
                            '    optionControl.Enabled = reportSheetOptions.CurrencyYear
                            Case "CreateNewPlan"
                                If OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(3) Then
                                    optionControl.Visible = True
                                    optionControl.Enabled = reportSheetOptions.CreateNewPlan
                                Else
                                    optionControl.Visible = False
                                End If
                            Case "SavePriorities"
                                If OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(3) Then
                                    optionControl.Visible = True
                                    optionControl.Enabled = reportSheetOptions.SavePriorities
                                Else
                                    optionControl.Visible = False
                                End If
                                'Case "BudgetVersion"
                                '    optionControl.Enabled = reportSheetOptions.BudgetVersion
                            Case "FormatGraph"
                                optionControl.Enabled = reportSheetOptions.FormatGraph
                            Case "SaveGroup"
                                optionControl.Enabled = reportSheetOptions.SaveGroup
                            Case "SaveBinTest"
                                optionControl.Enabled = reportSheetOptions.SaveBinTest
                            Case "SaveExcludedTypes"
                                optionControl.Enabled = reportSheetOptions.SaveExcludedTypes
                                'Case "UpdateSAP"
                                '    optionControl.Enabled = reportSheetOptions.UpdateSAP
                            Case "ShowStock"
                                optionControl.Enabled = reportSheetOptions.ShowStocks
                            Case "ShowMD04Data"
                                optionControl.Enabled = reportSheetOptions.ShowMD04Data
                            Case "ShelfLifeType"
                                optionControl.Enabled = reportSheetOptions.ShelfLifeTypes
                                'Case "MaterialsIncluded"
                                '    optionControl.Enabled = reportSheetOptions.MaterialsIncluded
                                'Case "SaveList"
                                '    optionControl.Enabled = reportSheetOptions.SaveList
                            Case "SaveManko"
                                optionControl.Enabled = reportSheetOptions.SaveManko
                            Case "KopiArk"
                                optionControl.Enabled = True
                        End Select
                    Next
                End Using
            Else
                For Each optionControl In Globals.Ribbons.OrklaRT.grpOptions.Items
                    If optionControl.Tag.Equals("KopiArk") Then
                        optionControl.Visible = True
                    Else
                        optionControl.Enabled = False
                    End If
                Next
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "Common", gUserId, gReportID)
        End Try
    End Sub
    Public Sub UpdateQueryTablePath()
        Dim qryActive As Excel.QueryTable

        Using entities = New DAL.SAPExlEntities()
            For Each sh In Application.Sheets
                If sh.QueryTables.Count > 0 Then
                    qryActive = sh.QueryTables(1)
                    qryActive.Connection = "TEXT;" + Path.GetTempPath + entities.vwCurrentUser.SingleOrDefault().SAPSystem + "\" + qryActive.Connection.ToString().Split("\").GetValue((qryActive.Connection.ToString().Split("\").Length) - 1)
                    Call RefreshQueryTable(sh.Name)
                End If
            Next sh
        End Using

    End Sub

    Public Function GetDataSourceFromFile(fileName As String) As System.Data.DataTable
        Dim dt As New System.Data.DataTable("data")
        Dim columns As String() = Nothing

        Dim lines = File.ReadAllLines(fileName)
        For i As Integer = 1 To lines.Count() - 1
            Dim dr As System.Data.DataRow = dt.NewRow()
            Dim values As String() = lines(i).Split(New Char() {vbTab})

            Dim j As Integer = 0
            While j < values.Count()
                If (i = 1) Then
                    dt.Columns.Add("Column" + j.ToString())
                End If
                dr(j) = values(j)
                j += 1
            End While

            dt.Rows.Add(dr)
        Next
        Return dt
    End Function
    Public Sub RefreshQueryTable(sheetName As String)

        Using entities = New DAL.SAPExlEntities()
            Application.Sheets(sheetName).QueryTables(1).TextFilePlatform = Excel.XlPlatform.xlWindows
            Application.Sheets(sheetName).QueryTables(1).TextFileDecimalSeparator = entities.vwCurrentUser.SingleOrDefault().DecimalSeparator
            Application.Sheets(sheetName).QueryTables(1).TextFileThousandsSeparator = entities.vwCurrentUser.SingleOrDefault().ThousandSeparator
            Application.Sheets(sheetName).QueryTables(1).TextFileTrailingMinusNumbers = True
            Application.Sheets(sheetName).QueryTables(1).refresh()
        End Using
    End Sub
    Public Function BDCInputTable() As System.Data.DataTable

        Dim table As New System.Data.DataTable

        table.Columns.Add("ValueFieldName", GetType(String))
        table.Columns.Add("FieldValue", GetType(String))

        Return table

    End Function

End Module
