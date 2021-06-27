Module MD04Overview
    Public md04Table = MD04DataTable()
    Public bdcTable = BDCUploadDataTable()
    Sub LocalUpdate()
        'Dim pi As Excel.PivotItem

        Try


            'If Application.ActiveWorkbook.CountA(Application.Sheet1.Range("Multiselect").Cells) = 1 Then
            '    MsgBox("The Product Group doesn't contain any materials.", , "Orkla SAP Integration")
            '    Exit Sub
            'End If

            Call GetAllMD04()

            'Application.Sheets("MD04").PivotTables(1).PivotCache.Refresh()
            Application.Sheets("Diagram 509").SelectData()
            'Application.Sheets("SAP_Update").Unprotect("next")
            'Call RefreshQueryTable("SAP_Update", ThisWorkbook)
            'Application.Sheets("SAP_Update").Range("UpdateValues").ClearContents()
            'Application.Sheets("SAP_Update").Range("SapRetValues").ClearContents()
            'Application.Sheets("SAP_Update").Protect("next")

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "MD04Overview", gUserId, gReportID)
        End Try
    End Sub

    Sub GetAllMD04()
        Dim c As Excel.Range
        Dim x As Integer
        Dim y As Integer
        Dim TListObject As Microsoft.Office.Tools.Excel.ListObject
        Dim mappedColumns As String()

        Application.Calculation = Excel.XlCalculation.xlCalculationManual


        Try

            'If Application.WorksheetFunction.CountA(Application.Range("MultiSelect")) = 0 Then
            x = Application.Sheets("Y149_Data").Range("PlantData").Columns(1).Cells.Count
            y = 0
            For Each c In Application.Sheets("Y149_Data").Range("PlantData").Columns(1).Cells
                If c.Value.ToString() = String.Empty Then Exit For
                y = y + 1
                Application.StatusBar = Format(y / x, "0 %") & " of materials read ..."
                If c.Row > 1 Then Call BAPI_MATERIAL_MRP_LIST_All(c.Value, OrklaRTBPL.SelectionFacade.MD04SelectionPlant)
            Next c
            'Else
            'x = Application.WorksheetFunction.CountA(Application.Range("MultiSelect"))
            'y = 0
            'For Each c In Application.Sheets("MultiSelect").Range("MultiSelect").Columns(1).Cells
            '    If c.Value = "" Then Exit For
            '    y = y + 1
            '    Application.StatusBar = Format(y / x, "0 %") & " of materials read ..."
            '    If c.Row > 1 Then Call Common.BAPI_MATERIAL_MRP_LIST_All(c.Value, OrklaRTBPL.SelectionFacade.MD04SelectionPlant)
            'Next c
            'End If

            If md04Table.Rows.Count > 0 Then
                For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Application.ActiveWorkbook.Sheets("DataBase").ListObjects
                    If listObject.Name.Equals("OrklaRTData") Then
                        Try
                            mappedColumns = New String() {}
                            ReDim mappedColumns(listObject.ListColumns.Count - 1)
                            For Each col In listObject.ListColumns
                                If col.Index - 1 < md04Table.Columns.Count Then
                                    md04Table.Columns(col.Index - 1).ColumnName = col.Name
                                    mappedColumns(col.Index - 1) = col.Name
                                Else
                                    mappedColumns(col.Index - 1) = String.Empty
                                End If
                            Next
                            TListObject = Globals.Factory.GetVstoObject(listObject)
                            TListObject.SetDataBinding(md04Table, String.Empty, mappedColumns)
                            TListObject.RefreshDataRows()
                            TListObject.Disconnect()
                            Application.StatusBar = String.Format("Data Successfully Loaded {0}", md04Table.Rows.Count)
                        Catch
                        End Try
                    End If
                Next
                'For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Application.ActiveWorkbook.Sheets("DataBase").ListObjects
                '    If listObject.Name.Equals("OrklaRTData") Then
                '        Try
                '            If Not listObject.DataBodyRange Is Nothing Then
                '                listObject.DataBodyRange.Delete()
                '            End If
                '            Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(resultTable)
                '            data.MoveFirst()
                '            Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, resultTable.Rows.Count, resultTable.Columns.Count)
                '            Application.StatusBar = String.Format("Data Successfully Loaded {0}", resultTable.Rows.Count)

                '        Catch
                '        End Try
                '    End If
                'Next
            End If

        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "MD04Overview", gUserId, gReportID)
        End Try
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic


    End Sub


    Sub UploadSAP()

        Dim strFileName As String
        Dim bolTest As Boolean
        'Dim Wb As Excel.Workbook
        Dim c As Excel.Range
        Dim Answ As Integer
        Dim bolRetry As Boolean
        Dim bolErr As Boolean

        On Error GoTo CleanUp

        Application.EnableEvents = False

        'Select Case sapConn.Connection.user
        '    Case "BSELLEVO", "BTOMMERB", "HSCHWAB", "SSUNDGOT"

        '        '      Case "JBATHOVA", "VCAPRATO", "TVATER"

        '    Case Else
        '        MsgBox("This function is only available to responsible Production Planner.", , "Orkla SAP Integration")
        '        GoTo CleanUp
        'End Select

        Answ = MsgBox("Are you sure you want to update SAP MM02?", vbYesNo, "Orkla SAP Integration")
        If Answ <> 6 Then GoTo CleanUp

        'Wb = Application.ActiveWorkbook

        Application.StatusBar = "Updating SAP Lot Size and Safety Time..."

        'Clear SAP Return values in column V
        Application.Sheets("SAP_Update").Range("SapRetValues").ClearContents()
        Application.ScreenUpdating = False

        If Application.Sheets("SAP_Update").Range("PlantData").Cells(1, 1).Value <> "Material" Then
            MsgBox("The first column doesn't contain Material numbers. SAP upload is cancelled.", , "Orkla SAP Integration")
            GoTo CleanUp
        End If

        If OrklaRTBPL.SelectionFacade.MD04SelectionPlant = "" Then
            MsgBox("The Plant field has to be filled out in the Selections sheet. SAP upload is cancelled.", , "Orkla SAP Integration")
            GoTo CleanUp
        End If

        For Each c In Application.Sheets("SAP_Update").Range("PlantData").Columns(1).Cells
            If c.Value = "Material" Then GoTo NextRow
            If c.Value = "" Then GoTo CleanUp
            If c.Offset(0, 19).Value = 0 Then GoTo NextRow

Retry:
            bolErr = True
            If Not IsError(Val(c.Value)) Then
                On Error GoTo NextRow
                'Material Number
                bdcTable.Rows(0)("Value1").Add(c.Value)
                'Lot Size
                If CLng(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 1).Column).Value) > 0 Then
                    If CLng(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 1).Column).Value) > 99 Then
                        bdcTable.Rows(0)("Value2").Add("99")
                    Else
                        bdcTable.Rows(0)("Value2").Add(String.Format("0:00", CLng(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 1).Column).Value)))
                    End If
                End If
                'Coverage Profile
                If CLng(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 2).Column).Value) > 0 Then
                    bdcTable.Rows(0)("Value3").Add(String.Format("0:000", CLng(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 2).Column).Value)))
                Else
                    bdcTable.Rows(0)("Value3").Add(String.Empty)
                End If
                'Safety Time Days
                If CLng(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 3).Column).Value) > 0 Then
                    bdcTable.Rows(0)("Value6").Add(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 3).Column).Value)
                Else
                    bdcTable.Rows(0)("Value6").Add("0")
                End If
                'Safety Time Indicator
                If CLng(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 4).Column).Value) > 0 Then
                    bdcTable.Rows(0)("Value5").Add(Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 4).Column).Value)
                Else
                    bdcTable.Rows(0)("Value5").Add(String.Empty)
                End If

                '         Exit For
                'Call SAP_Call_Trans_Data(Wb.Name, "BDC_Upload", "MM02")

                If Left(Application.Sheets("BDC_Upload").Range("RetValue").Value, 8) = "S:00-349" Then
                    bdcTable.Rows(0)("Value7").Add(String.Empty)
                    bdcTable.Rows(0)("Value8").Add("X")
                    If bolRetry = False Then
                        bolRetry = True
                        GoTo Retry
                    End If
                End If
                bdcTable.Rows(0)("Value7").Add("X")
                bdcTable.Rows(0)("Value8").Add(String.Empty)

                Application.ScreenUpdating = True
                'Application.Cells(c.Row, Application.Range("UpdateValues").Cells(1, 4).Column).Offset(0, 1).Value = Application.Sheets("BDC_Upload").Range("RetValue").Value
                Application.ScreenUpdating = False

            End If
            bolErr = False

NextRow:
            If bolErr = True Then
                c.Offset(0, 21).Value = "An error occured, update attempt cancelled."
            End If
        Next c

        '   Call RunUpdatePivot

        Application.ActiveSheet.Activate()

CleanUp:
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        'Wb = Nothing
    End Sub
    Function MD04DataTable() As System.Data.DataTable

        Dim dataTable As New System.Data.DataTable
        dataTable.Columns.Add("PLNGSEGMT", GetType(Int32))
        dataTable.Columns.Add("MRP_ELEMENT_IND", GetType(String))
        dataTable.Columns.Add("AVAIL_DATE", GetType(DateTime))
        dataTable.Columns.Add("PLD_IND_REQS", GetType(Double))
        dataTable.Columns.Add("REQMTS", GetType(Double))
        dataTable.Columns.Add("RECEIPTS", GetType(Double))
        dataTable.Columns.Add("AVAIL_QTY", GetType(Double))
        dataTable.Columns.Add("ACTL_COVERAGE", GetType(Double))
        dataTable.Columns.Add("Material", GetType(String))
        dataTable.Columns.Add("DateValue", GetType(DateTime))

        Return dataTable
    End Function
    Function BDCUploadDataTable() As System.Data.DataTable

        Dim dataTable As New System.Data.DataTable
        dataTable.Columns.Add("Value1", GetType(String))
        dataTable.Columns.Add("Value2", GetType(String))
        dataTable.Columns.Add("Value3", GetType(String))
        dataTable.Columns.Add("Value4", GetType(String))
        dataTable.Columns.Add("Value5", GetType(String))
        dataTable.Columns.Add("Value6", GetType(String))
        dataTable.Columns.Add("Value7", GetType(String))
        dataTable.Columns.Add("Value8", GetType(String))

        Return dataTable
    End Function


    Public Sub BAPI_MATERIAL_MRP_LIST_All(strMaterial As String, plant As String)
        Dim rfcTable As SAP.Middleware.Connector.IRfcTable


        Try

            rfcTable = BPL.RfcFunctions.GetBAPIMATERIALMRPLISTAll(strMaterial, plant)

            If rfcTable.RowCount > 0 Then
                For j As Integer = 0 To rfcTable.RowCount - 1
                    rfcTable.CurrentIndex = j
                    If CDate(rfcTable.GetValue("AVAIL_DATE")).Date < DateTime.Now.AddDays(186) Then
                        md04Table.Rows.Add(rfcTable.GetValue("PLNGSEGMT"), rfcTable.GetValue("MRP_ELEMENT_IND"), rfcTable.GetValue("AVAIL_DATE"), rfcTable.GetValue("PLD_IND_REQS"),
                                            rfcTable.GetValue("REQMTS"), rfcTable.GetValue("RECEIPTS"),
                                            rfcTable.GetValue("AVAIL_QTY"), rfcTable.GetValue("ACTL_COVERAGE"),
                                            strMaterial, DateTime.Now.Date)
                    End If
                Next j
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "MD04Overview", gUserId, gReportID)
        End Try
    End Sub

End Module
