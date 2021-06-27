Module OptimizedLotSize

    Sub LocalUpdate()

        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual
        Call FindDefault()
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.Sheets("Lot Sizes").Unprotect("next")
        Application.Sheets("Pivot").PivotTables(1).PivotCache.Refresh()
        Application.Sheets("Lot Sizes").Range("DataRange").Clear()
        Application.Sheets("Pivot").PivotTables(1).RowRange.Copy(Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0))
        Application.Sheets("Pivot").PivotTables(1).DataBodyRange.Copy(Application.Sheets("Lot Sizes").Range("PasteStart").Offset(0, 1))
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Interior.ColorIndex = 15
        Application.Sheets("Lot Sizes").Range("NewCO").Value = Application.Sheets("Lot Sizes").Range("CurrCO").Value
        Application.Sheets("Lot Sizes").Range("DataRange").Locked = False
        Application.Sheets("Lot Sizes").Range("DataRange").NumberFormat = "#,##0;[Red]-#,##0"
        Application.Sheets("Lot Sizes").Protect("next")
        Application.Sheets("All_Materials").PivotTables(1).PivotCache.Refresh()
        Application.EnableEvents = True

    End Sub

    Sub LocalPrepareSaving()
        Application.Sheets("Lot Sizes").Unprotect("next")
        Application.Sheets("Lot Sizes").Range("DataRange").Clear()
        Application.Sheets("Lot Sizes").Range("NewCO").ClearContents()
        Application.Sheets("Pivot").PivotTables(1).PivotCache.Refresh()
        Application.Sheets("All_Materials").PivotTables(1).PivotCache.Refresh()
        Application.Sheets("Lot Sizes").Protect("next")
    End Sub


    Sub MaterialsIncluded()
        Application.Sheets("Lot Sizes").Unprotect("next")
        Application.Sheets("Lot Sizes").Range("DataRange").Clear()
        Application.Sheets("Pivot").PivotTables(1).RowRange.Copy(Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0))
        Application.Sheets("Pivot").PivotTables(1).DataBodyRange.Copy(Application.Sheets("Lot Sizes").Range("PasteStart").Offset(0, 1))
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlDiagonalDown).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlDiagonalUp).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlEdgeLeft).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlEdgeTop).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Borders(Excel.XlBordersIndex.xlEdgeRight).LineStyle = Excel.Constants.xlNone
        Application.Sheets("Lot Sizes").Range("PasteStart").Offset(-1, 0).Interior.ColorIndex = 15
        Application.Sheets("Lot Sizes").Range("NewCO").Value = Application.Sheets("Lot Sizes").Range("CurrCO").Value
        Application.Sheets("Lot Sizes").Range("DataRange").Locked = False
        Application.Sheets("Lot Sizes").Range("DataRange").NumberFormat = "#,##0;[Red]-#,##0"
        Application.Sheets("Lot Sizes").Protect("next")
    End Sub


    Sub FindDefault()
        Dim c As Excel.Range
        Dim x As Long


        On Error Resume Next
        Application.Sheets("Y111_Data").Activate()
        Application.Sheets("Y111_Data").Range(Application.Cells(2, 57), Application.Cells(30000, 57)).ClearContents()
        x = 0
        For Each r In Application.Range("DefaultGroups").Rows
            x = x + 1
            If x > 1 Then
                With Application.Range("SapExlData")
                    c = .Columns(31).Find(String.Format("{0:000000}", Convert.ToInt32(r.Cells(1, 15).Value)) & " " & "Default" & r.Cells(1, 17).Value, LookIn:=Excel.XlFindLookIn.xlValues, Lookat:=Excel.XlLookAt.xlWhole)
                    If Not c Is Nothing Then
                        If c.Offset(0, 1).Value = r.Cells(1, 51).Value Then
                            r.Cells(1, 57).Value = "Default"
                        Else
                            r.Cells(1, 57).Value = "Optional"
                        End If
                    End If
                End With
            End If
        Next

CleanUp:
        c = Nothing

    End Sub


    Sub UploadSAP()
      
        '        Dim strFileName As String
        '        Dim bolTest As Boolean
        '        Dim Wb As Excel.Workbook
        '        Dim c As Excel.Range
        '        Dim Answ

        '        Application.ScreenUpdating = False
        '        Application.EnableEvents = False

        '        Select Case sapConn.Connection.user
        '            Case "NARDAL", "KLARSEN1", "BSELLEVO", "BTOMMERB"

        '                '      Case "JBATHOVA", "VCAPRATO", "TVATER"

        '            Case Else
        '                MsgBox("This function is only available to responsible Production Planner.", , "Orkla SAP Integration")
        '                GoTo CleanUp
        '        End Select

        '        Answ = MsgBox("Are you sure you want to update SAP MM02?", vbYesNo, "Orkla SAP Integration")
        '        If Answ <> 6 Then Exit Sub

        '        Wb = ThisWorkbook

        '        Application.StatusBar = "Updating SAP Lot Size and Costing Lot Size..."

        '        For Each c In Wb.Sheets("Lot Sizes").Range("DataRange").Columns(1).Cells
        '            If c.Value = "" Then Exit For
        '            If Not IsError(Val(Left(c.Value, 6))) Then
        '                'Debug.Print Val(Left(c.Value, 6))
        '                If Val(Left(c.Value, 6)) > 0 Then
        '                    Wb.Sheets("BDC_Upload").Range("Upload1").Value = Val(Left(c.Value, 6))
        '                End If
        '                If CLng(c.Offset(0, 8).Value) > 0 Then
        '                    If CLng(c.Offset(0, 8).Value) > 99 Then
        '                        Wb.Sheets("BDC_Upload").Range("Upload2").Value = "'99"
        '                    Else
        '                        Wb.Sheets("BDC_Upload").Range("Upload2").Value = "'" & Format(CLng(c.Offset(0, 8).Value), "00")
        '                    End If
        '                End If
        '                If CLng(c.Offset(0, 11).Value) > 0 Then
        '                    Wb.Sheets("BDC_Upload").Range("Upload3").Value = CLng(c.Offset(0, 11).Value)
        '                End If
        '                '         Exit For
        '                Call SAP_Call_Trans_Data(Wb.Name, "BDC_Upload", "MM02")
        '            End If
        '        Next c

        '        Call RunUpdatePivot()

        '        Wb.Activate()

        'CleanUp:
        '        Application.StatusBar = False
        '        Application.EnableEvents = True
        '        Application.ScreenUpdating = True
        '        Wb = Nothing
        '        Exit Sub

    End Sub


    Sub LocalEnd()
        Application.Sheets("Lot Sizes").Protect("next")
    End Sub



End Module
