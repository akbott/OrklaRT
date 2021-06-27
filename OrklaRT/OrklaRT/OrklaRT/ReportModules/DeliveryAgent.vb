Module DeliveryAgent
    Sub LocalUpdate()    

        Dim pi As Excel.PivotItem

        'Try
        Application.Sheets("QMLot").PivotTables(1).PivotCache.Refresh()

        On Error Resume Next
        For Each pi In Application.Sheets("QMLot").PivotTables(1).PivotFields("Test_QM").PivotItems
            If pi.Name = "1" Then
                Application.Sheets("QMLot").PivotTables(1).PivotFields("Test_QM").CurrentPage = "1"
            End If
        Next
        On Error GoTo 0
        Application.Sheets("Sikkerhetsdager ved levering").PivotTables(1).PivotCache.Refresh()
        Call FormatPivotTable()

        'Catch ex As Exception
        '    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "DeliveryAgent", gUserId, gReportID)
        'End Try

    End Sub

    Sub FormatPivotTable()
        Dim objActive As Excel.Range        

        On Error GoTo Cleanup
        Application.Sheets("Sikkerhetsdager ved levering").Activate()
        objActive = Application.ActiveCell

        On Error GoTo CleanUp
        Application.Sheets("Sikkerhetsdager ved levering").PivotTables(1).PivotFields("Tilgj.dato ").DataRange.Select()
        With Application.Selection
            .HorizontalAlignment = Excel.Constants.xlGeneral
            .VerticalAlignment = Excel.Constants.xlBottom
            .WrapText = False
            .Orientation = 45
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = Excel.Constants.xlContext
            .MergeCells = False
        End With

        Application.Sheets("Sikkerhetsdager ved levering").PivotTables(1).DataBodyRange.Select()
        Application.Selection.ColumnWidth = 3.2
        Application.Selection.Cells(1, 1).ColumnWidth = 7

        Application.Sheets("Sikkerhetsdager ved levering").PivotTables(1).DataBodyRange.Select()

        Application.Selection.FormatConditions.Delete()
        Application.Selection.FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlCellValue, Operator:=Excel.XlFormatConditionOperator.xlLess, Formula1:="0")
        Application.Selection.FormatConditions(1).Font.ColorIndex = 2
        Application.Selection.FormatConditions(1).Interior.ColorIndex = 3
        Application.Selection.FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlCellValue, Operator:=Excel.XlFormatConditionOperator.xlBetween, Formula1:="0,01", Formula2:=(Application.Range("GreenLimit").Value - 0.00001).ToString())
        Application.Selection.FormatConditions(2).Interior.ColorIndex = 44
        Application.Selection.FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlCellValue, Operator:=Excel.XlFormatConditionOperator.xlGreaterEqual, Formula1:=Application.Range("GreenLimit").Value.ToString())
        Application.Selection.FormatConditions(3).Interior.ColorIndex = 43

        Application.Sheets("Lagerdekning dager").Activate()
        Application.Sheets("Lagerdekning dager").PivotTables(1).DataBodyRange.Select()
        Application.Selection.Columns.AutoFit()
        Application.Selection.Cells(1, 1).Select()
        Application.Selection.Cells(1, 1).ColumnWidth = 7


CleanUp:
        Application.Sheets("Sikkerhetsdager ved levering").Activate()
        Call MarkQM()

        objActive.Select()

    End Sub

    Sub MarkQM()
        Dim c As Excel.Range
        Dim d As Excel.Range

        Try
            Application.Sheets("Sikkerhetsdager ved levering").PivotTables(1).RowRange.Interior.ColorIndex = Excel.Constants.xlNone
            'Application.Sheets("SafetyDays at Delivery").Activate()

            If Application.Sheets("QMLot").PivotTables(1).PivotFields("Test_QM").CurrentPage.value = "0" Then GoTo CleanUp
            For Each d In Application.Sheets("QMLot").PivotTables(1).RowRange.Cells
                If d.Value <> "Material Navn" Then
                    c = Nothing
                    c = Application.Sheets("Sikkerhetsdager ved levering").PivotTables(1).RowRange.Find(d.Value, LookIn:=Excel.XlFindLookIn.xlValues, lookat:=Excel.XlLookAt.xlPart)
                    If Not c Is Nothing Then
                        c.Interior.ColorIndex = 43
                    End If
                End If
            Next            

CleanUp:
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "DeliveryAgent", gUserId, gReportID)
        End Try
    End Sub
  
End Module
