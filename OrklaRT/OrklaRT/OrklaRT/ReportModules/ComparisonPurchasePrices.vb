Module ComparisonPurchasePrices

    Sub FormatPivotTable()

        If Application.Sheets("Price Control").PivotTables.Count = 0 Then Exit Sub

        Application.Sheets("Price Control").PivotTables(1).DataBodyRange.Select()

        Application.Selection.FormatConditions.Delete()
        Application.Selection.FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlCellValue, Operator:=Excel.XlFormatConditionOperator.xlLess, Formula1:=-Application.Range("DevLimit").Value.ToString())
        Application.Selection.FormatConditions(1).Interior.ColorIndex = 43
        Application.Selection.FormatConditions.Add(Type:=Excel.XlFormatConditionType.xlCellValue, Operator:=Excel.XlFormatConditionOperator.xlGreater, Formula1:=Application.Range("DevLimit").Value.ToString())
        Application.Selection.FormatConditions(2).Interior.ColorIndex = 22

CleanUp:

    End Sub



    'Sub reset()
    '    Application.ScreenUpdating = True
    '    Application.EnableEvents = True
    'End Sub


    '    Sub ToggleShowPercent()
    '        Dim obj As Excel.Range

    '        On Error GoTo CleanUp

    '        Application.ScreenUpdating = False
    '        Application.EnableEvents = False
    '        obj = Application.Selection
    '        Application.ActiveSheet.Shapes("Percent").Select()

    '        If Application.Selection.Characters.Text = "Show as numbers" Then
    '            obj.Select()
    '            With Application.ActiveSheet.PivotTables(1).PivotFields("Amount_BudRate ")
    '                .Calculation = .xlNormal
    '                .NumberFormat = "#,##0;[Red]-#,##0"
    '            End With
    '            Application.ActiveSheet.Shapes("Percent").Select()
    '            Application.Selection.Characters.Text = "Show as percent"
    '        Else
    '            obj.Select()
    '            With Application.ActiveSheet.PivotTables(1).PivotFields("Amount_BudRate ")
    '                .Calculation = Excel.XlPivotFieldCalculation.xlPercentOf
    '                .BaseField = "Cond. Type"
    '                .NumberFormat = "0.0 %;[Red]-0.0 %"
    '            End With
    '            Application.ActiveSheet.Shapes("Percent").Select()
    '            Application.Selection.Characters.Text = "Show as numbers"
    '        End If

    'CleanUp:
    '        On Error Resume Next

    '        obj.Select()
    '        Application.EnableEvents = True
    '        Application.ScreenUpdating = True

    '    End Sub

End Module
