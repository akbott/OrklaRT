Imports System.Linq.Expressions
Imports System.Xml
Imports System.IO


Module PivotFunctions
    Public Function ReturnPivotLayout() As String

        Dim pivotItemCount As Integer, seqNum As Integer, subTotals As Integer
        Dim firstPvt As Boolean, visible As Boolean
        Dim hiddenItems As String
        Dim pvt As Excel.PivotTable
        Dim stringWriter As New StringWriter()

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        Try

            firstPvt = True
            seqNum = 0
            Dim writer As New XmlTextWriter(stringWriter)
            'writer.WriteStartDocument(True)
            writer.Formatting = Formatting.Indented
            writer.Indentation = 2

            writer.WriteStartElement("XtraSerializer")
            WriteXmlAttribute("UserID", OrklaRTBPL.CommonFacade.GetUserID(), writer)
            WriteXmlAttribute("ReportID", gReportID, writer)
            WriteXmlAttribute("VariantID", "1", writer)
            For Each sheet In Application.ActiveWorkbook.Sheets
                If sheet.Tab.ColorIndex = 53 Then
                    If sheet.PivotTables.Count > 0 Then
                        pvt = sheet.PivotTables(1)
                        writer.WriteStartElement("property")
                        WriteXmlAttribute("SheetName", sheet.Name, writer)
                        WriteXmlAttribute("TableName", pvt.Name, writer)

                        writer.WriteStartElement("property")
                        WriteXmlAttribute("Seq", "1", writer)
                        WriteXmlAttribute("PivotElement", "Table", writer)
                        WriteXmlAttribute("SourceName", "Table", writer)
                        WriteXmlAttribute(GetPropertyName(Function() pvt.ColumnGrand), Convert.ToInt32(pvt.ColumnGrand).ToString(), writer)
                        WriteXmlAttribute(GetPropertyName(Function() pvt.RowGrand), Convert.ToInt32(pvt.RowGrand).ToString(), writer)
                        WriteXmlAttribute(GetPropertyName(Function() pvt.DisplayErrorString), Convert.ToInt32(pvt.DisplayErrorString).ToString(), writer)
                        WriteXmlAttribute("TableStart", pvt.TableRange1.Cells(1, 1).Address, writer)
                        WriteXmlAttribute("FieldList", Convert.ToInt32(Application.ActiveWorkbook.ShowPivotTableFieldList), writer)
                        WriteXmlAttribute("MissingItems", String.Empty, writer)
                        WriteXmlAttribute(GetPropertyName(Function() pvt.SaveData), pvt.SaveData, writer)
                        WriteXmlAttribute(GetPropertyName(Function() pvt.ShowDrillIndicators), pvt.ShowDrillIndicators, writer)
                        writer.WriteEndElement()

                        seqNum = (seqNum + 1)

                        If (firstPvt = True) Then
                            If (pvt.CalculatedFields().Count > 0) Then
                                For Each pf As Excel.PivotField In pvt.CalculatedFields()
                                    seqNum = (seqNum + 1)
                                    writer.WriteStartElement("property")
                                    WriteXmlAttribute("Seq", seqNum, writer)
                                    WriteXmlAttribute("PivotElement", "Formulas", writer)
                                    WriteXmlAttribute("SourceName", pf.SourceName, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Name), pf.Name, writer)
                                    WriteXmlAttribute("NumberFormat", "@", writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Orientation), Convert.ToInt32(pf.Orientation).ToString(), writer)
                                    If (pf.Formula.Substring(1, 1) = "(") Then
                                        WriteXmlAttribute(GetPropertyName(Function() pf.Formula), pf.Formula.Substring(1, pf.Formula.Length - 1), writer)
                                    Else
                                        WriteXmlAttribute(GetPropertyName(Function() pf.Formula), "(" & pf.Formula.Substring(2, pf.Formula.Length - 2) & ")", writer)
                                    End If
                                    writer.WriteEndElement()
                                Next
                            End If
                            firstPvt = False
                        End If

                        If (pvt.PageFields.Count > 0) Then
                            For Each pf As Excel.PivotField In pvt.PageFields
                                visible = False
                                seqNum = (seqNum + 1)
                                writer.WriteStartElement("property")
                                WriteXmlAttribute("Seq", seqNum, writer)
                                WriteXmlAttribute("PivotElement", "Page", writer)
                                WriteXmlAttribute("SourceName", pf.SourceName, writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption, writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.Name), pf.Name, writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.CurrentPage), pf.CurrentPage.Name, writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.Orientation), Convert.ToInt32(pf.Orientation).ToString(), writer)
                                WriteXmlAttribute("CpSourceName", pf.CurrentPage.SourceName, writer)

                                If (pf.AllItemsVisible = False) Then
                                    hiddenItems = ""
                                    pivotItemCount = 0
                                    For Each i As Excel.PivotItem In pf.PivotItems
                                        If (i.Visible = False) Then
                                            pivotItemCount = (pivotItemCount + 1)
                                        End If
                                    Next
                                    If (pivotItemCount > (pf.PivotItems().Count / 2)) Then
                                        visible = True
                                    Else
                                        visible = False
                                    End If
                                    For Each i As Excel.PivotItem In pf.PivotItems
                                        If (i.Visible = visible) Then
                                            If (i.SourceNameStandard = "(blank)") Then
                                                hiddenItems = (hiddenItems & (";" & "(blank)"))
                                            Else
                                                hiddenItems = (hiddenItems & (";" + i.SourceName))
                                            End If
                                        End If
                                    Next
                                    WriteXmlAttribute("PivotItems", hiddenItems & ";", writer)
                                    If visible = False Then
                                        WriteXmlAttribute("Visible", 0, writer)
                                    Else
                                        WriteXmlAttribute("Visible", 1, writer)
                                    End If
                                Else
                                    WriteXmlAttribute("PivotItems", String.Empty, writer)
                                    If visible = False Then
                                        WriteXmlAttribute("Visible", 0, writer)
                                    Else
                                        WriteXmlAttribute("Visible", 1, writer)
                                    End If
                                End If
                                writer.WriteEndElement()
                            Next
                        End If
                        If (pvt.RowFields.Count > 0) Then
                            For Each pf As Excel.PivotField In pvt.RowFields
                                If (pf.Name <> pvt.DataPivotField.Name) Then
                                    visible = False
                                    seqNum = (seqNum + 1)
                                    writer.WriteStartElement("property")
                                    WriteXmlAttribute("Seq", seqNum, writer)
                                    WriteXmlAttribute("PivotElement", "Row", writer)
                                    WriteXmlAttribute("SourceName", pf.SourceName, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Name), pf.Name, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.AutoSortOrder), pf.AutoSortOrder, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.AutoSortField), pf.AutoSortField, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Orientation), Convert.ToInt32(pf.Orientation).ToString(), writer)

                                    For subTotals = 1 To 12
                                        If pf.Subtotals(subTotals).Equals(True) Then
                                            WriteXmlAttribute("SubTotals", subTotals, writer)
                                            Exit For
                                        End If
                                    Next
                                    If (pf.AllItemsVisible = False) Then
                                        hiddenItems = ""
                                        pivotItemCount = 0
                                        For Each i As Excel.PivotItem In pf.PivotItems()
                                            If (i.Visible = False) Then
                                                pivotItemCount = (pivotItemCount + 1)
                                            End If
                                        Next
                                        If (pivotItemCount > (pf.PivotItems().Count / 2)) Then
                                            visible = True
                                        Else
                                            visible = False
                                        End If
                                        For Each i As Excel.PivotItem In pf.PivotItems()
                                            If (i.Visible = visible) Then
                                                If (i.SourceNameStandard = "(blank)") Then
                                                    hiddenItems = (hiddenItems & (";" & "(blank)"))
                                                Else
                                                    hiddenItems = (hiddenItems & (";" + i.SourceName))
                                                End If
                                            End If
                                        Next
                                        WriteXmlAttribute("PivotItems", hiddenItems & ";", writer)
                                        If visible = False Then
                                            WriteXmlAttribute("Visible", 0, writer)
                                        Else
                                            WriteXmlAttribute("Visible", 1, writer)
                                        End If
                                    Else
                                        WriteXmlAttribute("PivotItems", String.Empty, writer)
                                        If visible = False Then
                                            WriteXmlAttribute("Visible", 0, writer)
                                        Else
                                            WriteXmlAttribute("Visible", 1, writer)
                                        End If
                                    End If
                                End If
                                writer.WriteEndElement()
                            Next
                        End If
                        If (pvt.ColumnFields.Count > 0) Then
                            For Each pf As Excel.PivotField In pvt.ColumnFields
                                If (pf.Name <> pvt.DataPivotField.Name) Then
                                    visible = False
                                    seqNum = (seqNum + 1)
                                    writer.WriteStartElement("property")
                                    WriteXmlAttribute("Seq", seqNum, writer)
                                    WriteXmlAttribute("PivotElement", "Column", writer)
                                    WriteXmlAttribute("SourceName", pf.SourceName, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Name), pf.Name, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.AutoSortOrder), pf.AutoSortOrder, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.AutoSortField), pf.AutoSortField, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Orientation), Convert.ToInt32(pf.Orientation).ToString(), writer)

                                    subTotals = 1
                                    While (subTotals <= 12)
                                        If (pf.Subtotals(subTotals) = True) Then
                                            WriteXmlAttribute("SubTotals", subTotals, writer)
                                        End If
                                        subTotals += 1
                                    End While
                                    If (pf.AllItemsVisible = False) Then
                                        hiddenItems = ""
                                        pivotItemCount = 0
                                        For Each i As Excel.PivotItem In pf.PivotItems()
                                            If (i.Visible = False) Then
                                                pivotItemCount = (pivotItemCount + 1)
                                            End If
                                        Next
                                        If (pivotItemCount > (pf.PivotItems().Count / 2)) Then
                                            visible = True
                                        Else
                                            visible = False
                                        End If
                                        For Each i As Excel.PivotItem In pf.PivotItems()
                                            If (i.Visible = visible) Then
                                                If (i.SourceNameStandard = "(blank)") Then
                                                    hiddenItems = (hiddenItems & (";" & "(blank)"))
                                                Else
                                                    hiddenItems = (hiddenItems & (";" + i.SourceName))
                                                End If
                                            End If
                                        Next
                                        WriteXmlAttribute("PivotItems", hiddenItems & ";", writer)
                                        If visible = False Then
                                            WriteXmlAttribute("Visible", 0, writer)
                                        Else
                                            WriteXmlAttribute("Visible", 1, writer)
                                        End If
                                    Else
                                        WriteXmlAttribute("PivotItems", String.Empty, writer)
                                        If visible = False Then
                                            WriteXmlAttribute("Visible", 0, writer)
                                        Else
                                            WriteXmlAttribute("Visible", 1, writer)
                                        End If
                                    End If
                                    writer.WriteEndElement()
                                End If
                            Next
                        End If
                        If (pvt.DataFields.Count > 0) Then
                            For Each pf As Excel.PivotField In pvt.DataFields
                                seqNum = (seqNum + 1)
                                writer.WriteStartElement("property")
                                WriteXmlAttribute("Seq", seqNum, writer)
                                WriteXmlAttribute("PivotElement", "Data", writer)
                                WriteXmlAttribute("SourceName", pf.SourceName, writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.Name), pf.SourceName + " ", writer)

                                'If ((pf.[Function] = Excel.XlConsolidationFunction.xlSum) AndAlso (pf.Caption.Trim().Substring((pf.Caption.Trim().Length - pf.SourceName.Length)) = pf.SourceName)) Then
                                'WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption + " ", writer)
                                'Else
                                'WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption, writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.SourceName + " ", writer)
                                'End If


                                WriteXmlAttribute(GetPropertyName(Function() pf.[Function]), Convert.ToInt32(pf.[Function]).ToString(), writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.Calculation), Convert.ToInt32(pf.Calculation).ToString(), writer)

                                If (pf.Calculation <> Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation) And (pf.Calculation <> Excel.XlPivotFieldCalculation.xlPercentOfColumn) Then
                                    WriteXmlAttribute(GetPropertyName(Function() pf.BaseField), pf.BaseField, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.BaseItem), pf.BaseItem, writer)
                                Else
                                    WriteXmlAttribute(GetPropertyName(Function() pf.BaseField), String.Empty, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.BaseItem), String.Empty, writer)
                                End If


                                WriteXmlAttribute(GetPropertyName(Function() pf.NumberFormat), pf.NumberFormat, writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.Orientation), Convert.ToInt32(pf.Orientation).ToString(), writer)
                                writer.WriteEndElement()
                            Next
                        End If
                        seqNum = (seqNum + 1)
                        writer.WriteStartElement("property")
                        WriteXmlAttribute("Seq", seqNum, writer)
                        WriteXmlAttribute("PivotElement", "DataPivotField", writer)
                        WriteXmlAttribute("SourceName", pvt.DataPivotField.Name, writer)
                        WriteXmlAttribute(GetPropertyName(Function() pvt.DataPivotField.Name), pvt.DataPivotField.Name, writer)
                        WriteXmlAttribute(GetPropertyName(Function() pvt.DataPivotField.Caption), pvt.DataPivotField.Caption, writer)
                        WriteXmlAttribute(GetPropertyName(Function() pvt.DataPivotField.Orientation), Convert.ToInt32(pvt.DataPivotField.Orientation).ToString(), writer)
                        writer.WriteEndElement()

                        If (pvt.VisibleFields.Count > 0) Then
                            For Each pf As Excel.PivotField In pvt.VisibleFields
                                seqNum = (seqNum + 1)
                                writer.WriteStartElement("property")
                                WriteXmlAttribute("Seq", seqNum, writer)
                                WriteXmlAttribute("PivotElement", "Position", writer)
                                If ((pf.Orientation = Excel.XlPivotFieldOrientation.xlDataField) AndAlso (pf.Name <> pvt.DataPivotField.Name)) Then
                                    WriteXmlAttribute("SourceName", pf.SourceName, writer)
                                    'If ((pf.[Function] = Excel.XlConsolidationFunction.xlSum) AndAlso (pf.Caption.Trim().Substring((pf.Caption.Trim().Length - pf.SourceName.Length)) = pf.SourceName)) Then
                                    'WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption + " ", writer)
                                    'Else
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption, writer)
                                    'WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.SourceName + " ", writer)
                                    'End If
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Name), pf.Value, writer)
                                ElseIf (pf.Name = pvt.DataPivotField.Name) Then

                                    WriteXmlAttribute(GetPropertyName(Function() pf.SourceName), "DataField", writer)

                                Else
                                    WriteXmlAttribute(GetPropertyName(Function() pf.SourceName), pf.SourceName, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Caption), pf.Caption, writer)
                                    WriteXmlAttribute(GetPropertyName(Function() pf.Name), pf.Name, writer)

                                End If
                                WriteXmlAttribute(GetPropertyName(Function() pf.Position), pf.Position, writer)
                                WriteXmlAttribute(GetPropertyName(Function() pf.Orientation), pf.Orientation, writer)
                                writer.WriteEndElement()
                            Next
                        End If
                        writer.WriteEndElement()
                    End If
                End If
            Next
            'writer.WriteEndDocument()
            writer.WriteEndElement()
            writer.Flush()

            stringWriter.Close()
            writer.Close()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "PivotFunctions", gUserId, gReportID)
        End Try

        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True

        Return StringWriter.ToString()

    End Function
    Private Sub WriteXmlAttribute(ByVal Name As String, ByVal Value As String, ByVal writer As XmlTextWriter)

        writer.WriteAttributeString(Name, Value)

    End Sub
    Public Function ReturnDiffPivotLayout(ByVal PivotVariantId As Integer) As String

        Dim entities = New DAL.SAPExlEntities()
        Dim oldStream As New IO.MemoryStream()
        Dim newStream As New IO.MemoryStream()

        Dim xml = New XmlDocument()
        If PivotVariantId.Equals(0) Then
            Dim lastUsedVariantID = OrklaRTBPL.PivotFacade.GetCurrentUserReportPivotLayoutVariant(gUserId, gReportID).Rows(0)("PivotLayoutVariantID")
            xml.Load(New XmlTextReader(New IO.StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = gUserId And rp.VariantID = DirectCast(lastUsedVariantID, Integer)).PivotLayout)))
        Else
            xml.Load(New XmlTextReader(New IO.StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = 0 And rp.VariantID = 0).PivotLayout)))
        End If
        Dim xnlNodes = xml.SelectNodes("/XtraSerializer/property")
        Dim writer As New IO.StreamWriter(oldStream)
        For Each xln As XmlNode In xnlNodes
            For Each xln1 As XmlNode In xln.ChildNodes
                writer.Write(xln1.OuterXml + Environment.NewLine)
            Next
        Next
        writer.Flush()
        oldStream.Position = 0

        Dim xml1 = New XmlDocument()
        If PivotVariantId.Equals(0) Then
            'Dim lastUsedVariantID = OrklaRTBPL.PivotFacade.GetCurrentUserReportPivotLayoutVariant(gUserId, gReportID).Rows(0)("PivotLayoutVariantID")
            'xml1.Load(New XmlTextReader(New IO.StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = gUserId And rp.VariantID = DirectCast(lastUsedVariantID, Integer)).PivotLayout)))
            xml1.Load(New XmlTextReader(New IO.StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = 0 And rp.VariantID = 0).PivotLayout)))
        Else
            xml1.Load(New XmlTextReader(New IO.StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = gUserId And rp.VariantID = PivotVariantId).PivotLayout)))
        End If
        Dim xnlNodes1 = xml1.SelectNodes("/XtraSerializer/property")
        Dim writer1 As New IO.StreamWriter(newStream)
        For Each xln2 As XmlNode In xnlNodes1
            For Each xln3 As XmlNode In xln2.ChildNodes
                writer1.Write(xln3.OuterXml + Environment.NewLine)
            Next
        Next
        writer1.Flush()
        newStream.Position = 0

        Dim xmlDiff = xml1.InnerXml
        Dim cc = New DifferenceEngine.DiffList_TextFile(New IO.StreamReader(oldStream))
        Dim dd = New DifferenceEngine.DiffList_TextFile(New IO.StreamReader(newStream))
        Dim time As Double = 0
        Dim de = New DifferenceEngine.DiffEngine()

        If PivotVariantId.Equals(0) Then
            time = de.ProcessDiff(dd, cc, DifferenceEngine.DiffEngineLevel.SlowPerfect)
        Else
            time = de.ProcessDiff(dd, cc, DifferenceEngine.DiffEngineLevel.SlowPerfect)
        End If
        Dim rep = de.DiffReport()

        For Each drs As DifferenceEngine.DiffResultSpan In rep
            If drs.Status = DifferenceEngine.DiffResultSpanStatus.NoChange Then
                'If PivotVariantId.Equals(0) Then
                '    For i = 0 To drs.Length - 1
                '        xmlDiff = xmlDiff.Replace(DirectCast(cc.GetByIndex(drs.DestIndex + i), DifferenceEngine.TextLine).Line, String.Empty)
                '    Next
                'Else
                For i = 0 To drs.Length - 1
                    xmlDiff = xmlDiff.Replace(DirectCast(dd.GetByIndex(drs.DestIndex + i), DifferenceEngine.TextLine).Line, String.Empty)
                Next
                'End If
                'ElseIf drs.Status = DifferenceEngine.DiffResultSpanStatus.DeleteSource Then
                '    If PivotVariantId.Equals(0) Then
                '        For i = 0 To drs.Length - 1
                '            xmlDiff = xmlDiff.Replace(DirectCast(cc.GetByIndex(drs.DestIndex + i), DifferenceEngine.TextLine).Line, String.Empty)
                '        Next
                '    Else
                '        For i = 0 To drs.Length - 1
                '            xmlDiff = xmlDiff.Replace(DirectCast(cc.GetByIndex(drs.SourceIndex + i), DifferenceEngine.TextLine).Line, String.Empty)
                '        Next
                '    End If
                'ElseIf drs.Status = DifferenceEngine.DiffResultSpanStatus.AddDestination Then
                '    If PivotVariantId.Equals(0) Then
                '        For i = 0 To drs.Length - 1
                '            xmlDiff = xmlDiff.Insert(drs.DestIndex + i, DirectCast(cc.GetByIndex(drs.DestIndex + i), DifferenceEngine.TextLine).Line)
                '        Next
                '    Else
                '        For i = 0 To drs.Length - 1
                '            xmlDiff = xmlDiff.Insert(drs.SourceIndex + i, DirectCast(dd.GetByIndex(drs.SourceIndex + i), DifferenceEngine.TextLine).Line)
                '        Next
                '    End If
            ElseIf drs.Status = DifferenceEngine.DiffResultSpanStatus.Replace Then
                'If PivotVariantId.Equals(0) Then
                '    For i = 0 To drs.Length - 1
                '        If DirectCast(dd.GetByIndex(drs.SourceIndex + i), DifferenceEngine.TextLine).Line.ToUpper.Equals(DirectCast(cc.GetByIndex(drs.DestIndex + i), DifferenceEngine.TextLine).Line.ToUpper) Then
                '            xmlDiff = xmlDiff.Replace(DirectCast(cc.GetByIndex(drs.DestIndex + i), DifferenceEngine.TextLine).Line, String.Empty)
                '        End If
                '    Next
                'Else
                For i = 0 To drs.Length - 1
                    If DirectCast(cc.GetByIndex(drs.SourceIndex + i), DifferenceEngine.TextLine).Line.ToUpper.Equals(DirectCast(dd.GetByIndex(drs.DestIndex + i), DifferenceEngine.TextLine).Line.ToUpper) Then
                        xmlDiff = xmlDiff.Replace(DirectCast(dd.GetByIndex(drs.DestIndex + i), DifferenceEngine.TextLine).Line, String.Empty)
                    End If
                Next
                'End If
            End If
        Next
        Return xmlDiff
    End Function

    Public Sub LoadPivotLayout()
        Dim bPvtFirst As Boolean
        Dim pvt As Excel.PivotTable
        Dim pf As Excel.PivotField
        Dim pi As Excel.PivotItem
        Dim sItem As String, sAllItems As String
        Dim y As Integer
        Dim shActive As Worksheet

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = Excel.XlCalculation.xlCalculationManual

        Try
            For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Application.ActiveWorkbook.Sheets("PvtTableDef").ListObjects
                bPvtFirst = True
                If listObject.Name.Equals("PvtTableDef") Then
                    For i = 1 To listObject.ListRows.Count
                        If i > 1 Then
                            If listObject.ListRows(i - 1).Range(1, listObject.ListColumns("SheetName").Index).value.ToString() <> listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value.ToString() Then
                                pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
                                pvt.ClearAllFilters()
                            End If
                        Else
                            pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
                            pvt.ClearAllFilters()
                        End If
                        Select Case listObject.ListRows(i).Range(1, listObject.ListColumns("PivotElement").Index).value
                            Case "Table"

                                If pvt.DataFields.Count <= 1 Then
                                    pvt.PivotFields(1).Orientation = Excel.XlPivotFieldOrientation.xlDataField
                                    pvt.PivotFields(2).Orientation = Excel.XlPivotFieldOrientation.xlDataField
                                End If
                                pvt.DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlHidden
                                For Each pf In pvt.VisibleFields
                                    pf.Orientation = Excel.XlPivotFieldOrientation.xlHidden
                                Next pf
                                pvt.ColumnGrand = listObject.ListRows(i).Range(1, listObject.ListColumns("ColumnGrand").Index).value
                                pvt.RowGrand = listObject.ListRows(i).Range(1, listObject.ListColumns("RowGrand").Index).value
                                pvt.DisplayErrorString = listObject.ListRows(i).Range(1, listObject.ListColumns("DisplayErrorString").Index).value
                                Application.ActiveWorkbook.ShowPivotTableFieldList = listObject.ListRows(i).Range(1, listObject.ListColumns("FieldList").Index).value
                                pvt.PivotCache.MissingItemsLimit = listObject.ListRows(i).Range(1, listObject.ListColumns("MissingItems").Index).value
                                pvt.SaveData = listObject.ListRows(i).Range(1, listObject.ListColumns("SaveData").Index).value
                                pvt.ShowDrillIndicators = listObject.ListRows(i).Range(1, listObject.ListColumns("ShowDrillIndicators").Index).value
                                bPvtFirst = False
                            Case "Formulas"
                                Try
                                    'pvt.CalculatedFields.Add(listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value, "=" & listObject.ListRows(i).Range(1, listObject.ListColumns("Formula").Index).value & "", True)
                                    pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value).Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
                                Catch ex As Exception
                                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "PivotFunctions - Case Formulas", gUserId, gReportID)
                                End Try
                            Case "Page"
                                'pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
                                pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Name").Index).value)
                                pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
                                pf.Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
                                pf.CurrentPage = listObject.ListRows(i).Range(1, listObject.ListColumns("CurrentPage").Index).value

                                If listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value <> "" Then
                                    y = 0
                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
                                        pf.ClearAllFilters()
                                        pf.CurrentPage = "(All)"
                                    Else
                                        For Each pi In pf.PivotItems
                                            y = y + 1
                                            If y > 1 Then
                                                pi.Visible = False
                                            Else
                                                pi.Visible = True
                                            End If
                                        Next pi
                                    End If
                                    sAllItems = listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value
                                    If sAllItems <> Nothing Then
                                        y = 2

                                        Do While InStr(y, sAllItems, ";") > 0
                                            pf.EnableMultiplePageItems = True
                                            sItem = Mid(sAllItems, y, InStr(y, sAllItems, ";") - y)
                                            y = y + Len(sItem) + 1
                                            If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
                                                pf.PivotItems(sItem).Visible = False
                                            Else
                                                pf.PivotItems(sItem).Visible = True
                                            End If

                                            If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 1 Then 'Check if to include or exclude items.
                                                If sItem <> "" And sItem <> pf.PivotItems(1).SourceName Then 'Check if the only item left is the first one.
                                                    pf.PivotItems(1).Visible = False
                                                End If
                                            End If
                                        Loop
                                        pf.EnableMultiplePageItems = True
                                    Else
                                        pf.ClearAllFilters()
                                        pf.EnableMultiplePageItems = False
                                    End If
                                Else
                                    pf.ClearAllFilters()
                                    pf.EnableMultiplePageItems = False
                                End If
                            Case "Column"
                                'pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
                                pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value)
                                pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
                                pf.Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
                                Select Case listObject.ListRows(i).Range(1, listObject.ListColumns("SubTotals").Index).value
                                    Case Nothing
                                    Case 0
                                        If pf.Subtotals(1) = True Then
                                            pf.Subtotals = New Boolean() {False, False, False, False, False, False, False, False, False, False, False, False}
                                        End If
                                    Case Is > 0
                                        pf.Subtotals(listObject.ListRows(i).Range(1, listObject.ListColumns("SubTotals").Index).value) = True
                                End Select

                                If listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value <> "" Then
                                    y = 0
                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
                                        pf.ClearAllFilters()
                                    Else
                                        For Each pi In pf.PivotItems
                                            y = y + 1
                                            If y > 1 Then
                                                pi.Visible = False
                                            Else
                                                pi.Visible = True
                                            End If
                                        Next pi
                                    End If
                                    sAllItems = listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value
                                    If sAllItems <> Nothing Then
                                        y = 2

                                        Do While InStr(y, sAllItems, ";") > 0
                                            sItem = Mid(sAllItems, y, InStr(y, sAllItems, ";") - y)
                                            y = y + Len(sItem) + 1

                                            If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
                                                pf.PivotItems(sItem).Visible = False
                                            Else

                                                pf.PivotItems(sItem).Visible = True
                                            End If



                                            If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 1 Then 'Check if to include or exclude items.
                                                If sItem <> "" And sItem <> pf.PivotItems(1).SourceName Then 'Check if the only item left is the first one.
                                                    pf.PivotItems(1).Visible = False
                                                End If
                                            End If
                                        Loop
                                        pf.EnableMultiplePageItems = True
                                    Else
                                        pf.ClearAllFilters()
                                        pf.EnableMultiplePageItems = False
                                    End If
                                Else
                                    pf.ClearAllFilters()
                                    pf.EnableMultiplePageItems = False
                                End If
                            Case "Row"
                                'pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
                                pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Name").Index).value)
                                pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
                                pf.Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
                                Select Case listObject.ListRows(i).Range(1, listObject.ListColumns("SubTotals").Index).value
                                    Case Nothing
                                    Case 0
                                        pf.Subtotals(1) = True
                                        pf.Subtotals = New Boolean() {False, False, False, False, False, False, False, False, False, False, False, False}
                                    Case Is > 0
                                        pf.Subtotals(listObject.ListRows(i).Range(1, listObject.ListColumns("SubTotals").Index).value) = True
                                End Select
                                If listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value <> "" Then
                                    y = 0
                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
                                        pf.ClearAllFilters()
                                    Else
                                        For Each pi In pf.PivotItems
                                            y = y + 1
                                            If y > 1 Then
                                                pi.Visible = False
                                            Else
                                                pi.Visible = True
                                            End If
                                        Next pi
                                    End If
                                    sAllItems = listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value
                                    If sAllItems <> Nothing Then
                                        y = 2

                                        Do While InStr(y, sAllItems, ";") > 0
                                            sItem = Mid(sAllItems, y, InStr(y, sAllItems, ";") - y)
                                            y = y + Len(sItem) + 1

                                            If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
                                                pf.PivotItems(sItem).Visible = False
                                            Else

                                                pf.PivotItems(sItem).Visible = True
                                            End If



                                            If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 1 Then 'Check if to include or exclude items.
                                                If sItem <> "" And sItem <> pf.PivotItems(1).SourceName Then 'Check if the only item left is the first one.
                                                    pf.PivotItems(1).Visible = False
                                                End If
                                            End If
                                        Loop
                                        pf.EnableMultiplePageItems = True
                                    Else
                                        pf.ClearAllFilters()
                                        pf.EnableMultiplePageItems = False
                                    End If



                                Else
                                    pf.ClearAllFilters()
                                    pf.EnableMultiplePageItems = False
                                End If
                            Case "DataPivotField"
                                'pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
                                If listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value <> 0 Then
                                    pf = pvt.DataPivotField
                                    pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
                                End If

                            Case "Data"
                                Try
                                    'pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
                                    pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value)

                                    pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
                                    pf.Function = listObject.ListRows(i).Range(1, listObject.ListColumns("Function").Index).value

                                    'pf.Name = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
                                    'pf.Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value

                                    pf.Calculation = listObject.ListRows(i).Range(1, listObject.ListColumns("Calculation").Index).value
                                    If pf.Calculation <> Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation Then
                                        pf.BaseField = listObject.ListRows(i).Range(1, listObject.ListColumns("BaseField").Index).value
                                        pf.BaseItem = CStr(listObject.ListRows(i).Range(1, listObject.ListColumns("BaseItem").Index).value)
                                    End If
                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("NumberFormat").Index).value <> " " Then
                                        pf.NumberFormat = listObject.ListRows(i).Range(1, listObject.ListColumns("NumberFormat").Index).value
                                    End If
                                Catch ex As Exception
                                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "PivotFunctions - Case Data", gUserId, gReportID)
                                End Try
                            Case "Position"
                                'pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
                                'If ((listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value = Excel.XlPivotFieldOrientation.xlDataField) AndAlso (listObject.ListRows(i).Range(1, listObject.ListColumns("Name").Index).value <> pvt.DataPivotField.Name)) Then
                                If listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value = Excel.XlPivotFieldOrientation.xlDataField Then
                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value.ToString().TrimEnd(String.Empty).ToUpper() <> listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value.ToString().TrimEnd(String.Empty).ToUpper() Then
                                        pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value)
                                        pf.Position = listObject.ListRows(i).Range(1, listObject.ListColumns("Position").Index).value
                                    End If
                                Else
                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value <> Nothing Then
                                        pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value)
                                        pf.Position = listObject.ListRows(i).Range(1, listObject.ListColumns("Position").Index).value
                                    End If
                                End If

                                If listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value = "DataPivotField" Then
                                    'pf = pvt.DataPivotField
                                    'pf.Position = listObject.ListRows(i).Range(1, listObject.ListColumns("Position").Index).value
                                    ' Else
                                    pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value)
                                    pf.Position = listObject.ListRows(i).Range(1, listObject.ListColumns("Position").Index).value
                                End If
                        End Select

                        If listObject.ListRows(i).Range(1, listObject.ListColumns("AutoSortOrder").Index).value <> 0 Then
                            pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)

                            pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Name").Index).value)
                            pf.AutoSort(listObject.ListRows(i).Range(1, listObject.ListColumns("AutoSortOrder").Index).value, listObject.ListRows(i).Range(1, listObject.ListColumns("AutoSortField").Index).value)
                        End If
                    Next i


                End If
            Next

            Application.ActiveWorkbook.ShowPivotTableFieldList = False
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "PivotFunctions", gUserId, gReportID)
        End Try

        pvt = Nothing
        pf = Nothing
        Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Exit Sub

    End Sub
    'Public Sub LoadPivotLayout()
    '    Dim bPvtFirst As Boolean
    '    Dim pvt As Excel.PivotTable
    '    Dim pf As Excel.PivotField
    '    Dim pi As Excel.PivotItem
    '    Dim sItem As String, sAllItems As String
    '    Dim y As Integer
    '    Dim shActive As Worksheet

    '    Application.ScreenUpdating = False
    '    Application.EnableEvents = False
    '    Application.Calculation = Excel.XlCalculation.xlCalculationManual

    '    Try
    '        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Application.ActiveWorkbook.Sheets("PvtTableDef").ListObjects
    '            bPvtFirst = True
    '            If listObject.Name.Equals("PvtTableDef") Then
    '                For i = 1 To listObject.ListRows.Count
    '                    Select Case listObject.ListRows(i).Range(1, listObject.ListColumns("PivotElement").Index).value
    '                        Case "Table"
    '                            pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
    '                            If pvt.DataFields.Count <= 1 Then
    '                                pvt.PivotFields(1).Orientation = Excel.XlPivotFieldOrientation.xlDataField
    '                                pvt.PivotFields(2).Orientation = Excel.XlPivotFieldOrientation.xlDataField
    '                            End If
    '                            pvt.DataPivotField.Orientation = Excel.XlPivotFieldOrientation.xlHidden
    '                            For Each pf In pvt.VisibleFields
    '                                pf.Orientation = Excel.XlPivotFieldOrientation.xlHidden
    '                            Next pf
    '                            pvt.ColumnGrand = listObject.ListRows(i).Range(1, listObject.ListColumns("ColumnGrand").Index).value
    '                            pvt.RowGrand = listObject.ListRows(i).Range(1, listObject.ListColumns("RowGrand").Index).value
    '                            pvt.DisplayErrorString = listObject.ListRows(i).Range(1, listObject.ListColumns("DisplayErrorString").Index).value
    '                            Application.ActiveWorkbook.ShowPivotTableFieldList = listObject.ListRows(i).Range(1, listObject.ListColumns("FieldList").Index).value
    '                            pvt.PivotCache.MissingItemsLimit = listObject.ListRows(i).Range(1, listObject.ListColumns("MissingItems").Index).value
    '                            pvt.SaveData = listObject.ListRows(i).Range(1, listObject.ListColumns("SaveData").Index).value
    '                            pvt.ShowDrillIndicators = listObject.ListRows(i).Range(1, listObject.ListColumns("ShowDrillIndicators").Index).value
    '                            bPvtFirst = False
    '                        Case "Formulas"
    '                            Try
    '                                'pvt.CalculatedFields.Add(listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value, "=" & listObject.ListRows(i).Range(1, listObject.ListColumns("Formula").Index).value & "", True)
    '                                pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value).Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
    '                            Catch ex As Exception
    '                                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "PivotFunctions - Case Formulas", gUserId, gReportID)
    '                            End Try
    '                        Case "Page"
    '                            pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Name").Index).value)
    '                            pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
    '                            pf.Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
    '                            pf.CurrentPage = listObject.ListRows(i).Range(1, listObject.ListColumns("CurrentPage").Index).value

    '                            If listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value <> "" Then
    '                                y = 0
    '                                If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
    '                                    pf.ClearAllFilters()
    '                                    pf.CurrentPage = "(All)"
    '                                Else
    '                                    For Each pi In pf.PivotItems
    '                                        y = y + 1
    '                                        If y > 1 Then
    '                                            pi.Visible = False
    '                                        Else
    '                                            pi.Visible = True
    '                                        End If
    '                                    Next pi
    '                                End If
    '                                sAllItems = listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value
    '                                y = 2
    '                                Do While InStr(y, sAllItems, ";") > 0
    '                                    pf.EnableMultiplePageItems = True
    '                                    sItem = Mid(sAllItems, y, InStr(y, sAllItems, ";") - y)
    '                                    y = y + Len(sItem) + 1
    '                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
    '                                        pf.PivotItems(sItem).Visible = False
    '                                    Else
    '                                        pf.PivotItems(sItem).Visible = True
    '                                    End If

    '                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 1 Then 'Check if to include or exclude items.
    '                                        If sItem <> "" And sItem <> pf.PivotItems(1).SourceName Then 'Check if the only item left is the first one.
    '                                            pf.PivotItems(1).Visible = False
    '                                        End If
    '                                    End If
    '                                Loop
    '                            Else
    '                                pf.ClearAllFilters()
    '                                pf.EnableMultiplePageItems = False
    '                            End If
    '                        Case "Column"
    '                                pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value)
    '                                pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
    '                                pf.Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
    '                                Select Case listObject.ListRows(i).Range(1, listObject.ListColumns("SubTotals").Index).value
    '                                    Case Nothing
    '                                    Case 0
    '                                        If pf.Subtotals(1) = True Then
    '                                            pf.Subtotals = New Boolean() {False, False, False, False, False, False, False, False, False, False, False, False}
    '                                        End If
    '                                    Case Is > 0
    '                                        pf.Subtotals(listObject.ListRows(i).Range(1, listObject.ListColumns("SubTotals").Index).value) = True
    '                                End Select

    '                                If listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value <> "" Then
    '                                    y = 0
    '                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
    '                                        pf.ClearAllFilters()
    '                                    Else
    '                                        For Each pi In pf.PivotItems
    '                                            y = y + 1
    '                                            If y > 1 Then
    '                                                pi.Visible = False
    '                                            Else
    '                                                pi.Visible = True
    '                                            End If
    '                                        Next pi
    '                                    End If
    '                                    sAllItems = listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value
    '                                    y = 2
    '                                    Do While InStr(y, sAllItems, ";") > 0
    '                                        sItem = Mid(sAllItems, y, InStr(y, sAllItems, ";") - y)
    '                                        y = y + Len(sItem) + 1
    '                                        If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
    '                                            pf.PivotItems(sItem).Visible = False
    '                                        Else
    '                                            pf.PivotItems(sItem).Visible = True
    '                                        End If
    '                                Loop
    '                                'If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 1 Then 'Check if to include or exclude items.
    '                                '    If sItem <> "" And sItem <> pf.PivotItems(1).SourceName Then 'Check if the only item left is the first one.
    '                                '        pf.PivotItems(1).Visible = False
    '                                '    End If
    '                                'End If
    '                                pf.EnableMultiplePageItems = True
    '                            Else
    '                                pf.ClearAllFilters()
    '                                pf.EnableMultiplePageItems = False
    '                            End If
    '                        Case "Row"
    '                                pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Name").Index).value)
    '                                pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
    '                                pf.Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
    '                                Select Case listObject.ListRows(i).Range(1, listObject.ListColumns("SubTotals").Index).value
    '                                    Case Nothing
    '                                    Case 0
    '                                        pf.Subtotals(1) = True
    '                                        pf.Subtotals = New Boolean() {False, False, False, False, False, False, False, False, False, False, False, False}
    '                                    Case Is > 0
    '                                        pf.Subtotals(listObject.ListRows(i).Range(1, listObject.ListColumns("SubTotals").Index).value) = True
    '                                End Select
    '                                If listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value <> "" Then
    '                                    y = 0
    '                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
    '                                        pf.ClearAllFilters()
    '                                    Else
    '                                        For Each pi In pf.PivotItems
    '                                            y = y + 1
    '                                            If y > 1 Then
    '                                                pi.Visible = False
    '                                            Else
    '                                                pi.Visible = True
    '                                            End If
    '                                        Next pi
    '                                    End If
    '                                    sAllItems = listObject.ListRows(i).Range(1, listObject.ListColumns("PivotItems").Index).value
    '                                    y = 2
    '                                    Do While InStr(y, sAllItems, ";") > 0
    '                                        sItem = Mid(sAllItems, y, InStr(y, sAllItems, ";") - y)
    '                                        y = y + Len(sItem) + 1
    '                                        If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 0 Then
    '                                            pf.PivotItems(sItem).Visible = False
    '                                        Else
    '                                            pf.PivotItems(sItem).Visible = True
    '                                        End If
    '                                    Loop
    '                                    'If listObject.ListRows(i).Range(1, listObject.ListColumns("Visible").Index).value = 1 Then 'Check if to include or exclude items.
    '                                    '    If sItem <> "" And sItem <> pf.PivotItems(1).SourceName Then 'Check if the only item left is the first one.
    '                                    '        pf.PivotItems(1).Visible = False
    '                                    '    End If
    '                                    'End If
    '                                    pf.EnableMultiplePageItems = True
    '                                Else
    '                                    pf.ClearAllFilters()
    '                                    pf.EnableMultiplePageItems = False
    '                                End If
    '                        Case "DataPivotField"
    '                                If listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value <> 0 Then
    '                                    pf = pvt.DataPivotField
    '                                    pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
    '                                End If

    '                        Case "Data"
    '                                Try
    '                                    pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value)

    '                                    pf.Orientation = listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value
    '                                    pf.Function = listObject.ListRows(i).Range(1, listObject.ListColumns("Function").Index).value

    '                                    'pf.Name = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value
    '                                    'pf.Caption = listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value

    '                                    pf.Calculation = listObject.ListRows(i).Range(1, listObject.ListColumns("Calculation").Index).value
    '                                    If pf.Calculation <> Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation Then
    '                                        pf.BaseField = listObject.ListRows(i).Range(1, listObject.ListColumns("BaseField").Index).value
    '                                        pf.BaseItem = CStr(listObject.ListRows(i).Range(1, listObject.ListColumns("BaseItem").Index).value)
    '                                    End If
    '                                    If listObject.ListRows(i).Range(1, listObject.ListColumns("NumberFormat").Index).value <> " " Then
    '                                        pf.NumberFormat = listObject.ListRows(i).Range(1, listObject.ListColumns("NumberFormat").Index).value
    '                                    End If
    '                                Catch ex As Exception
    '                                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "PivotFunctions - Case Data", gUserId, gReportID)
    '                                End Try
    '                        Case "Position"
    '                                'If ((listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value = Excel.XlPivotFieldOrientation.xlDataField) AndAlso (listObject.ListRows(i).Range(1, listObject.ListColumns("Name").Index).value <> pvt.DataPivotField.Name)) Then
    '                                '    'If listObject.ListRows(i).Range(1, listObject.ListColumns("Orientation").Index).value = Excel.XlPivotFieldOrientation.xlDataField Then
    '                                '    pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value)
    '                                '    pf.Position = listObject.ListRows(i).Range(1, listObject.ListColumns("Position").Index).value
    '                                'Else
    '                                If listObject.ListRows(i).Range(1, listObject.ListColumns("SourceName").Index).value = "DataPivotField" Then
    '                                    'pf = pvt.DataPivotField
    '                                    'pf.Position = listObject.ListRows(i).Range(1, listObject.ListColumns("Position").Index).value
    '                                    ' Else
    '                                    pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Caption").Index).value)
    '                                    pf.Position = listObject.ListRows(i).Range(1, listObject.ListColumns("Position").Index).value
    '                                End If
    '                    End Select

    '                    If listObject.ListRows(i).Range(1, listObject.ListColumns("AutoSortOrder").Index).value <> 0 Then
    '                        pvt = Application.ActiveWorkbook.Sheets(listObject.ListRows(i).Range(1, listObject.ListColumns("SheetName").Index).value).PivotTables(listObject.ListRows(i).Range(1, listObject.ListColumns("TableName").Index).value)
    '                        pf = pvt.PivotFields(listObject.ListRows(i).Range(1, listObject.ListColumns("Name").Index).value)
    '                        pf.AutoSort(listObject.ListRows(i).Range(1, listObject.ListColumns("AutoSortOrder").Index).value, listObject.ListRows(i).Range(1, listObject.ListColumns("AutoSortField").Index).value)
    '                    End If
    '                Next i
    '            End If
    '        Next

    '        Application.ActiveWorkbook.ShowPivotTableFieldList = False
    '    Catch ex As Exception
    '        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "PivotFunctions", gUserId, gReportID)
    '    End Try

    '    pvt = Nothing
    '    pf = Nothing
    '    Application.Calculation = Excel.XlCalculation.xlCalculationAutomatic
    '    Application.EnableEvents = True
    '    Application.ScreenUpdating = True
    '    Exit Sub

    'End Sub
    Public Function GetPropertyName(Of T)(expression As Expression(Of Func(Of T))) As String
        Dim body As MemberExpression = DirectCast(expression.Body, MemberExpression)
        Return body.Member.Name
    End Function

    Public Sub PivotShowDataPercentOfColumn()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Calculation = Excel.XlPivotFieldCalculation.xlPercentOfColumn
        Application.ActiveCell.PivotField.NumberFormat = "0.0 %;[Red]-0.0 %"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotShowDataPercentOfRow()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Calculation = Excel.XlPivotFieldCalculation.xlPercentOfRow
        Application.ActiveCell.PivotField.NumberFormat = "0.0 %;[Red]-0.0 %"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotShowDataNormal()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Calculation = Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation
        Application.ActiveCell.PivotField.NumberFormat = "#,##0;[Red]-#,##0"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotNumberFormatStandard()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Calculation = Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation
        Application.ActiveCell.PivotField.NumberFormat = "#,##0;[Red]-#,##0"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotNumberFormatPercent()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Calculation = Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation
        Application.ActiveCell.PivotField.NumberFormat = "0.0 %;[Red]-0.0 %"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotNumberFormatThousands()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Calculation = Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation
        Application.ActiveCell.PivotField.NumberFormat = "#,##0,;[Red]-#,##0,"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotNumberFormatMillions()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Calculation = Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation
        Application.ActiveCell.PivotField.NumberFormat = "#,##0.0,,;[Red]-#,##0.0,,"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotNumberFormatPrice()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Calculation = Excel.XlPivotFieldCalculation.xlNoAdditionalCalculation
        Application.ActiveCell.PivotField.NumberFormat = "#,##0.00;[Red]-#,##0.00"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotFieldFunctionSum()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Function = Excel.XlConsolidationFunction.xlSum
        Application.ActiveCell.PivotField.NumberFormat = "#,##0;[Red]-#,##0"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotFieldFunctionCount()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Function = Excel.XlConsolidationFunction.xlCount
        Application.ActiveCell.PivotField.NumberFormat = "#,##0;[Red]-#,##0"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub PivotFieldFunctionAverage()
        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        Application.ActiveCell.PivotField.Function = Excel.XlConsolidationFunction.xlAverage
        Application.ActiveCell.PivotField.NumberFormat = "#,##0;[Red]-#,##0"
CleanUp:
        Application.ScreenUpdating = True
    End Sub

    Public Sub AutofitColumns()
        Dim rActiveCell As Excel.Range
        Dim iMinimumColWidth As Integer
        Dim rColumn As Excel.Range

        Application.ScreenUpdating = False
        On Error GoTo CleanUp
        rActiveCell = Application.ActiveCell
        iMinimumColWidth = 6
        Application.ActiveCell.PivotTable.TableRange1.Select()
        Application.Selection.AutoFit()
        For Each rColumn In Application.Selection.Columns
            If rColumn.ColumnWidth < iMinimumColWidth Then
                rColumn.ColumnWidth = iMinimumColWidth
            End If
        Next
        rActiveCell.Select()
CleanUp:
        Application.ScreenUpdating = True
    End Sub
End Module

