Imports System.IO
Imports Microsoft.Office.Tools.Ribbon
Imports System.IO.Packaging
Imports System.Xml
Imports System.ComponentModel
Imports System.Net
Imports System.Net.Sockets

Public Class OrklaRT
    Public selectionTaskPane As Microsoft.Office.Tools.CustomTaskPane
    Private Sub OrklaRT_Load(ByVal sender As System.Object, ByVal e As RibbonUIEventArgs) Handles MyBase.Load
        Try
            If CheckServerAvailablity() Then
                Using entities = New DAL.SAPExlEntities()
                    'edtSAPSystem.Text = "R3P"
                    If Not entities.vwCurrentUser.SingleOrDefault().SAPSystem Is Nothing Then
                        edtSAPSystem.Text = entities.vwCurrentUser.SingleOrDefault().SAPSystem
                    End If
                    SQLDataHandler.GetConnection.InitializeConnection()
                    gUserId = OrklaRTBPL.CommonFacade.GetUserID()
                    'gUserId = 55
                End Using
                'Using entities = New DAL.SAPExlEntities()
                'SQLDataHandler.GetConnection.InitializeConnection()
                'Dim currentUser = OrklaRTBPL.CommonFacade.GetCurrentUser()
                'If currentUser.Rows.Count > 0 Then
                '    If Not currentUser.Rows(0)("SAPSystem") Is Nothing Then
                '        edtSAPSystem.Text = currentUser.Rows(0)("SAPSystem")
                '    End If
                'End If
                'gUserId = OrklaRTBPL.CommonFacade.GetUserID()
                'End Using
            Else
                Globals.Ribbons.OrklaRT.group5.Visible = False
                Globals.ThisAddIn.Application.EnableEvents = False
                Globals.Ribbons.OrklaRT.grpLabelMessage.Visible = True
            End If

        Catch
            'OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId)
        End Try
    End Sub
    Public Function CheckServerAvailablity() As Boolean
        Try
            Dim TcpClient As New TcpClient()
            'TcpClient.Connect("10.195.9.174", 1433)
            TcpClient.Connect(System.Configuration.ConfigurationManager.AppSettings("Server").ToString(), 1433)
            TcpClient.Close()
            Return True
        Catch ex As Exception
            'OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId)
            Return False
        End Try
    End Function
    Friend Sub LoadReportMenu()
        Try
            Using entities = New DAL.SAPExlEntities()
                For Each ribbonGroup As RibbonGroup In tabOrklaRT.Groups
                    Dim reportGroup As Integer = Convert.ToInt32(ribbonGroup.Tag)
                    Dim reportMenus = entities.ReportGroups.Where(Function(rg) rg.ReportGroup = reportGroup).OrderBy(Function(o) o.ReportSubGroupID)
                    For Each reportMenu In reportMenus
                        For Each ribbonMenu In ribbonGroup.Items
                            If TypeOf ribbonMenu Is RibbonMenu Then
                                If ribbonMenu.Name.Contains(reportMenu.ReportSubGroupID.ToString()) Then
                                    DirectCast(ribbonMenu, RibbonMenu).Tag = reportMenu.ReportSubGroupID
                                    DirectCast(ribbonMenu, RibbonMenu).Label = reportMenu.ReportSubGroupName
                                    DirectCast(ribbonMenu, RibbonMenu).Visible = True
                                End If
                            End If
                        Next
                    Next
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Friend Sub LoadReportMenuItems()

        Try
            Using entities = New DAL.SAPExlEntities()
                For Each ribbonGroup As RibbonGroup In tabOrklaRT.Groups
                    If ribbonGroup.Tag IsNot Nothing Then
                        For Each ribbonMenu In ribbonGroup.Items
                            If TypeOf ribbonMenu Is RibbonMenu Then
                                DirectCast(ribbonMenu, RibbonMenu).Items.Clear()                                
                                Dim reportMenus = OrklaRTBPL.CommonFacade.GetReports(Convert.ToInt32(ribbonMenu.Tag))
                                For Each reportMenu As Data.DataRow In reportMenus.Tables(0).Rows
                                    Dim ribbonButton As RibbonButton = Me.Factory.CreateRibbonButton()
                                    ribbonButton.Tag = reportMenu("ReportID")
                                    If reportMenu("BeginGroup") = True Then
                                        Dim ribbonSeparator As RibbonSeparator = Me.Factory.CreateRibbonSeparator()
                                        DirectCast(ribbonMenu, RibbonMenu).Items.Add(ribbonSeparator)
                                    End If
                                    ribbonButton.Label = reportMenu("ReportName")
                                    ribbonButton.Visible = reportMenu("Enabled")
                                    AddHandler ribbonButton.Click, AddressOf ribbonButton_Click
                                    DirectCast(ribbonMenu, RibbonMenu).Items.Add(ribbonButton)
                                Next
                            End If
                        Next
                    End If
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub

    Private Sub ribbonButton_Click(sender As Object, e As RibbonControlEventArgs)
        Try
            Dim reportID As Integer
            reportID = Convert.ToInt32(DirectCast(sender, RibbonButton).Tag)
            OpenReport(reportID, False)
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub LoadMaterialPrices()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlMaterialPrice.Items.Clear()
                For Each view In entities.vwMaterialPrice
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlMaterialPrice.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub LoadSalesValue()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlSalesValue.Items.Clear()
                For Each view In entities.vwSalesValue
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlSalesValue.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub LoadQuantityUnit()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlQuantityUnit.Items.Clear()
                For Each view In entities.vwQuantityUnit
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlQuantityUnit.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub LoadCurrency()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlCurrency.Items.Clear()
                For Each view In entities.vwCurrency
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlCurrency.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub LoadBudgetVersions()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlBudgetVersion.Items.Clear()
                For Each view In entities.vwBudgetVersions
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlBudgetVersion.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub LoadShowOptions()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlShowOptions.Items.Clear()
                For Each view In entities.vwShowOptions
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlShowOptions.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Public Sub LoadShelfLifeTypes()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlShelfLifeTypes.Items.Clear()
                For Each view In entities.vwShelfLifeTypes
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlShelfLifeTypes.Items.Add(rdi)
                Next
            End Using
            ddlShelfLifeTypes.SelectedItemIndex = 1
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub LoadShowStocks()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlShowStocks.Items.Clear()
                For Each view In entities.vwShowStocks
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlShowStocks.Items.Add(rdi)
                Next
            End Using            
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Public Sub LoadShowMD04Data()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlShowMD04Data.Items.Clear()
                For Each view In entities.vwShowMD04Data
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlShowMD04Data.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub LoadMaterialsIncluded()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlMaterialsIncluded.Items.Clear()
                For Each view In entities.vwMaterialsIncluded
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = view.ID
                    rdi.Label = view.Text
                    ddlMaterialsIncluded.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, Me.ToString(), gUserId, gReportID)
        End Try
    End Sub
    Private Sub ddlMaterialPrice_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlMaterialPrice.SelectionChanged
        gwbReport.Sheets("ReportOptions").Range(ddlMaterialPrice.Tag).Value = ddlMaterialPrice.SelectedItem.Tag
        Call RefreshStandardPivot()
    End Sub
    Private Sub ddlSalesValue_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlSalesValue.SelectionChanged
        gwbReport.Sheets("ReportOptions").Range(ddlSalesValue.Tag).Value = ddlSalesValue.SelectedItem.Tag
        Call RefreshStandardPivot()
    End Sub

    Private Sub ddlQuantityUnit_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlQuantityUnit.SelectionChanged
        gwbReport.Sheets("ReportOptions").Range(ddlQuantityUnit.Tag).Value = ddlQuantityUnit.SelectedItem.Tag
        Call RefreshStandardPivot()
    End Sub

    Private Sub ddlCurrency_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlCurrency.SelectionChanged
        gwbReport.Sheets("ReportOptions").Range(ddlCurrency.Tag).Value = ddlCurrency.SelectedItem.Tag
        Call RefreshStandardPivot()
    End Sub

    Private Sub ddlBudgetVersion_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlBudgetVersion.SelectionChanged
        gwbReport.Sheets("ReportOptions").Range(ddlBudgetVersion.Tag).Value = ddlBudgetVersion.SelectedItem.Tag
        Call RefreshStandardPivot()
    End Sub

    Private Sub edbCurrencyYear_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles edbCurrencyYear.TextChanged
        GetExchangeRates(edbCurrencyYear.Text)
        Call RefreshStandardPivot()
    End Sub
    Public Sub GetLockedOrders()
        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("LocalLists").ListObjects
            If listObject.Name.Equals("LockedTable") Then
                Try
                    If Not listObject.DataBodyRange Is Nothing Then
                        listObject.DataBodyRange.Delete()
                    End If
                    Dim lockedOrders = OrklaRTBPL.CommonFacade.GetLockedOrders(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant)
                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(lockedOrders.Tables(0))
                    data.MoveFirst()
                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, lockedOrders.Tables(0).Rows.Count, lockedOrders.Tables(0).Columns.Count)
                Catch
                End Try
            End If
        Next
    End Sub
    Private Sub GetExchangeRates(year As String)
        For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("ExchangeRates").ListObjects
            If listObject.Name.Equals("ExchangeRates") Then
                Try
                    Dim exchangeRates = OrklaRTBPL.CommonFacade.GetExchangeRates(year)
                    Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(exchangeRates.Tables(0))
                    data.MoveFirst()
                    Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, exchangeRates.Tables(0).Rows.Count, exchangeRates.Tables(0).Columns.Count)
                Catch
                End Try
            End If
        Next
    End Sub
    Private Sub btnSaveLayout_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSaveLayout.Click
        Try
            Dim pivotLayoutVariantForm As New PivotLayoutVariant(gUserId, gReportID, Convert.ToInt32(ddlPivotLayout.SelectedItem.Tag))
            Call pivotLayoutVariantForm.Show()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "OrklaRT-Ribbon", gUserId, gReportID)
        End Try
    End Sub

    Public Sub LoadPivotLayouts()
        Try
            Using entities = New DAL.SAPExlEntities()
                ddlPivotLayout.Items.Clear()
                Dim newLayout As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                newLayout.Tag = 0
                newLayout.Label = "Standard"
                ddlPivotLayout.Items.Add(newLayout)
                Dim reportPivotLayouts = entities.PivotLayoutVariants.Where(Function(rp) rp.ReportID = gReportID And rp.UserID = gUserId)
                For Each reportPivotLayout In reportPivotLayouts
                    Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
                    rdi.Tag = reportPivotLayout.ID
                    rdi.Label = reportPivotLayout.VariantName
                    ddlPivotLayout.Items.Add(rdi)
                Next
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "OrklaRT-Ribbon", gUserId, gReportID)
        End Try
    End Sub

    Private Sub ddlPivotLayout_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlPivotLayout.SelectionChanged
        Try
            Using entities = New DAL.SAPExlEntities()
                Dim xml = New XmlDocument()
                If (ddlPivotLayout.SelectedItem.Tag.Equals(0)) Then
                    xml.Load(New XmlTextReader(New StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = 0 And rp.VariantID = 0).PivotLayout)))
                Else
                    xml.Load(New XmlTextReader(New StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = gUserId And rp.VariantID = DirectCast(ddlPivotLayout.SelectedItem.Tag, Integer)).PivotLayout)))
                End If
                Application.DisplayAlerts = False
                Application.ActiveWorkbook.XmlMaps.Add(xml.InnerXml, "XtraSerializer")
                For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("PvtTableDef").ListObjects
                    If listObject.Name.Equals("PvtTableDef") Then
                        Try
                            'If (ddlPivotLayout.SelectedItem.Tag.Equals(0)) Then
                            '    listObject.XmlMap.ImportXml(xml.InnerXml, True)
                            'Else
                            listObject.XmlMap.ImportXml(ReturnDiffPivotLayout(DirectCast(ddlPivotLayout.SelectedItem.Tag, Integer)), True)
                            'End If
                        Catch
                        End Try
                    End If
                Next
                Application.DisplayAlerts = True
                Call PivotFunctions.LoadPivotLayout()
                OrklaRTBPL.PivotFacade.UpdateCurrentUserReportPivotLayoutVariant(gUserId, gReportID, DirectCast(ddlPivotLayout.SelectedItem.Tag, Integer))
            End Using
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "OrklaRT-Ribbon", gUserId, gReportID)
        End Try
    End Sub
    Public Sub GetPivotLayout()
        'Using entities = New DAL.SAPExlEntities()
        '    Dim xml = New XmlDocument()
        '    xml.Load(New XmlTextReader(New StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = 0 And rp.VariantID = 0).PivotLayout)))
        '    Globals.ThisAddIn.Application.DisplayAlerts = False
        '    Globals.ThisAddIn.Application.ActiveWorkbook.XmlMaps.Add(xml.InnerXml, "XtraSerializer")
        '    For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("PvtTableDef").ListObjects
        '        If listObject.Name.Equals("PvtTableDef") Then
        '            Try
        '                listObject.XmlMap.ImportXml(xml.InnerXml, True)
        '            Catch
        '            End Try
        '        End If
        '    Next
        '    Globals.ThisAddIn.Application.DisplayAlerts = True
        'End Using
    End Sub
    Private Sub GetReportComments(reportID As Int32)
        Try
            For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("Comments").ListObjects
                If listObject.Name.Equals("Comments") Then
                    Try
                        Dim reportComments = OrklaRTBPL.CommonFacade.GetReportComments(reportID)
                        Dim data = OrklaRTBPL.CommonFacade.ConvertToRecordset(reportComments.Tables(0))
                        data.MoveFirst()
                        Dim i As Integer = listObject.Range(2, 1).CopyFromRecordset(data, reportComments.Tables(0).Rows.Count, reportComments.Tables(0).Columns.Count)
                    Catch
                    End Try
                End If
            Next
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "OrklaRT-Ribbon", gUserId, gReportID)
        End Try
    End Sub

    'Friend Sub LoadSAPSystems()
    '    cboSAPSystems.Items.Clear()
    '    Dim entities = New DAL.SAPExlEntities()
    '    For Each sapSystem In entities.vwUserSAPSystems
    '        Dim rdi As RibbonDropDownItem = Globals.Factory.GetRibbonFactory().CreateRibbonDropDownItem()
    '        rdi.Tag = sapSystem.SAPSystem
    '        rdi.Label = sapSystem.SAPSystem
    '        cboSAPSystems.Items.Add(rdi)
    '    Next
    '    If Not IsDBNull(entities.CurrentUsers.Where(Function(cu) cu.UserName = Environment.UserDomainName + "\" + Environment.UserName).SingleOrDefault().SAPSystem) Then
    '        cboSAPSystems.Text = entities.CurrentUsers.Where(Function(cu) cu.UserName = Environment.UserDomainName + "\" + Environment.UserName).SingleOrDefault().SAPSystem
    '    End If
    'End Sub
    Private Sub btnCreateNewPlan_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCreateNewPlan.Click
        Call FixedProductionPlan.CreateNewPlan()
    End Sub
    Private Sub btnSavePriorities_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSavePriorities.Click
        Call FixedProductionPlan.SavePriorities()
    End Sub

    Private Sub btnFormatGraph_Click(sender As Object, e As RibbonControlEventArgs) Handles btnFormatGraph.Click
        Call CapacityLevelling.FormatGraph()
    End Sub

    Private Sub btnSaveGroup_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSaveGroup.Click
        Call StockTransfer.WriteGroups()
    End Sub

    Private Sub btnSaveBinTest_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSaveBinTest.Click
        Call StockTransfer.WriteBinTest()
    End Sub

    Private Sub btnSaveExcludedTypes_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSaveExcludedTypes.Click
        Call StockTransfer.WriteExclTypes()
    End Sub

    Private Sub btnUpdateSAP_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUpdateSAP.Click
        Call OptimizedLotSize.UploadSAP()
    End Sub

    Private Sub ddlShowOptions_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlShowOptions.SelectionChanged
        If ddlShowOptions.SelectedItem.Tag.Equals("Total") Then
            Call AllReports.StockDeviationShowTotal()
        Else
            Call AllReports.StockDeviationShowDeviation()
        End If
    End Sub

    Private Sub ddlShelfLifeTypes_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlShelfLifeTypes.SelectionChanged
        gwbReport.Sheets("ReportOptions").Range(ddlShelfLifeTypes.Tag).Value = ddlShelfLifeTypes.SelectedItem.Tag
        Call RefreshStandardPivot()
    End Sub
    Private Sub ddlShowStocks_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlShowStocks.SelectionChanged
        gwbReport.Sheets("ReportOptions").Range(ddlShowStocks.Tag).Value = ddlShowStocks.SelectedItem.Tag
        Call RefreshStandardPivot()
    End Sub
    Private Sub ddlShowMD04Data_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlShowMD04Data.SelectionChanged
        Try
            gwbReport.Sheets("ReportOptions").Range(ddlShowMD04Data.Tag).Value = ddlShowMD04Data.SelectedItem.Tag
            If ddlShowMD04Data.SelectedItem.Tag.ToString().Equals("Planl.behov") Then
                gwbReport.ActiveSheet.Range("Datafor").Value = "Data for planlagt behov vises(for ferdigvarer vil det være prognoser)"
            ElseIf ddlShowMD04Data.SelectedItem.Tag.ToString().Equals("Inngang") Then
                gwbReport.ActiveSheet.Range("Datafor").Value = "Data for inngang vises(for ferdigvarer vil det være planordrer)"
            ElseIf ddlShowMD04Data.SelectedItem.Tag.ToString().Equals("Behov") Then
                gwbReport.ActiveSheet.Range("Datafor").Value = "Data for behov vises(for ferdigvarer vil det være ordrer/leveringer)"
            End If
            Call RefreshStandardPivot()
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "OrklaRT-Ribbon", gUserId, gReportID)
        End Try
    End Sub

    Private Sub ddlMaterialsIncluded_SelectionChanged(sender As Object, e As RibbonControlEventArgs) Handles ddlMaterialsIncluded.SelectionChanged
        Application.Sheets("Pivot").PivotTables(1).PivotFields("Default_Test").CurrentPage = ddlMaterialsIncluded.SelectedItem.Tag
        Call OptimizedLotSize.MaterialsIncluded()
    End Sub

    Private Sub edtSAPSystem_TextChanged(sender As Object, e As RibbonControlEventArgs) Handles edtSAPSystem.TextChanged

    End Sub

    Private Sub btnSaveList_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSaveList.Click
        Call MixingPlan.WriteMixWC()
    End Sub

    Private Sub btnSaveManko_Click(sender As Object, e As RibbonControlEventArgs) Handles btnSaveManko.Click
        If Application.ActiveSheet.Name.Equals("MankoKunder") Then
            Call SalesOrder.MakeSheetCopyLocal()
        ElseIf Application.ActiveSheet.Name.Equals("MankoUke") Then
            Call SalesOrder.MakeSheetCopyWeek()
        Else
            Call SalesOrder.MakeSheetCopyManko()
        End If
    End Sub

    Public Sub OpenReport(reportID As Integer, fromRightClick As Boolean)

        Dim report = OrklaRTBPL.CommonFacade.GetReportDefinition(reportID)
        Dim currentUser = OrklaRTBPL.CommonFacade.GetCurrentUser()
        gReportID = reportID
        'Using entities = New DAL.SAPExlEntities()
        '    Dim reportDefinition = entities.Reports.Where(Function(r) r.ReportID = reportID).SingleOrDefault()
        '    reportDefinition.ReportDefinition = File.ReadAllBytes("C:\OrklaRT-v4.2\OrklaRTReports\" + reportDefinition.ReportName + ".xlsm")
        '    entities.SaveChanges()
        'End Using
        If Not IsNothing(report.Rows(0)("ReportDefinition")) Then
            Dim bytes As Byte() = report.Rows(0)("ReportDefinition")
            Dim temporaryFile = Path.GetTempPath + currentUser.Rows(0)("SAPSystem").ToString() + "\" + report.Rows(0)("ReportName").ToString() + ".xlsm"
            If (Not Directory.Exists(Path.GetTempPath + currentUser.Rows(0)("SAPSystem").ToString())) Then Directory.CreateDirectory(Path.GetTempPath + currentUser.Rows(0)("SAPSystem").ToString())
            Try
                File.WriteAllBytes(temporaryFile, bytes)
                File.SetAttributes(temporaryFile, File.GetAttributes(temporaryFile) Or FileAttributes.Temporary)
            Catch ex As Exception
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod.Name, "OrklaRT-Ribbon", gUserId, gReportID)
            End Try

            For Each workbook As Excel.Workbook In Application.Workbooks
                If workbook.Name = report.Rows(0)("ReportName").ToString() + ".xlsm" Or workbook.FullName = report.Rows(0)("ReportName").ToString() + ".xlsm" Then
                    MessageBox.Show(report.Rows(0)("ReportName").ToString() + " rapport is allerede åpnet!")
                    Exit Sub
                End If
            Next

            Dim wb As Excel.Workbook = Application.Workbooks.Open(temporaryFile, , False)
            OrklaRTBPL.SelectionFacade.UpdateCurrentUserReportCount(reportID, gUserId)

            If selectionTaskPane Is Nothing Then
                If wb.ActiveSheet.Name.Equals("Rapport info") Then
                    selectionTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(New Selection(reportID, currentUser.Rows(0)("ID"), fromRightClick), report.Rows(0)("ReportName").ToString())
                    Dim win8version As New Version(6, 2, 9200, 0)
                    If Environment.OSVersion.Platform = PlatformID.Win32NT AndAlso Environment.OSVersion.Version >= win8version Then
                        selectionTaskPane.Width = 450
                    Else
                        selectionTaskPane.Width = 450
                    End If
                    selectionTaskPane.Visible = True
                    selectionTaskPane.DockPositionRestrict = Microsoft.Office.Core.MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange
                End If
            End If

            If reportID <> 0 Then
                Call Common.ShowReportOptions()
                grpPivotLayout.Visible = True
                LoadPivotLayouts()
                If (reportID = 7 Or reportID = 63) Then GetLockedOrders()
                If reportID = 8 Or reportID = 24 Or reportID = 33 Or reportID = 39 Then
                    GetReportComments(reportID)
                End If
            End If

            GC.Collect()
            GC.WaitForPendingFinalizers()
            selectionTaskPane = Nothing
        End If
    End Sub


    Private Sub btnR3PLogon_Click(sender As Object, e As RibbonControlEventArgs) Handles btnR3PLogon.Click
        Try
            Globals.Ribbons.OrklaRT.grpLabelMessage.Visible = False
            SQLDataHandler.GetConnection.InitializeConnection()
            OrklaRTBPL.CommonFacade.UpdateOrklaRTVersion(gUserId, tabOrklaRT.Label)
            Using entities = New DAL.SAPExlEntities()
                If BPL.RfcConnection.CheckR3PConnection() Then
                    Application.Cursor = Excel.XlMousePointer.xlWait
                    Dim reportGroups = entities.vwReportGroups.OrderBy(Function(vRG) vRG.ID)
                    For Each reportGroup In reportGroups
                        For Each ribbonGroup As RibbonGroup In tabOrklaRT.Groups
                            If ribbonGroup.Name.Contains(reportGroup.ID) Then
                                ribbonGroup.Visible = True
                                ribbonGroup.Tag = reportGroup.ID
                                ribbonGroup.Label = reportGroup.Text
                            End If
                        Next
                    Next
                    LoadReportMenu()
                    LoadReportMenuItems()
                    grpSettings.Visible = True
                    If OrklaRTBPL.CommonFacade.GetUserGroup(gUserId).Equals(2) Then
                        Globals.Ribbons.OrklaRT.Menu10.Visible = True
                    Else
                        Globals.Ribbons.OrklaRT.Menu10.Visible = False
                    End If
                    Application.Cursor = Excel.XlMousePointer.xlDefault
                Else
                    Dim sapLoginForm As New SAPLogin(edtSAPSystem.Text, True)
                    Call sapLoginForm.ShowDialog()
                    Globals.Ribbons.OrklaRT.group1.Visible = False
                    Application.Cursor = Excel.XlMousePointer.xlWait
                    Dim reportGroups = entities.vwReportGroups.OrderBy(Function(vRG) vRG.ID)
                    For Each reportGroup In reportGroups
                        For Each ribbonGroup As RibbonGroup In tabOrklaRT.Groups
                            If ribbonGroup.Name.Contains(reportGroup.ID) Then
                                ribbonGroup.Visible = True
                                ribbonGroup.Tag = reportGroup.ID
                                ribbonGroup.Label = reportGroup.Text
                            End If
                        Next
                    Next
                    LoadReportMenu()
                    LoadReportMenuItems()
                    Application.Cursor = Excel.XlMousePointer.xlDefault
                    If edtSAPSystem.Text.Equals(String.Empty) Then
                        edtSAPSystem.Text = entities.vwCurrentUser.SingleOrDefault().SAPSystem
                    End If
                End If
            End Using
        Catch
        End Try
    End Sub

    Private Sub btnCopySheet_Click(sender As Object, e As RibbonControlEventArgs) Handles btnCopySheet.Click
        Utilities.MakeSheetCopy()
    End Sub

   
    Private Sub btnShowAllSheets_Click(sender As Object, e As RibbonControlEventArgs) Handles btnShowAllSheets.Click
        Call Utilities.ShowAllSheets()
    End Sub

    Private Sub btnHideRedSheets_Click(sender As Object, e As RibbonControlEventArgs) Handles btnHideRedSheets.Click
        Call Utilities.HideAllSheets()
    End Sub

    Private Sub btnUnprotectSheet_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUnprotectSheet.Click
        Call Utilities.SheetUnprotect()
    End Sub

    Private Sub btnProtectSheet_Click(sender As Object, e As RibbonControlEventArgs) Handles btnProtectSheet.Click
        Call Utilities.SheetProtect()
    End Sub

    Private Sub btnTransformPivotToTable_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTransformPivotToTable.Click
        Call Utilities.TransformPivotToTable()
    End Sub

    Private Sub btnTransformPivotToList_Click(sender As Object, e As RibbonControlEventArgs) Handles btnTransformPivotToList.Click
        Call Utilities.TransformPivotToList()
    End Sub

    Private Sub ppTimer_Tick(sender As Object, e As EventArgs) Handles ppTimer.Tick
        Try
            Call System.Windows.Forms.SendKeys.Send("{F8}")
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Timer", gUserId, gReportID)
        End Try
    End Sub

    Private Sub btnUserSettings_Click(sender As Object, e As RibbonControlEventArgs) Handles btnUserSettings.Click
        Dim setting As New UserSettings()
        Call setting.Show()
    End Sub

    Private Sub btnDeletePivotLayout_Click(sender As Object, e As RibbonControlEventArgs) Handles btnDeletePivotLayout.Click
        Try
            If Not (ddlPivotLayout.SelectedItem.Tag.Equals(0)) Then
                Using entities = New DAL.SAPExlEntities()
                    Dim xml = New XmlDocument()
                    xml.Load(New XmlTextReader(New StringReader(entities.PivotLayouts.SingleOrDefault(Function(rp) rp.ReportID = gReportID And rp.UserID = 0 And rp.VariantID = 0).PivotLayout)))
                    Application.DisplayAlerts = False
                    Application.ActiveWorkbook.XmlMaps.Add(xml.InnerXml, "XtraSerializer")
                    For Each listObject As Microsoft.Office.Interop.Excel.ListObject In Globals.ThisAddIn.Application.ActiveWorkbook.Sheets("PvtTableDef").ListObjects
                        If listObject.Name.Equals("PvtTableDef") Then
                            Try
                                listObject.XmlMap.ImportXml(ReturnDiffPivotLayout(0), True)
                            Catch
                            End Try
                        End If
                    Next
                    Application.DisplayAlerts = True
                    OrklaRTBPL.PivotFacade.DeletePivotLayout(ddlPivotLayout.SelectedItem.Tag)
                    Call PivotFunctions.LoadPivotLayout()
                    LoadPivotLayouts()
                    OrklaRTBPL.PivotFacade.UpdateCurrentUserReportPivotLayoutVariant(gUserId, gReportID, 0)
                End Using
            End If
        Catch ex As Exception
            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Pivot Layout", gUserId, gReportID)
        End Try
    End Sub

    Private Sub btnAboutOrklaRT_Click(sender As Object, e As RibbonControlEventArgs) Handles btnAboutOrklaRT.Click
        Dim aboutOrklaRT As New AboutOrklaRT()
        aboutOrklaRT.Text = "Om OrklaRT v4.2"
        aboutOrklaRT.Show()
    End Sub
End Class