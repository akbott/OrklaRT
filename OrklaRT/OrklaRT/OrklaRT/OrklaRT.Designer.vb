Partial Class OrklaRT
    Inherits Microsoft.Office.Tools.Ribbon.RibbonBase

    <System.Diagnostics.DebuggerNonUserCode()> _
   Public Sub New(ByVal container As System.ComponentModel.IContainer)
        MyClass.New()

        'Required for Windows.Forms Class Composition Designer support
        If (container IsNot Nothing) Then
            container.Add(Me)
        End If

    End Sub

    <System.Diagnostics.DebuggerNonUserCode()> _
    Public Sub New()
        MyBase.New(Globals.Factory.GetRibbonFactory())

        'This call is required by the Component Designer.
        InitializeComponent()

    End Sub

    'Component overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Component Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Component Designer
    'It can be modified using the Component Designer.
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(OrklaRT))
        Dim RibbonDropDownItemImpl1 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl2 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl3 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl4 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl5 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl6 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl7 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl8 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl9 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl10 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl11 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl12 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl13 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl14 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl15 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl16 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl17 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl18 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl19 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl20 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl21 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl22 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl23 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl24 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl25 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl26 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl27 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl28 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl29 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl30 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl31 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl32 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Dim RibbonDropDownItemImpl33 As Microsoft.Office.Tools.Ribbon.RibbonDropDownItem = Me.Factory.CreateRibbonDropDownItem
        Me.tabOrklaRT = Me.Factory.CreateRibbonTab
        Me.group5 = Me.Factory.CreateRibbonGroup
        Me.edtSAPSystem = Me.Factory.CreateRibbonEditBox
        Me.cboSAPSystems = Me.Factory.CreateRibbonComboBox
        Me.btnR3PLogon = Me.Factory.CreateRibbonButton
        Me.group1 = Me.Factory.CreateRibbonGroup
        Me.menu1 = Me.Factory.CreateRibbonMenu
        Me.menu2 = Me.Factory.CreateRibbonMenu
        Me.menu3 = Me.Factory.CreateRibbonMenu
        Me.Separator1 = Me.Factory.CreateRibbonSeparator
        Me.menu4 = Me.Factory.CreateRibbonMenu
        Me.menu5 = Me.Factory.CreateRibbonMenu
        Me.menu6 = Me.Factory.CreateRibbonMenu
        Me.menu7 = Me.Factory.CreateRibbonMenu
        Me.menu8 = Me.Factory.CreateRibbonMenu
        Me.menu9 = Me.Factory.CreateRibbonMenu
        Me.grpOptions = Me.Factory.CreateRibbonGroup
        Me.ddlMaterialPrice = Me.Factory.CreateRibbonDropDown
        Me.ddlSalesValue = Me.Factory.CreateRibbonDropDown
        Me.ddlQuantityUnit = Me.Factory.CreateRibbonDropDown
        Me.ddlCurrency = Me.Factory.CreateRibbonDropDown
        Me.edbCurrencyYear = Me.Factory.CreateRibbonEditBox
        Me.btnCreateNewPlan = Me.Factory.CreateRibbonButton
        Me.btnSavePriorities = Me.Factory.CreateRibbonButton
        Me.ddlBudgetVersion = Me.Factory.CreateRibbonDropDown
        Me.btnFormatGraph = Me.Factory.CreateRibbonButton
        Me.btnSaveGroup = Me.Factory.CreateRibbonButton
        Me.btnSaveBinTest = Me.Factory.CreateRibbonButton
        Me.btnSaveExcludedTypes = Me.Factory.CreateRibbonButton
        Me.btnUpdateSAP = Me.Factory.CreateRibbonButton
        Me.ddlShowOptions = Me.Factory.CreateRibbonDropDown
        Me.ddlShelfLifeTypes = Me.Factory.CreateRibbonDropDown
        Me.ddlMaterialsIncluded = Me.Factory.CreateRibbonDropDown
        Me.btnSaveList = Me.Factory.CreateRibbonButton
        Me.btnSaveManko = Me.Factory.CreateRibbonButton
        Me.btnCopySheet = Me.Factory.CreateRibbonButton
        Me.btnTransformPivotToTable = Me.Factory.CreateRibbonButton
        Me.ddlShowStocks = Me.Factory.CreateRibbonDropDown
        Me.ddlShowMD04Data = Me.Factory.CreateRibbonDropDown
        Me.btnTransformPivotToList = Me.Factory.CreateRibbonButton
        Me.grpPivotLayout = Me.Factory.CreateRibbonGroup
        Me.btnSaveLayout = Me.Factory.CreateRibbonButton
        Me.ddlPivotLayout = Me.Factory.CreateRibbonDropDown
        Me.btnDeletePivotLayout = Me.Factory.CreateRibbonButton
        Me.grpSettings = Me.Factory.CreateRibbonGroup
        Me.Menu10 = Me.Factory.CreateRibbonMenu
        Me.btnShowAllSheets = Me.Factory.CreateRibbonButton
        Me.btnHideRedSheets = Me.Factory.CreateRibbonButton
        Me.btnUnprotectSheet = Me.Factory.CreateRibbonButton
        Me.btnProtectSheet = Me.Factory.CreateRibbonButton
        Me.btnUserSettings = Me.Factory.CreateRibbonButton
        Me.btnAboutOrklaRT = Me.Factory.CreateRibbonButton
        Me.grpLabelMessage = Me.Factory.CreateRibbonGroup
        Me.lblMessage = Me.Factory.CreateRibbonLabel
        Me.ppTimer = New System.Windows.Forms.Timer(Me.components)
        Me.tabOrklaRT.SuspendLayout()
        Me.group5.SuspendLayout()
        Me.group1.SuspendLayout()
        Me.grpOptions.SuspendLayout()
        Me.grpPivotLayout.SuspendLayout()
        Me.grpSettings.SuspendLayout()
        Me.grpLabelMessage.SuspendLayout()
        Me.SuspendLayout()
        '
        'tabOrklaRT
        '
        Me.tabOrklaRT.Groups.Add(Me.group5)
        Me.tabOrklaRT.Groups.Add(Me.group1)
        Me.tabOrklaRT.Groups.Add(Me.grpOptions)
        Me.tabOrklaRT.Groups.Add(Me.grpPivotLayout)
        Me.tabOrklaRT.Groups.Add(Me.grpSettings)
        Me.tabOrklaRT.Groups.Add(Me.grpLabelMessage)
        Me.tabOrklaRT.Label = "OrklaRT v4.2"
        Me.tabOrklaRT.Name = "tabOrklaRT"
        '
        'group5
        '
        Me.group5.Items.Add(Me.edtSAPSystem)
        Me.group5.Items.Add(Me.cboSAPSystems)
        Me.group5.Items.Add(Me.btnR3PLogon)
        Me.group5.Label = "Tilkobling"
        Me.group5.Name = "group5"
        '
        'edtSAPSystem
        '
        Me.edtSAPSystem.Enabled = False
        Me.edtSAPSystem.Image = CType(resources.GetObject("edtSAPSystem.Image"), System.Drawing.Image)
        Me.edtSAPSystem.Label = "System"
        Me.edtSAPSystem.Name = "edtSAPSystem"
        Me.edtSAPSystem.ShowImage = True
        Me.edtSAPSystem.Text = Nothing
        '
        'cboSAPSystems
        '
        Me.cboSAPSystems.Image = CType(resources.GetObject("cboSAPSystems.Image"), System.Drawing.Image)
        Me.cboSAPSystems.Label = "System"
        Me.cboSAPSystems.Name = "cboSAPSystems"
        Me.cboSAPSystems.ShowImage = True
        Me.cboSAPSystems.Text = Nothing
        Me.cboSAPSystems.Visible = False
        '
        'btnR3PLogon
        '
        Me.btnR3PLogon.Label = "Logg på"
        Me.btnR3PLogon.Name = "btnR3PLogon"
        Me.btnR3PLogon.ShowImage = True
        '
        'group1
        '
        Me.group1.Items.Add(Me.menu1)
        Me.group1.Items.Add(Me.menu2)
        Me.group1.Items.Add(Me.menu3)
        Me.group1.Items.Add(Me.Separator1)
        Me.group1.Items.Add(Me.menu4)
        Me.group1.Items.Add(Me.menu5)
        Me.group1.Items.Add(Me.menu6)
        Me.group1.Items.Add(Me.menu7)
        Me.group1.Items.Add(Me.menu8)
        Me.group1.Items.Add(Me.menu9)
        Me.group1.Label = "group1"
        Me.group1.Name = "group1"
        Me.group1.Visible = False
        '
        'menu1
        '
        Me.menu1.Dynamic = True
        Me.menu1.Image = CType(resources.GetObject("menu1.Image"), System.Drawing.Image)
        Me.menu1.Label = "menu1"
        Me.menu1.Name = "menu1"
        Me.menu1.ShowImage = True
        Me.menu1.Visible = False
        '
        'menu2
        '
        Me.menu2.Dynamic = True
        Me.menu2.Image = CType(resources.GetObject("menu2.Image"), System.Drawing.Image)
        Me.menu2.Label = "menu2"
        Me.menu2.Name = "menu2"
        Me.menu2.OfficeImageId = "PivotTableChangeDataSource"
        Me.menu2.ShowImage = True
        Me.menu2.Visible = False
        '
        'menu3
        '
        Me.menu3.Dynamic = True
        Me.menu3.Image = CType(resources.GetObject("menu3.Image"), System.Drawing.Image)
        Me.menu3.Label = "menu3"
        Me.menu3.Name = "menu3"
        Me.menu3.OfficeImageId = "PivotTableChangeDataSource"
        Me.menu3.ShowImage = True
        Me.menu3.Visible = False
        '
        'Separator1
        '
        Me.Separator1.Name = "Separator1"
        '
        'menu4
        '
        Me.menu4.Dynamic = True
        Me.menu4.Image = CType(resources.GetObject("menu4.Image"), System.Drawing.Image)
        Me.menu4.Label = "menu4"
        Me.menu4.Name = "menu4"
        Me.menu4.OfficeImageId = "PivotTableChangeDataSource"
        Me.menu4.ShowImage = True
        Me.menu4.Visible = False
        '
        'menu5
        '
        Me.menu5.Dynamic = True
        Me.menu5.Image = CType(resources.GetObject("menu5.Image"), System.Drawing.Image)
        Me.menu5.Label = "menu5"
        Me.menu5.Name = "menu5"
        Me.menu5.OfficeImageId = "PivotTableChangeDataSource"
        Me.menu5.ShowImage = True
        Me.menu5.Visible = False
        '
        'menu6
        '
        Me.menu6.Dynamic = True
        Me.menu6.Image = CType(resources.GetObject("menu6.Image"), System.Drawing.Image)
        Me.menu6.Label = "menu6"
        Me.menu6.Name = "menu6"
        Me.menu6.OfficeImageId = "PivotTableChangeDataSource"
        Me.menu6.ShowImage = True
        Me.menu6.Visible = False
        '
        'menu7
        '
        Me.menu7.Dynamic = True
        Me.menu7.Image = CType(resources.GetObject("menu7.Image"), System.Drawing.Image)
        Me.menu7.Label = "menu7"
        Me.menu7.Name = "menu7"
        Me.menu7.OfficeImageId = "PivotTableChangeDataSource"
        Me.menu7.ShowImage = True
        Me.menu7.Visible = False
        '
        'menu8
        '
        Me.menu8.Dynamic = True
        Me.menu8.Image = CType(resources.GetObject("menu8.Image"), System.Drawing.Image)
        Me.menu8.Label = "menu8"
        Me.menu8.Name = "menu8"
        Me.menu8.OfficeImageId = "PivotTableChangeDataSource"
        Me.menu8.ShowImage = True
        Me.menu8.Visible = False
        '
        'menu9
        '
        Me.menu9.Dynamic = True
        Me.menu9.Image = CType(resources.GetObject("menu9.Image"), System.Drawing.Image)
        Me.menu9.Label = "menu9"
        Me.menu9.Name = "menu9"
        Me.menu9.OfficeImageId = "PivotTableChangeDataSource"
        Me.menu9.ShowImage = True
        Me.menu9.Visible = False
        '
        'grpOptions
        '
        Me.grpOptions.Items.Add(Me.ddlMaterialPrice)
        Me.grpOptions.Items.Add(Me.ddlSalesValue)
        Me.grpOptions.Items.Add(Me.ddlQuantityUnit)
        Me.grpOptions.Items.Add(Me.ddlCurrency)
        Me.grpOptions.Items.Add(Me.edbCurrencyYear)
        Me.grpOptions.Items.Add(Me.btnCreateNewPlan)
        Me.grpOptions.Items.Add(Me.btnSavePriorities)
        Me.grpOptions.Items.Add(Me.ddlBudgetVersion)
        Me.grpOptions.Items.Add(Me.btnFormatGraph)
        Me.grpOptions.Items.Add(Me.btnSaveGroup)
        Me.grpOptions.Items.Add(Me.btnSaveBinTest)
        Me.grpOptions.Items.Add(Me.btnSaveExcludedTypes)
        Me.grpOptions.Items.Add(Me.btnUpdateSAP)
        Me.grpOptions.Items.Add(Me.ddlShowOptions)
        Me.grpOptions.Items.Add(Me.ddlShelfLifeTypes)
        Me.grpOptions.Items.Add(Me.ddlMaterialsIncluded)
        Me.grpOptions.Items.Add(Me.btnSaveList)
        Me.grpOptions.Items.Add(Me.btnSaveManko)
        Me.grpOptions.Items.Add(Me.btnCopySheet)
        Me.grpOptions.Items.Add(Me.btnTransformPivotToTable)
        Me.grpOptions.Items.Add(Me.ddlShowStocks)
        Me.grpOptions.Items.Add(Me.ddlShowMD04Data)
        Me.grpOptions.Items.Add(Me.btnTransformPivotToList)
        Me.grpOptions.Label = "Alternativer"
        Me.grpOptions.Name = "grpOptions"
        Me.grpOptions.Visible = False
        '
        'ddlMaterialPrice
        '
        RibbonDropDownItemImpl1.Label = "Budget price"
        RibbonDropDownItemImpl1.Tag = "Budget"
        RibbonDropDownItemImpl2.Label = "Moving avg."
        RibbonDropDownItemImpl2.Tag = "Moving"
        RibbonDropDownItemImpl3.Label = "Standard price"
        RibbonDropDownItemImpl3.Tag = "Standard"
        Me.ddlMaterialPrice.Items.Add(RibbonDropDownItemImpl1)
        Me.ddlMaterialPrice.Items.Add(RibbonDropDownItemImpl2)
        Me.ddlMaterialPrice.Items.Add(RibbonDropDownItemImpl3)
        Me.ddlMaterialPrice.Label = "Material price"
        Me.ddlMaterialPrice.Name = "ddlMaterialPrice"
        Me.ddlMaterialPrice.Tag = "MaterialPrice"
        Me.ddlMaterialPrice.Visible = False
        '
        'ddlSalesValue
        '
        RibbonDropDownItemImpl4.Label = "Gross"
        RibbonDropDownItemImpl4.Tag = "Gross"
        RibbonDropDownItemImpl5.Label = "Net"
        RibbonDropDownItemImpl5.Tag = "Net"
        Me.ddlSalesValue.Items.Add(RibbonDropDownItemImpl4)
        Me.ddlSalesValue.Items.Add(RibbonDropDownItemImpl5)
        Me.ddlSalesValue.Label = "Sales value"
        Me.ddlSalesValue.Name = "ddlSalesValue"
        Me.ddlSalesValue.Tag = "SalesValue"
        Me.ddlSalesValue.Visible = False
        '
        'ddlQuantityUnit
        '
        RibbonDropDownItemImpl6.Label = "Antall"
        RibbonDropDownItemImpl6.Tag = "Antall"
        RibbonDropDownItemImpl7.Label = "KG"
        RibbonDropDownItemImpl7.Tag = "KG"
        Me.ddlQuantityUnit.Items.Add(RibbonDropDownItemImpl6)
        Me.ddlQuantityUnit.Items.Add(RibbonDropDownItemImpl7)
        Me.ddlQuantityUnit.Label = "Kvantum"
        Me.ddlQuantityUnit.Name = "ddlQuantityUnit"
        Me.ddlQuantityUnit.Tag = "QuantityUnit"
        Me.ddlQuantityUnit.Visible = False
        '
        'ddlCurrency
        '
        RibbonDropDownItemImpl8.Label = "Default"
        RibbonDropDownItemImpl8.Tag = "Default"
        RibbonDropDownItemImpl9.Label = "CZK"
        RibbonDropDownItemImpl9.Tag = "CZK"
        RibbonDropDownItemImpl10.Label = "DKK"
        RibbonDropDownItemImpl10.Tag = "DKK"
        RibbonDropDownItemImpl11.Label = "EUR"
        RibbonDropDownItemImpl11.Tag = "EUR"
        RibbonDropDownItemImpl12.Label = "PLN"
        RibbonDropDownItemImpl12.Tag = "PLN"
        RibbonDropDownItemImpl13.Label = "NOK"
        RibbonDropDownItemImpl13.Tag = "NOK"
        RibbonDropDownItemImpl14.Label = "SEK"
        RibbonDropDownItemImpl14.Tag = "SEK"
        RibbonDropDownItemImpl15.Label = "USD"
        RibbonDropDownItemImpl15.Tag = "USD"
        Me.ddlCurrency.Items.Add(RibbonDropDownItemImpl8)
        Me.ddlCurrency.Items.Add(RibbonDropDownItemImpl9)
        Me.ddlCurrency.Items.Add(RibbonDropDownItemImpl10)
        Me.ddlCurrency.Items.Add(RibbonDropDownItemImpl11)
        Me.ddlCurrency.Items.Add(RibbonDropDownItemImpl12)
        Me.ddlCurrency.Items.Add(RibbonDropDownItemImpl13)
        Me.ddlCurrency.Items.Add(RibbonDropDownItemImpl14)
        Me.ddlCurrency.Items.Add(RibbonDropDownItemImpl15)
        Me.ddlCurrency.Label = "Valuta"
        Me.ddlCurrency.Name = "ddlCurrency"
        Me.ddlCurrency.Tag = "Currency"
        Me.ddlCurrency.Visible = False
        '
        'edbCurrencyYear
        '
        Me.edbCurrencyYear.Label = "Curr Year"
        Me.edbCurrencyYear.MaxLength = 4
        Me.edbCurrencyYear.Name = "edbCurrencyYear"
        Me.edbCurrencyYear.Tag = "CurrencyYear"
        Me.edbCurrencyYear.Text = Nothing
        Me.edbCurrencyYear.Visible = False
        '
        'btnCreateNewPlan
        '
        Me.btnCreateNewPlan.Label = "Lagre nyplan"
        Me.btnCreateNewPlan.Name = "btnCreateNewPlan"
        Me.btnCreateNewPlan.Tag = "CreateNewPlan"
        Me.btnCreateNewPlan.Visible = False
        '
        'btnSavePriorities
        '
        Me.btnSavePriorities.Label = "Lagre priorities"
        Me.btnSavePriorities.Name = "btnSavePriorities"
        Me.btnSavePriorities.Tag = "SavePriorities"
        Me.btnSavePriorities.Visible = False
        '
        'ddlBudgetVersion
        '
        RibbonDropDownItemImpl16.Label = "Budget price"
        RibbonDropDownItemImpl16.Tag = "Budget"
        RibbonDropDownItemImpl17.Label = "Moving avg."
        RibbonDropDownItemImpl17.Tag = "Moving"
        RibbonDropDownItemImpl18.Label = "Standard price"
        RibbonDropDownItemImpl18.Tag = "Standard"
        Me.ddlBudgetVersion.Items.Add(RibbonDropDownItemImpl16)
        Me.ddlBudgetVersion.Items.Add(RibbonDropDownItemImpl17)
        Me.ddlBudgetVersion.Items.Add(RibbonDropDownItemImpl18)
        Me.ddlBudgetVersion.Label = "Budget Versions"
        Me.ddlBudgetVersion.Name = "ddlBudgetVersion"
        Me.ddlBudgetVersion.Tag = "BudgetVersion"
        Me.ddlBudgetVersion.Visible = False
        '
        'btnFormatGraph
        '
        Me.btnFormatGraph.Label = "Format Graph"
        Me.btnFormatGraph.Name = "btnFormatGraph"
        Me.btnFormatGraph.Tag = "FormatGraph"
        Me.btnFormatGraph.Visible = False
        '
        'btnSaveGroup
        '
        Me.btnSaveGroup.Label = "Lagre gruppe"
        Me.btnSaveGroup.Name = "btnSaveGroup"
        Me.btnSaveGroup.Tag = "SaveGroup"
        Me.btnSaveGroup.Visible = False
        '
        'btnSaveBinTest
        '
        Me.btnSaveBinTest.Label = "Lagre bintest"
        Me.btnSaveBinTest.Name = "btnSaveBinTest"
        Me.btnSaveBinTest.Tag = "SaveBinTest"
        Me.btnSaveBinTest.Visible = False
        '
        'btnSaveExcludedTypes
        '
        Me.btnSaveExcludedTypes.Label = "Save ExcludedTypes"
        Me.btnSaveExcludedTypes.Name = "btnSaveExcludedTypes"
        Me.btnSaveExcludedTypes.Tag = "SaveExcludedTypes"
        Me.btnSaveExcludedTypes.Visible = False
        '
        'btnUpdateSAP
        '
        Me.btnUpdateSAP.Label = "Update SAP"
        Me.btnUpdateSAP.Name = "btnUpdateSAP"
        Me.btnUpdateSAP.Tag = "UpdateSAP"
        Me.btnUpdateSAP.Visible = False
        '
        'ddlShowOptions
        '
        RibbonDropDownItemImpl19.Label = "Budget price"
        RibbonDropDownItemImpl19.Tag = "Budget"
        RibbonDropDownItemImpl20.Label = "Moving avg."
        RibbonDropDownItemImpl20.Tag = "Moving"
        RibbonDropDownItemImpl21.Label = "Standard price"
        RibbonDropDownItemImpl21.Tag = "Standard"
        Me.ddlShowOptions.Items.Add(RibbonDropDownItemImpl19)
        Me.ddlShowOptions.Items.Add(RibbonDropDownItemImpl20)
        Me.ddlShowOptions.Items.Add(RibbonDropDownItemImpl21)
        Me.ddlShowOptions.Label = "Show Options"
        Me.ddlShowOptions.Name = "ddlShowOptions"
        Me.ddlShowOptions.Tag = "ShowOptions"
        Me.ddlShowOptions.Visible = False
        '
        'ddlShelfLifeTypes
        '
        RibbonDropDownItemImpl22.Label = "Budget price"
        RibbonDropDownItemImpl22.Tag = "Budget"
        RibbonDropDownItemImpl23.Label = "Moving avg."
        RibbonDropDownItemImpl23.Tag = "Moving"
        RibbonDropDownItemImpl24.Label = "Standard price"
        RibbonDropDownItemImpl24.Tag = "Standard"
        Me.ddlShelfLifeTypes.Items.Add(RibbonDropDownItemImpl22)
        Me.ddlShelfLifeTypes.Items.Add(RibbonDropDownItemImpl23)
        Me.ddlShelfLifeTypes.Items.Add(RibbonDropDownItemImpl24)
        Me.ddlShelfLifeTypes.Label = "Holdbarhet Type"
        Me.ddlShelfLifeTypes.Name = "ddlShelfLifeTypes"
        Me.ddlShelfLifeTypes.Tag = "ShelfLifeType"
        Me.ddlShelfLifeTypes.Visible = False
        '
        'ddlMaterialsIncluded
        '
        RibbonDropDownItemImpl25.Label = "Budget price"
        RibbonDropDownItemImpl25.Tag = "Budget"
        RibbonDropDownItemImpl26.Label = "Moving avg."
        RibbonDropDownItemImpl26.Tag = "Moving"
        RibbonDropDownItemImpl27.Label = "Standard price"
        RibbonDropDownItemImpl27.Tag = "Standard"
        Me.ddlMaterialsIncluded.Items.Add(RibbonDropDownItemImpl25)
        Me.ddlMaterialsIncluded.Items.Add(RibbonDropDownItemImpl26)
        Me.ddlMaterialsIncluded.Items.Add(RibbonDropDownItemImpl27)
        Me.ddlMaterialsIncluded.Label = "Materials Included"
        Me.ddlMaterialsIncluded.Name = "ddlMaterialsIncluded"
        Me.ddlMaterialsIncluded.Tag = "MaterialsIncluded"
        Me.ddlMaterialsIncluded.Visible = False
        '
        'btnSaveList
        '
        Me.btnSaveList.Label = "Lagre blanderessurs"
        Me.btnSaveList.Name = "btnSaveList"
        Me.btnSaveList.Tag = "SaveList"
        Me.btnSaveList.Visible = False
        '
        'btnSaveManko
        '
        Me.btnSaveManko.Label = "Lagre Manko Kopi"
        Me.btnSaveManko.Name = "btnSaveManko"
        Me.btnSaveManko.Tag = "SaveManko"
        Me.btnSaveManko.Visible = False
        '
        'btnCopySheet
        '
        Me.btnCopySheet.Label = "Kopi Ark"
        Me.btnCopySheet.Name = "btnCopySheet"
        Me.btnCopySheet.Tag = "KopiArk"
        Me.btnCopySheet.Visible = False
        '
        'btnTransformPivotToTable
        '
        Me.btnTransformPivotToTable.Label = "Lagre tabell fra pivot"
        Me.btnTransformPivotToTable.Name = "btnTransformPivotToTable"
        Me.btnTransformPivotToTable.Tag = "TransformPivotToTable"
        Me.btnTransformPivotToTable.Visible = False
        '
        'ddlShowStocks
        '
        RibbonDropDownItemImpl28.Label = "Budget price"
        RibbonDropDownItemImpl28.Tag = "Budget"
        RibbonDropDownItemImpl29.Label = "Moving avg."
        RibbonDropDownItemImpl29.Tag = "Moving"
        RibbonDropDownItemImpl30.Label = "Standard price"
        RibbonDropDownItemImpl30.Tag = "Standard"
        Me.ddlShowStocks.Items.Add(RibbonDropDownItemImpl28)
        Me.ddlShowStocks.Items.Add(RibbonDropDownItemImpl29)
        Me.ddlShowStocks.Items.Add(RibbonDropDownItemImpl30)
        Me.ddlShowStocks.Label = "Vis beholdning i"
        Me.ddlShowStocks.Name = "ddlShowStocks"
        Me.ddlShowStocks.Tag = "ShowStock"
        Me.ddlShowStocks.Visible = False
        '
        'ddlShowMD04Data
        '
        RibbonDropDownItemImpl31.Label = "Budget price"
        RibbonDropDownItemImpl31.Tag = "Budget"
        RibbonDropDownItemImpl32.Label = "Moving avg."
        RibbonDropDownItemImpl32.Tag = "Moving"
        RibbonDropDownItemImpl33.Label = "Standard price"
        RibbonDropDownItemImpl33.Tag = "Standard"
        Me.ddlShowMD04Data.Items.Add(RibbonDropDownItemImpl31)
        Me.ddlShowMD04Data.Items.Add(RibbonDropDownItemImpl32)
        Me.ddlShowMD04Data.Items.Add(RibbonDropDownItemImpl33)
        Me.ddlShowMD04Data.Label = "Vis data for"
        Me.ddlShowMD04Data.Name = "ddlShowMD04Data"
        Me.ddlShowMD04Data.Tag = "ShowMD04Data"
        Me.ddlShowMD04Data.Visible = False
        '
        'btnTransformPivotToList
        '
        Me.btnTransformPivotToList.Label = "Lagre liste fra pivot"
        Me.btnTransformPivotToList.Name = "btnTransformPivotToList"
        Me.btnTransformPivotToList.Tag = "TransformPivotToList"
        Me.btnTransformPivotToList.Visible = False
        '
        'grpPivotLayout
        '
        Me.grpPivotLayout.Items.Add(Me.btnSaveLayout)
        Me.grpPivotLayout.Items.Add(Me.ddlPivotLayout)
        Me.grpPivotLayout.Items.Add(Me.btnDeletePivotLayout)
        Me.grpPivotLayout.Label = "PivotLayout"
        Me.grpPivotLayout.Name = "grpPivotLayout"
        Me.grpPivotLayout.Visible = False
        '
        'btnSaveLayout
        '
        Me.btnSaveLayout.Label = "Lagre PivotLayout"
        Me.btnSaveLayout.Name = "btnSaveLayout"
        '
        'ddlPivotLayout
        '
        Me.ddlPivotLayout.Label = "PivotLayout"
        Me.ddlPivotLayout.Name = "ddlPivotLayout"
        '
        'btnDeletePivotLayout
        '
        Me.btnDeletePivotLayout.Label = "Slett PivotLayout"
        Me.btnDeletePivotLayout.Name = "btnDeletePivotLayout"
        '
        'grpSettings
        '
        Me.grpSettings.Items.Add(Me.Menu10)
        Me.grpSettings.Items.Add(Me.btnUserSettings)
        Me.grpSettings.Items.Add(Me.btnAboutOrklaRT)
        Me.grpSettings.Label = "Instillinger"
        Me.grpSettings.Name = "grpSettings"
        Me.grpSettings.Visible = False
        '
        'Menu10
        '
        Me.Menu10.Items.Add(Me.btnShowAllSheets)
        Me.Menu10.Items.Add(Me.btnHideRedSheets)
        Me.Menu10.Items.Add(Me.btnUnprotectSheet)
        Me.Menu10.Items.Add(Me.btnProtectSheet)
        Me.Menu10.Label = "Ark Instillinger"
        Me.Menu10.Name = "Menu10"
        '
        'btnShowAllSheets
        '
        Me.btnShowAllSheets.Label = "Show All Sheets"
        Me.btnShowAllSheets.Name = "btnShowAllSheets"
        Me.btnShowAllSheets.ShowImage = True
        '
        'btnHideRedSheets
        '
        Me.btnHideRedSheets.Label = "Hide Red Sheets"
        Me.btnHideRedSheets.Name = "btnHideRedSheets"
        Me.btnHideRedSheets.ShowImage = True
        '
        'btnUnprotectSheet
        '
        Me.btnUnprotectSheet.Label = "Unprotect Sheet"
        Me.btnUnprotectSheet.Name = "btnUnprotectSheet"
        Me.btnUnprotectSheet.ShowImage = True
        '
        'btnProtectSheet
        '
        Me.btnProtectSheet.Label = "Protect Sheet"
        Me.btnProtectSheet.Name = "btnProtectSheet"
        Me.btnProtectSheet.ShowImage = True
        '
        'btnUserSettings
        '
        Me.btnUserSettings.Label = "Bruker oppsett"
        Me.btnUserSettings.Name = "btnUserSettings"
        '
        'btnAboutOrklaRT
        '
        Me.btnAboutOrklaRT.Label = "Om OrklaRT"
        Me.btnAboutOrklaRT.Name = "btnAboutOrklaRT"
        '
        'grpLabelMessage
        '
        Me.grpLabelMessage.Items.Add(Me.lblMessage)
        Me.grpLabelMessage.Name = "grpLabelMessage"
        Me.grpLabelMessage.Visible = False
        '
        'lblMessage
        '
        Me.lblMessage.Label = "Server ikke tilgjengelig"
        Me.lblMessage.Name = "lblMessage"
        '
        'ppTimer
        '
        Me.ppTimer.Interval = 300000
        '
        'OrklaRT
        '
        Me.Name = "OrklaRT"
        Me.RibbonType = "Microsoft.Excel.Workbook"
        Me.Tabs.Add(Me.tabOrklaRT)
        Me.tabOrklaRT.ResumeLayout(False)
        Me.tabOrklaRT.PerformLayout()
        Me.group5.ResumeLayout(False)
        Me.group5.PerformLayout()
        Me.group1.ResumeLayout(False)
        Me.group1.PerformLayout()
        Me.grpOptions.ResumeLayout(False)
        Me.grpOptions.PerformLayout()
        Me.grpPivotLayout.ResumeLayout(False)
        Me.grpPivotLayout.PerformLayout()
        Me.grpSettings.ResumeLayout(False)
        Me.grpSettings.PerformLayout()
        Me.grpLabelMessage.ResumeLayout(False)
        Me.grpLabelMessage.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents menu1 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents menu2 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents menu3 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents menu4 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents menu5 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents menu6 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents menu7 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents menu8 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents menu9 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents grpOptions As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents ddlMaterialPrice As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ddlSalesValue As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ddlQuantityUnit As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ddlCurrency As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents edbCurrencyYear As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents group1 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents group5 As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents grpPivotLayout As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents btnSaveLayout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ddlPivotLayout As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents Separator1 As Microsoft.Office.Tools.Ribbon.RibbonSeparator
    Public WithEvents tabOrklaRT As Microsoft.Office.Tools.Ribbon.RibbonTab
    Friend WithEvents grpLabelMessage As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents lblMessage As Microsoft.Office.Tools.Ribbon.RibbonLabel
    Friend WithEvents cboSAPSystems As Microsoft.Office.Tools.Ribbon.RibbonComboBox
    Friend WithEvents btnCreateNewPlan As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSavePriorities As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ddlBudgetVersion As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents btnFormatGraph As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSaveGroup As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSaveBinTest As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSaveExcludedTypes As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUpdateSAP As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ddlShowOptions As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ddlShelfLifeTypes As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ddlMaterialsIncluded As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents edtSAPSystem As Microsoft.Office.Tools.Ribbon.RibbonEditBox
    Friend WithEvents btnSaveList As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnSaveManko As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnR3PLogon As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnCopySheet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ddlShowStocks As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents ddlShowMD04Data As Microsoft.Office.Tools.Ribbon.RibbonDropDown
    Friend WithEvents grpSettings As Microsoft.Office.Tools.Ribbon.RibbonGroup
    Friend WithEvents Menu10 As Microsoft.Office.Tools.Ribbon.RibbonMenu
    Friend WithEvents btnShowAllSheets As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnHideRedSheets As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnUnprotectSheet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnProtectSheet As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnTransformPivotToTable As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnTransformPivotToList As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents ppTimer As System.Windows.Forms.Timer
    Friend WithEvents btnUserSettings As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnDeletePivotLayout As Microsoft.Office.Tools.Ribbon.RibbonButton
    Friend WithEvents btnAboutOrklaRT As Microsoft.Office.Tools.Ribbon.RibbonButton

End Class

Partial Class ThisRibbonCollection

    <System.Diagnostics.DebuggerNonUserCode()> _
    Friend ReadOnly Property OrklaRT() As OrklaRT
        Get
            Return Me.GetRibbon(Of OrklaRT)()
        End Get
    End Property
End Class
