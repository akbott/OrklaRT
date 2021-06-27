using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Controls.Primitives;
using System.Globalization;
using System.Windows.Threading;
using OrklaRTBPL;
using System.ComponentModel;
using System.Reflection;

namespace SelectionPane
{
    /// <summary>
    /// Interaction logic for Selection.xaml
    /// </summary>
    public partial class Selection : UserControl
    {
        public int globalReportID, globalUserID, globalVariantID, datePickerCount;
        public bool newVariant = false;
        public bool rightClick;
        public DateTime fromDate, toDate;
        public Selection(int reportID, int userID, bool fromRightClick)
        {
            InitializeComponent();
            globalReportID = reportID;
            globalUserID = userID;
            tbiMultipleSelectionOptions.IsEnabled = false;
            btnDeleteSelection.Visibility = Visibility.Hidden;
            try
            {
                if (fromRightClick.Equals(false))
                {
                    cboSelectionVariant.ItemsSource = SelectionFacade.GetUserReportVariants(userID, reportID).Tables[0].DefaultView;
                    cboSelectionVariant.SelectedValue = SelectionFacade.GetCurrentUserVariant(reportID, userID);
                    SelectionFacade.DeleteCurrentUserReportSelections(reportID, userID);
                    if (cboSelectionVariant.SelectedValue.Equals(0))
                    {
                        SelectionFacade.InsertReportSelectionToCurrentUserReportSelections(reportID, userID);

                    }
                    else
                    {
                        SelectionFacade.InsertCurrentUserReportSelections(reportID, userID, SelectionFacade.GetCurrentUserVariant(reportID, userID));
                    }
                }
                else
                {

                    rightClick = fromRightClick;
                }

                switch (reportID)
                {
                    case 7:
                        SelectionFacade.InsertReportSelectionToCurrentUserReportSelections(52, userID);
                        SelectionFacade.UpdateCurrentUserReportSelectionHighValue(52, globalUserID, "ZVR021", "S", "I", "BT", "99991231");
                        break;
                    case 8:
                        SelectionFacade.MixingPlanProdPlanSelectionDate = DateTime.Now.ToString("yyyyMMdd");
                        break;
                    case 11:
                        SelectionFacade.InsertReportSelectionToCurrentUserReportSelections(52, userID);

                        SelectionFacade.UpdateCurrentUserReportSelectionLowValue(52, globalUserID, "ZVR021", "S", "I", "BT", DateTime.Now.Date.ToString("yyyyMMdd"));
                        SelectionFacade.UpdateCurrentUserReportSelectionHighValue(52, globalUserID, "ZVR021", "S", "I", "BT", "99991231");
                        break;
                    case 35:
                        SelectionFacade.InsertReportSelectionToCurrentUserReportSelections(60, userID);

                        SelectionFacade.UpdateCurrentUserReportSelectionLowValue(60, globalUserID, "ZVU016", "S", "I", "BT", DateTime.Now.AddYears(-2).Date.ToString("yyyyMM"));
                        SelectionFacade.UpdateCurrentUserReportSelectionHighValue(60, globalUserID, "ZVU016", "S", "I", "BT", DateTime.Now.Date.ToString("yyyyMM"));
                        break;
                    case 63:
                        goto case 7;
                }

                LoadControls(reportID);
                LoadUserSelectionData(reportID, userID, 0);
                if (rightClick.Equals(false))
                { LoadReportDefaultValues(); }

            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        public void LoadControls(int reportID)
        {
            using (var entities = new DAL.SAPExlEntities())
            {
                var report = entities.Reports.Where(p => p.ReportID == reportID).SingleOrDefault();
                txtWeek.IsEnabled = report.Week;
                txtMonth.IsEnabled = report.Month;
                txtYear.IsEnabled = report.Year;
                if (report.MultipleSelection.Equals(true))
                {
                    tbiMultipleSelectionOptions.Visibility = Visibility.Visible;
                }
                //if (report.MaterialHelp.Equals(true))
                //{
                //    grpMaterialSearch.Visibility = Visibility.Visible;
                //    cboPlants.ItemsSource = entities.vwPlants.ToList();
                //    cboMaterialTypes.ItemsSource = entities.vwMaterialTypes.ToList();
                //}
                //else
                //{
                //    grpSelection.Margin = new Thickness(grdSelection.Margin.Left, 170, grdSelection.Margin.Right, grdSelection.Margin.Bottom);
                //}
                if (report.Year.Equals(true))
                {
                    txtYear.Text = DateTime.Now.Year.ToString();
                }
                var controls = entities.ReportSelections.Where(p => p.ReportID == reportID && p.ControlType != null).OrderBy(o => o.SortOrder);
                foreach (var row in controls)
                {
                    LoadLabelText(Convert.ToInt32(row.SortOrder), row.ScreenID, row.FieldName);
                    dynamic controlSource = null;
                    switch (row.FieldName)
                    {
                        case "Merke":
                            controlSource = OrklaRTBPL.SelectionFacade.GetBrands().Tables[0].DefaultView;
                            break;
                        case "Fabrikk":
                            controlSource = OrklaRTBPL.SelectionFacade.GetPlants().Tables[0].DefaultView;
                            break;
                        case "Materialtype":
                            controlSource = OrklaRTBPL.SelectionFacade.GetMaterialTypes().Tables[0].DefaultView;
                            break;
                        case "Firmakode":
                            controlSource = OrklaRTBPL.SelectionFacade.GetCompanyCodes().Tables[0].DefaultView;
                            break;
                        case "Salesorganisasjon":
                            controlSource = OrklaRTBPL.SelectionFacade.GetSalesOrganizations().Tables[0].DefaultView;
                            break;
                        case "Materialgruppe":
                            controlSource = OrklaRTBPL.SelectionFacade.GetMaterialGroups().Tables[0].DefaultView;
                            break;
                        case "Produksjonsplanlegger":
                            controlSource = OrklaRTBPL.SelectionFacade.GetProductionScheduler().Tables[0].DefaultView;
                            break;
                        //case "Arbeidsstasjongruppe":
                        //    controlSource = entities.vwWorkCenterGroups.ToList();
                        //    break;
                        //case "Arbeidsstasjon":
                        //    controlSource = entities.vwWorkCenters.ToList();
                        //    break;
                        //case "Valuation Class":
                        //    controlSource = entities.vwValuationClass.ToList();
                        //    break;
                        case "Innkjøpsgruppe":
                            controlSource = OrklaRTBPL.SelectionFacade.GetPurchasingGroups().Tables[0].DefaultView;
                            break;
                        case "Lagernummer":
                            controlSource = OrklaRTBPL.SelectionFacade.GetWarehouseNumbers().Tables[0].DefaultView;
                            break;
                        case "Lager":
                            controlSource = OrklaRTBPL.SelectionFacade.GetStorageLocations().Tables[0].DefaultView;
                            break;
                        case "Materialplanlegger":
                            controlSource = OrklaRTBPL.SelectionFacade.GetMRPControllers().Tables[0].DefaultView;
                            break;
                        case "Produktansvarlig":
                            controlSource = OrklaRTBPL.SelectionFacade.GetProductResponsibles().Tables[0].DefaultView;
                            break;
                        case "MaterialArt":
                            controlSource = OrklaRTBPL.SelectionFacade.GetMaterialArts().Tables[0].DefaultView;
                            break;
                            //case "Merke":
                            //    controlSource = entities.vwBrands.ToList();
                            //    break;
                            //case "Fabrikk":
                            //    controlSource = entities.vwPlants.ToList();
                            //    break;
                            //case "Materialtype":
                            //    controlSource = entities.vwMaterialTypes.ToList();
                            //    break;
                            //case "Firmakode":
                            //    controlSource = entities.vwCompanyCodes.ToList();
                            //    break;
                            //case "Salesorganisasjon":
                            //    controlSource = entities.vwSalesOrganizations.ToList();
                            //    break;
                            //case "Materialgruppe":
                            //    controlSource = entities.vwMaterialGroups.ToList();
                            //    break;
                            //case "Produksjonsplanlegger":
                            //    controlSource = entities.vwProductionScheduler.ToList();
                            //    break;
                            ////case "Arbeidsstasjongruppe":
                            ////    controlSource = entities.vwWorkCenterGroups.ToList();
                            ////    break;
                            ////case "Arbeidsstasjon":
                            ////    controlSource = entities.vwWorkCenters.ToList();
                            ////    break;
                            ////case "Valuation Class":
                            ////    controlSource = entities.vwValuationClass.ToList();
                            ////    break;
                            //case "Innkjøpsgruppe":
                            //    controlSource = entities.vwPurchasingGroups.ToList();
                            //    break;
                            //case "Lagernummer":
                            //    controlSource = entities.vwWarehouseNumbers.ToList();
                            //    break;
                            //case "Lager":
                            //    controlSource = entities.vwStorageLocations.ToList();
                            //    break;
                            //case "Materialplanlegger":
                            //    controlSource = entities.vwMRPControllers.ToList();
                            //    break;
                    }
                    if (row.ControlType.Equals("ComboBox"))
                    {
                        if (row.SelectionOption.Equals("BT"))
                        {
                            ComboBox comboBox = null;
                            if (Environment.OSVersion.Platform.Equals(PlatformID.Win32NT) && Environment.OSVersion.Version >= new Version(6, 2, 9200, 0))
                            {
                                comboBox = new ComboBoxWin8();
                            }
                            else
                            {
                                comboBox = new ComboBox();
                            }
                            //comboBox.Name = "LV" + row.FieldName;
                            comboBox.Name = "LowValue";
                            comboBox.Tag = row.ScreenID;
                            comboBox.Width = 125;
                            comboBox.Height = 20;
                            comboBox.IsEditable = true;
                            comboBox.IsTextSearchEnabled = true;
                            comboBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                            comboBox.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            comboBox.ItemsSource = controlSource;
                            comboBox.SelectedValuePath = "ID";

                            switch (row.Mandatory)
                            {
                                case 1:
                                    comboBox.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                    break;
                                case 2:
                                    comboBox.Background = Brushes.Yellow;
                                    break;
                                default:
                                    comboBox.Background = Brushes.LightYellow;
                                    break;
                            }

                            comboBox.SelectionChanged += comboBox_SelectionChanged;

                            TextSearch.SetTextPath(comboBox, "ID");
                            DataTemplate dataTemplete = new DataTemplate();

                            FrameworkElementFactory stackPanel = new FrameworkElementFactory(typeof(StackPanel));
                            stackPanel.Name = "comboBoxStackPanel";
                            stackPanel.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);

                            FrameworkElementFactory id = new FrameworkElementFactory(typeof(TextBlock));
                            id.SetBinding(TextBlock.TextProperty, new Binding("ID"));
                            id.SetValue(TextBlock.WidthProperty, Double.Parse("75"));
                            stackPanel.AppendChild(id);

                            FrameworkElementFactory text = new FrameworkElementFactory(typeof(TextBlock));
                            text.SetBinding(TextBlock.TextProperty, new Binding("Text"));
                            stackPanel.AppendChild(text);

                            dataTemplete.VisualTree = stackPanel;
                            comboBox.ItemTemplate = dataTemplete;
                            comboBox.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            grdSelection.Children.Add(comboBox);


                            ComboBox comboBox1 = null;
                            if (Environment.OSVersion.Platform.Equals(PlatformID.Win32NT) && Environment.OSVersion.Version >= new Version(6, 2, 9200, 0))
                            {
                                comboBox1 = new ComboBoxWin8();
                            }
                            else
                            {
                                comboBox1 = new ComboBox();
                            }
                            //comboBox1.Name = "HV" + row.FieldName;
                            comboBox1.Name = "HighValue";
                            comboBox1.Tag = row.ScreenID;
                            comboBox1.Width = 125;
                            comboBox1.Height = 20;
                            comboBox1.IsEditable = true;
                            comboBox1.IsTextSearchEnabled = true;
                            comboBox1.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                            comboBox1.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            comboBox1.ItemsSource = controlSource;
                            comboBox1.SelectedValuePath = "ID";

                            switch (row.Mandatory)
                            {
                                case 1:
                                    comboBox1.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                    break;
                                case 2:
                                    comboBox1.Background = Brushes.Yellow;
                                    break;
                                default:
                                    comboBox1.Background = Brushes.LightYellow;
                                    break;
                            }

                            comboBox1.SelectionChanged += comboBox_SelectionChanged;

                            TextSearch.SetTextPath(comboBox1, "ID");
                            DataTemplate dataTemplete1 = new DataTemplate();

                            FrameworkElementFactory stackPanel1 = new FrameworkElementFactory(typeof(StackPanel));
                            stackPanel1.Name = "comboBoxStackPanel";
                            stackPanel1.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);

                            FrameworkElementFactory id1 = new FrameworkElementFactory(typeof(TextBlock));
                            id1.SetBinding(TextBlock.TextProperty, new Binding("ID"));
                            id1.SetValue(TextBlock.WidthProperty, Double.Parse("75"));
                            stackPanel1.AppendChild(id1);

                            FrameworkElementFactory text1 = new FrameworkElementFactory(typeof(TextBlock));
                            text1.SetBinding(TextBlock.TextProperty, new Binding("Text"));
                            stackPanel1.AppendChild(text1);

                            dataTemplete1.VisualTree = stackPanel1;
                            comboBox1.ItemTemplate = dataTemplete1;

                            comboBox1.Margin = new Thickness(285, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            grdSelection.Children.Add(comboBox1);

                        }
                        else
                        {
                            ComboBox comboBox = null;
                            if (Environment.OSVersion.Platform.Equals(PlatformID.Win32NT) && Environment.OSVersion.Version >= new Version(6, 2, 9200, 0))
                            {
                                comboBox = new ComboBoxWin8();
                            }
                            else
                            {
                                comboBox = new ComboBox();
                            }
                            comboBox.Name = "LowValue";
                            comboBox.Tag = row.ScreenID;
                            comboBox.Width = 125;
                            comboBox.Height = 20;
                            comboBox.IsEditable = true;
                            comboBox.IsTextSearchEnabled = true;
                            comboBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                            comboBox.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            comboBox.ItemsSource = controlSource;
                            comboBox.SelectedValuePath = "ID";

                            switch (row.Mandatory)
                            {
                                case 1:
                                    comboBox.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                    break;
                                case 2:
                                    comboBox.Background = Brushes.Yellow;
                                    break;
                                default:
                                    comboBox.Background = Brushes.LightYellow;
                                    break;
                            }

                            comboBox.SelectionChanged += comboBox_SelectionChanged;
                            comboBox.PreviewTextInput += comboBox_PreviewTextInput;

                            switch (row.FieldName)
                            {
                                case "Logical System":
                                    comboBox.DisplayMemberPath = "Text";
                                    break;
                                default:
                                    TextSearch.SetTextPath(comboBox, "ID");
                                    DataTemplate dataTemplete = new DataTemplate();

                                    FrameworkElementFactory stackPanel = new FrameworkElementFactory(typeof(StackPanel));
                                    stackPanel.Name = "comboBoxStackPanel";
                                    stackPanel.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);

                                    FrameworkElementFactory id = new FrameworkElementFactory(typeof(TextBlock));
                                    id.SetBinding(TextBlock.TextProperty, new Binding("ID"));
                                    id.SetValue(TextBlock.WidthProperty, Double.Parse("75"));
                                    stackPanel.AppendChild(id);

                                    FrameworkElementFactory text = new FrameworkElementFactory(typeof(TextBlock));
                                    text.SetBinding(TextBlock.TextProperty, new Binding("Text"));
                                    stackPanel.AppendChild(text);

                                    dataTemplete.VisualTree = stackPanel;
                                    comboBox.ItemTemplate = dataTemplete;
                                    break;
                            }

                            comboBox.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            grdSelection.Children.Add(comboBox);

                            if (row.MultipleSelection.Equals(true))
                            {
                                Button button = new Button();
                                button.Tag = row.ScreenID;
                                button.Width = 20;
                                button.Height = 20;
                                button.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                button.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                button.Content = ">>";
                                button.Background = Brushes.WhiteSmoke;
                                if (SelectionFacade.GetCurrentUserMultipleSelectionValue(reportID, globalUserID, row.ScreenID).Equals(true))
                                {
                                    button.BorderBrush = Brushes.DarkGreen;
                                    button.Foreground = Brushes.DarkGreen;
                                }
                                else
                                {
                                    button.BorderBrush = Brushes.Black;
                                }
                                button.FontWeight = FontWeights.Heavy;
                                button.Click += button_Click;
                                button.Margin = new Thickness(285, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                                grdSelection.Children.Add(button);
                            }
                        }
                    }
                    if (row.ControlType.Equals("DatePicker"))
                    {
                        datePickerCount = datePickerCount + 1;
                        if (row.SelectionOption.Equals("BT"))
                        {
                            if (row.LowValueVisible.Equals(true))
                            {
                                DatePicker datePicker = new DatePicker();
                                datePicker.Name = "LowValue";
                                datePicker.Tag = row.ScreenID;
                                datePicker.Width = 125;
                                datePicker.Height = 22;

                                switch (row.Mandatory)
                                {
                                    case 1:
                                        datePicker.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                        break;
                                    case 2:
                                        datePicker.Background = Brushes.Yellow;
                                        break;
                                    default:
                                        datePicker.Background = Brushes.LightYellow;
                                        break;
                                }

                                datePicker.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                datePicker.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                datePicker.SelectedDateChanged += datePicker_SelectedDateChanged;
                                datePicker.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                                grdSelection.Children.Add(datePicker);
                            }
                            else if (row.HighValueVisible.Equals(true))
                            {
                                DatePicker datePicker = new DatePicker();
                                datePicker.Name = "HighValue";
                                datePicker.Tag = row.ScreenID;
                                datePicker.Width = 125;
                                datePicker.Height = 22;

                                switch (row.Mandatory)
                                {
                                    case 1:
                                        datePicker.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                        break;
                                    case 2:
                                        datePicker.Background = Brushes.Yellow;
                                        break;
                                    default:
                                        datePicker.Background = Brushes.LightYellow;
                                        break;
                                }

                                datePicker.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                datePicker.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                datePicker.SelectedDateChanged += datePicker_SelectedDateChanged;
                                datePicker.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                                grdSelection.Children.Add(datePicker);
                            }
                            else
                            {
                                if (!row.UpdateLowValue.Equals(true) && !row.UpdateHighValue.Equals(true))
                                {
                                    DatePicker datePicker = new DatePicker();
                                    datePicker.Name = "LowValue";
                                    datePicker.Tag = row.ScreenID;
                                    datePicker.Width = 125;
                                    datePicker.Height = 22;

                                    switch (row.Mandatory)
                                    {
                                        case 1:
                                            datePicker.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                            break;
                                        case 2:
                                            datePicker.Background = Brushes.Yellow;
                                            break;
                                        default:
                                            datePicker.Background = Brushes.LightYellow;
                                            break;
                                    }

                                    datePicker.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                    datePicker.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                    datePicker.SelectedDateChanged += datePicker_SelectedDateChanged;
                                    datePicker.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                                    grdSelection.Children.Add(datePicker);
                                    DatePicker datePicker1 = new DatePicker();
                                    datePicker1.Name = "HighValue";
                                    datePicker1.Tag = row.ScreenID;
                                    datePicker1.Width = 125;
                                    datePicker1.Height = 22;

                                    switch (row.Mandatory)
                                    {
                                        case 1:
                                            datePicker1.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                            break;
                                        case 2:
                                            datePicker1.Background = Brushes.Yellow;
                                            break;
                                        default:
                                            datePicker1.Background = Brushes.LightYellow;
                                            break;
                                    }

                                    datePicker1.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                    datePicker1.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                    datePicker1.SelectedDateChanged += datePicker_SelectedDateChanged;
                                    if (row.HighValueVisible.Equals(true))
                                    {
                                        datePicker.Visibility = Visibility.Hidden;
                                        datePicker1.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                                    }
                                    else
                                    {
                                        datePicker1.Margin = new Thickness(285, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                                    }
                                    grdSelection.Children.Add(datePicker1);
                                }
                            }
                        }
                        else
                        {
                            DatePicker datePicker = new DatePicker();
                            datePicker.Name = "LowValue";
                            datePicker.Tag = row.ScreenID;
                            datePicker.Width = 125;
                            datePicker.Height = 22;

                            switch (row.Mandatory)
                            {
                                case 1:
                                    datePicker.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                    break;
                                case 2:
                                    datePicker.Background = Brushes.Yellow;
                                    break;
                                default:
                                    datePicker.Background = Brushes.LightYellow;
                                    break;
                            }

                            datePicker.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                            datePicker.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            datePicker.SelectedDateChanged += datePicker_SelectedDateChanged;
                            //if (datePickerCount.Equals(2))
                            //{
                            //    datePicker.Margin = new Thickness(285, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            //}
                            //else
                            //{
                            datePicker.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            //}
                            grdSelection.Children.Add(datePicker);

                            //if (row.MultipleSelection.Equals(true))
                            //{
                            //    Button button = new Button();
                            //    button.Tag = row.ScreenID;
                            //    button.Width = 20;
                            //    button.Height = 20;
                            //    button.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                            //    button.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            //    button.Content = ">>";
                            //    if (SelectionFacade.GetCurrentUserMultipleSelectionValue(reportID, globalUserID, row.ScreenID).Equals(true))
                            //    {
                            //        button.BorderBrush = Brushes.DarkGreen;
                            //        button.Foreground = Brushes.DarkGreen;
                            //    }
                            //    else
                            //    {
                            //        button.BorderBrush = Brushes.Black;
                            //    }
                            //    button.Background = Brushes.WhiteSmoke;
                            //    button.FontWeight = FontWeights.Heavy;
                            //    button.Click += button_Click;
                            //    button.Margin = new Thickness(200, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            //    grdSelection.Children.Add(button);
                            //}
                        }
                    }
                    if (row.ControlType.Equals("TextBox"))
                    {
                        if (row.SelectionOption.Equals("BT"))
                        {
                            TextBox textBox = new TextBox();
                            textBox.Name = "LowValue";
                            textBox.Tag = row.ScreenID;
                            if (row.FieldName.Equals("Periode"))
                                textBox.ToolTip = "YYYYMM";

                            switch (row.Mandatory)
                            {
                                case 1:
                                    textBox.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                    break;
                                case 2:
                                    textBox.Background = Brushes.Yellow;
                                    break;
                                default:
                                    textBox.Background = Brushes.LightYellow;
                                    break;
                            }


                            textBox.Width = 125;
                            textBox.Height = 20;
                            textBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                            textBox.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            textBox.TextChanged += textBox_TextChanged;
                            textBox.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            grdSelection.Children.Add(textBox);
                            TextBox textBox1 = new TextBox();
                            textBox1.Name = "HighValue";
                            textBox1.Tag = row.ScreenID;
                            if (row.FieldName.Equals("Periode"))
                                textBox1.ToolTip = "YYYYMM";

                            switch (row.Mandatory)
                            {
                                case 1:
                                    textBox1.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                    break;
                                case 2:
                                    textBox1.Background = Brushes.Yellow;
                                    break;
                                default:
                                    textBox1.Background = Brushes.LightYellow;
                                    break;
                            }

                            textBox1.Width = 125;
                            textBox1.Height = 20;
                            textBox1.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                            textBox1.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            textBox1.TextChanged += textBox_TextChanged;
                            textBox1.Margin = new Thickness(285, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            grdSelection.Children.Add(textBox1);

                        }
                        else
                        {
                            TextBox textBox = new TextBox();
                            textBox.Name = "LowValue";
                            textBox.Tag = row.ScreenID;
                            if (row.FieldName.Equals("Periode"))
                                textBox.ToolTip = "YYYYMM";

                            switch (row.Mandatory)
                            {
                                case 1:
                                    textBox.Background = new SolidColorBrush(Color.FromRgb(255, 128, 128));
                                    break;
                                case 2:
                                    textBox.Background = Brushes.Yellow;
                                    break;
                                default:
                                    textBox.Background = Brushes.LightYellow;
                                    break;
                            }

                            textBox.Width = 125;
                            textBox.Height = 20;
                            textBox.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                            textBox.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                            textBox.TextChanged += textBox_TextChanged;
                            if (row.MultipleSelection.Equals(true))
                            {
                                textBox.LostFocus += textBox_LostFocus;
                            }
                            textBox.Margin = new Thickness(150, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                            grdSelection.Children.Add(textBox);

                            if (row.MultipleSelection.Equals(true))
                            {
                                Button button = new Button();
                                button.Tag = row.ScreenID;
                                button.Width = 20;
                                button.Height = 20;
                                button.HorizontalAlignment = System.Windows.HorizontalAlignment.Left;
                                button.VerticalAlignment = System.Windows.VerticalAlignment.Top;
                                button.Content = ">>";
                                if (SelectionFacade.GetCurrentUserMultipleSelectionValue(reportID, globalUserID, row.ScreenID).Equals(true))
                                {
                                    button.BorderBrush = Brushes.DarkGreen;
                                    button.Foreground = Brushes.DarkGreen;
                                }
                                else
                                {
                                    button.BorderBrush = Brushes.Black;
                                }
                                button.Background = Brushes.WhiteSmoke;
                                button.FontWeight = FontWeights.Heavy;
                                button.Click += button_Click;
                                button.Margin = new Thickness(285, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Top, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Right, ((Label)grdSelection.FindName("lblName" + row.SortOrder)).Margin.Bottom);
                                grdSelection.Children.Add(button);
                            }
                        }
                    }
                }
            }
        }

        void button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(e.OriginalSource is System.Windows.Controls.Primitives.ToggleButton))
                {
                    if (!((Button)e.OriginalSource).Name.Equals("btnSaveSelection"))
                    {
                        tbiMultipleSelectionOptions.IsEnabled = true;
                        tbiMultipleSelectionOptions.Focus();
                        lblScreenID.Content = ((Button)e.OriginalSource).Tag.ToString();
                        using (var entities = new DAL.SAPExlEntities())
                        {
                            var controls = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((Button)e.OriginalSource).Tag).SingleOrDefault();
                            if (dgSelectSingleValues.Columns.Count > 1)
                            {
                                dgSelectSingleValues.Columns.RemoveAt(1);
                                if (controls.ScreenID.ToString().Equals("ZVR013") || controls.ScreenID.ToString().Equals("ZVS005") || controls.ScreenID.ToString().Equals("ZVR028"))
                                {
                                    dgSelectSingleValues.UpdateLayout();
                                    DataGridTextColumn dgCmbCol = new DataGridTextColumn();
                                    dgCmbCol.Header = "Enkeltverdier";
                                    dgCmbCol.Binding = new Binding("LowValue");
                                    dgSelectSingleValues.Columns.Add(dgCmbCol);
                                }
                                else
                                {
                                    dgSelectSingleValues.UpdateLayout();
                                    DataGridComboBoxColumn dgCmbCol = new DataGridComboBoxColumn();
                                    dgCmbCol.Header = "Enkeltverdier";
                                    LoadSingleValue(dgCmbCol, controls.FieldName);
                                    dgCmbCol.DisplayMemberPath = "CombinedText";
                                    dgCmbCol.SelectedValuePath = "ID";
                                    dgCmbCol.SelectedValueBinding = new Binding("LowValue");
                                    dgSelectSingleValues.Columns.Add(dgCmbCol);
                                }
                            }
                            else
                            {
                                if (controls.ScreenID.ToString().Equals("ZVR013") || controls.ScreenID.ToString().Equals("ZVS005") || controls.ScreenID.ToString().Equals("ZVR028"))
                                {
                                    dgSelectSingleValues.UpdateLayout();
                                    DataGridTextColumn dgCmbCol = new DataGridTextColumn();
                                    dgCmbCol.Header = "Enkeltverdier";
                                    dgCmbCol.Binding = new Binding("LowValue");
                                    dgSelectSingleValues.Columns.Add(dgCmbCol);
                                }
                                else
                                {
                                    DataGridComboBoxColumn dgCmbCol = new DataGridComboBoxColumn();
                                    dgCmbCol.Header = "Enkeltverdier";
                                    LoadSingleValue(dgCmbCol, controls.FieldName);
                                    dgCmbCol.DisplayMemberPath = "CombinedText";
                                    dgCmbCol.SelectedValuePath = "ID";
                                    dgCmbCol.SelectedValueBinding = new Binding("LowValue");
                                    dgSelectSingleValues.Columns.Add(dgCmbCol);
                                }
                            }
                            dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == ((Button)e.OriginalSource).Tag).ToList();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        void comboBox_PreviewTextInput(object sender, TextCompositionEventArgs e)
        {

        }

        void comboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!globalReportID.Equals(0) && !globalUserID.Equals(0))
            {
                try
                {
                    if (e.AddedItems.Count > 0 && ((ComboBox)sender).SelectedValue != null)
                    {
                        using (var entities = new DAL.SAPExlEntities())
                        {
                            var controls = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((ComboBox)sender).Tag).SingleOrDefault();
                            if (((ComboBox)sender).Name.Equals("LowValue"))
                            {
                                SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, ((ComboBox)sender).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                                if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((ComboBox)sender).Tag).SingleOrDefault().FieldName.Equals("Fabrikk"))
                                {
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Plant", ((ComboBox)sender).SelectedValue.ToString(), globalUserID, globalReportID);
                                    //if (rightClick.Equals(true))
                                    //{
                                        OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("Plant", ((ComboBox)sender).SelectedValue.ToString(), globalUserID, 2);
                                    //}
                                }
                                else if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((ComboBox)sender).Tag).SingleOrDefault().FieldName.Equals("Material Type"))
                                {
                                    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("MaterialType", ((ComboBox)sender).SelectedValue.ToString(), globalUserID, globalReportID);
                                }

                                if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((ComboBox)sender).Tag).SingleOrDefault().MultipleSelection.Equals(true))
                                {
                                    //SelectionFacade.InsertCurrentUserReportMultipleSelections(globalReportID, globalUserID, ((ComboBox)sender).Tag.ToString(), ((ComboBox)sender).SelectedValue.ToString());                                  
                                    SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(globalReportID, globalUserID, ((ComboBox)sender).Tag.ToString(), ((ComboBox)sender).SelectedValue.ToString(), true);
                                    dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == ((ComboBox)sender).Tag).ToList();
                                }

                                if (globalReportID.Equals(13) || globalReportID.Equals(62))
                                {
                                    foreach (Control control in grdSelection.Children)
                                    {
                                        if (control is ComboBox && entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((ComboBox)sender).Tag).SingleOrDefault().FieldName.Equals("Fabrikk"))
                                        {
                                            try
                                            {
                                                if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().FieldName.Equals("Produksjonsplanlegger"))
                                                {
                                                    ((ComboBox)control).ItemsSource = entities.vwProductionScheduler.Where(ps => ps.FilterField == ((ComboBox)sender).SelectedValue).ToList();
                                                }
                                            }
                                            catch { }
                                        }
                                    }
                                }
                            }
                            else
                            {
                                SelectionFacade.UpdateCurrentUserReportSelectionHighValue(globalReportID, globalUserID, ((ComboBox)sender).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                            }

                            if (entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == ((ComboBox)sender).Tag).Count() > 0)
                            {
                                SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(((DAL.CurrentUserReportMultipleSelections)dgSelectSingleValues.Items[0]).ID, ((ComboBox)sender).SelectedValue.ToString());
                                dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == ((ComboBox)sender).Tag).ToList();
                            }


                            if (entities.ReportsLinkedQuery.Where(rlq => rlq.ReportID == globalReportID).Count() > 0)
                            {
                                foreach (var linkedReport in entities.ReportsLinkedQuery.Where(rlq => rlq.ReportID == globalReportID))
                                {
                                    foreach (DataRow dRow in SelectionFacade.GetSubReportSelection(linkedReport.SubReports).Tables[0].Rows)
                                    {
                                        if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((ComboBox)sender).Tag).SingleOrDefault().FieldName.Equals(dRow["FieldName"].ToString()))
                                        {
                                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(linkedReport.SubReports, globalUserID, dRow["ScreenID"].ToString(), "S", dRow["Sign"].ToString(), dRow["SelectionOption"].ToString(), (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                                        }
                                    }
                                }
                            }

                            //// Material NUmber
                            //if (globalReportID.Equals(2))
                            //{
                            //    if (((ComboBox)sender).Tag.Equals("SP$00006"))
                            //    {
                            //        cboPlants.SelectedValue = ((ComboBox)sender).SelectedValue;
                            //    }
                            //}
                            ////Material Number

                            //For Mixing Plan 
                            if (globalReportID.Equals(8))
                            {
                                OrklaRTBPL.SelectionFacade.MixingPlanSelectionPlant = (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString());
                            }
                            //For Mixing Plan 

                            //For Supply area Stocks and Requirements Report
                            if (globalReportID.Equals(9))
                            {
                                if (((ComboBox)sender).Tag.Equals("ZVU006"))
                                {
                                    foreach (Control control in grdSelection.Children)
                                    {
                                        if (control is ComboBox)
                                        {
                                            if (control.Tag.ToString().Equals("ZVU010"))
                                            {
                                                ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetStorageTypes(((ComboBox)sender).SelectedValue.ToString()).Tables[0].DefaultView;
                                            }
                                        }
                                    }
                                }

                            }
                            //End Supply area Stocks and Requirements Report

                            //For Production Plan Report
                            if (globalReportID.Equals(7) || globalReportID.Equals(63))
                            {
                                if (((ComboBox)sender).Tag.Equals("0P_PLANT"))
                                {
                                    OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant = (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString());
                                    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(52, globalUserID, "ZVS009", "S", "I", "EQ", (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                                    //foreach (Control control in grdSelection.Children)
                                    //{
                                    //    if (control is ComboBox)
                                    //    {
                                    //        if (control.Tag.ToString().Equals("ZVU015"))
                                    //        {
                                    //            ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetWorkCenters(((ComboBox)sender).SelectedValue.ToString()).Tables[0].DefaultView;
                                    //        }
                                    //    }
                                    //}
                                }
                            }
                            //End Production Plan Report                           

                            //For Capacity Levelling Report
                            if (globalReportID.Equals(11))
                            {
                                if (((ComboBox)sender).Tag.Equals("ZVU047"))
                                {
                                    OrklaRTBPL.SelectionFacade.CapacityLevellingWorkGroupCenter = (((ComboBox)sender).SelectedValue == null ? String.Empty : ((DataRowView)((ComboBox)sender).SelectedItem).Row.ItemArray[0].ToString());
                                    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(52, globalUserID, "ZVS019", "S", "I", "EQ", OrklaRTBPL.SelectionFacade.CapacityLevellingWorkGroupCenter);
                                }
                                else if (((ComboBox)sender).Tag.Equals("ZVS009"))
                                {
                                    OrklaRTBPL.SelectionFacade.CapacityLevellingSelectionPlant = (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString());
                                    foreach (Control control in grdSelection.Children)
                                    {
                                        if (control is ComboBox)
                                        {
                                            if (control.Tag.ToString().Equals("ZVU015"))
                                            {
                                                ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetWorkCenters(((ComboBox)sender).SelectedValue.ToString()).Tables[0].DefaultView;
                                            }
                                            else if (control.Tag.ToString().Equals("ZVU047"))
                                            {
                                                ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetWorkCenterGroups(((ComboBox)sender).SelectedValue.ToString()).Tables[0].DefaultView;
                                            }
                                        }
                                    }
                                    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(52, globalUserID, "ZVS009", "S", "I", "EQ", (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                                }
                                else if (((ComboBox)sender).Tag.Equals("ZVU015"))
                                {
                                    OrklaRTBPL.SelectionFacade.CapacityLevellingSelectionWorkCenter = (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString());
                                }
                            }
                            //End Capacity Levelling Report

                            //For Stock Transfer Report

                            if (globalReportID.Equals(14))
                            {
                                if (((ComboBox)sender).Tag.Equals("ZVU006"))
                                {
                                    OrklaRTBPL.SelectionFacade.StockTransferSelectionWarehouse = (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString());
                                }
                            }

                            // End Stock Transfer report

                            //For Optimized LotSize Report
                            if (globalReportID.Equals(15))
                            {
                                //if (((ComboBox)sender).Tag.Equals("SP$00002"))
                                //{
                                //    foreach (Control control in grdSelection.Children)
                                //    {
                                //        if (control is ComboBox)
                                //        {
                                //            if (control.Tag.ToString().Equals("SP$00000"))
                                //            {
                                //                ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetWorkCenters(((ComboBox)sender).SelectedValue.ToString()).Tables[0].DefaultView;
                                //            }
                                //        }
                                //    }
                                //    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(57, globalUserID, "SP$00004", "S", "I", "EQ", (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                                //}
                                //else if (((ComboBox)sender).Tag.Equals("SP$00009"))
                                //{
                                //    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(57, globalUserID, "SP$00013", "S", "I", "EQ", (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                                //}
                                //else if (((ComboBox)sender).Tag.Equals("SP$00000"))
                                //{
                                //    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(57, globalUserID, "SP$00018", "S", "I", "EQ", (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                                //}

                            }
                            //End Optimized LotSize Report

                            //For DeliveryAgent Report
                            if (globalReportID.Equals(16))
                            {
                                if (((ComboBox)sender).Tag.Equals("0P_PLANT"))
                                {
                                    OrklaRTBPL.SelectionFacade.DeliveryAgentSelectionPlant = (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString());
                                    foreach (Control control in grdSelection.Children)
                                    {
                                        if (control is ComboBox)
                                        {
                                            if (control.Tag.ToString().Equals("ZVS011"))
                                            {
                                                ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMRPControllers(((ComboBox)sender).SelectedValue.ToString()).Tables[0].DefaultView;
                                            }
                                        }
                                    }
                                }
                            }
                            //End DeliveryAgent Report

                            //For MD04 Report
                            if (globalReportID.Equals(10))
                            {
                                if (((ComboBox)sender).Tag.Equals("0P_PLANT"))
                                {
                                    OrklaRTBPL.SelectionFacade.MD04SelectionPlant = (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString());
                                    foreach (Control control in grdSelection.Children)
                                    {
                                        if (control is ComboBox)
                                        {
                                            if (control.Tag.ToString().Equals("SP$11111"))
                                            {
                                                ((ComboBox)control).ItemsSource = BPL.RfcFunctions.GetProductGroups(((ComboBox)sender).SelectedValue.ToString()).DefaultView;
                                            }
                                        }
                                        if (control is ComboBox)
                                        {
                                            if (control.Tag.ToString().Equals("ZVS011"))
                                            {
                                                ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMRPControllers(((ComboBox)sender).SelectedValue.ToString()).Tables[0].DefaultView;
                                            }
                                        }
                                    }
                                }
                                else if (((ComboBox)sender).Tag.Equals("SP$11111"))
                                {
                                    OrklaRTBPL.SelectionFacade.DeleteCurrentUserMultipleSelections(10, globalUserID, "ZVR013");
                                    //((ComboBox)sender).ItemsSource = BPL.RfcFunctions.GetProductGroupMaterias(((ComboBox)sender).SelectedValue.ToString(), OrklaRTBPL.SelectionFacade.MD04SelectionPlant).DefaultView;
                                    foreach (DataRow dRow in OrklaRTBPL.SelectionFacade.GetProductGroupMaterials(OrklaRTBPL.SelectionFacade.MD04SelectionPlant, ((ComboBox)sender).SelectedValue.ToString()).Tables[0].Rows)
                                    {
                                        if (OrklaRTBPL.CommonFacade.IsNumeric(dRow["Material"].ToString()))
                                        {
                                            SelectionFacade.InsertCurrentUserReportMultipleSelections(10, globalUserID, "ZVR013", dRow["Material"].ToString());
                                        }
                                        else
                                        {
                                            foreach (DataRow dRow1 in OrklaRTBPL.SelectionFacade.GetProductGroupMaterials(OrklaRTBPL.SelectionFacade.MD04SelectionPlant, dRow["Material"].ToString()).Tables[0].Rows)
                                            {
                                                if (OrklaRTBPL.CommonFacade.IsNumeric(dRow1["Material"].ToString()))
                                                {
                                                    SelectionFacade.InsertCurrentUserReportMultipleSelections(10, globalUserID, "ZVR013", dRow1["Material"].ToString());
                                                }
                                                else
                                                {
                                                    foreach (DataRow dRow2 in OrklaRTBPL.SelectionFacade.GetProductGroupMaterials(OrklaRTBPL.SelectionFacade.MD04SelectionPlant, dRow1["Material"].ToString()).Tables[0].Rows)
                                                    {
                                                        if (OrklaRTBPL.CommonFacade.IsNumeric(dRow2["Material"].ToString()))
                                                        {
                                                            SelectionFacade.InsertCurrentUserReportMultipleSelections(10, globalUserID, "ZVR013", dRow2["Material"].ToString());
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    if (OrklaRTBPL.SelectionFacade.GetCurrentUserMultipleValueCount(10, globalUserID, "ZVR013").Equals(0))
                                    {
                                        MessageBox.Show("Produktgruppen inneholder ingen materialer.", "Orkla SAP Integration");
                                    }
                                }
                            }
                            //End MD04 Report


                            //For SalesOrder

                            if (globalReportID.Equals(24))
                            {
                                if (((ComboBox)sender).Tag.Equals("ZVS003"))
                                {
                                    OrklaRTBPL.SelectionFacade.SalesOrderSelectionSalesOrg = ((ComboBox)sender).SelectedValue.ToString();
                                }
                            }

                            //End SalesOrder

                            //For Shelf Life Report
                            if (globalReportID.Equals(39))
                            {
                                switch (((ComboBox)sender).Tag.ToString())
                                {
                                    case "ZVS013":
                                        OrklaRTBPL.SelectionFacade.ShelfLifeSelectionFirmakode = ((ComboBox)sender).SelectedValue.ToString();
                                        break;
                                    case "ZVU021":
                                        OrklaRTBPL.SelectionFacade.ShelfLifeSelectionPlant = ((ComboBox)sender).SelectedValue.ToString();
                                        break;
                                    case "ZVU005":
                                        OrklaRTBPL.SelectionFacade.ShelfLifeSelectionStorageLocation = ((ComboBox)sender).SelectedValue.ToString();
                                        break;
                                    case "ZVR038":
                                        OrklaRTBPL.SelectionFacade.ShelfLifeSelectionMaterialType = ((ComboBox)sender).SelectedValue.ToString();
                                        break;
                                }
                            }
                            // End Shelf Life Report

                            //For Stock History Report
                            if (globalReportID.Equals(34))
                            {
                                if (((ComboBox)sender).Tag.Equals("ZVS009"))
                                {
                                    OrklaRTBPL.SelectionFacade.StockHistorySelectionPlant = ((ComboBox)sender).SelectedValue.ToString();
                                }
                            }
                            // End Stock History Report

                            //For Stock Simulation Report
                            if (globalReportID.Equals(35))
                            {
                                if (((ComboBox)sender).Tag.Equals("ZVS009"))
                                {
                                    OrklaRTBPL.SelectionFacade.StockSimulationSelectionPlant = ((ComboBox)sender).SelectedValue.ToString();
                                    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(60, globalUserID, "ZVS009", "S", "I", "EQ", (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                                }
                            }
                            // End Stock Simulation Report

                            //For Stock Values and Coverage
                            if (globalReportID.Equals(38))
                            {
                                if (((ComboBox)sender).Tag.Equals("ZVS009"))
                                {
                                    OrklaRTBPL.SelectionFacade.StockValuesAndCoverageProdPlanSelectionPlant = ((ComboBox)sender).SelectedValue.ToString();
                                }
                            }
                            //End Stock Values and Coverage         

                            //For Daily Production Plan
                            if (globalReportID.Equals(62))
                            {
                                if (((ComboBox)sender).Tag.Equals("0P_PLANT"))
                                {
                                    OrklaRTBPL.SelectionFacade.DailyProductionPlanPlant = ((ComboBox)sender).SelectedValue.ToString();
                                }
                            }
                            //End Daily Production Plan
                        }
                    }
                    else if (e.RemovedItems.Count > 0 && ((ComboBox)sender).SelectedValue == null)
                    {
                        SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, ((ComboBox)sender).Tag.ToString(), "S", "I", "BT", (((ComboBox)sender).SelectedValue == null ? String.Empty : ((ComboBox)sender).SelectedValue.ToString()));
                    }
                }
                catch (Exception ex)
                {
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                }
            }
        }

        void datePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (!globalReportID.Equals(0) && !globalUserID.Equals(0))
            {
                try
                {
                    using (var entities = new DAL.SAPExlEntities())
                    {
                        var row = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((DatePicker)sender).Tag).SingleOrDefault();
                        if (((DatePicker)sender).Name.Equals("LowValue"))
                        {
                            fromDate = ((DatePicker)sender).SelectedDate.Value;
                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, ((DatePicker)sender).Tag.ToString(), "S", "I", row.SelectionOption.ToString(), ((DatePicker)sender).SelectedDate.Value.ToString("yyyyMMdd"));
                        }
                        else
                        {
                            toDate = ((DatePicker)sender).SelectedDate.Value;
                            SelectionFacade.UpdateCurrentUserReportSelectionHighValue(globalReportID, globalUserID, ((DatePicker)sender).Tag.ToString(), "S", "I", row.SelectionOption.ToString(), ((DatePicker)sender).SelectedDate.Value.ToString("yyyyMMdd"));
                        }
                        if (!toDate.Equals(OrklaRTBPL.CommonFacade.GetDateTime()) && !fromDate.Equals(OrklaRTBPL.CommonFacade.GetDateTime()))
                        {
                            if (toDate.Subtract(fromDate).Days < row.MaxDateLimit && row.MaxDateLimit > 0)
                            {
                                var row1 = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.Required == true).SingleOrDefault();
                                foreach (Control control in grdSelection.Children)
                                {
                                    if (!(control is Label) && !(control is Button) && !(row != null))
                                    {
                                        if (control.Tag.ToString().Equals(row1.ScreenID))
                                        {
                                            lblWarning.Foreground = Brushes.Black;
                                            lblWarning.Content = String.Empty;
                                            control.BorderBrush = Brushes.Transparent;
                                            control.ToolTip = null;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                if (!row.MaxDateLimit.Equals(0))
                                {
                                    string requireFields = String.Empty;
                                    string requireFieldNames = String.Empty;
                                    var requiredFields = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.Required == true);
                                    foreach (var field in requiredFields)
                                    {
                                        requireFields += field.ScreenID + ",";
                                        requireFieldNames += field.FieldName + ",";
                                    }
                                    foreach (Control control in grdSelection.Children)
                                    {
                                        if (!(control is Label) && !(control is Button))
                                        {
                                            if (requireFields.Contains(control.Tag.ToString()))
                                            {
                                                control.BorderBrush = Brushes.Red;
                                                control.ToolTip = "Obligatorisk";
                                            }
                                        }
                                    }
                                    lblWarning.Foreground = Brushes.Red;
                                    lblWarning.Content = "Datointervall overgår maksimums dato grensen på " + row.MaxDateLimit + " arbeidsdager, " + Environment.NewLine +
                                                         "Fyll inn gyldig dato eller Obligatorisk felt " + requireFieldNames.TrimEnd(',');
                                }
                            }
                        }


                        //For Production Plan Report

                        if (globalReportID.Equals(7) || globalReportID.Equals(63))
                        {
                            //Update Capacity Planning
                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, "ZVR0062", "S", "I", "BT", ((DatePicker)sender).SelectedDate.Value.AddDays(-25).ToString("yyyyMMdd"));
                            SelectionFacade.UpdateCurrentUserReportSelectionHighValue(globalReportID, globalUserID, "ZVR0062", "S", "I", "BT", ((DatePicker)sender).SelectedDate.Value.AddDays(+10).ToString("yyyyMMdd"));
                            OrklaRTBPL.SelectionFacade.ProductionPlanSelectionDate = ((DatePicker)sender).SelectedDate.Value.ToShortDateString();
                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(52, globalUserID, "ZVR021", "S", "I", "BT", ((DatePicker)sender).SelectedDate.Value.ToString("yyyyMMdd"));
                        }

                        // End Production Plan report

                        //For Mixing Plan Report

                        if (globalReportID.Equals(8))
                        {
                            if (((DatePicker)sender).Tag.ToString().Equals("SP$00000"))
                            {
                                OrklaRTBPL.SelectionFacade.MixingPlanProdPlanSelectionDate = ((DatePicker)sender).SelectedDate.Value.ToShortDateString();
                            }
                        }

                        // End Mixing Plan report

                        //For Stock History Report
                        if (globalReportID.Equals(34))
                        {
                            OrklaRTBPL.SelectionFacade.StockHistorySelectionFromDate = ((DatePicker)sender).SelectedDate.Value.AddYears(-1).ToShortDateString();
                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, "0I_DAYIN", "S", "I", "BT", ((DatePicker)sender).SelectedDate.Value.AddYears(-1).ToString("yyyyMMdd"));
                        }
                        // End Stock History report

                        //For Stock Simuation Report

                        if (globalReportID.Equals(35))
                        {
                            OrklaRTBPL.SelectionFacade.StockSimulationSelectionFromDate = ((DatePicker)sender).SelectedDate.Value.AddYears(-1).ToShortDateString();
                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, "0I_DAYIN", "S", "I", "BT", ((DatePicker)sender).SelectedDate.Value.AddYears(-1).ToString("yyyyMMdd"));
                        }

                        // End Stock Simuation report                       

                        //For Slow Moving Goods

                        if (globalReportID.Equals(36))
                        {
                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, "ZVR025", "S", "I", "BT", DateTime.Now.AddDays(-1000).ToString("yyyyMMdd"));

                        }
                        // End Slow Moving Goods                       
                    }
                }
                catch (Exception ex)
                {
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                }
            }
        }

        void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!globalReportID.Equals(0) && !globalUserID.Equals(0))
            {
                try
                {
                    using (var entities = new DAL.SAPExlEntities())
                    {
                        var controls = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((TextBox)sender).Tag).SingleOrDefault();
                        if (controls.FieldName.Equals("Periode"))
                        {
                            if (((TextBox)sender).Text.Length.Equals(6))
                            {
                                if (CheckPeriods().Equals(false))
                                {
                                    ClearPeriods();
                                    lblWarning.Foreground = Brushes.Red;
                                    lblWarning.Content = "Date range exceeds maximum period limit of 731 calendar days, " + Environment.NewLine +
                                                         "Please enter valid date - Required atleast one of the fields Plant  " + Environment.NewLine +
                                                         "Material Group or Company Code";

                                    var row1 = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.Required == true);
                                    foreach (var item in row1)
                                    {
                                        foreach (Control control in grdSelection.Children)
                                        {
                                            if (!(control is Label) && !(control is Button))
                                            {
                                                if (control.Tag.ToString().StartsWith("SP$") && control.Tag.ToString().Equals(item.ScreenID))
                                                {
                                                    control.BorderBrush = Brushes.Red;
                                                    control.ToolTip = "Required atleast one";
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    if (((TextBox)sender).Name.Equals("LowValue"))
                                    {
                                        //if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((TextBox)sender).Tag.ToString()).SingleOrDefault().FieldName.Equals("Language Key"))
                                        //{
                                        //    OrklaRTBPL.CommonFacade.UpdateCurrentUserReportFields("LanguageKey", ((TextBox)sender).Text, globalUserID, globalReportID);
                                        //}
                                        SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, ((TextBox)sender).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), ((TextBox)sender).Text);
                                    }
                                    else
                                    {
                                        SelectionFacade.UpdateCurrentUserReportSelectionHighValue(globalReportID, globalUserID, ((TextBox)sender).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), ((TextBox)sender).Text);
                                    }
                                    //var row1 = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.Required == true);
                                    //foreach (var item in row1)
                                    //{
                                    //    foreach (Control control in grdSelection.Children)
                                    //    {

                                    //        if (!(control is Label) && !(control is Button))
                                    //        {
                                    //            if (control.Tag.ToString().StartsWith("SP$") && control.Tag.ToString().Equals(item.ScreenID))
                                    //            {
                                    //                lblWarning.Foreground = Brushes.Black;
                                    //                lblWarning.Content = String.Empty;
                                    //                control.BorderBrush = Brushes.Transparent;
                                    //                control.ToolTip = null;
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                }

                            }
                        }
                        else
                        {
                            if (!((TextBox)sender).Text.Equals(String.Empty))
                            {
                                if (((TextBox)sender).Name.Equals("LowValue"))
                                {
                                    if (((TextBox)sender).Tag.ToString().Equals("ZVS001") || ((TextBox)sender).Tag.ToString().Equals("SP$22222"))
                                    {
                                        OrklaRTBPL.SelectionFacade.ReportSelectionLanguage = ((TextBox)sender).Text;
                                    }
                                    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, ((TextBox)sender).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), ((TextBox)sender).Text);
                                }
                                else
                                {
                                    SelectionFacade.UpdateCurrentUserReportSelectionHighValue(globalReportID, globalUserID, ((TextBox)sender).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), ((TextBox)sender).Text);
                                }
                            }
                            //if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((TextBox)sender).Tag).SingleOrDefault().MultipleSelection.Equals(true))
                            //{                               
                            //    SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(globalReportID, globalUserID, ((TextBox)sender).Tag.ToString(), ((TextBox)sender).Text, false);
                            //    dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == ((TextBox)sender).Tag).ToList();
                            //}
                        }

                    }

                    //For Lot Size Report

                    if (globalReportID.Equals(15))
                    {
                        //if (((TextBox)sender).Tag.ToString().Equals("SP$11111"))
                        //{
                        //    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(57, globalUserID, "SP$00005", "S", "I", "EQ", (((TextBox)sender).Text.Length == 4 ? ("2" + ((TextBox)sender).Text.Substring(2).ToString()) : String.Empty));
                        //}
                        //else if (((TextBox)sender).Tag.ToString().Equals("SP$00011"))
                        //{
                        //    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, "SP$00008", "S", "I", "EQ", ((TextBox)sender).Text);
                        //    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(57, globalUserID, "SP$00016", "S", "I", "EQ", ((TextBox)sender).Text);
                        //    SelectionFacade.UpdateCurrentUserReportSelectionLowValue(57, globalUserID, "SP$00019", "S", "I", "EQ", ((TextBox)sender).Text);
                        //}
                    }

                    // End Lot Size report                   

                    //For Scrapping Overview Report

                    if (globalReportID.Equals(33))
                    {
                        if (((TextBox)sender).Tag.ToString().Equals("ZVS001"))
                        {
                            OrklaRTBPL.SelectionFacade.ScrappingOverviewSelectionLanguage = ((TextBox)sender).Text;
                        }
                    }

                    // End Scrapping Overview report

                    if (globalReportID.Equals(34))
                    {
                        if (((TextBox)sender).Tag.ToString().Equals("ZVS005"))
                        {
                            OrklaRTBPL.SelectionFacade.StockHistorySelectionMaterial = ((TextBox)sender).Text;
                        }
                    }


                    // For Stock Simulation Report

                    if (globalReportID.Equals(35))
                    {
                        if (((TextBox)sender).Tag.ToString().Equals("ZVS001"))
                        {
                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(60, globalUserID, "ZVS001", "S", "I", "EQ", ((TextBox)sender).Text);
                        }
                        else if (((TextBox)sender).Tag.ToString().Equals("ZVS005"))
                        {
                            OrklaRTBPL.SelectionFacade.StockSimulationSelectionMaterial = ((TextBox)sender).Text;
                            SelectionFacade.UpdateCurrentUserReportSelectionLowValue(60, globalUserID, "ZVR013", "S", "I", "EQ", ((TextBox)sender).Text);
                        }
                    }

                    // End Stock Simulation Report

                }
                catch (Exception ex)
                {
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                }
            }
        }

        void textBox_LostFocus(object sender, RoutedEventArgs e)
        {
            if (!globalReportID.Equals(0) && !globalUserID.Equals(0))
            {
                try
                {
                    using (var entities = new DAL.SAPExlEntities())
                    {
                        if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == ((TextBox)sender).Tag).SingleOrDefault().MultipleSelection.Equals(true))
                        {
                            SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(globalReportID, globalUserID, ((TextBox)sender).Tag.ToString(), ((TextBox)sender).Text, false);
                            dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == ((TextBox)sender).Tag).ToList();
                        }
                    }
                }
                catch (Exception ex)
                {
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                }
            }
        }

        public void LoadLabelText(int sortOrder, string screenId, string fieldName)
        {
            switch (sortOrder)
            {
                case 1:
                    lblName1.Content = fieldName;
                    lblName1.Tag = screenId;
                    break;
                case 2:
                    lblName2.Content = fieldName;
                    lblName2.Tag = screenId;
                    break;
                case 3:
                    lblName3.Content = fieldName;
                    lblName3.Tag = screenId;
                    break;
                case 4:
                    if (globalReportID.Equals(7) || globalReportID.Equals(63))
                    {
                        lblName4.Visibility = Visibility.Hidden;
                    }
                    lblName4.Content = fieldName;
                    lblName4.Tag = screenId;
                    break;
                case 5:
                    lblName5.Content = fieldName;
                    lblName5.Tag = screenId;
                    break;
                case 6:
                    lblName6.Content = fieldName;
                    lblName6.Tag = screenId;
                    break;
                case 7:
                    lblName7.Content = fieldName;
                    lblName7.Tag = screenId;
                    break;
                case 8:
                    lblName8.Content = fieldName;
                    lblName8.Tag = screenId;
                    break;
                case 9:
                    lblName9.Content = fieldName;
                    lblName9.Tag = screenId;
                    break;
            }
        }

        public void LoadUserSelectionData(int reportID, int userID, int variantID)
        {
            try
            {
                foreach (Control control in grdSelection.Children)
                {
                    if (!(control is Label) && !(control is Button))
                    {
                        if (control.Tag.ToString().StartsWith("ZV") || control.Tag.ToString().StartsWith("0P") || control.Tag.ToString().StartsWith("0I"))
                        {
                            using (var entities = new DAL.SAPExlEntities())
                            {
                                var userReportValues = entities.CurrentUserReportSelections.Where(urv => urv.ReportID == reportID && urv.UserID == userID && urv.VariantID == variantID && urv.ScreenID == control.Tag).SingleOrDefault();
                                if (control is ComboBox)
                                {
                                    if (userReportValues != null)
                                    {
                                        if (control.Name.Equals("LowValue"))
                                        {
                                            ((ComboBox)control).Text = userReportValues.LowValue;
                                        }
                                        else
                                        {
                                            ((ComboBox)control).Text = userReportValues.HighValue;
                                        }
                                    }
                                    else
                                    {
                                        ((ComboBox)control).Text = String.Empty;
                                    }
                                }
                                else if (control is TextBox)
                                {
                                    if (userReportValues != null)
                                    {
                                        if (control.Name.Equals("LowValue"))
                                        {
                                            ((TextBox)control).Text = userReportValues.LowValue;
                                            if (rightClick.Equals(true))
                                            {
                                                try
                                                {
                                                    if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == userReportValues.ScreenID).SingleOrDefault().MultipleSelection.Equals(true) && userReportValues.LowValue != null)
                                                    {
                                                        SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(globalReportID, globalUserID, userReportValues.ScreenID, userReportValues.LowValue, false);
                                                        dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == userReportValues.ScreenID).ToList();
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                                                }

                                            }
                                        }
                                        else
                                        {
                                            ((TextBox)control).Text = userReportValues.HighValue;
                                        }
                                    }
                                    else
                                    {
                                        ((TextBox)control).Text = String.Empty;
                                    }
                                }
                                else if (control is DatePicker)
                                {
                                    LoadReportDefaultValues();
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }
        public DateTime AddBusinessDays(DateTime date, int days)
        {
            if (days == 0) return date;

            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = date.AddDays(2);
                days -= 1;
            }
            else if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = date.AddDays(1);
                days -= 1;
            }

            date = date.AddDays(days / 5 * 7);
            int extraDays = days % 5;

            if ((int)date.DayOfWeek + extraDays > 5)
            {
                extraDays += 2;
            }

            return date.AddDays(extraDays);

        }

        private void txtWeek_TextChanged(object sender, TextChangedEventArgs e)
        {
            try
            {
                var week = Enumerable.Range(0, 7).Select(d => FirstDateOfWeek(Convert.ToInt32(txtYear.Text), Convert.ToInt32(txtWeek.Text)).AddDays(d)).ToList();
                foreach (Control control in grdSelection.Children)
                {
                    if (control is DatePicker)
                    {
                        try
                        {
                            if (control.Name.Equals("LowValue"))
                            {
                                if (globalReportID.Equals(14))
                                {
                                    if (control.Tag.ToString().Equals("ZVR0062"))
                                    {
                                        ((DatePicker)control).Text = week[0].ToString();
                                    }
                                }
                                else
                                {
                                    ((DatePicker)control).Text = week[0].ToString();
                                }
                            }
                            else
                            {
                                if (globalReportID.Equals(14))
                                {
                                    if (control.Tag.ToString().Equals("ZVR0062"))
                                    {
                                        ((DatePicker)control).Text = week[6].ToString();
                                    }
                                }
                                else
                                {
                                    ((DatePicker)control).Text = week[6].ToString();
                                }
                            }
                        }
                        catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }
        public DateTime FirstDateOfWeek(int year, int weekOfYear)
        {
            DateTime jan1 = new DateTime(year, 1, 1);
            int daysOffset = Convert.ToInt32(System.Globalization.CultureInfo.CurrentCulture.DateTimeFormat.FirstDayOfWeek) - Convert.ToInt32(jan1.DayOfWeek);
            DateTime firstWeekDay = jan1.AddDays(daysOffset);
            System.Globalization.CultureInfo curCulture = System.Globalization.CultureInfo.CurrentCulture;
            int firstWeek = curCulture.Calendar.GetWeekOfYear(jan1, curCulture.DateTimeFormat.CalendarWeekRule, curCulture.DateTimeFormat.FirstDayOfWeek);
            if (firstWeek <= 1)
            {
                weekOfYear -= 1;
            }
            return firstWeekDay.AddDays(weekOfYear * 7);
        }

        private void txtMonth_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (!txtMonth.Text.Equals(String.Empty))
            {
                foreach (Control control in grdSelection.Children)
                {
                    if (control is DatePicker)
                    {
                        try
                        {
                            if (control.Name.Equals("LowValue"))
                            {
                                ((DatePicker)control).Text = new DateTime(DateTime.Now.Year, Convert.ToInt32(txtMonth.Text), 1).ToShortDateString();
                            }
                            else
                            {
                                ((DatePicker)control).Text = new DateTime(DateTime.Now.Year, Convert.ToInt32(txtMonth.Text), 1).AddMonths(1).AddDays(-1).ToShortDateString();
                            }
                        }
                        catch (Exception ex)
                        {
                            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                        }
                    }
                    else if (control is TextBox)
                    {
                        try
                        {
                            using (var entities = new DAL.SAPExlEntities())
                            {
                                var controls = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault();
                                if (controls.FieldName.Equals("Periode"))
                                {
                                    if (!txtMonth.Text.Equals("0") && !txtMonth.Text.Equals("1") && txtMonth.Text.Length.Equals(1) && Convert.ToInt32(txtMonth.Text) < 10)
                                    {
                                        txtMonth.Text = txtMonth.Text.PadLeft(2, '0');
                                    }

                                    if (!txtYear.Text.Equals(String.Empty))
                                    {
                                        if (txtYear.Text.Length.Equals(6))
                                        {
                                            if (txtMonth.Text.Equals("1"))
                                            {
                                                ((TextBox)control).Text = ((TextBox)control).Text.Remove(((TextBox)control).Text.Length - 2, 2) + "0" + txtMonth.Text;
                                            }
                                            else
                                            {
                                                ((TextBox)control).Text = ((TextBox)control).Text.Remove(((TextBox)control).Text.Length - 2, 2) + txtMonth.Text;
                                            }
                                        }
                                        else
                                        {
                                            if (txtMonth.Text.Equals("1"))
                                            {
                                                ((TextBox)control).Text = txtYear.Text + "0" + txtMonth.Text;
                                            }
                                            else
                                            {
                                                ((TextBox)control).Text = txtYear.Text + txtMonth.Text;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (txtMonth.Text.Equals("1"))
                                        {
                                            ((TextBox)control).Text = DateTime.Now.Year.ToString() + "0" + txtMonth.Text;
                                            txtYear.Text = DateTime.Now.Year.ToString();
                                        }
                                        else
                                        {
                                            ((TextBox)control).Text = DateTime.Now.Year.ToString() + txtMonth.Text;
                                            txtYear.Text = DateTime.Now.Year.ToString();
                                        }

                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                        }
                    }
                }
            }
        }

        private void txtYear_TextChanged(object sender, TextChangedEventArgs e)
        {
            foreach (Control control in grdSelection.Children)
            {
                if (control is DatePicker)
                {
                    try
                    {
                        if (txtYear.Text.Length.Equals(4))
                        {
                            if (control.Name.Equals("LowValue"))
                            {
                                if (!txtMonth.Text.Equals(String.Empty))
                                {
                                    ((DatePicker)control).Text = new DateTime(Convert.ToInt32(txtYear.Text), Convert.ToInt32(txtMonth.Text), 1).ToShortDateString();
                                }
                                else
                                {
                                    ((DatePicker)control).Text = new DateTime(Convert.ToInt32(txtYear.Text), 1, 1).ToShortDateString();
                                }
                            }
                            else
                            {
                                if (!txtMonth.Text.Equals(String.Empty))
                                {
                                    ((DatePicker)control).Text = new DateTime(Convert.ToInt32(txtYear.Text), Convert.ToInt32(txtMonth.Text), 31).ToShortDateString();
                                }
                                else
                                {
                                    ((DatePicker)control).Text = new DateTime(Convert.ToInt32(txtYear.Text), 12, 31).ToShortDateString();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                    }
                }
                else if (control is TextBox)
                {
                    try
                    {
                        if (txtYear.Text.Length.Equals(4))
                        {
                            using (var entities = new DAL.SAPExlEntities())
                            {
                                var controls = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault();
                                if (controls.FieldName.Equals("Periode"))
                                {
                                    if (control.Name.Equals("LowValue"))
                                    {
                                        if (!txtMonth.Text.Equals(String.Empty) && txtMonth.Text.Length.Equals(2))
                                        {
                                            ((TextBox)control).Text = txtYear.Text + txtMonth.Text;
                                        }
                                        else
                                        {
                                            ((TextBox)control).Text = txtYear.Text + "01";
                                        }
                                    }
                                    else
                                    {
                                        if (!txtMonth.Text.Equals(String.Empty) && txtMonth.Text.Length.Equals(2))
                                        {
                                            ((TextBox)control).Text = txtYear.Text + txtMonth.Text;
                                        }
                                        else
                                        {
                                            ((TextBox)control).Text = txtYear.Text + "12";
                                        }
                                    }

                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                    }
                }
            }
        }
        public bool CheckPeriods()
        {
            bool ret = false;
            DateTime fromDate = new DateTime();
            DateTime toDate = new DateTime();
            foreach (Control control in grdSelection.Children)
            {
                if (control is TextBox)
                {
                    try
                    {
                        using (var entities = new DAL.SAPExlEntities())
                        {
                            var controls = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault();
                            if (controls.FieldName.Equals("Period"))
                            {
                                if (control.Name.Equals("LowValue"))
                                {
                                    fromDate = new DateTime(Convert.ToInt32(((TextBox)control).Text.Substring(0, 4)), Convert.ToInt32(((TextBox)control).Text.Substring(4)), 1);
                                }
                                else
                                {
                                    toDate = new DateTime(Convert.ToInt32(((TextBox)control).Text.Substring(0, 4)), Convert.ToInt32(((TextBox)control).Text.Substring(4)), 1);
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                    }
                }
            }
            if ((((toDate.Year - fromDate.Year) * 12) + toDate.Month - fromDate.Month) > 24)
            {
                ret = false;
            }
            else
            {
                ret = true;
            }
            return ret;
        }
        public void ClearPeriods()
        {
            foreach (Control control in grdSelection.Children)
            {
                if (control is TextBox)
                {
                    try
                    {
                        using (var entities = new DAL.SAPExlEntities())
                        {
                            var controls = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault();
                            if (controls.FieldName.Equals("Periode"))
                            {
                                txtMonth.Text = String.Empty;
                                txtYear.Text = String.Empty;
                                ((TextBox)control).Text = String.Empty;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                    }
                }
            }
        }
        public DataTable GetPeriods(bool fromPeriod, string period)
        {
            DataTable periodList = new DataTable();
            periodList.TableName = "Periods";
            periodList.Columns.Add("ID", typeof(string));
            periodList.Columns.Add("Text", typeof(string));

            try
            {
                DateTime date = new DateTime(Convert.ToInt32(period.Substring(0, 4)), Convert.ToInt32(period.Substring(4)), 1);

                if (fromPeriod.Equals(false))
                {
                    DateTime fromDate = date.Date.AddYears(-2);
                    while (date.Date > fromDate.Date)
                    {
                        periodList.Rows.Add(date.Year.ToString() + (date.Month.ToString().Length.Equals(2) ? date.Month.ToString() : "0" + date.Month.ToString()), date.Year.ToString() + (date.Month.ToString().Length.Equals(2) ? date.Month.ToString() : "0" + date.Month.ToString()));
                        date = date.AddMonths(-1);
                    }
                }
                else
                {
                    DateTime ToDate = date.Date.AddYears(2);
                    while (date.Date < ToDate.Date)
                    {
                        periodList.Rows.Add(date.Year.ToString() + (date.Month.ToString().Length.Equals(2) ? date.Month.ToString() : "0" + date.Month.ToString()), date.Year.ToString() + (date.Month.ToString().Length.Equals(2) ? date.Month.ToString() : "0" + date.Month.ToString()));
                        date = date.AddMonths(1);
                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }

            return periodList;
        }

        private void btnRun_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                System.Windows.Forms.Application.DoEvents();
                System.Windows.Forms.Application.EnableVisualStyles();
                btnRun.Dispatcher.Invoke(DispatcherPriority.Background, new Action(() => { System.Windows.Forms.SendKeys.Send("{F8}"); }));
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        private void btnDeleteSelection_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!cboSelectionVariant.SelectedValue.ToString().Equals("0"))
                {
                    SelectionFacade.DeleteUserReportVariant(globalReportID, globalUserID, Convert.ToInt32(cboSelectionVariant.SelectedValue));
                    cboSelectionVariant.ItemsSource = SelectionFacade.GetUserReportVariants(globalUserID, globalReportID).Tables[0].DefaultView;
                    cboSelectionVariant.SelectedValue = 0;
                }

            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        private void expVariants_Expanded(object sender, RoutedEventArgs e)
        {
            txtVariantName.Text = cboSelectionVariant.Text;
            if (cboSelectionVariant.SelectedValue.ToString().Equals("0"))
            {
                btnOK.Visibility = Visibility.Hidden;
                btnCancel.Visibility = Visibility.Hidden;
                txtVariantName.IsEnabled = false;
                txtDescription.IsEnabled = false;
            }
            else
            {
                btnOK.Visibility = Visibility.Visible;
                btnCancel.Visibility = Visibility.Visible;
                txtVariantName.IsEnabled = true;
                txtDescription.IsEnabled = true;
            }
            //btnSaveSelection.Margin = new Thickness(btnSaveSelection.Margin.Left, 130, btnSaveSelection.Margin.Right, btnSaveSelection.Margin.Bottom);
            lblMessage.Margin = new Thickness(lblMessage.Margin.Left, 125, lblMessage.Margin.Right, lblMessage.Margin.Bottom);
            btnYes.Margin = new Thickness(btnYes.Margin.Left, 155, btnYes.Margin.Right, btnYes.Margin.Bottom);
            btnNew.Margin = new Thickness(btnNew.Margin.Left, 155, btnNew.Margin.Right, btnNew.Margin.Bottom);
            btnCancel1.Margin = new Thickness(btnCancel1.Margin.Left, 155, btnCancel1.Margin.Right, btnCancel1.Margin.Bottom);
            grpEasyDateSelection.Margin = new Thickness(grpEasyDateSelection.Margin.Left, 185, grpEasyDateSelection.Margin.Right, grpEasyDateSelection.Margin.Bottom);
            grpSelection.Margin = new Thickness(grpSelection.Margin.Left, 275, grpSelection.Margin.Right, grpSelection.Margin.Bottom);
        }

        private void expVariants_Collapsed(object sender, RoutedEventArgs e)
        {
            //btnSaveSelection.Margin = new Thickness(btnSaveSelection.Margin.Left,30, btnSaveSelection.Margin.Right, btnSaveSelection.Margin.Bottom);
            lblMessage.Margin = new Thickness(lblMessage.Margin.Left, 25, lblMessage.Margin.Right, lblMessage.Margin.Bottom);
            btnYes.Margin = new Thickness(btnYes.Margin.Left, 50, btnYes.Margin.Right, btnYes.Margin.Bottom);
            btnNew.Margin = new Thickness(btnNew.Margin.Left, 50, btnNew.Margin.Right, btnNew.Margin.Bottom);
            btnCancel1.Margin = new Thickness(btnCancel1.Margin.Left, 50, btnCancel1.Margin.Right, btnCancel1.Margin.Bottom);
            grpEasyDateSelection.Margin = new Thickness(grpEasyDateSelection.Margin.Left, 80, grpEasyDateSelection.Margin.Right, grpEasyDateSelection.Margin.Bottom);
            grpSelection.Margin = new Thickness(grpSelection.Margin.Left, 170, grpSelection.Margin.Right, grpSelection.Margin.Bottom);
        }


        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            expVariants.IsExpanded = false;
            btnSaveSelection.Visibility = Visibility.Visible;
            lblMessage.Visibility = Visibility.Hidden;
            btnYes.Visibility = Visibility.Hidden;
            btnNew.Visibility = Visibility.Hidden;
            btnCancel1.Visibility = Visibility.Hidden;
        }

        private void cboSelectionVariant_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (e.AddedItems.Count > 0)
                {
                    globalVariantID = Convert.ToInt32(((DataRowView)e.AddedItems[0]).Row.ItemArray[0].ToString());
                    if (!globalReportID.Equals(0) && !globalUserID.Equals(0))
                    {
                        SelectionFacade.UpdateCurrentUserReportVariants(globalReportID, globalUserID, globalVariantID);
                    }
                    txtWeek.Text = String.Empty;
                    txtMonth.Text = String.Empty;
                    txtYear.Text = String.Empty;
                    if (!globalVariantID.Equals(0))
                    {
                        DataSet variantsDataSet = SelectionFacade.GetUserReportVariants(globalVariantID);
                        txtVariantName.Text = variantsDataSet.Tables[0].Rows[0]["VariantName"].ToString();
                        txtDescription.Text = variantsDataSet.Tables[0].Rows[0]["VariantDescription"].ToString();
                        SelectionFacade.DeleteCurrentUserReportSelections(globalReportID, globalUserID);
                        SelectionFacade.InsertCurrentUserReportSelections(globalReportID, globalUserID, globalVariantID);
                        LoadUserSelectionData(globalReportID, globalUserID, 0);
                        foreach (Control control in grdSelection.Children)
                        {
                            if (control is Button)
                            {
                                if (SelectionFacade.GetCurrentUserMultipleSelectionValue(globalReportID, globalUserID, ((Button)control).Tag.ToString()).Equals(true))
                                {
                                    ((Button)control).Foreground = Brushes.DarkGreen;
                                    ((Button)control).BorderBrush = Brushes.DarkGreen;
                                }
                                else
                                {
                                    ((Button)control).Foreground = Brushes.Black;
                                    ((Button)control).BorderBrush = Brushes.Black;
                                }
                            }
                        }
                    }
                    else
                    {
                        txtVariantName.Text = "Standard";
                        txtDescription.Text = String.Empty;
                        SelectionFacade.DeleteCurrentUserReportSelections(globalReportID, globalUserID);
                        foreach (Control control in grdSelection.Children)
                        {
                            if (control is Button)
                            {
                                ((Button)control).Foreground = Brushes.Black;
                                ((Button)control).BorderBrush = Brushes.Black;
                            }
                        }
                        LoadUserSelectionData(globalReportID, globalUserID, 0);
                        LoadReportDefaultValues();
                        SelectionFacade.DeleteEmptyCurrentUserReportSelections(globalReportID, globalUserID);
                    }

                    if (!((DataRowView)e.AddedItems[0]).Row.ItemArray[0].ToString().Equals("0"))
                    {
                        txtVariantName.IsEnabled = true;
                        txtDescription.IsEnabled = true;
                        btnOK.Visibility = Visibility.Visible;
                        btnCancel.Visibility = Visibility.Visible;
                        //btnSaveSelection.Visibility = Visibility.Visible;
                        btnDeleteSelection.Visibility = Visibility.Visible;
                    }
                    else
                    {
                        txtVariantName.IsEnabled = false;
                        txtDescription.IsEnabled = false;
                        btnOK.Visibility = Visibility.Hidden;
                        btnCancel.Visibility = Visibility.Hidden;
                        //btnSaveSelection.Visibility = Visibility.Hidden;
                        btnDeleteSelection.Visibility = Visibility.Hidden;
                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }
        private void btnPaste_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                List<string[]> rowData = ClipboardHelper.ParseClipboardData();
                if (rowData != null && rowData.Count > 1)
                {
                    SelectionFacade.DeleteCurrentUserMultipleSelections(globalReportID, globalUserID, lblScreenID.Content.ToString());
                    foreach (Control control in grdSelection.Children)
                    {
                        if (control is ComboBox)
                        {
                            if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                            {
                                ((ComboBox)control).SelectedValue = rowData[0].GetValue(0).ToString();
                            }
                        }
                        if (control is Button)
                        {
                            if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                            {
                                ((Button)control).Foreground = Brushes.DarkGreen;
                                ((Button)control).BorderBrush = Brushes.DarkGreen;
                            }
                        }
                    }
                    foreach (var item in rowData)
                    {
                        SelectionFacade.InsertCurrentUserReportMultipleSelections(globalReportID, globalUserID, lblScreenID.Content.ToString(), item[0]);
                        //SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(globalReportID, globalUserID, lblScreenID.Content.ToString(), item[0]);
                    }
                    foreach (Control control in grdSelection.Children)
                    {
                        if (control is ComboBox)
                        {
                            if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                            {
                                ((ComboBox)control).SelectedValue = rowData[0].GetValue(0).ToString();
                            }
                        }
                        if (control is Button)
                        {
                            if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                            {
                                ((Button)control).Foreground = Brushes.DarkGreen;
                                ((Button)control).BorderBrush = Brushes.DarkGreen;
                            }
                        }
                    }
                    using (var entities = new DAL.SAPExlEntities())
                    {
                        SelectionFacade.UpdateCurrentUserReportSelectionMultiplSelected(globalReportID, globalUserID, lblScreenID.Content.ToString(), true);
                        dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == lblScreenID.Content).ToList();

                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        private void dgSelectSingleValues_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            if (!globalReportID.Equals(0) && !globalUserID.Equals(0))
            {
                try
                {
                    using (var entities = new DAL.SAPExlEntities())
                    {
                        if (((DAL.CurrentUserReportMultipleSelections)e.Row.Item).ID.Equals(0))
                        {
                            if (e.Column.ToString().Equals("System.Windows.Controls.DataGridTextColumn"))
                            {
                                SelectionFacade.InsertCurrentUserReportMultipleSelections(globalReportID, globalUserID, lblScreenID.Content.ToString(), ((TextBox)e.EditingElement).Text);
                            }
                            else
                            {
                                SelectionFacade.InsertCurrentUserReportMultipleSelections(globalReportID, globalUserID, lblScreenID.Content.ToString(), ((ComboBox)e.EditingElement).SelectedValue.ToString());
                            }
                            //SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(globalReportID, globalUserID, lblScreenID.Content.ToString(), ((ComboBox)e.EditingElement).SelectedValue.ToString());
                        }
                        else
                        {
                            if (e.Column.ToString().Equals("System.Windows.Controls.DataGridTextColumn"))
                            {
                                SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(((DAL.CurrentUserReportMultipleSelections)e.Row.Item).ID, ((TextBox)e.EditingElement).Text);
                            }
                            else
                            {
                                SelectionFacade.UpdateCurrentUserReportMultipleSelectionValue(((DAL.CurrentUserReportMultipleSelections)e.Row.Item).ID, ((ComboBox)e.EditingElement).SelectedValue.ToString());
                            }
                        }
                        foreach (Control control in grdSelection.Children)
                        {
                            if (control is ComboBox)
                            {
                                if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                                {
                                    ((ComboBox)control).Text = SelectionFacade.GetCurrentUserFirstMultipleValue(globalReportID, globalUserID, lblScreenID.Content.ToString());
                                }
                            }
                            if (control is TextBox)
                            {
                                if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                                {
                                    ((TextBox)control).Text = SelectionFacade.GetCurrentUserFirstMultipleValue(globalReportID, globalUserID, lblScreenID.Content.ToString());
                                }
                            }
                            if (control is Button)
                            {
                                if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                                {
                                    ((Button)control).Foreground = Brushes.DarkGreen;
                                    ((Button)control).BorderBrush = Brushes.DarkGreen;
                                }
                            }
                        }
                        dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == lblScreenID.Content).ToList();
                        SelectionFacade.UpdateCurrentUserReportSelectionMultiplSelected(globalReportID, globalUserID, lblScreenID.Content.ToString(), true);
                    }

                }
                catch (Exception ex)
                {
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                }

            }
        }
        private void dgSelectSingleValues_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Delete)
            {
                try
                {
                    SelectionFacade.DeleteCurrentUserMultipleSelections(((DAL.CurrentUserReportMultipleSelections)dgSelectSingleValues.CurrentItem).ID);
                    using (var entities = new DAL.SAPExlEntities())
                    {
                        dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == lblScreenID.Content).ToList();
                    }
                }
                catch (Exception ex)
                {
                    OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
                }
            }

        }
        private void btnSaveSelection_Click(object sender, RoutedEventArgs e)
        {
            lblMessage.Content = String.Empty;
            btnSaveSelection.Visibility = Visibility.Hidden;
            lblMessage.Visibility = Visibility.Visible;

            if (!cboSelectionVariant.SelectedValue.ToString().Equals("0"))
            {
                lblMessage.Content = "Er du sikker på at du vil overskrive valgt variant -  " + cboSelectionVariant.Text;
                btnNew.Visibility = Visibility.Visible;
                btnCancel1.Visibility = Visibility.Visible;
                btnYes.Visibility = Visibility.Visible;
            }
            else
            {
                expVariants.IsExpanded = true;
                lblMessage.Visibility = Visibility.Hidden;
                btnYes.Visibility = Visibility.Hidden;
                btnNew.Visibility = Visibility.Hidden;
                btnCancel1.Visibility = Visibility.Hidden;
                txtVariantName.IsEnabled = true;
                txtDescription.IsEnabled = true;
                txtVariantName.Text = String.Empty;
                txtDescription.Text = String.Empty;
                btnOK.Visibility = Visibility.Visible;
                btnCancel.Visibility = Visibility.Visible;
            }
        }

        private void btnDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SelectionFacade.DeleteCurrentUserMultipleSelections(globalReportID, globalUserID, lblScreenID.Content.ToString());
                using (var entities = new DAL.SAPExlEntities())
                {
                    dgSelectSingleValues.ItemsSource = entities.CurrentUserReportMultipleSelections.Where(urms => urms.ReportID == globalReportID && urms.UserID == globalUserID && urms.ScreenID == lblScreenID.Content).ToList();
                }
                foreach (Control control in grdSelection.Children)
                {
                    if (control is ComboBox)
                    {
                        if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                        {
                            ((ComboBox)control).SelectedValue = null;
                        }
                    }
                    if (control is Button)
                    {
                        if (control.Tag.ToString().Equals(lblScreenID.Content.ToString()))
                        {
                            ((Button)control).Foreground = Brushes.Black;
                            ((Button)control).BorderBrush = Brushes.Black;
                        }
                    }
                }
                //SelectionFacade.UpdateCurrentUserReportSelectionMultiplSelected(globalReportID, globalUserID, lblScreenID.Content.ToString(), false);
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!txtVariantName.Text.Equals(String.Empty))
                {
                    newVariant = true;
                    SelectionFacade.InsertUserReportVariant(globalReportID, globalUserID, txtVariantName.Text, txtDescription.Text);
                    globalVariantID = SelectionFacade.GetLastVariantID(globalUserID, globalReportID);
                    btnSaveSelection.Visibility = Visibility.Visible;
                    btnDeleteSelection.Visibility = Visibility.Visible;
                    expVariants.IsExpanded = false;
                    SelectionFacade.InsertUserReportSelections(globalReportID, globalUserID, SelectionFacade.GetLastVariantID(globalUserID, globalReportID));
                    using (var entities = new DAL.SAPExlEntities())
                    {
                        if (entities.ReportsLinkedQuery.Where(rlq => rlq.ReportID == globalReportID).Count() > 0)
                        {
                            foreach (var row in entities.ReportsLinkedQuery.Where(rlq => rlq.ReportID == globalReportID))
                            {
                                SelectionFacade.InsertUserReportSelections(row.SubReports, globalUserID, SelectionFacade.GetLastVariantID(globalUserID, globalReportID));
                            }
                        }
                    }
                    cboSelectionVariant.ItemsSource = SelectionFacade.GetUserReportVariants(globalUserID, globalReportID).Tables[0].DefaultView;
                    cboSelectionVariant.SelectedValue = SelectionFacade.GetLastVariantID(globalUserID, globalReportID);
                    //SelectionFacade.DeleteCurrentUserReportSelections(globalReportID, globalUserID);
                }
                else
                {
                    MessageBox.Show("Vennligst fyll ut Variant Navn !");
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        private void btnYes_Click(object sender, RoutedEventArgs e)
        {
            btnSaveSelection.Visibility = Visibility.Visible;
            btnYes.Visibility = Visibility.Hidden;
            btnNew.Visibility = Visibility.Hidden;
            btnCancel1.Visibility = Visibility.Hidden;
            lblMessage.Content = String.Empty;
            lblMessage.Visibility = Visibility.Hidden;
            SelectionFacade.InsertUserReportSelections(globalReportID, globalUserID, globalVariantID);
            //SelectionFacade.DeleteCurrentUserReportSelections(globalReportID, globalUserID);
        }

        private void btnNew_Click(object sender, RoutedEventArgs e)
        {
            expVariants.IsExpanded = true;
            lblMessage.Visibility = Visibility.Hidden;
            btnYes.Visibility = Visibility.Hidden;
            btnNew.Visibility = Visibility.Hidden;
            btnCancel1.Visibility = Visibility.Hidden;
            txtVariantName.IsEnabled = true;
            txtDescription.IsEnabled = true;
            txtVariantName.Text = String.Empty;
            txtDescription.Text = String.Empty;
            btnOK.Visibility = Visibility.Visible;
            btnCancel.Visibility = Visibility.Visible;
        }

        private void btnCancel1_Click(object sender, RoutedEventArgs e)
        {
            expVariants.IsExpanded = false;
            btnSaveSelection.Visibility = Visibility.Visible;
            lblMessage.Visibility = Visibility.Hidden;
            btnYes.Visibility = Visibility.Hidden;
            btnNew.Visibility = Visibility.Hidden;
            btnCancel1.Visibility = Visibility.Hidden;
        }
        public void UpdateSelections()
        {
            //foreach (Control control in grdSelection.Children)
            //{
            //    if (!(control is Label) && !(control is Button))
            //    {
            //        if (control.Tag.ToString().StartsWith("SP$"))
            //        {
            //            using (var entities = new DAL.SAPExlEntities())
            //            {
            //                var controls = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault();
            //                {
            //                    if (control is ComboBox)
            //                    {
            //                        if (((ComboBox)control).Name.Equals("LowValue"))
            //                        {

            //                            SelectionFacade.UpdateUserReportSelectionLowValue(globalReportID, globalUserID, globalVariantID ,((ComboBox)control).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), (((ComboBox)control).SelectedValue == null ? String.Empty : ((ComboBox)control).SelectedValue.ToString()));
            //                        }
            //                        else
            //                        {
            //                            SelectionFacade.UpdateUserReportSelectionHighValue(globalReportID, globalUserID, globalVariantID, ((ComboBox)control).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), (((ComboBox)control).SelectedValue == null ? String.Empty : ((ComboBox)control).SelectedValue.ToString()));
            //                        }
            //                    }
            //                    else if (control is TextBox)
            //                    {
            //                        if (!controls.FieldName.Equals("Period"))
            //                        {
            //                            if (((TextBox)control).Name.Equals("LowValue"))
            //                            {
            //                                SelectionFacade.UpdateUserReportSelectionLowValue(globalReportID, globalUserID, globalVariantID,((TextBox)control).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), (((TextBox)control).Text == null ? String.Empty : ((TextBox)control).Text));
            //                            }
            //                            else

            //                                SelectionFacade.UpdateUserReportSelectionHighValue(globalReportID, globalUserID,  ((TextBox)control).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), (((TextBox)control).Text == null ? String.Empty : ((TextBox)control).Text));
            //                            }
            //                        }
            //                    }
            //                    else if (control is DatePicker)
            //                    {
            //                        if (((DatePicker)control).Name.Equals("LowValue"))
            //                        {

            //                            SelectionFacade.UpdateUserReportSelectionLowValue(globalReportID, globalUserID, globalVariantID, ((DatePicker)control).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), (((DatePicker)control).SelectedDate.Value == null ? String.Empty : ((DatePicker)control).SelectedDate.Value.Year.ToString() +
            //                                (((DatePicker)control).SelectedDate.Value.Month.ToString().Length == 1 ? "0" + ((DatePicker)control).SelectedDate.Value.Month.ToString() : ((DatePicker)control).SelectedDate.Value.Month.ToString()) +
            //                                (((DatePicker)control).SelectedDate.Value.Day.ToString().Length == 1 ? "0" + ((DatePicker)control).SelectedDate.Value.Day.ToString() : ((DatePicker)control).SelectedDate.Value.Day.ToString())));
            //                        }
            //                        else
            //                        {
            //                            SelectionFacade.UpdateUserReportSelectionHighValue(globalReportID, globalUserID, globalVariantID, ((DatePicker)control).Tag.ToString(), "S", "I", controls.SelectionOption.ToString(), (((DatePicker)control).SelectedDate.Value == null ? String.Empty : ((DatePicker)control).SelectedDate.Value.Year.ToString() +
            //                                (((DatePicker)control).SelectedDate.Value.Month.ToString().Length == 1 ? "0" + ((DatePicker)control).SelectedDate.Value.Month.ToString() : ((DatePicker)control).SelectedDate.Value.Month.ToString()) +
            //                                (((DatePicker)control).SelectedDate.Value.Day.ToString().Length == 1 ? "0" + ((DatePicker)control).SelectedDate.Value.Day.ToString() : ((DatePicker)control).SelectedDate.Value.Day.ToString())));
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }
            //}
            //SelectionFacade.DeleteEmptyCurrentUserReportSelections(globalReportID, globalUserID);
            //if (newVariant.Equals(true))
            //{
            //    SelectionFacade.InsertMultipleSelections(SelectionFacade.GetLastVariantID(globalUserID, globalReportID), Convert.ToInt32(cboSelectionVariant.SelectedValue));
            //    newVariant = false;
            //}
        }
        private void tbcSelection_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            try
            {
                if (e.AddedItems.Count > 0)
                {
                    if (e.AddedItems[0] is TabItem)
                    {
                        if (((TabItem)e.AddedItems[0]).Header.Equals("Utvalg"))
                        {
                            tbiMultipleSelectionOptions.IsEnabled = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        public void LoadSingleValue(DataGridComboBoxColumn dgComb, string fieldName)
        {
            try
            {
                using (var entities = new DAL.SAPExlEntities())
                {
                    dgComb.ItemsSource = null;
                    switch (fieldName)
                    {
                        case "Merke":
                            dgComb.ItemsSource = entities.vwBrands.ToList();
                            break;
                        case "Fabrikk":
                            dgComb.ItemsSource = entities.vwPlants.ToList();
                            break;
                        case "Materialtype":
                            dgComb.ItemsSource = entities.vwMaterialTypes.ToList();
                            break;
                        case "Firmakode":
                            dgComb.ItemsSource = entities.vwCompanyCodes.ToList();
                            break;
                        case "Salesorganisasjon":
                            dgComb.ItemsSource = entities.vwSalesOrganizations.ToList();
                            break;
                        case "Materialgruppe":
                            dgComb.ItemsSource = entities.vwMaterialGroups.ToList();
                            break;
                        case "Produksjonsplanlegger":
                            dgComb.ItemsSource = entities.vwProductionScheduler.ToList();
                            break;
                        case "Arbeidsstasjongruppe":
                            if (globalReportID.Equals(11))
                            {
                                dgComb.ItemsSource = OrklaRTBPL.SelectionFacade.GetWorkCenterGroups(OrklaRTBPL.SelectionFacade.CapacityLevellingSelectionPlant).Tables[0].DefaultView;
                            }
                            break;
                        case "Arbeidsstasjon":
                            if (globalReportID.Equals(11))
                            {
                                dgComb.ItemsSource = OrklaRTBPL.SelectionFacade.GetWorkCenters(OrklaRTBPL.SelectionFacade.CapacityLevellingSelectionPlant).Tables[0].DefaultView;
                            }
                            //else if (globalReportID.Equals(7))
                            //{
                            //    cboSingleValues.ItemsSource = OrklaRTBPL.SelectionFacade.GetWorkCenters(OrklaRTBPL.SelectionFacade.ProductionPlanSelectionPlant).Tables[0].DefaultView;
                            //}
                            break;
                        //case "Valuation Class":
                        //    cboSingleValues.ItemsSource = entities.vwValuationClass.ToList();
                        //    break;
                        case "Innkjøpsgruppe":
                            dgComb.ItemsSource = entities.vwPurchasingGroups.ToList();
                            break;
                        case "Lagernummer":
                            dgComb.ItemsSource = entities.vwWarehouseNumbers.ToList();
                            break;
                        case "Lager":
                            dgComb.ItemsSource = entities.vwStorageLocations.ToList();
                            break;
                        case "Materialplanlegger":
                            dgComb.ItemsSource = entities.vwMRPControllers.ToList();
                            break;
                        case "MaterialArt":
                            dgComb.ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterialArts().Tables[0].DefaultView;
                            break;
                            //default:
                            //    cboSingleValues.ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
                            //    break;
                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }

        }

        public bool CheckReportMandatoryInput(int reportID, int userID)
        {
            bool ret = false;
            try
            {
                if (SelectionFacade.CheckMandatoryFields(reportID, globalUserID).Equals(0))
                {
                    ret = true;
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
            return ret;
        }

        public bool CheckReportRequiredInput(int reportID, int userID)
        {
            bool ret = false;
            try
            {
                ret = SelectionFacade.CheckRequiredFields(reportID, globalUserID);
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
            return ret;
        }

        public void LoadReportDefaultValues()
        {
            try
            {
                foreach (Control control in grdSelection.Children)
                {
                    if (!(control is Label) && !(control is Button))
                    {
                        //if (control.Tag.ToString().StartsWith("SP$"))
                        //{
                        using (var entities = new DAL.SAPExlEntities())
                        {
                            if (control is TextBox)
                            {
                                if (globalVariantID.Equals(0))
                                {
                                    if (control.Name.Equals("LowValue"))
                                    {
                                        ((TextBox)control).Text = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().LowValue;
                                    }
                                    else
                                    {
                                        ((TextBox)control).Text = entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().HighValue;
                                    }
                                }
                                //if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().FieldName.Equals("Period"))
                                //{
                                //    txtYear.Text = DateTime.Now.Year.ToString();
                                //    txtMonth.Text = DateTime.Now.AddMonths(-1).Month.ToString().Length == 1 ? "0" + DateTime.Now.AddMonths(-1).Month.ToString() : DateTime.Now.AddMonths(-1).Month.ToString();
                                //    ((TextBox)control).Text = DateTime.Now.Year.ToString() + (DateTime.Now.AddMonths(-1).Month.ToString().Length == 1 ? "0" + DateTime.Now.AddMonths(-1).Month.ToString() : DateTime.Now.AddMonths(-1).Month.ToString());
                                //}
                            }
                            else if (control is DatePicker)
                            {
                                if (control.Name.Equals("LowValue"))
                                {
                                    if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().LowValue != null)
                                    {
                                        switch (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().LowValue.ToString())
                                        {
                                            //case "0":
                                            //    ((DatePicker)control).SelectedDate = DateTime.Now.Date;
                                            //    break;
                                            case "-1WF":
                                                ((DatePicker)control).SelectedDate = OrklaRTBPL.CommonFacade.GetStartOfLastWeek();
                                                break;
                                            case "-1WL":
                                                ((DatePicker)control).SelectedDate = OrklaRTBPL.CommonFacade.GetEndOfLastWeek();
                                                break;
                                            case "-1MF":
                                                ((DatePicker)control).SelectedDate = OrklaRTBPL.CommonFacade.GetStartOfLastMonth();
                                                break;
                                            case "MF":
                                                ((DatePicker)control).SelectedDate = OrklaRTBPL.CommonFacade.GetStartOfCurrentMonth();
                                                break;
                                            case "-2M":
                                                ((DatePicker)control).SelectedDate = OrklaRTBPL.CommonFacade.Get2MonthsPeriod();
                                                break;
                                            case "-1Y":
                                                ((DatePicker)control).SelectedDate = DateTime.Now.AddYears(-1);
                                                break;
                                            default:
                                                if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().LowValue.Equals("0"))
                                                {
                                                    ((DatePicker)control).SelectedDate = DateTime.Now.Date;
                                                }
                                                else
                                                {
                                                    ((DatePicker)control).SelectedDate = DateTime.Now.AddDays(Convert.ToDouble(entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().LowValue)).Date;

                                                }
                                                break;
                                        }
                                    }
                                    if (((DatePicker)control).SelectedDate != null)
                                    {
                                        SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, ((DatePicker)control).Tag.ToString(), "S", "I", entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().SelectionOption.ToString(), ((DatePicker)control).SelectedDate.Value.ToString("yyyyMMdd"));
                                    }
                                    else
                                    {
                                        SelectionFacade.UpdateCurrentUserReportSelectionLowValue(globalReportID, globalUserID, ((DatePicker)control).Tag.ToString(), "S", "I", entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().SelectionOption.ToString(), DBNull.Value.ToString());
                                    }
                                }
                                else
                                {
                                    if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().HighValue != null)
                                    {
                                        switch (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().HighValue.ToString())
                                        {
                                            case "-1WL":
                                                ((DatePicker)control).SelectedDate = OrklaRTBPL.CommonFacade.GetEndOfLastWeek();
                                                break;
                                            case "ML":
                                                ((DatePicker)control).SelectedDate = OrklaRTBPL.CommonFacade.GetEndOfCurrentMonth();
                                                break;
                                            case "3M":
                                                ((DatePicker)control).SelectedDate = OrklaRTBPL.CommonFacade.Get3MonthsPeriod();
                                                break;
                                            default:
                                                if (entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().HighValue.Equals("0"))
                                                {
                                                    ((DatePicker)control).SelectedDate = DateTime.Now.Date;
                                                }
                                                else
                                                {
                                                    ((DatePicker)control).SelectedDate = DateTime.Now.AddDays(Convert.ToDouble(entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().HighValue)).Date;
                                                }
                                                break;
                                        }
                                    }
                                    if (((DatePicker)control).SelectedDate != null)
                                    {
                                        SelectionFacade.UpdateCurrentUserReportSelectionHighValue(globalReportID, globalUserID, ((DatePicker)control).Tag.ToString(), "S", "I", entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().SelectionOption.ToString(), ((DatePicker)control).SelectedDate.Value.ToString("yyyyMMdd"));
                                    }
                                    else
                                    {
                                        SelectionFacade.UpdateCurrentUserReportSelectionHighValue(globalReportID, globalUserID, ((DatePicker)control).Tag.ToString(), "S", "I", entities.ReportSelections.Where(p => p.ReportID == globalReportID && p.ScreenID == control.Tag).SingleOrDefault().SelectionOption.ToString(), DBNull.Value.ToString());
                                    }
                                    if (globalReportID.Equals(7) || globalReportID.Equals(63))
                                    {
                                        SelectionFacade.UpdateCurrentUserReportSelectionHighValue(globalReportID, globalUserID, "ZVR0062", "S", "I", "BT", ((DatePicker)control).SelectedDate.Value.AddDays(+10).ToString("yyyyMMdd"));
                                    }

                                }
                            }
                            //}
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
            }
        }

        //private void btnRefresh_Click(object sender, RoutedEventArgs e)
        //{
        //    try
        //    {
        //        if (cboPlants.SelectedValue != null && cboMaterialTypes.SelectedValue != null)
        //        {
        //            if (globalReportID.Equals(2))
        //            {
        //                foreach (Control control in grdSelection.Children)
        //                {
        //                    if (control is ComboBox)
        //                    {
        //                        if (control.Tag.ToString().Equals("ZVR013") || control.Tag.ToString().Equals("0MATERIAL1"))
        //                        {
        //                            ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //                        }                                                                                            
        //                    }
        //                }
        //            }
        //            else if (globalReportID.Equals(10))
        //            {
        //                foreach (Control control in grdSelection.Children)
        //                {
        //                    if (control is ComboBox)
        //                    {
        //                        if (control.Tag.ToString().Equals("ZVR013"))
        //                        {
        //                            ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //                            cboSingleValues.ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //                        }
        //                    }
        //                }
        //            }
        //            else if (globalReportID.Equals(18))
        //            {
        //                foreach (Control control in grdSelection.Children)
        //                {
        //                    if (control is ComboBox)
        //                    {
        //                        if (control.Tag.ToString().Equals("ZVR013"))
        //                        {
        //                            ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //                        }
        //                    }
        //                }
        //            }                  
        //            else if (globalReportID.Equals(19))
        //            {
        //                //foreach (Control control in grdSelection.Children)
        //                //{
        //                //    if (control is ComboBox)
        //                //    {
        //                //        if (control.Tag.ToString().Equals("SP$00010"))
        //                //        {
        //                //            ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //                //        }
        //                //    }
        //                //}
        //            }
        //            else if (globalReportID.Equals(33))
        //            {
        //                foreach (Control control in grdSelection.Children)
        //                {
        //                    if (control is ComboBox)
        //                    {
        //                        if (control.Tag.ToString().Equals("ZVR013"))
        //                        {
        //                            ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //                        }
        //                    }
        //                }
        //            }
        //            //else if (globalReportID.Equals(34))
        //            //{
        //            //    foreach (Control control in grdSelection.Children)
        //            //    {
        //            //        if (control is ComboBox)
        //            //        {
        //            //            if (control.Tag.ToString().Equals("ZVS005"))
        //            //            {
        //            //                ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //            //            }
        //            //        }
        //            //    }
        //            //}
        //            //else if (globalReportID.Equals(35))
        //            //{
        //            //    foreach (Control control in grdSelection.Children)
        //            //    {
        //            //        if (control is ComboBox)
        //            //        {
        //            //            if (control.Tag.ToString().Equals("ZVS005"))
        //            //            {
        //            //                ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //            //            }
        //            //        }
        //            //    }
        //            //}
        //            else if (globalReportID.Equals(38))
        //            {
        //                foreach (Control control in grdSelection.Children)
        //                {
        //                    if (control is ComboBox)
        //                    {
        //                        if (control.Tag.ToString().Equals("ZVR013"))
        //                        {
        //                            ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //                        }
        //                    }
        //                }
        //            }
        //            else if (globalReportID.Equals(39))
        //            {
        //                foreach (Control control in grdSelection.Children)
        //                {
        //                    if (control is ComboBox)
        //                    {
        //                        if (control.Tag.ToString().Equals("ZVR013"))
        //                        {
        //                            ((ComboBox)control).ItemsSource = OrklaRTBPL.SelectionFacade.GetMaterials(cboPlants.SelectedValue.ToString(), cboMaterialTypes.SelectedValue.ToString()).Tables[0].DefaultView;
        //                        }
        //                    }
        //                }
        //            }
        //        }
        //        else
        //        {
        //            MessageBox.Show("Vennligst fyll ut Fabrikk & Materialtype !");
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID, globalReportID);
        //    }
        //}

        //private void cboPlants_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    try
        //    {
        //        if (e.AddedItems.Count > 0)
        //        {
        //            CommonFacade.UpdateCurrentUserReportFields("SearchPlant", ((DataRowView)e.AddedItems[0]).Row.ItemArray[0].ToString(), globalUserID);
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID);
        //    }
        //}

        //private void cboMaterialTypes_SelectionChanged(object sender, SelectionChangedEventArgs e)
        //{
        //    try
        //    {
        //        if (e.AddedItems.Count > 0)
        //        {
        //            CommonFacade.UpdateCurrentUserReportFields("SearchMaterialType", ((DataRowView)e.AddedItems[0]).Row.ItemArray[0].ToString(), globalUserID);
        //        }
        //    }
        //    catch(Exception ex)
        //    {
        //        OrklaRTBPL.CommonFacade.InsertErrorLog(ex.Message, System.Reflection.MethodBase.GetCurrentMethod().Name, "Selection", globalUserID);
        //    }

        //}
    }
}

