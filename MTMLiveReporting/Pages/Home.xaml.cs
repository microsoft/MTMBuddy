using System;
using System.Collections.Generic;
using System.Configuration;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using MTMIntegration;
using ReportingLayer;
using WinForm = System.Windows.Forms;
using FirstFloor.ModernUI.Windows.Navigation;


namespace MTMLiveReporting.Pages
{
    /// <summary>
    ///     Interaction logic for Home.xaml
    /// </summary>
    public partial class Home : UserControl
    {
        private string currentreport = string.Empty;
        private int gridsuiteid;
        private List<resultsummary> rawdata = new List<resultsummary>();

        public Home()
        {
            InitializeComponent();
           
               
            
                try
            {

                if (string.IsNullOrEmpty(MTMInteraction.planName))
                {
                    MTMInteraction.initializeVSTFUpdate(new Uri(ConfigurationManager.AppSettings["TFSUrl"]),
                            ConfigurationManager.AppSettings["TeamProject"],
                            int.Parse(ConfigurationManager.AppSettings["TestPlanID"]),
                            ConfigurationManager.AppSettings["BuildNumber"]);
                    DataGetter.diagnostic.AppendLine("TFS URL: " + ConfigurationManager.AppSettings["TFSUrl"]);
                    DataGetter.diagnostic.AppendLine("Team Project: " + ConfigurationManager.AppSettings["TeamProject"]);
                    DataGetter.diagnostic.AppendLine("---------------------------------------------------");
                }
                MTMInteraction.getwpfsuitetree(int.Parse(ConfigurationManager.AppSettings["TestPlanID"]), tvMTM, true);
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtPlanId.Text = ConfigurationManager.AppSettings["TestPlanID"];
            var reports = new List<string>();
            reports.Add("Summary");
            reports.Add("OneLineSummary");
            reports.Add("Module");
            reports.Add("Tester");
            reports.Add("ModuleGroup");
            reports.Add("TesterGroup");
            reports.Add("Automation");
            reports.Add("TestList");
            cmbReportSelection.ItemsSource = reports;
            cmbReportSelection.SelectedIndex = 0;
            var AutomationStatus = new List<string>();
            AutomationStatus.Add("Both");
            AutomationStatus.Add("Automated");
            AutomationStatus.Add("Manual");
            cmbAutomationStaus.ItemsSource = AutomationStatus;
            int planid = 0;
            if (int.TryParse(ConfigurationManager.AppSettings["TestPlanID"], out planid))
            {
                if(planid!=0)
                planName.Text = MTMInteraction.getPlanName(planid);
            }
        }

        


   


        private void buttonExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
               
                if (ResultDataGrid.Items.Count.Equals(0) || string.IsNullOrEmpty(currentreport))
                {
                    MessageBox.Show("Nothing to export.Please generate a report", "OOPS!", MessageBoxButton.OK);
                    return;
                }

                switch (currentreport.ToUpper())
                {
                    case "SUMMARY":
                        var expExlSum = new ExportToExcel<SummaryReport, List<SummaryReport>>();
                        expExlSum.dataToExport = (List<SummaryReport>) ResultDataGrid.ItemsSource;
                        expExlSum.GenerateExcel();
                        break;
                    case "MODULE":
                        var expExlMod = new ExportToExcel<ModuleLevelReport, List<ModuleLevelReport>>();
                        expExlMod.dataToExport = (List<ModuleLevelReport>) ResultDataGrid.ItemsSource;
                        expExlMod.GenerateExcel();
                        break;
                    case "ONELINESUMMARY":
                        var expExlOLS = new ExportToExcel<OneLineSummary, List<OneLineSummary>>();
                        expExlOLS.dataToExport = (List<OneLineSummary>) ResultDataGrid.ItemsSource;
                        expExlOLS.GenerateExcel();
                        break;
                    case "MODULEGROUP":
                        var expExlMLG = new ExportToExcel<ModuleLevelReportGroup, List<ModuleLevelReportGroup>>();
                        expExlMLG.dataToExport = (List<ModuleLevelReportGroup>) ResultDataGrid.ItemsSource;
                        expExlMLG.GenerateExcel();
                        break;
                    case "TESTERGROUP":
                        var expExlTG = new ExportToExcel<TesterLevelReportGroup, List<TesterLevelReportGroup>>();
                        expExlTG.dataToExport = (List<TesterLevelReportGroup>) ResultDataGrid.ItemsSource;
                        expExlTG.GenerateExcel();
                        break;
                    case "TESTER":
                        var expExlTLR = new ExportToExcel<TesterLevelReport, List<TesterLevelReport>>();
                        expExlTLR.dataToExport = (List<TesterLevelReport>) ResultDataGrid.ItemsSource;
                        expExlTLR.GenerateExcel();
                        break;
                    case "EXECUTIONTREND":
                        var expExlexec = new ExportToExcel<ExecutionTrend, List<ExecutionTrend>>();
                        expExlexec.dataToExport = (List<ExecutionTrend>) ResultDataGrid.ItemsSource;
                        expExlexec.GenerateExcel();
                        break;
                    case "TESTLIST":
                        var tlist = new ExportToExcel<TestList, List<TestList>>();
                        tlist.dataToExport = (List<TestList>) ResultDataGrid.ItemsSource;
                        tlist.GenerateExcel();
                        break;
                    case "ISSUELIST":
                        var ilist = new ExportToExcel<resultsummary, List<resultsummary>>();
                        ilist.dataToExport = (List<resultsummary>) ResultDataGrid.ItemsSource;
                        ilist.GenerateExcel();
                        break;
                    case "AUTOMATION":
                        var alist = new ExportToExcel<AutomationReport, List<AutomationReport>>();
                        alist.dataToExport = (List<AutomationReport>) ResultDataGrid.ItemsSource;
                        alist.GenerateExcel();
                        break;
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                
            }
        }

        private void btnGenerateSummaryReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (tvMTM.SelectedItem == null)
                {
                    MessageBox.Show("Please select a suite", "Error");
                    return;
                }

                var selectedItem = (TreeViewItem) tvMTM.SelectedItem;


                var suiteid = int.Parse(selectedItem.Tag.ToString());


                if (!gridsuiteid.Equals(suiteid))
                {
                    rawdata.Clear();
                    try
                    {
                        Mouse.OverrideCursor = Cursors.Wait;
                        rawdata = DataGetter.GetResultSummaryList(suiteid, selectedItem.Header.ToString());
                    }
                    catch (Exception exp)
                    {
                        MessageBox.Show(
                            "We are not able to get data from VSTF.Please check if you are able to connect using Visual studio.Please send us the below error details if you are able to connect." +
                            Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);

                        return;
                    }
                    finally
                    {
                        Mouse.OverrideCursor = null;
                    }
                    gridsuiteid = suiteid;
                }


                var moduleinclusion = chkModuleInclusion.IsChecked.Value;
                var modulefilter = txtmodulefilter.Text;
                var testerinclusion = chkTesterInclusion.IsChecked.Value;
                var testerfilter = txttesterfilter.Text;
                var automationstatus = cmbAutomationStaus.SelectedItem.ToString();


                


                currentreport = cmbReportSelection.SelectedItem.ToString();
                switch (currentreport.ToUpper())
                {
                    case "SUMMARY":
                        ResultDataGrid.ItemsSource = SummaryReport.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "MODULE":
                        ResultDataGrid.ItemsSource = ModuleLevelReport.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "TESTER":
                        ResultDataGrid.ItemsSource = TesterLevelReport.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "MODULEGROUP":
                        ResultDataGrid.ItemsSource = ModuleLevelReportGroup.Generate(rawdata, modulefilter,
                            moduleinclusion, testerfilter, testerinclusion, automationstatus);
                        
                        break;
                    case "TESTERGROUP":
                        ResultDataGrid.ItemsSource = TesterLevelReportGroup.Generate(rawdata, modulefilter,
                            moduleinclusion, testerfilter, testerinclusion, automationstatus);
                        
                        break;
                    case "ONELINESUMMARY":
                        ResultDataGrid.ItemsSource = OneLineSummary.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        
                        break;
                    case "EXECUTIONTREND":
                        ResultDataGrid.ItemsSource = ExecutionTrend.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "TESTLIST":
                        ResultDataGrid.ItemsSource = TestList.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "AUTOMATION":
                        ResultDataGrid.ItemsSource = AutomationReport.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                      
                        break;
                }
                if (Utilities.GetIssueList(rawdata).Count > 0)
                {
                    MessageBox.Show(
                        "Your report is ready but we found some test cases which did not fit the report criteria. Please click on Show Issues Button for more details.");
                    buttonIssueList.Visibility = Visibility.Visible;
                    
                }
                else
                {
                    MessageBox.Show("Your report is ready!");
                    buttonIssueList.Visibility = Visibility.Hidden;
                }
               
               
                  
                
            }
            catch (Exception exp)
            {
                currentreport = string.Empty;
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                
            }
        }


        private void btnGenerateTree_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var planId = int.Parse(txtPlanId.Text);
                planName.Text = MTMInteraction.getPlanName(planId);
                if (planName.Text == null)
                {
                    MessageBox.Show("Invalid planId. Retry!");
                }
                if (int.TryParse(txtPlanId.Text, out planId))
                {
                    tvMTM.Items.Clear();
                    MTMInteraction.getwpfsuitetree(planId, tvMTM, true);
                }
                else
                {
                    MessageBox.Show("Invalid planid.You can get the planid from MTM in the Select Test Plan Window.",
                        "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);

              
            }
        }

        

        private void btnHelp_OnMouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Help;
            var current = sender as Button;
            current.Opacity = 1.0;
        }

        private void btnHelp_OnMouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Arrow;
            var current = sender as Button;
            current.Opacity = 0.8;
        }

        private void buttonIssueList_Click(object sender, RoutedEventArgs e)
        {
            ResultDataGrid.ItemsSource = Utilities.GetIssueList(rawdata);
            currentreport = "issuelist";
        }

        private void btnRefreshData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (tvMTM.SelectedItem == null)
                {
                    MessageBox.Show("Please select a suite", "Error");
                    return;
                }
                var selectedItem = (TreeViewItem) tvMTM.SelectedItem;
                var suiteid = int.Parse(selectedItem.Tag.ToString());

                rawdata.Clear();
                try
                {
                    rawdata = DataGetter.GetResultSummaryList(suiteid, selectedItem.Header.ToString());
                }
                catch (Exception exp)
                {
                    MessageBox.Show(
                        "We are not able to get data from VSTF.Please check if you are able to connect using Visual studio.Please send us the below error details if you are able to connect." +
                        Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);

                    return;
                }
                gridsuiteid = suiteid;
                
                var moduleinclusion = chkModuleInclusion.IsChecked.Value;
                var modulefilter = txtmodulefilter.Text;
                var testerinclusion = chkTesterInclusion.IsChecked.Value;
                var testerfilter = txttesterfilter.Text;
                var automationstatus = cmbAutomationStaus.SelectedItem.ToString();
                currentreport = cmbReportSelection.SelectedItem.ToString();
                switch (currentreport.ToUpper())
                {
                    case "SUMMARY":
                        ResultDataGrid.ItemsSource = SummaryReport.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                    case "MODULE":
                        ResultDataGrid.ItemsSource = ModuleLevelReport.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                    case "TESTER":
                        ResultDataGrid.ItemsSource = TesterLevelReport.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                    case "MODULEGROUP":
                        ResultDataGrid.ItemsSource = ModuleLevelReportGroup.Generate(rawdata, modulefilter,
                            moduleinclusion, testerfilter, testerinclusion, automationstatus);
                        break;
                    case "TESTERGROUP":
                        ResultDataGrid.ItemsSource = TesterLevelReportGroup.Generate(rawdata, modulefilter,
                            moduleinclusion, testerfilter, testerinclusion, automationstatus);
                        break;
                    case "ONELINESUMMARY":
                        ResultDataGrid.ItemsSource = OneLineSummary.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                    case "EXECUTIONTREND":
                        ResultDataGrid.ItemsSource = ExecutionTrend.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                    case "TESTLIST":
                        ResultDataGrid.ItemsSource = TestList.Generate(rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                }
                if (Utilities.GetIssueList(rawdata).Count > 0)
                {
                    MessageBox.Show(
                        "Your report is ready but we found some test cases which did not fit the report criteria. Please click on Show Issues Button for more details.");
                    buttonIssueList.Visibility = Visibility.Visible;
                }
                else
                {
                    MessageBox.Show("Your report is ready!");
                    buttonIssueList.Visibility = Visibility.Hidden;
                }
               
            }
            catch (Exception exp)
            {
                currentreport = string.Empty;
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
               
            }
        }

        }
}