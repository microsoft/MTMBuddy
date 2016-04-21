//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
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
        private string _currentreport = string.Empty;
        private int _gridsuiteid;
        private List<ResultSummary> _rawdata = new List<ResultSummary>();

        public Home()
        {
            InitializeComponent();
           
               
            
                try
            {

                if (string.IsNullOrEmpty(MtmInteraction.SelectedPlanName))
                {
                    MtmInteraction.Initialize_VSTF(new Uri(ConfigurationManager.AppSettings["TFSUrl"]),
                            ConfigurationManager.AppSettings["TeamProject"],
                            int.Parse(ConfigurationManager.AppSettings["TestPlanID"]));
                    DataGetter.Diagnostic.AppendLine("TFS URL: " + ConfigurationManager.AppSettings["TFSUrl"]);
                    DataGetter.Diagnostic.AppendLine("Team Project: " + ConfigurationManager.AppSettings["TeamProject"]);
                    DataGetter.Diagnostic.AppendLine("---------------------------------------------------");
                }
                MtmInteraction.Getwpfsuitetree(int.Parse(ConfigurationManager.AppSettings["TestPlanID"]), TvMtm, true);
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            TxtPlanId.Text = ConfigurationManager.AppSettings["TestPlanID"];
            var reports = new List<string>();
            reports.Add("Summary");
            reports.Add("OneLineSummary");
            reports.Add("Module");
            reports.Add("Tester");
            reports.Add("ModuleGroup");
            reports.Add("TesterGroup");
            reports.Add("Automation");
            reports.Add("TestList");
            CmbReportSelection.ItemsSource = reports;
            CmbReportSelection.SelectedIndex = 0;
            var automationStatus = new List<string>();
            automationStatus.Add("Both");
            automationStatus.Add("Automated");
            automationStatus.Add("Manual");
            CmbAutomationStaus.ItemsSource = automationStatus;
            int planid = 0;
            if (int.TryParse(ConfigurationManager.AppSettings["TestPlanID"], out planid))
            {
                if(planid!=0)
                PlanName.Text = MtmInteraction.GetPlanName(planid);
            }
        }

        


   


        private void buttonExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
               
                if (ResultDataGrid.Items.Count.Equals(0) || string.IsNullOrEmpty(_currentreport))
                {
                    MessageBox.Show("Nothing to export.Please generate a report", "OOPS!", MessageBoxButton.OK);
                    return;
                }

                switch (_currentreport.ToUpper())
                {
                    case "SUMMARY":
                        var expExlSum = new ExportToExcel<SummaryReport, List<SummaryReport>>();
                        expExlSum.DataToExport = (List<SummaryReport>) ResultDataGrid.ItemsSource;
                        expExlSum.GenerateExcel();
                        break;
                    case "MODULE":
                        var expExlMod = new ExportToExcel<ModuleLevelReport, List<ModuleLevelReport>>();
                        expExlMod.DataToExport = (List<ModuleLevelReport>) ResultDataGrid.ItemsSource;
                        expExlMod.GenerateExcel();
                        break;
                    case "ONELINESUMMARY":
                        var expExlOls = new ExportToExcel<OneLineSummary, List<OneLineSummary>>();
                        expExlOls.DataToExport = (List<OneLineSummary>) ResultDataGrid.ItemsSource;
                        expExlOls.GenerateExcel();
                        break;
                    case "MODULEGROUP":
                        var expExlMlg = new ExportToExcel<ModuleLevelReportGroup, List<ModuleLevelReportGroup>>();
                        expExlMlg.DataToExport = (List<ModuleLevelReportGroup>) ResultDataGrid.ItemsSource;
                        expExlMlg.GenerateExcel();
                        break;
                    case "TESTERGROUP":
                        var expExlTg = new ExportToExcel<TesterLevelReportGroup, List<TesterLevelReportGroup>>();
                        expExlTg.DataToExport = (List<TesterLevelReportGroup>) ResultDataGrid.ItemsSource;
                        expExlTg.GenerateExcel();
                        break;
                    case "TESTER":
                        var expExlTlr = new ExportToExcel<TesterLevelReport, List<TesterLevelReport>>();
                        expExlTlr.DataToExport = (List<TesterLevelReport>) ResultDataGrid.ItemsSource;
                        expExlTlr.GenerateExcel();
                        break;
                  
                    case "TESTLIST":
                        var tlist = new ExportToExcel<TestList, List<TestList>>();
                        tlist.DataToExport = (List<TestList>) ResultDataGrid.ItemsSource;
                        tlist.GenerateExcel();
                        break;
                    case "ISSUELIST":
                        var ilist = new ExportToExcel<ResultSummary, List<ResultSummary>>();
                        ilist.DataToExport = (List<ResultSummary>) ResultDataGrid.ItemsSource;
                        ilist.GenerateExcel();
                        break;
                    case "AUTOMATION":
                        var alist = new ExportToExcel<AutomationReport, List<AutomationReport>>();
                        alist.DataToExport = (List<AutomationReport>) ResultDataGrid.ItemsSource;
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
                if (TvMtm.SelectedItem == null)
                {
                    MessageBox.Show("Please select a suite", "Error");
                    return;
                }

                var selectedItem = (TreeViewItem) TvMtm.SelectedItem;


                var suiteid = int.Parse(selectedItem.Tag.ToString());


                if (!_gridsuiteid.Equals(suiteid))
                {
                    _rawdata.Clear();
                    try
                    {
                        Mouse.OverrideCursor = Cursors.Wait;
                        _rawdata = DataGetter.GetResultSummaryList(suiteid, selectedItem.Header.ToString());
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
                    _gridsuiteid = suiteid;
                }


                var moduleinclusion = ChkModuleInclusion.IsChecked.Value;
                var modulefilter = Txtmodulefilter.Text;
                var testerinclusion = ChkTesterInclusion.IsChecked.Value;
                var testerfilter = Txttesterfilter.Text;
                var automationstatus = CmbAutomationStaus.SelectedItem.ToString();


                


                _currentreport = CmbReportSelection.SelectedItem.ToString();
                switch (_currentreport.ToUpper())
                {
                    case "SUMMARY":
                        ResultDataGrid.ItemsSource = SummaryReport.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "MODULE":
                        ResultDataGrid.ItemsSource = ModuleLevelReport.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "TESTER":
                        ResultDataGrid.ItemsSource = TesterLevelReport.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "MODULEGROUP":
                        ResultDataGrid.ItemsSource = ModuleLevelReportGroup.Generate(_rawdata, modulefilter,
                            moduleinclusion, testerfilter, testerinclusion, automationstatus);
                        
                        break;
                    case "TESTERGROUP":
                        ResultDataGrid.ItemsSource = TesterLevelReportGroup.Generate(_rawdata, modulefilter,
                            moduleinclusion, testerfilter, testerinclusion, automationstatus);
                        
                        break;
                    case "ONELINESUMMARY":
                        ResultDataGrid.ItemsSource = OneLineSummary.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        
                        break;
                    
                       
                       
                    case "TESTLIST":
                        ResultDataGrid.ItemsSource = TestList.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                       
                        break;
                    case "AUTOMATION":
                        ResultDataGrid.ItemsSource = AutomationReport.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                      
                        break;
                }
                if (Utilities.GetIssueList(_rawdata).Count > 0)
                {
                    MessageBox.Show(
                        "Your report is ready but we found some test cases which did not fit the report criteria. Please click on Show Issues Button for more details.");
                    ButtonIssueList.Visibility = Visibility.Visible;
                    
                }
                else
                {
                    MessageBox.Show("Your report is ready!");
                    ButtonIssueList.Visibility = Visibility.Hidden;
                }
               
               
                  
                
            }
            catch (Exception exp)
            {
                _currentreport = string.Empty;
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                
            }
        }


        private void btnGenerateTree_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var planId = int.Parse(TxtPlanId.Text);
                PlanName.Text = MtmInteraction.GetPlanName(planId);
                if (PlanName.Text == null)
                {
                    MessageBox.Show("Invalid planId. Retry!");
                }
                if (int.TryParse(TxtPlanId.Text, out planId))
                {
                    TvMtm.Items.Clear();
                    MtmInteraction.Getwpfsuitetree(planId, TvMtm, true);
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
            ResultDataGrid.ItemsSource = Utilities.GetIssueList(_rawdata);
            _currentreport = "issuelist";
        }

        private void btnRefreshData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (TvMtm.SelectedItem == null)
                {
                    MessageBox.Show("Please select a suite", "Error");
                    return;
                }
                var selectedItem = (TreeViewItem) TvMtm.SelectedItem;
                var suiteid = int.Parse(selectedItem.Tag.ToString());

                _rawdata.Clear();
                try
                {
                    _rawdata = DataGetter.GetResultSummaryList(suiteid, selectedItem.Header.ToString());
                }
                catch (Exception exp)
                {
                    MessageBox.Show(
                        "We are not able to get data from VSTF.Please check if you are able to connect using Visual studio.Please send us the below error details if you are able to connect." +
                        Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);

                    return;
                }
                _gridsuiteid = suiteid;
                
                var moduleinclusion = ChkModuleInclusion.IsChecked.Value;
                var modulefilter = Txtmodulefilter.Text;
                var testerinclusion = ChkTesterInclusion.IsChecked.Value;
                var testerfilter = Txttesterfilter.Text;
                var automationstatus = CmbAutomationStaus.SelectedItem.ToString();
                _currentreport = CmbReportSelection.SelectedItem.ToString();
                switch (_currentreport.ToUpper())
                {
                    case "SUMMARY":
                        ResultDataGrid.ItemsSource = SummaryReport.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                    case "MODULE":
                        ResultDataGrid.ItemsSource = ModuleLevelReport.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                    case "TESTER":
                        ResultDataGrid.ItemsSource = TesterLevelReport.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                    case "MODULEGROUP":
                        ResultDataGrid.ItemsSource = ModuleLevelReportGroup.Generate(_rawdata, modulefilter,
                            moduleinclusion, testerfilter, testerinclusion, automationstatus);
                        break;
                    case "TESTERGROUP":
                        ResultDataGrid.ItemsSource = TesterLevelReportGroup.Generate(_rawdata, modulefilter,
                            moduleinclusion, testerfilter, testerinclusion, automationstatus);
                        break;
                    case "ONELINESUMMARY":
                        ResultDataGrid.ItemsSource = OneLineSummary.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                
                    case "TESTLIST":
                        ResultDataGrid.ItemsSource = TestList.Generate(_rawdata, modulefilter, moduleinclusion,
                            testerfilter, testerinclusion, automationstatus);
                        break;
                }
                if (Utilities.GetIssueList(_rawdata).Count > 0)
                {
                    MessageBox.Show(
                        "Your report is ready but we found some test cases which did not fit the report criteria. Please click on Show Issues Button for more details.");
                    ButtonIssueList.Visibility = Visibility.Visible;
                }
                else
                {
                    MessageBox.Show("Your report is ready!");
                    ButtonIssueList.Visibility = Visibility.Hidden;
                }
               
            }
            catch (Exception exp)
            {
                _currentreport = string.Empty;
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
               
            }
        }

        }
}