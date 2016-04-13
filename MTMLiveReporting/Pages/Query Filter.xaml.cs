using System;
using System.Collections.Generic;
using System.Configuration;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using MTMIntegration;
using ReportingLayer;
using Application = Microsoft.Office.Interop.Excel.Application;
using WinForm = System.Windows.Forms;

namespace MTMLiveReporting.Pages
{
    /// <summary>
    ///     Interaction logic for Query_Filter.xaml
    /// </summary>
    public partial class Query_Filter : UserControl
    {
        private int planId;

        private List<QueryInterface> QueryResult = new List<QueryInterface>();
        private List<resultsummary> rawdata = new List<resultsummary>();

        private int suiteid;


        public Query_Filter()
        {
            DataContext = this;

            InitializeComponent();

            try
            {
                var Actions = new List<string>();
                Actions.Add("Assign Tester");
                Actions.Add("Playlist");
                Actions.Add("Result Update");

                testerAction.ItemsSource = Actions;


                testerAction.SelectedIndex = 0;
                var AutomationStatus = new List<string>();
                AutomationStatus.Add("Both");
                AutomationStatus.Add("Automated");
                AutomationStatus.Add("Manual");
                cmbAutomationStaus.ItemsSource = AutomationStatus;
                var Priorities = new List<string>();
                Priorities.Add("All");
                Priorities.Add("1");
                Priorities.Add("2");
                Priorities.Add("3");
                Priorities.Add("4");
                cmbPriority.ItemsSource = Priorities;
                var Outcomes = new List<string>();
                Outcomes.Add("All");
                Outcomes.Add("Active");
                Outcomes.Add("Blocked");
                Outcomes.Add("Failed");
                Outcomes.Add("Passed");
                cmbOutcome.ItemsSource = Outcomes;

                var validOutcomes = Outcomes;

                validOutcomes.Remove("All");
                resultoptions.ItemsSource = validOutcomes;
               
                
                
              
                MTMInteraction.getwpfsuitetree(int.Parse(ConfigurationManager.AppSettings["TestPlanID"]), tvMTM, true);
                var tester = new List<string>();
                tester = MTMInteraction.getTester();
                tester.Sort();
                testerName.ItemsSource = tester;

            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            txtPlanId.Text = ConfigurationManager.AppSettings["TestPlanID"];
        }


        public void getlist_Click(object sender, RoutedEventArgs e)
        {
            try {
                var plans = new Dictionary<int, string>();
                plans = MTMInteraction.getPlanId();

                if (int.TryParse(txtPlanId.Text, out planId))
                {
                    tvMTM.Items.Clear();
                    MTMInteraction.getwpfsuitetree(planId, tvMTM, true);
                    var tester = new List<string>();
                    tester = MTMInteraction.getTester();
                    tester.Sort();
                    testerName.ItemsSource = tester;
                }

                else
                {
                    MessageBox.Show("Invalid PlanId");
                }
            
             }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
}

      
        public void getFilter_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedItem = (TreeViewItem) tvMTM.SelectedItem;
                if (suiteid != int.Parse(selectedItem.Tag.ToString()))
                {
                    suiteid = int.Parse(selectedItem.Tag.ToString());
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
                }
                var Filters = new List<MTMInteraction.filter>();

                var moduleinclusion = chkModuleInclusion.IsChecked.Value;
                var modulefilter = txtmodulefilter.Text;
                var testerinclusion = chkTesterInclusion.IsChecked.Value;
                var testerfilter = txttesterfilter.Text;
                var automationstatus = cmbAutomationStaus.SelectedItem.ToString();
                if (!string.IsNullOrEmpty(txtTitlefilter.Text))
                {
                    
                    var TitleFilter = new MTMInteraction.filter();
                    if (chkTitleInclusion.IsChecked.Value)
                    {
                        TitleFilter.op = "Contains";
                    }
                    else
                    {
                        TitleFilter.op = "Not Contains";
                    }
                    TitleFilter.name = "Title";
                    TitleFilter.value = txtTitlefilter.Text;
                    Filters.Add(TitleFilter);
                }
                if (!cmbPriority.SelectedValue.ToString().Equals("All", StringComparison.InvariantCultureIgnoreCase))
                {
                   
                    var TitleFilter = new MTMInteraction.filter();
                    
                    TitleFilter.name = "Priority";
                    TitleFilter.value = cmbPriority.SelectedValue.ToString();
                    Filters.Add(TitleFilter);
                }
                if (!cmbOutcome.SelectedValue.ToString().Equals("All", StringComparison.InvariantCultureIgnoreCase))
                {
                    
                    var TitleFilter = new MTMInteraction.filter();

                    TitleFilter.name = "Outcome";
                    TitleFilter.value = cmbOutcome.SelectedValue.ToString();
                    Filters.Add(TitleFilter);
                }

                QueryResult = QueryInterface.Generate(rawdata, modulefilter, moduleinclusion, testerfilter,
                    testerinclusion, automationstatus, Filters);
                    ResultDataGrid.ItemsSource = QueryResult;
                   
                MessageBox.Show("All done. Found " + QueryResult.Count + " Cases");
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
               
            }
            finally
            {
                Mouse.OverrideCursor = null;
                ;
            }
        }

        private void Assign_Tester()
        {
            try {
                var name = testerName.SelectedItem.ToString();
                var selectedID = new List<int>();
                foreach (var i in QueryResult)
                {
                    if (i.Selected)
                    {
                        selectedID.Add(i.TCId);
                    }
                }
                if (selectedID.Count == 0)
                {
                    MessageBox.Show("Hey! It seems you forgot to select cases");
                    return;
                }
                MTMInteraction.assignTester(suiteid, selectedID, name);
               
                MessageBox.Show("All changes done!");
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
        }
        

        private void Hyperlink_Click(object sender, RoutedEventArgs e)
        {
            try {
                foreach (var i in QueryResult)
                {
                    i.Selected = true;
                }
                ResultDataGrid.Items.Refresh();
               
            
             }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
}

        private void Updateresults()
        {
            try
            {
                //string name = testerName.SelectedItem.ToString();
                Mouse.OverrideCursor = Cursors.Wait;
                var errflag = false;
                foreach (var i in QueryResult)
                {
                    if (i.Selected)
                    {
                        try
                        {
                            AssociateTestCase.updateResult(i.TCId.ToString(), "Updated from MTM Buddy",
                                i.SuiteName.Substring(i.SuiteName.LastIndexOf("\\") + 1), false,
                                resultoptions.SelectedValue.ToString());
                        }
                        catch (Exception ex)
                        {
                            DataGetter.diagnostic.AppendLine("Error occured in migrating testcase " + i.TCId);
                            DataGetter.diagnostic.AppendLine(ex.Message);
                            DataGetter.diagnostic.AppendLine("---------------------------------------------------");
                            errflag = true;
                        
                        }
                    }
                }
                if (errflag)
                {
                    MessageBox.Show("Sorry! Something went wrong. Please check the diagnostic log for details.");
                }
                else
                {
                    MessageBox.Show("All done");
                }

                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sorry! Something went wrong. Please send the diagnostic log to us.");
               
                DataGetter.diagnostic.AppendLine("Error occured in migrating testcases.");
                DataGetter.diagnostic.AppendLine(ex.Message);
                DataGetter.diagnostic.AppendLine("---------------------------------------------------");
                
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }


        

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try { 
            Mouse.OverrideCursor = Cursors.Wait;
            if (testerAction.SelectedItem.ToString()
                .Equals("Assign Tester", StringComparison.InvariantCultureIgnoreCase))
            {
                Assign_Tester();
            }

            else if (testerAction.SelectedItem.ToString().Equals("Playlist", StringComparison.InvariantCultureIgnoreCase))
            {
                Create_Playlist();
            }

           

            else if (testerAction.SelectedItem.ToString()
                .Equals("Result Update", StringComparison.InvariantCultureIgnoreCase))
            {
                Updateresults();
            }

           
            Mouse.OverrideCursor = Cursors.Arrow;
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
        }

        private void Create_Playlist()
        {
            try
            {
                //string name = testerName.SelectedItem.ToString();
                var selectedID = new List<int>();
                foreach (var i in QueryResult)
                {
                    if (i.Selected && i.AutomationStatus)
                    {
                        selectedID.Add(i.TCId);
                    }
                }

                if (selectedID.Count == 0)
                {
                    MessageBox.Show("Hey! It seems there are no automated cases here to generate the playlist");
                    return;
                }

                var dlg = new SaveFileDialog();
                dlg.FileName = "Playlist1"; // Default file name
                dlg.DefaultExt = ".playlist"; // Default file extension
                dlg.Filter = "Playlist (.playlist)|*.playlist"; // Filter files by extension

                var result = dlg.ShowDialog();
                var fileLocation = "";
                if (result == true)
                {
                    fileLocation = dlg.FileName;
                    filePath.Text = fileLocation;
                    Mouse.OverrideCursor = Cursors.Wait;
                    MTMInteraction.createPlaylist(suiteid, selectedID, fileLocation);
                    
                    DataGetter.diagnostic.AppendLine("Total time taken for playlist generation: " +
                                                     (MTMInteraction.AutomationMethodTime1 +
                                                      MTMInteraction.AutomationMethodTime));
                    DataGetter.diagnostic.AppendLine("Time taken for playlist data fetch " +
                                                     MTMInteraction.AutomationMethodTime);
                    DataGetter.diagnostic.AppendLine("Time taken for playlist file generation: " +
                                                     MTMInteraction.AutomationMethodTime1);
                    DataGetter.diagnostic.AppendLine("---------------------------------------------------");

                    MessageBox.Show("Playlist generated!");
                }              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sorry! Something went wrong. Please send the diagnostic log to us.");
                DataGetter.diagnostic.AppendLine("Error occured in generating playlist."); 
                DataGetter.diagnostic.AppendLine(ex.Message);
                DataGetter.diagnostic.AppendLine("---------------------------------------------------");
                
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void testerAction_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {
            if (testerAction.SelectedItem.ToString()
                .Equals("Assign Tester", StringComparison.InvariantCultureIgnoreCase))
            {
                testerName.Visibility = Visibility.Visible;
                lblActionParam.Visibility = Visibility.Visible;
                
                resultoptions.Visibility = Visibility.Hidden;
                testActionsoptions.Visibility = Visibility.Hidden;
                //filePath.Visibility =  System.Windows.Visibility.Hidden;
                actionButton.Content = "Assign";
            }
            else if (testerAction.SelectedItem.ToString().Equals("Playlist", StringComparison.InvariantCultureIgnoreCase))
            {
                testerName.Visibility = Visibility.Hidden;
                lblActionParam.Visibility = Visibility.Hidden;
               
                resultoptions.Visibility = Visibility.Hidden;
                testActionsoptions.Visibility = Visibility.Hidden;
                //filePath.Visibility = System.Windows.Visibility.Visible;
                actionButton.Content = "Generate";
            }
           
            else if (testerAction.SelectedItem.ToString()
                .Equals("Result Update", StringComparison.InvariantCultureIgnoreCase))
            {
                testerName.Visibility = Visibility.Hidden;
                lblActionParam.Visibility = Visibility.Visible ;
                lblActionParam.Text = "Outcome";
                resultoptions.Visibility = Visibility.Visible;
                testActionsoptions.Visibility = Visibility.Hidden;
                actionButton.Content = "Update";
            }
           
        }

        private void Hyperlink_Click_1(object sender, RoutedEventArgs e)
        {

            try
            {
                if (ResultDataGrid.Items.Count.Equals(0))
                {
                    MessageBox.Show("Nothing to export.Please generate a report", "OOPS!", MessageBoxButton.OK);
                    return;
                }
                var expExlSum = new ExportToExcel<QueryInterface, List<QueryInterface>>();
                expExlSum.dataToExport = (List<QueryInterface>)ResultDataGrid.ItemsSource;
                expExlSum.GenerateExcel();
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
        }
    }
}
