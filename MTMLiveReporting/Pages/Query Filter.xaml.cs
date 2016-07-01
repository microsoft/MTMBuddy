//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using Microsoft.Win32;
using MTMIntegration;
using ReportingLayer;


namespace MTMLiveReporting.Pages
{
    /// <summary>
    ///     Interaction logic for Query_Filter.xaml
    /// </summary>
    public partial class QueryFilter 
    {
        private int _planId;

        private List<QueryInterface> _queryResult = new List<QueryInterface>();
        private List<ResultSummary> _rawdata = new List<ResultSummary>();

        private int _suiteid;


        public QueryFilter()
        {
            DataContext = this;

            InitializeComponent();

            try
            {
                var actions = new List<string> {"Assign Tester", "Playlist", "Result Update","Result History"};

                TesterAction.ItemsSource = actions;


                TesterAction.SelectedIndex = 0;
                var automationStatus = new List<string> {"Both", "Automated", "Manual"};
                CmbAutomationStaus.ItemsSource = automationStatus;
                var priorities = new List<string> {"All", "1", "2", "3", "4"};
                CmbPriority.ItemsSource = priorities;
                var outcomes = new List<string> {"All", "Active", "Blocked", "Failed", "Passed", "In progress"};
                CmbOutcome.ItemsSource = outcomes;

                //Valid Outcomes are used for result update.
                var validOutcomes = outcomes;

                validOutcomes.Remove("All");
                Resultoptions.ItemsSource = validOutcomes;
               
                
                
              
                MtmInteraction.Getwpfsuitetree(int.Parse(ConfigurationManager.AppSettings["TestPlanID"]), TvMtm, true);
                var tester = MtmInteraction.GetTester();
               
                tester.Sort();
                TesterName.ItemsSource = tester;

            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            TxtPlanId.Text = ConfigurationManager.AppSettings["TestPlanID"];
        }


        public void GetPlanDetails_Click(object sender, RoutedEventArgs e)
        {
            try {
                

                if (int.TryParse(TxtPlanId.Text, out _planId))
                {
                    TvMtm.Items.Clear();
                    MtmInteraction.Getwpfsuitetree(_planId, TvMtm, true);
                    var tester = MtmInteraction.GetTester();
                    tester.Sort();
                    TesterName.ItemsSource = tester;
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
                DataGetter.Diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
}

      
        public void getResults_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                //Get data for the selected suite
                var selectedItem = (TreeViewItem) TvMtm.SelectedItem;
               
                    _suiteid = int.Parse(selectedItem.Tag.ToString());
                    try
                    {
                        Mouse.OverrideCursor = Cursors.Wait;
                        _rawdata = DataGetter.GetResultSummaryList(_suiteid, selectedItem.Header.ToString());
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
                //Add Filters
                var filters = new List<Filter>();

                var moduleinclusion = ChkModuleInclusion.IsChecked != null && ChkModuleInclusion.IsChecked.Value;
                var modulefilter = Txtmodulefilter.Text;
                var testerinclusion = ChkTesterInclusion.IsChecked != null && ChkTesterInclusion.IsChecked.Value;
                var testerfilter = Txttesterfilter.Text;
                var automationstatus = CmbAutomationStaus.SelectedItem.ToString();
                if (!string.IsNullOrEmpty(TxtTitlefilter.Text))
                {
                    
                    var titleFilter = new Filter();
                    if (ChkTitleInclusion.IsChecked != null && ChkTitleInclusion.IsChecked.Value)
                    {
                        titleFilter.Op = "Contains";
                    }
                    else
                    {
                        titleFilter.Op = "Not Contains";
                    }
                    titleFilter.Name = "Title";
                    titleFilter.Value = TxtTitlefilter.Text;
                    filters.Add(titleFilter);
                }
                if (!CmbPriority.SelectedValue.ToString().Equals("All", StringComparison.InvariantCultureIgnoreCase))
                {

                    var titleFilter = new Filter {Name = "Priority", Value = CmbPriority.SelectedValue.ToString()};

                    filters.Add(titleFilter);
                }
                if (!CmbOutcome.SelectedValue.ToString().Equals("All", StringComparison.InvariantCultureIgnoreCase))
                {

                    var titleFilter = new Filter {Name = "Outcome", Value = CmbOutcome.SelectedValue.ToString()};

                    filters.Add(titleFilter);
                }

                _queryResult = QueryInterface.Generate(_rawdata, modulefilter, moduleinclusion, testerfilter,
                    testerinclusion, automationstatus, filters);
                    ResultDataGrid.ItemsSource = _queryResult;
                   
                MessageBox.Show("All done. Found " + _queryResult.Count + " Cases");
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
                
            }
        }

        private void Assign_Tester()
        {
            try {
                var name = TesterName.Text;
                var selectedId = (from i in _queryResult where i.Selected select i.TcId).ToList();
                if (selectedId.Count == 0)
                {
                    MessageBox.Show("Hey! It seems you forgot to select cases");
                    return;
                }
                MtmInteraction.AssignTester(_suiteid, selectedId, name);
               
                MessageBox.Show("All changes done!");
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.Diagnostic.AppendLine("Error while Assigning Testers: " + exp.Message);
            }
        }


        private void FetchHistory()
        {
            try
            {
                Mouse.OverrideCursor = Cursors.Wait;
                var selectedId = (from i in _queryResult where i.Selected select i.TcId).ToList();
                if (selectedId.Count == 0)
                {
                    MessageBox.Show("Hey! It seems you forgot to select cases", "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
                if (MtmInteraction.PlanIdsDictionary.Count*selectedId.Count > 100)
                    MessageBox.Show("We are going through " + MtmInteraction.PlanIdsDictionary.Count +
                                    " plans\r\n and searching for " + selectedId.Count + " cases to get the history" +
                                    Environment.NewLine + "This'll take a while so please bear with us","MTMBuddy",MessageBoxButton.OK,MessageBoxImage.Information);
                List<ResultHistorySummary> reshistory = MtmInteraction.GetResultHistory(selectedId);
                ResultDataGrid.ItemsSource = reshistory;


            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.Diagnostic.AppendLine("Error while Assigning Testers: " + exp.Message);
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }


        private void SelectAll_Click(object sender, RoutedEventArgs e)
        {
            try {
                foreach (var i in _queryResult)
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
                DataGetter.Diagnostic.AppendLine("Select All: " + exp.Message);
            }
}

        private void Updateresults()
        {
            try
            {
                //string name = testerName.SelectedItem.ToString();
                Mouse.OverrideCursor = Cursors.Wait;
                var errflag = false;
                foreach (var i in _queryResult.Where(i => i.Selected))
                {
                    try
                    {
                        MtmInteraction.UpdateResult(i.TcId.ToString(), "Updated from MTM Buddy",
                            i.SuiteName.Substring(i.SuiteName.LastIndexOf("\\",StringComparison.OrdinalIgnoreCase ) + 1), false,
                            Resultoptions.SelectedValue.ToString());
                    }
                    catch (Exception ex)
                    {
                        DataGetter.Diagnostic.AppendLine("Error occured in updating results " + i.TcId);
                        DataGetter.Diagnostic.AppendLine(ex.Message);
                        DataGetter.Diagnostic.AppendLine("---------------------------------------------------");
                        errflag = true;
                        
                    }
                }
                MessageBox.Show(errflag
                    ? "Sorry! Something went wrong. Please check the diagnostic log for details."
                    : "All done");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sorry! Something went wrong. Please send the diagnostic log to us.");
               
                DataGetter.Diagnostic.AppendLine("Error occured in updating results.");
                DataGetter.Diagnostic.AppendLine(ex.Message);
                DataGetter.Diagnostic.AppendLine("---------------------------------------------------");
                
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }


        

        private void ActionButton_Click(object sender, RoutedEventArgs e)
        {
            try { 
            Mouse.OverrideCursor = Cursors.Wait;
            if (TesterAction.SelectedItem.ToString()
                .Equals("Assign Tester", StringComparison.InvariantCultureIgnoreCase))
            {
                Assign_Tester();
            }

            else if (TesterAction.SelectedItem.ToString().Equals("Playlist", StringComparison.InvariantCultureIgnoreCase))
            {
                Create_Playlist();
            }

           

            else if (TesterAction.SelectedItem.ToString()
                .Equals("Result Update", StringComparison.InvariantCultureIgnoreCase))
            {
                Updateresults();
            }
                else if (TesterAction.SelectedItem.ToString()
                    .Equals("Result History", StringComparison.InvariantCultureIgnoreCase))
                {
                    FetchHistory();
                }

                    Mouse.OverrideCursor = Cursors.Arrow;
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.Diagnostic.AppendLine("Error while "+TesterAction.SelectedItem+": " + exp.Message);
            }
        }

        private void Create_Playlist()
        {
            try
            {
                //string name = testerName.SelectedItem.ToString();
                var selectedId = (from i in _queryResult where i.Selected && i.AutomationStatus select i.TcId).ToList();

                if (selectedId.Count == 0)
                {
                    MessageBox.Show("Hey! It seems there are no automated cases here to generate the playlist");
                    return;
                }

                var dlg = new SaveFileDialog
                {
                    FileName = "Playlist1",
                    DefaultExt = ".playlist",
                    Filter = "Playlist (.playlist)|*.playlist"
                };
               

                var result = dlg.ShowDialog();
              
                if (result == true)
                {
                    var fileLocation = dlg.FileName;
                    FilePath.Text = fileLocation;
                    Mouse.OverrideCursor = Cursors.Wait;
                    MtmInteraction.CreatePlaylist( selectedId, fileLocation);
                    
                    DataGetter.Diagnostic.AppendLine("Total time taken for playlist generation: " +
                                                     (MtmInteraction.AutomationPlaylistAddition +
                                                      MtmInteraction.AutomationMethodTime));
                    DataGetter.Diagnostic.AppendLine("Time taken for playlist data fetch " +
                                                     MtmInteraction.AutomationMethodTime);
                    DataGetter.Diagnostic.AppendLine("Time taken for playlist file generation: " +
                                                     MtmInteraction.AutomationPlaylistAddition);
                    DataGetter.Diagnostic.AppendLine("---------------------------------------------------");

                    MessageBox.Show("Playlist generated!");
                }              
            }
            catch (Exception ex)
            {
                MessageBox.Show("Sorry! Something went wrong. Please send the diagnostic log to us.");
                DataGetter.Diagnostic.AppendLine("Error occured in generating playlist."); 
                DataGetter.Diagnostic.AppendLine(ex.Message);
                DataGetter.Diagnostic.AppendLine("---------------------------------------------------");
                
            }
            finally
            {
                Mouse.OverrideCursor = null;
            }
        }

        private void testerAction_SelectionChanged_1(object sender, SelectionChangedEventArgs e)
        {
            if (TesterAction.SelectedItem.ToString()
                .Equals("Assign Tester", StringComparison.InvariantCultureIgnoreCase))
            {
                TesterName.Visibility = Visibility.Visible;
                LblActionParam.Visibility = Visibility.Visible;
                
                Resultoptions.Visibility = Visibility.Hidden;
                TestActionsoptions.Visibility = Visibility.Hidden;
                
                ActionButton.Content = "Assign";
            }
            else if (TesterAction.SelectedItem.ToString().Equals("Playlist", StringComparison.InvariantCultureIgnoreCase))
            {
                TesterName.Visibility = Visibility.Hidden;
                LblActionParam.Visibility = Visibility.Hidden;
               
                Resultoptions.Visibility = Visibility.Hidden;
                TestActionsoptions.Visibility = Visibility.Hidden;
              
                ActionButton.Content = "Generate";
            }
           
            else if (TesterAction.SelectedItem.ToString()
                .Equals("Result Update", StringComparison.InvariantCultureIgnoreCase))
            {
                TesterName.Visibility = Visibility.Hidden;
                LblActionParam.Visibility = Visibility.Visible ;
                LblActionParam.Text = "Outcome";
                Resultoptions.Visibility = Visibility.Visible;
                TestActionsoptions.Visibility = Visibility.Hidden;
                ActionButton.Content = "Update";
            }
            else if (TesterAction.SelectedItem.ToString()
                .Equals("Result History", StringComparison.InvariantCultureIgnoreCase))
            {
                TesterName.Visibility = Visibility.Hidden;
                LblActionParam.Visibility = Visibility.Hidden ;
                LblActionParam.Text = "Outcome";
                Resultoptions.Visibility = Visibility.Hidden ;
                TestActionsoptions.Visibility = Visibility.Hidden;
                ActionButton.Content = "Generate";
            }
        }

        private void ExportToExcel_Click(object sender, RoutedEventArgs e)
        {

            try
            {
                if (ResultDataGrid.Items.Count.Equals(0))
                {
                    MessageBox.Show("Nothing to export.Please generate a report", "OOPS!", MessageBoxButton.OK);
                    return;
                }
                if (TesterAction.SelectedItem.ToString()
                .Equals("Result History", StringComparison.InvariantCultureIgnoreCase))
                {
                    var expExlSum = new ExportToExcel<ResultHistorySummary >
                    {
                        DataToExport = (List<ResultHistorySummary>)ResultDataGrid.ItemsSource
                    };
                    expExlSum.GenerateExcel();
                }
                else
                {
                    var expExlSum = new ExportToExcel<QueryInterface>
                    {
                        DataToExport = (List<QueryInterface>) ResultDataGrid.ItemsSource
                    };
                    expExlSum.GenerateExcel();
                }
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.Diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
        }
    }
}
