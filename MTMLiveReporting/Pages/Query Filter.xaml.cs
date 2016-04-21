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
    public partial class QueryFilter : UserControl
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
                var actions = new List<string>();
                actions.Add("Assign Tester");
                actions.Add("Playlist");
                actions.Add("Result Update");

                TesterAction.ItemsSource = actions;


                TesterAction.SelectedIndex = 0;
                var automationStatus = new List<string>();
                automationStatus.Add("Both");
                automationStatus.Add("Automated");
                automationStatus.Add("Manual");
                CmbAutomationStaus.ItemsSource = automationStatus;
                var priorities = new List<string>();
                priorities.Add("All");
                priorities.Add("1");
                priorities.Add("2");
                priorities.Add("3");
                priorities.Add("4");
                CmbPriority.ItemsSource = priorities;
                var outcomes = new List<string>();
                outcomes.Add("All");
                outcomes.Add("Active");
                outcomes.Add("Blocked");
                outcomes.Add("Failed");
                outcomes.Add("Passed");
                outcomes.Add("In progress");
                CmbOutcome.ItemsSource = outcomes;

                var validOutcomes = outcomes;

                validOutcomes.Remove("All");
                Resultoptions.ItemsSource = validOutcomes;
               
                
                
              
                MtmInteraction.Getwpfsuitetree(int.Parse(ConfigurationManager.AppSettings["TestPlanID"]), TvMtm, true);
                var tester = new List<string>();
                tester = MtmInteraction.GetTester();
               
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


        public void getlist_Click(object sender, RoutedEventArgs e)
        {
            try {
                var plans = new Dictionary<int, string>();
                plans = MtmInteraction.GetPlanIdsDictionary();

                if (int.TryParse(TxtPlanId.Text, out _planId))
                {
                    TvMtm.Items.Clear();
                    MtmInteraction.Getwpfsuitetree(_planId, TvMtm, true);
                    var tester = new List<string>();
                    tester = MtmInteraction.GetTester();
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

      
        public void getFilter_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedItem = (TreeViewItem) TvMtm.SelectedItem;
                if (_suiteid != int.Parse(selectedItem.Tag.ToString()))
                {
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
                }
                var filters = new List<MtmInteraction.Filter>();

                var moduleinclusion = ChkModuleInclusion.IsChecked.Value;
                var modulefilter = Txtmodulefilter.Text;
                var testerinclusion = ChkTesterInclusion.IsChecked.Value;
                var testerfilter = Txttesterfilter.Text;
                var automationstatus = CmbAutomationStaus.SelectedItem.ToString();
                if (!string.IsNullOrEmpty(TxtTitlefilter.Text))
                {
                    
                    var titleFilter = new MtmInteraction.Filter();
                    if (ChkTitleInclusion.IsChecked.Value)
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
                   
                    var titleFilter = new MtmInteraction.Filter();
                    
                    titleFilter.Name = "Priority";
                    titleFilter.Value = CmbPriority.SelectedValue.ToString();
                    filters.Add(titleFilter);
                }
                if (!CmbOutcome.SelectedValue.ToString().Equals("All", StringComparison.InvariantCultureIgnoreCase))
                {
                    
                    var titleFilter = new MtmInteraction.Filter();

                    titleFilter.Name = "Outcome";
                    titleFilter.Value = CmbOutcome.SelectedValue.ToString();
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
                ;
            }
        }

        private void Assign_Tester()
        {
            try {
                var name = TesterName.SelectedItem.ToString();
                var selectedId = new List<int>();
                foreach (var i in _queryResult)
                {
                    if (i.Selected)
                    {
                        selectedId.Add(i.TcId);
                    }
                }
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
                DataGetter.Diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
        }
        

        private void Hyperlink_Click(object sender, RoutedEventArgs e)
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
                DataGetter.Diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
}

        private void Updateresults()
        {
            try
            {
                //string name = testerName.SelectedItem.ToString();
                Mouse.OverrideCursor = Cursors.Wait;
                var errflag = false;
                foreach (var i in _queryResult)
                {
                    if (i.Selected)
                    {
                        try
                        {
                            MTMIntegration.MtmInteraction.UpdateResult(i.TcId.ToString(), "Updated from MTM Buddy",
                                i.SuiteName.Substring(i.SuiteName.LastIndexOf("\\") + 1), false,
                                Resultoptions.SelectedValue.ToString());
                        }
                        catch (Exception ex)
                        {
                            DataGetter.Diagnostic.AppendLine("Error occured in migrating testcase " + i.TcId);
                            DataGetter.Diagnostic.AppendLine(ex.Message);
                            DataGetter.Diagnostic.AppendLine("---------------------------------------------------");
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
               
                DataGetter.Diagnostic.AppendLine("Error occured in migrating testcases.");
                DataGetter.Diagnostic.AppendLine(ex.Message);
                DataGetter.Diagnostic.AppendLine("---------------------------------------------------");
                
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

           
            Mouse.OverrideCursor = Cursors.Arrow;
            }
            catch (Exception exp)
            {
                MessageBox.Show(
                    "It seems something has gone wrong. Please send us the below information so that we can resolve the issue." +
                    Environment.NewLine + exp.Message, "OOPS!", MessageBoxButton.OK, MessageBoxImage.Warning);
                DataGetter.Diagnostic.AppendLine("Error while exporting to Excel: " + exp.Message);
            }
        }

        private void Create_Playlist()
        {
            try
            {
                //string name = testerName.SelectedItem.ToString();
                var selectedId = new List<int>();
                foreach (var i in _queryResult)
                {
                    if (i.Selected && i.AutomationStatus)
                    {
                        selectedId.Add(i.TcId);
                    }
                }

                if (selectedId.Count == 0)
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
                    FilePath.Text = fileLocation;
                    Mouse.OverrideCursor = Cursors.Wait;
                    MtmInteraction.CreatePlaylist(_suiteid, selectedId, fileLocation);
                    
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
                //filePath.Visibility =  System.Windows.Visibility.Hidden;
                ActionButton.Content = "Assign";
            }
            else if (TesterAction.SelectedItem.ToString().Equals("Playlist", StringComparison.InvariantCultureIgnoreCase))
            {
                TesterName.Visibility = Visibility.Hidden;
                LblActionParam.Visibility = Visibility.Hidden;
               
                Resultoptions.Visibility = Visibility.Hidden;
                TestActionsoptions.Visibility = Visibility.Hidden;
                //filePath.Visibility = System.Windows.Visibility.Visible;
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
                expExlSum.DataToExport = (List<QueryInterface>)ResultDataGrid.ItemsSource;
                expExlSum.GenerateExcel();
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
