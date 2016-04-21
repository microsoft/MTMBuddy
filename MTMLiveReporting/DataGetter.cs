//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Security.Principal;
using System.Text;
using MTMIntegration;


namespace MTMLiveReporting
{
    public static class DataGetter
    {
       

       

        public static bool FirstRun = false;

      

        public static StringBuilder Diagnostic = new StringBuilder();

        private static readonly Stopwatch Stp = new Stopwatch();

        public static List<ResultSummary> GetResultSummaryList(int suiteId, string suiteName)
        {
            var resDetail = new ConcurrentBag<ResultSummary>();
            Stp.Restart();
            MtmInteraction.Clear_performance_counters();
            MtmInteraction.Getsuiteresults(suiteId, resDetail, suiteName);
            Stp.Stop();
            Diagnostic.AppendLine("Fetch Suite data for " + suiteName + "-" + suiteId + ":    " +
                                  Stp.Elapsed.TotalSeconds);
            Diagnostic.AppendLine("Outcometime:   " + MtmInteraction.OutcomeTime);
            Diagnostic.AppendLine("Priority Time:   " + MtmInteraction.PriorityTime);
            Diagnostic.AppendLine("Tester Time:   " + MtmInteraction.TesterTime);
            Diagnostic.AppendLine("Title Time:   " + MtmInteraction.Titletime);
            Diagnostic.AppendLine("Initialize Time:   " + MtmInteraction.Initialize);
            Diagnostic.AppendLine("Test Case Id Time:   " + MtmInteraction.TCidtime);
            Diagnostic.AppendLine("Automation Status Time:   " + MtmInteraction.AutomationTime);
            Diagnostic.AppendLine("Plan:   " + MtmInteraction.SelectedPlanName);
            Diagnostic.AppendLine("Total count:  " + resDetail.Count() + "    Blocked Count:   " +
                                  resDetail.Where(
                                      l => l.Outcome.Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                                      .Count());
         
            Diagnostic.AppendLine("---------------------------------------------------");
           
            return resDetail.ToList();
        }

        /// <summary>
        ///     This method will update the calling applications config file for the given name with given valu and uppdate to the
        ///     disk
        /// </summary>
        /// <param name="name">
        ///     Name/key of the configuration setting
        /// </param>
        /// <param name="value">
        ///     New value for the configuration setting
        /// </param>
        public static void SaveConfig(string name, string value)
        {
            var cnfg = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);
            cnfg.AppSettings.Settings[name].Value = value;
            cnfg.Save(ConfigurationSaveMode.Modified, true);
        }

      
    }
}