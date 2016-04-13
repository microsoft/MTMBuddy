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

      

        public static StringBuilder diagnostic = new StringBuilder();

        private static readonly Stopwatch stp = new Stopwatch();

        public static List<resultsummary> GetResultSummaryList(int suiteId, string suiteName)
        {
            var resDetail = new ConcurrentBag<resultsummary>();
            stp.Restart();
            MTMInteraction.clearperfcounters();
            MTMInteraction.getsuiteresults(suiteId, resDetail, suiteName);
            stp.Stop();
            diagnostic.AppendLine("Fetch Suite data for " + suiteName + "-" + suiteId + ":    " +
                                  stp.Elapsed.TotalSeconds);
            diagnostic.AppendLine("Outcometime:   " + MTMInteraction.OutcomeTime);
            diagnostic.AppendLine("Priority Time:   " + MTMInteraction.PriorityTime);
            diagnostic.AppendLine("Tester Time:   " + MTMInteraction.TesterTime);
            diagnostic.AppendLine("Title Time:   " + MTMInteraction.Titletime);
            diagnostic.AppendLine("Initialize Time:   " + MTMInteraction.Initialize);
            diagnostic.AppendLine("Test Case Id Time:   " + MTMInteraction.tcidtime);
            diagnostic.AppendLine("Automation Status Time:   " + MTMInteraction.AutomationTime);
            diagnostic.AppendLine("Plan:   " + MTMInteraction.planName);
            diagnostic.AppendLine("Total count:  " + resDetail.Count() + "    Blocked Count:   " +
                                  resDetail.Where(
                                      l => l.Outcome.Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                                      .Count());
            diagnostic.AppendLine("Automation Test Name Time:" + MTMInteraction.AutomationTestNameFetchTime);
            diagnostic.AppendLine("---------------------------------------------------");
           
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