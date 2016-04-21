//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Collections.Generic;
using System.Linq;
using MTMIntegration;

namespace ReportingLayer
{
    public class AutomationReport
    {
        public string Priority { get; set; }

        public int Automated { get; set; }
        public int AutomationNew { get; set; }
        public int AutomatedRegression { get; set; }

        public int Manual { get; set; }
        public int ManualNew { get; set; }

        public int ManualRegression { get; set; }


        public float AutomationRatio { get; set; }

        public float RegressionAutoRatio { get; set; }

        public float NewAutoRatio { get; set; }
        public int Total { get; set; }


        public static List<AutomationReport> Generate(List<ResultSummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();

            var reportList = new List<AutomationReport>();
            var rd = new List<ResultSummary>();
            var rd1 = new List<ResultSummary>();
            var filtereddata = Utilities.FilterData(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            filtereddata = filtereddata.GroupBy(p => p.TCID).Select(g => g.First()).ToList();
            for (var i = 1; i <= 3; i++)
            {
                var sr = new AutomationReport();

                rd = filtereddata.Where(l => l.Priority.Equals(i)).ToList();
                sr.Total = rd.Count;
                sr.Priority = "P" + i;

                sr.Automated = rd.Where(l => l.AutomationStatus).Count();
                sr.Manual = rd.Where(l => !l.AutomationStatus).Count();

                if (sr.Total > 0)
                {
                    sr.AutomationRatio = (float) Math.Round((float) sr.Automated/(sr.Automated + sr.Manual)*100, 2);


                    rd1 = Utilities.FilterData(rd, "Regression", true, tester, testerinclusion, automationstatus);
                    sr.ManualRegression = rd1.Where(l => !l.AutomationStatus).Count();
                    sr.AutomatedRegression = rd1.Where(l => l.AutomationStatus).Count();
                    if (rd1.Count > 0)
                        sr.RegressionAutoRatio =
                            (float)
                                Math.Round(
                                    (float) sr.AutomatedRegression/(sr.AutomatedRegression + sr.ManualRegression)*100, 2);
                    rd1 = Utilities.FilterData(rd, "Regression", false, tester, testerinclusion, automationstatus);
                    sr.ManualNew = rd1.Where(l => !l.AutomationStatus).Count();
                    sr.AutomationNew = rd1.Where(l => l.AutomationStatus).Count();
                    if (rd1.Count > 0)
                        sr.NewAutoRatio =
                            (float) Math.Round((float) sr.AutomationNew/(sr.AutomationNew + sr.ManualNew)*100, 2);
                }

                reportList.Add(sr);
            }

            var sr1 = new AutomationReport();
            sr1.Priority = "Total";
            sr1.Total = filtereddata.Count;
            sr1.Automated = filtereddata.Where(l => l.AutomationStatus).Count();
            sr1.Manual = filtereddata.Where(l => !l.AutomationStatus).Count();

            if (sr1.Total > 0)
            {
                sr1.AutomationRatio = (float) Math.Round((float) sr1.Automated/(sr1.Automated + sr1.Manual)*100, 2);
                rd1 = Utilities.FilterData(filtereddata, "Regression", true, tester, testerinclusion, automationstatus);
                sr1.ManualRegression = rd1.Where(l => !l.AutomationStatus).Count();
                sr1.AutomatedRegression = rd1.Where(l => l.AutomationStatus).Count();
                if (rd1.Count > 0)
                    sr1.RegressionAutoRatio =
                        (float)
                            Math.Round(
                                (float) sr1.AutomatedRegression/(sr1.AutomatedRegression + sr1.ManualRegression)*100, 2);
                rd1 = Utilities.FilterData(filtereddata, "Regression", false, tester, testerinclusion, automationstatus);
                sr1.ManualNew = rd1.Where(l => !l.AutomationStatus).Count();
                sr1.AutomationNew = rd1.Where(l => l.AutomationStatus).Count();
                if (rd1.Count > 0)
                    sr1.NewAutoRatio =
                        (float) Math.Round((float) sr1.AutomationNew/(sr1.AutomationNew + sr1.ManualNew)*100, 2);
            }
            reportList.Add(sr1);

            return reportList;
        }
    }
}