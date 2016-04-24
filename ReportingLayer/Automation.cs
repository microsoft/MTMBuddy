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
            List<ResultSummary> rd1;
            var filtereddata = Utilities.FilterData(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            filtereddata = filtereddata.GroupBy(p => p.TcId).Select(g => g.First()).ToList();
            for (var i = 1; i <= 4; i++)
            {
                var sr = new AutomationReport();

                var rd = filtereddata.Where(l => l.Priority.Equals(i)).ToList();
                sr.Total = rd.Count;
                sr.Priority = "P" + i;

                sr.Automated = rd.Count(l => l.AutomationStatus);
                sr.Manual = rd.Count(l => !l.AutomationStatus);

                if (sr.Total > 0)
                {
                    sr.AutomationRatio = (float) Math.Round((float) sr.Automated/(sr.Automated + sr.Manual)*100, 2);


                    rd1 = Utilities.FilterData(rd, "Regression", true, tester, testerinclusion, automationstatus);
                    sr.ManualRegression = rd1.Count(l => !l.AutomationStatus);
                    sr.AutomatedRegression = rd1.Count(l => l.AutomationStatus);
                    if (rd1.Count > 0)
                        sr.RegressionAutoRatio =
                            (float)
                                Math.Round(
                                    (float) sr.AutomatedRegression/(sr.AutomatedRegression + sr.ManualRegression)*100, 2);
                    rd1 = Utilities.FilterData(rd, "Regression", false, tester, testerinclusion, automationstatus);
                    sr.ManualNew = rd1.Count(l => !l.AutomationStatus);
                    sr.AutomationNew = rd1.Count(l => l.AutomationStatus);
                    if (rd1.Count > 0)
                        sr.NewAutoRatio =
                            (float) Math.Round((float) sr.AutomationNew/(sr.AutomationNew + sr.ManualNew)*100, 2);
                }

                reportList.Add(sr);
            }

            var sr1 = new AutomationReport
            {
                Priority = "Total",
                Total = filtereddata.Count,
                Automated = filtereddata.Count(l => l.AutomationStatus),
                Manual = filtereddata.Count(l => !l.AutomationStatus)
            };

            if (sr1.Total > 0)
            {
                sr1.AutomationRatio = (float) Math.Round((float) sr1.Automated/(sr1.Automated + sr1.Manual)*100, 2);
                rd1 = Utilities.FilterData(filtereddata, "Regression", true, tester, testerinclusion, automationstatus);
                sr1.ManualRegression = rd1.Count(l => !l.AutomationStatus);
                sr1.AutomatedRegression = rd1.Count(l => l.AutomationStatus);
                if (rd1.Count > 0)
                    sr1.RegressionAutoRatio =
                        (float)
                            Math.Round(
                                (float) sr1.AutomatedRegression/(sr1.AutomatedRegression + sr1.ManualRegression)*100, 2);
                rd1 = Utilities.FilterData(filtereddata, "Regression", false, tester, testerinclusion, automationstatus);
                sr1.ManualNew = rd1.Count(l => !l.AutomationStatus);
                sr1.AutomationNew = rd1.Count(l => l.AutomationStatus);
                if (rd1.Count > 0)
                    sr1.NewAutoRatio =
                        (float) Math.Round((float) sr1.AutomationNew/(sr1.AutomationNew + sr1.ManualNew)*100, 2);
            }
            reportList.Add(sr1);

            return reportList;
        }
    }
}