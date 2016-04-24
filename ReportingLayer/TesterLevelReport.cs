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
    public class TesterLevelReport
    {
        public string Tester { get; set; }
        public int Total { get; set; }
        public int Active { get; set; }
        public int Blocked { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }

        public float ProgressRate { get; set; }
        public int P1Total { get; set; }
        public int P1Active { get; set; }
        public int P1Blocked { get; set; }
        public int P1Passed { get; set; }
        public int P1Failed { get; set; }

        public float P1ProgressRate { get; set; }
        public int P2Total { get; set; }
        public int P2Active { get; set; }
        public int P2Blocked { get; set; }
        public int P2Passed { get; set; }
        public int P2Failed { get; set; }

        public float P2ProgressRate { get; set; }
        public int P3Total { get; set; }
        public int P3Active { get; set; }
        public int P3Blocked { get; set; }
        public int P3Passed { get; set; }
        public int P3Failed { get; set; }

        public float P3ProgressRate { get; set; }


        public static List<TesterLevelReport> Generate(List<ResultSummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();

            new List<SummaryReport>();
            var modreplist = new List<TesterLevelReport>();
            var filtereddata = Utilities.FilterData(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            //get distinct modules
            var testers = filtereddata.Select(p => p.Tester).Distinct().ToList();


            foreach (var reptester in testers)
            {
                var modrep = new TesterLevelReport {Tester = reptester};
                var moduledata = Utilities.FilterData(rawData, module, moduleinclusion, reptester, true,
                    automationstatus);
                //Overall data
                modrep.Total = moduledata.Count;
                modrep.Active = moduledata.Count(l => l.Outcome.Equals("Active"));
                modrep.Passed = moduledata.Count(l => l.Outcome.Equals("Passed"));
                modrep.Failed = moduledata.Count(l => l.Outcome.Equals("Failed"));
                modrep.Blocked = moduledata.Count(l => l.Outcome.Equals("Blocked"));
                if (modrep.Total > 0)
                {
                    modrep.ProgressRate =
                        (float) Math.Round((float) (modrep.Passed + modrep.Failed)/modrep.Total*100, 2);
                }
                //P1
                var rd = moduledata.Where(l => l.Priority.Equals(1)).ToList();
                modrep.P1Total = rd.Count;
                modrep.P1Active = rd.Count(l => l.Outcome.Equals("Active"));
                modrep.P1Passed = rd.Count(l => l.Outcome.Equals("Passed"));
                modrep.P1Failed = rd.Count(l => l.Outcome.Equals("Failed"));
                modrep.P1Blocked = rd.Count(l => l.Outcome.Equals("Blocked"));
                if (modrep.P1Total > 0)
                {
                    modrep.P1ProgressRate =
                        (float) Math.Round((float) (modrep.P1Passed + modrep.P1Failed)/modrep.P1Total*100, 2);
                }

                //P2
                rd = moduledata.Where(l => l.Priority.Equals(2)).ToList();
                modrep.P2Total = rd.Count;
                modrep.P2Active = rd.Count(l => l.Outcome.Equals("Active"));
                modrep.P2Passed = rd.Count(l => l.Outcome.Equals("Passed"));
                modrep.P2Failed = rd.Count(l => l.Outcome.Equals("Failed"));
                modrep.P2Blocked = rd.Count(l => l.Outcome.Equals("Blocked"));
                if (modrep.P2Total > 0)
                {
                    modrep.P2ProgressRate =
                        (float) Math.Round((float) (modrep.P2Passed + modrep.P2Failed)/modrep.P2Total*100, 2);
                }
                //P3
                rd = moduledata.Where(l => l.Priority.Equals(3)).ToList();
                modrep.P3Total = rd.Count;
                modrep.P3Active = rd.Count(l => l.Outcome.Equals("Active"));
                modrep.P3Passed = rd.Count(l => l.Outcome.Equals("Passed"));
                modrep.P3Failed = rd.Count(l => l.Outcome.Equals("Failed"));
                modrep.P3Blocked = rd.Count(l => l.Outcome.Equals("Blocked"));
                if (modrep.P3Total > 0)
                {
                    modrep.P3ProgressRate =
                        (float) Math.Round((float) (modrep.P3Passed + modrep.P3Failed)/modrep.P3Total*100, 2);
                }
                modreplist.Add(modrep);
            }


            return modreplist;
        }
    }
}