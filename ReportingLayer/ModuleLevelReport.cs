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
    public class ModuleLevelReport
    {
        public string Module { get; set; }
        public int Total { get; set; }

        public float AutomationRatio { get; set; }
        public int Active { get; set; }
        public int Blocked { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public float FailRate { get; set; }
        public float PassRate { get; set; }
        public float ProgressRate { get; set; }
        public int P1Total { get; set; }
        public int P1Active { get; set; }
        public int P1Blocked { get; set; }
        public int P1Passed { get; set; }
        public int P1Failed { get; set; }
        public float P1FailRate { get; set; }
        public float P1PassRate { get; set; }
        public float P1ProgressRate { get; set; }
        public int P2Total { get; set; }
        public int P2Active { get; set; }
        public int P2Blocked { get; set; }
        public int P2Passed { get; set; }
        public int P2Failed { get; set; }
        public float P2FailRate { get; set; }
        public float P2PassRate { get; set; }
        public float P2ProgressRate { get; set; }
        public int P3Total { get; set; }
        public int P3Active { get; set; }
        public int P3Blocked { get; set; }
        public int P3Passed { get; set; }
        public int P3Failed { get; set; }
        public float P3FailRate { get; set; }
        public float P3PassRate { get; set; }
        public float P3ProgressRate { get; set; }


        public static List<ModuleLevelReport> Generate(List<ResultSummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();


            var modreplist = new List<ModuleLevelReport>();
            var filtereddata = Utilities.FilterData(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            //get distinct modules
            var modules = filtereddata.Select(p => p.SuiteName).Distinct().ToList();


            foreach (var repmodule in modules)
            {
                var modrep = new ModuleLevelReport {Module = repmodule};
                var moduledata = Utilities.FilterData(rawData, repmodule, true, tester, testerinclusion,
                    automationstatus,true);
                //Overall data
                modrep.Total = moduledata.Count;
                modrep.AutomationRatio =
                    (float) Math.Round((float) moduledata.Count(l => l.AutomationStatus)*100/modrep.Total, 2);
                modrep.Active = moduledata.Count(l => l.Outcome.Equals("Active"));
                modrep.Passed = moduledata.Count(l => l.Outcome.Equals("Passed"));
                modrep.Failed = moduledata.Count(l => l.Outcome.Equals("Failed"));
                modrep.Blocked = moduledata.Count(l => l.Outcome.Equals("Blocked"));
                if (modrep.Passed + modrep.Failed > 0)
                {
                    modrep.FailRate = (float) Math.Round((float) modrep.Failed/(modrep.Passed + modrep.Failed)*100, 2);
                    modrep.PassRate = (float) Math.Round((float) modrep.Passed/(modrep.Passed + modrep.Failed)*100, 2);
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
                if (modrep.P1Passed + modrep.P1Failed > 0)
                {
                    modrep.P1FailRate =
                        (float) Math.Round((float) modrep.P1Failed/(modrep.P1Passed + modrep.P1Failed)*100, 2);
                    modrep.P1PassRate =
                        (float) Math.Round((float) modrep.P1Passed/(modrep.P1Passed + modrep.P1Failed)*100, 2);
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
                if (modrep.P2Passed + modrep.P2Failed > 0)
                {
                    modrep.P2FailRate =
                        (float) Math.Round((float) modrep.P2Failed/(modrep.P2Passed + modrep.P2Failed)*100, 2);
                    modrep.P2PassRate =
                        (float) Math.Round((float) modrep.P2Passed/(modrep.P2Passed + modrep.P2Failed)*100, 2);
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
                if (modrep.P3Passed + modrep.P3Failed > 0)
                {
                    modrep.P3FailRate =
                        (float) Math.Round((float) modrep.P3Failed/(modrep.P3Passed + modrep.P3Failed)*100, 2);
                    modrep.P3PassRate =
                        (float) Math.Round((float) modrep.P3Passed/(modrep.P3Passed + modrep.P3Failed)*100, 2);
                    modrep.P3ProgressRate =
                        (float) Math.Round((float) (modrep.P3Passed + modrep.P3Failed)/modrep.P3Total*100, 2);
                }
                modreplist.Add(modrep);
            }


            return modreplist;
        }
    }
}