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
    public class TesterLevelReportGroup
    {
        public string Tester { get; set; }
        public string Priority { get; set; }
        public int Total { get; set; }
        public int Active { get; set; }
        public int Blocked { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }

        public float ProgressRate { get; set; }

        public static List<TesterLevelReportGroup> Generate(List<resultsummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();

            var modreplist = new List<TesterLevelReportGroup>();
            var rd = new List<resultsummary>();
            var filtereddata = Utilities.filterdata(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            //get distinct modules
            var modules = filtereddata.Select(p => p.Tester).Distinct().ToList();


            foreach (var repmodule in modules)
            {
                var modrep = new TesterLevelReportGroup();
                modrep.Tester = repmodule;
                modrep.Priority = "Total";
                var moduledata = Utilities.filterdata(rawData, module, moduleinclusion, repmodule, true,
                    automationstatus);
                //Overall data
                modrep.Total = moduledata.Count;
                modrep.Active = moduledata.Where(l => l.Outcome.Equals("Active")).Count();
                modrep.Passed = moduledata.Where(l => l.Outcome.Equals("Passed")).Count();
                modrep.Failed = moduledata.Where(l => l.Outcome.Equals("Failed")).Count();
                modrep.Blocked = moduledata.Where(l => l.Outcome.Equals("Blocked")).Count();
                if (modrep.Total > 0)
                {
                    modrep.ProgressRate =
                        (float) Math.Round((float) (modrep.Passed + modrep.Failed)/modrep.Total*100, 2);
                }
                modreplist.Add(modrep);
                //P1
                for (var i = 1; i <= 3; i++)
                {
                    var modrepP1 = new TesterLevelReportGroup();
                    modrepP1.Priority = "   P" + i;
                    modrepP1.Tester = repmodule;

                    rd = moduledata.Where(l => l.Priority.Equals(i)).ToList();
                    modrepP1.Total = rd.Count;
                    modrepP1.Active = rd.Where(l => l.Outcome.Equals("Active")).Count();
                    modrepP1.Passed = rd.Where(l => l.Outcome.Equals("Passed")).Count();
                    modrepP1.Failed = rd.Where(l => l.Outcome.Equals("Failed")).Count();
                    modrepP1.Blocked = rd.Where(l => l.Outcome.Equals("Blocked")).Count();
                    if (modrepP1.Total > 0)
                    {
                        modrepP1.ProgressRate =
                            (float) Math.Round((float) (modrepP1.Passed + modrepP1.Failed)/modrepP1.Total*100, 2);
                    }
                    modreplist.Add(modrepP1);
                }
            }


            return modreplist;
        }
    }
}