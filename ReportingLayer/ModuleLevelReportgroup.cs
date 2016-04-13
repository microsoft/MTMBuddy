using System;
using System.Collections.Generic;
using System.Linq;
using MTMIntegration;

namespace ReportingLayer
{
    public class ModuleLevelReportGroup
    {
        public string Module { get; set; }
        public string Priority { get; set; }
        public int Total { get; set; }
        public int Active { get; set; }
        public int Blocked { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public float FailRate { get; set; }
        public float PassRate { get; set; }
        public float ProgressRate { get; set; }

        public static List<ModuleLevelReportGroup> Generate(List<resultsummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();

            var modreplist = new List<ModuleLevelReportGroup>();
            var rd = new List<resultsummary>();
            var filtereddata = Utilities.filterdata(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            //get distinct modules
            var modules = filtereddata.Select(p => p.SuiteName).Distinct().ToList();


            foreach (var repmodule in modules)
            {
                var modrep = new ModuleLevelReportGroup();
                modrep.Module = repmodule;
                modrep.Priority = "Total";
                var moduledata = Utilities.filterdata(rawData, repmodule, true, tester, testerinclusion,
                    automationstatus);
                //Overall data
                modrep.Total = moduledata.Count;
                modrep.Active = moduledata.Where(l => l.Outcome.Equals("Active")).Count();
                modrep.Passed = moduledata.Where(l => l.Outcome.Equals("Passed")).Count();
                modrep.Failed = moduledata.Where(l => l.Outcome.Equals("Failed")).Count();
                modrep.Blocked = moduledata.Where(l => l.Outcome.Equals("Blocked")).Count();
                if (modrep.Passed + modrep.Failed > 0)
                {
                    modrep.FailRate = (float) Math.Round((float) modrep.Failed/(modrep.Passed + modrep.Failed)*100, 2);
                    modrep.PassRate = (float) Math.Round((float) modrep.Passed/(modrep.Passed + modrep.Failed)*100, 2);
                    modrep.ProgressRate =
                        (float) Math.Round((float) (modrep.Passed + modrep.Failed)/modrep.Total*100, 2);
                }
                modreplist.Add(modrep);
                //P1
                for (var i = 1; i <= 3; i++)
                {
                    var modrepP1 = new ModuleLevelReportGroup();
                    modrepP1.Priority = "   P" + i;
                    modrepP1.Module = repmodule;

                    rd = moduledata.Where(l => l.Priority.Equals(i)).ToList();
                    modrepP1.Total = rd.Count;
                    modrepP1.Active = rd.Where(l => l.Outcome.Equals("Active")).Count();
                    modrepP1.Passed = rd.Where(l => l.Outcome.Equals("Passed")).Count();
                    modrepP1.Failed = rd.Where(l => l.Outcome.Equals("Failed")).Count();
                    modrepP1.Blocked = rd.Where(l => l.Outcome.Equals("Blocked")).Count();
                    if (modrepP1.Passed + modrepP1.Failed > 0)
                    {
                        modrepP1.FailRate =
                            (float) Math.Round((float) modrepP1.Failed/(modrepP1.Passed + modrepP1.Failed)*100, 2);
                        modrepP1.PassRate =
                            (float) Math.Round((float) modrepP1.Passed/(modrepP1.Passed + modrepP1.Failed)*100, 2);
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