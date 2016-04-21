//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System.Collections.Generic;
using System.Linq;
using MTMIntegration;

namespace ReportingLayer
{
    public class OneLineSummary
    {
        public int P1Total { get; set; }
        public int P2Total { get; set; }
        public int P3Total { get; set; }
        public int Total { get; set; }

        public int P1Passed { get; set; }
        public int P2Passed { get; set; }
        public int P3Passed { get; set; }
        public int Passed { get; set; }

        public int P1Failed { get; set; }
        public int P2Failed { get; set; }
        public int P3Failed { get; set; }
        public int Failed { get; set; }

        public int P1Blocked { get; set; }
        public int P2Blocked { get; set; }
        public int P3Blocked { get; set; }
        public int Blocked { get; set; }

        public int P1Active { get; set; }
        public int P2Active { get; set; }
        public int P3Active { get; set; }
        public int Active { get; set; }


        public static List<OneLineSummary> Generate(List<resultsummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();

            var reportList = new List<SummaryReport>();
            var modreplist = new List<OneLineSummary>();
            var rd = new List<resultsummary>();
            var filtereddata = Utilities.filterdata(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);


            var modrep = new OneLineSummary();

            //Overall data
            modrep.Total = filtereddata.Count;
            modrep.Active = filtereddata.Where(l => l.Outcome.Equals("Active")).Count();
            modrep.Passed = filtereddata.Where(l => l.Outcome.Equals("Passed")).Count();
            modrep.Failed = filtereddata.Where(l => l.Outcome.Equals("Failed")).Count();
            modrep.Blocked = filtereddata.Where(l => l.Outcome.Equals("Blocked")).Count();

            //P1
            rd = filtereddata.Where(l => l.Priority.Equals(1)).ToList();
            modrep.P1Total = rd.Count;
            modrep.P1Active = rd.Where(l => l.Outcome.Equals("Active")).Count();
            modrep.P1Passed = rd.Where(l => l.Outcome.Equals("Passed")).Count();
            modrep.P1Failed = rd.Where(l => l.Outcome.Equals("Failed")).Count();
            modrep.P1Blocked = rd.Where(l => l.Outcome.Equals("Blocked")).Count();

            //P2
            rd = filtereddata.Where(l => l.Priority.Equals(2)).ToList();
            modrep.P2Total = rd.Count;
            modrep.P2Active = rd.Where(l => l.Outcome.Equals("Active")).Count();
            modrep.P2Passed = rd.Where(l => l.Outcome.Equals("Passed")).Count();
            modrep.P2Failed = rd.Where(l => l.Outcome.Equals("Failed")).Count();
            modrep.P2Blocked = rd.Where(l => l.Outcome.Equals("Blocked")).Count();

            //P3
            rd = filtereddata.Where(l => l.Priority.Equals(3)).ToList();
            modrep.P3Total = rd.Count;
            modrep.P3Active = rd.Where(l => l.Outcome.Equals("Active")).Count();
            modrep.P3Passed = rd.Where(l => l.Outcome.Equals("Passed")).Count();
            modrep.P3Failed = rd.Where(l => l.Outcome.Equals("Failed")).Count();
            modrep.P3Blocked = rd.Where(l => l.Outcome.Equals("Blocked")).Count();

            modreplist.Add(modrep);

            rd = filtereddata.Where(p => p.Priority != 3 && p.Priority != 1 && p.Priority != 2).ToList();


            return modreplist;
        }
    }
}