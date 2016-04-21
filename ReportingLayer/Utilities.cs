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
    public static class Utilities
    {
        public static List<ResultSummary> GetIssueList(List<ResultSummary> rawData)
        {
            var issuelist = new List<ResultSummary>();
            issuelist.Clear();
            issuelist.AddRange(rawData.Where(p => p.Priority < 1).ToList());
            issuelist.AddRange(rawData.Where(p => p.Priority > 3).ToList());
            issuelist.AddRange(
                rawData.Where(
                    p =>
                        !(p.Outcome.Equals("Active", StringComparison.InvariantCultureIgnoreCase) ||
                          p.Outcome.Equals("Passed", StringComparison.InvariantCultureIgnoreCase) ||
                          p.Outcome.Equals("Failed", StringComparison.InvariantCultureIgnoreCase) ||
                          p.Outcome.Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))).ToList());
            return issuelist;
        }

        public static List<ResultSummary> FilterData(List<ResultSummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "BOTH")
        {
            var filtereddata = new List<ResultSummary>();

            //filter data


            filtereddata.Clear();
            var inclusionmodules = module.Split(',');

            foreach (var resitem in rawData)
            {
                var foundflag = false;
                foreach (var inclusionmodule in inclusionmodules)
                {
                    if (resitem.SuiteName.ToUpperInvariant().Contains(inclusionmodule.ToUpperInvariant()))
                    {
                        foundflag = true;
                        break;
                    }
                }
                if (!string.IsNullOrEmpty(module))
                {
                    if (!foundflag && !moduleinclusion)
                        filtereddata.Add(resitem);
                    else if (foundflag && moduleinclusion)
                        filtereddata.Add(resitem);
                }
                else
                    filtereddata.Add(resitem);
            }


            if (!string.IsNullOrEmpty(tester))
            {
                if (testerinclusion)
                {
                    filtereddata =
                        filtereddata.Where(l => tester.ToUpperInvariant().Contains(l.Tester.ToUpperInvariant()))
                            .ToList();
                }
                else
                {
                    filtereddata =
                        filtereddata.Where(l => !tester.ToUpperInvariant().Contains(l.Tester.ToUpperInvariant()))
                            .ToList();
                }
            }

            switch (automationstatus.ToUpperInvariant())
            {
                case "AUTOMATED":
                    filtereddata = filtereddata.Where(l => l.AutomationStatus).ToList();
                    break;
                case "MANUAL":
                    filtereddata = filtereddata.Where(l => !l.AutomationStatus).ToList();
                    break;
            }

            return filtereddata;
        }

      
    }
}