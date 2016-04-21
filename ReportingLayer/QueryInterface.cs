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
    public class QueryInterface

    {
        public bool Selected { get; set; }
        public int Priority { get; set; }

        public string Outcome { get; set; }
        public int TcId { get; set; }

        public string Title { get; set; }

        public bool AutomationStatus { get; set; }

        public string Tester { get; set; }

        public string SuiteName { get; set; }


        public static List<QueryInterface> Generate(List<ResultSummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both", List<MtmInteraction.Filter> filterCriteria = null)
        {
            var reportList = new List<QueryInterface>();
            var rd = new List<ResultSummary>();
            var filtereddata = Utilities.FilterData(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            var queryResult = filtereddata;

            foreach (var filter in filterCriteria)
            {
                switch (filter.Name)
                {
                    case "Priority":
                        var priority = int.Parse(filter.Value);
                        queryResult = queryResult.Where(p => p.Priority == priority).ToList();

                        break;
                    case "Title":
                        if (filter.Op.Equals("Contains", StringComparison.InvariantCultureIgnoreCase))
                        {
                            queryResult =
                                queryResult.Where(
                                    p => p.Title.ToLowerInvariant().Contains(filter.Value.ToLowerInvariant())).ToList();
                        }
                        else if (filter.Op.Equals("Not Contains", StringComparison.InvariantCultureIgnoreCase))
                        {
                            queryResult = queryResult.Where(p => !p.Title.Contains(filter.Value)).ToList();
                        }
                        break;
                    case "Outcome":
                        queryResult =
                            queryResult.Where(
                                p => p.Outcome.Equals(filter.Value, StringComparison.InvariantCultureIgnoreCase))
                                .ToList();
                        break;
                }
            }

            foreach (var res in queryResult)
            {
                var tlitem = new QueryInterface();
                tlitem.AutomationStatus = res.AutomationStatus;
                tlitem.Outcome = res.Outcome;
                tlitem.Priority = res.Priority;
                tlitem.SuiteName = res.SuiteName;
                tlitem.TcId = res.TCID;
                tlitem.Tester = res.Tester;
                tlitem.Title = res.Title;
                reportList.Add(tlitem);
            }
            return reportList;
        }
    }
}