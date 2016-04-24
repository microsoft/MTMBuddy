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
            string automationstatus = "both", List<Filter> filterCriteria = null)
        {
            var filtereddata = Utilities.FilterData(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            var queryResult = filtereddata;

            if (filterCriteria == null)
                return queryResult.Select(res => new QueryInterface
                {
                    AutomationStatus = res.AutomationStatus,
                    Outcome = res.Outcome,
                    Priority = res.Priority,
                    SuiteName = res.SuiteName,
                    TcId = res.TcId,
                    Tester = res.Tester,
                    Title = res.Title
                }).ToList();
            foreach (var filter in filterCriteria)
            {
                switch (filter.Name)
                {
                    case "Priority":
                        var priority = int.Parse(filter.Value);
                        queryResult = queryResult.Where(p => p.Priority == priority).ToList();

                        break;
                    case "Title":
                        if (filter.Op.Equals("Contains", StringComparison.OrdinalIgnoreCase))
                        {
                            queryResult =
                                queryResult.Where(
                                    p => p.Title.ToLowerInvariant().Contains(filter.Value.ToLowerInvariant())).ToList();
                        }
                        else if (filter.Op.Equals("Not Contains", StringComparison.OrdinalIgnoreCase))
                        {
                            queryResult = queryResult.Where(p => !p.Title.Contains(filter.Value)).ToList();
                        }
                        break;
                    case "Outcome":
                        queryResult =
                            queryResult.Where(
                                p => p.Outcome.Equals(filter.Value, StringComparison.OrdinalIgnoreCase))
                                .ToList();
                        break;
                }
            }

            return queryResult.Select(res => new QueryInterface
            {
                AutomationStatus = res.AutomationStatus, Outcome = res.Outcome, Priority = res.Priority, SuiteName = res.SuiteName, TcId = res.TcId, Tester = res.Tester, Title = res.Title
            }).ToList();
        }
    }

    public class Filter
    {
        public string Name { get; set; }
        public string Op { get; set; }
        public string Value { get; set; }
    }
}