//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Collections.Generic;
using MTMIntegration;

namespace ReportingLayer
{
    public class TestList
    {
        public string SuiteName { get; set; }
        public int TcId { get; set; }

        public string Title { get; set; }
        public int Priority { get; set; }
        public string Outcome { get; set; }
        public DateTime Date { get; set; }
        public string Tester { get; set; }
        public bool AutomationStatus { get; set; }


        public static List<TestList> Generate(List<ResultSummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();

            var reportList = new List<TestList>();
            new List<ResultSummary>();
            var filtereddata = Utilities.FilterData(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            foreach (var res in filtereddata)
            {
                TestList tlitem = new TestList
                {
                    AutomationStatus = res.AutomationStatus,
                    Date = res.Date,
                    Outcome = res.Outcome,
                    Priority = res.Priority,
                    SuiteName = res.SuiteName,
                    TcId = res.TcId,
                    Tester = res.Tester,
                    Title = res.Title
                };
                reportList.Add(tlitem);
            }
            return reportList;
        }
    }
}


