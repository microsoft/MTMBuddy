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
        public int TCId { get; set; }

        public string Title { get; set; }

        public bool AutomationStatus { get; set; }

        public string Tester { get; set; }

        public string SuiteName { get; set; }


        public static List<QueryInterface> Generate(List<resultsummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both", List<MTMInteraction.filter> FilterCriteria = null)
        {
            var reportList = new List<QueryInterface>();
            var rd = new List<resultsummary>();
            var filtereddata = Utilities.filterdata(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            var QueryResult = filtereddata;

            foreach (var Filter in FilterCriteria)
            {
                switch (Filter.name)
                {
                    case "Priority":
                        var Priority = int.Parse(Filter.value);
                        QueryResult = QueryResult.Where(p => p.Priority == Priority).ToList();

                        break;
                    case "Title":
                        if (Filter.op.Equals("Contains", StringComparison.InvariantCultureIgnoreCase))
                        {
                            QueryResult =
                                QueryResult.Where(
                                    p => p.Title.ToLowerInvariant().Contains(Filter.value.ToLowerInvariant())).ToList();
                        }
                        else if (Filter.op.Equals("Not Contains", StringComparison.InvariantCultureIgnoreCase))
                        {
                            QueryResult = QueryResult.Where(p => !p.Title.Contains(Filter.value)).ToList();
                        }
                        break;
                    case "Outcome":
                        QueryResult =
                            QueryResult.Where(
                                p => p.Outcome.Equals(Filter.value, StringComparison.InvariantCultureIgnoreCase))
                                .ToList();
                        break;
                }
            }

            foreach (var res in QueryResult)
            {
                var tlitem = new QueryInterface();
                tlitem.AutomationStatus = res.AutomationStatus;
                tlitem.Outcome = res.Outcome;
                tlitem.Priority = res.Priority;
                tlitem.SuiteName = res.SuiteName;
                tlitem.TCId = res.TcId;
                tlitem.Tester = res.Tester;
                tlitem.Title = res.Title;
                reportList.Add(tlitem);
            }
            return reportList;
        }
    }
}