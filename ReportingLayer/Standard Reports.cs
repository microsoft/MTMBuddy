using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using MTMIntegration;

namespace ReportingLayer
{
    public class SummaryReport
    {
        public string Priority { get; set; }
        public int Active { get; set; }
        public int Blocked { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public int Total { get; set; }
        public float FailRate { get; set; }
        public float PassRate { get; set; }
        public float ProgressRate { get; set; }

        public string LastRefreshDate { get; set; }
    
        public static List<SummaryReport> Generate(List<resultsummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();

            var reportList = new List<SummaryReport>();
            var rd = new List<resultsummary>();
            var filtereddata = Utilities.filterdata(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);
            for (var i = 1; i <= 3; i++)
            {
                var sr = new SummaryReport();

                rd = filtereddata.Where(l => l.Priority.Equals(i)).ToList();
                sr.Total = rd.Count;
                sr.Priority = "P" + i;
                sr.LastRefreshDate = DateTime.Now.ToString(@"dMMyyyy", CultureInfo.InvariantCulture);
                sr.Active = rd.Where(l => l.Outcome.Equals("Active")).Count();
                sr.Passed = rd.Where(l => l.Outcome.Equals("Passed")).Count();
                sr.Failed = rd.Where(l => l.Outcome.Equals("Failed")).Count();
                sr.Blocked = rd.Where(l => l.Outcome.Equals("Blocked")).Count();
                if (sr.Passed + sr.Failed > 0)
                {
                    sr.FailRate = (float) Math.Round((float) sr.Failed/(sr.Passed + sr.Failed)*100, 2);
                    sr.PassRate = (float) Math.Round((float) sr.Passed/(sr.Passed + sr.Failed)*100, 2);
                    sr.ProgressRate = (float) Math.Round((float) (sr.Passed + sr.Failed)/sr.Total*100, 2);
                }

                reportList.Add(sr);
            }

            var sr1 = new SummaryReport();
            sr1.Priority = "Total";
            sr1.LastRefreshDate = DateTime.Now.ToString(@"dMMyyyy", CultureInfo.InvariantCulture);
            sr1.Total = filtereddata.Count;
            sr1.Active = filtereddata.Where(l => l.Outcome.Equals("Active")).Count();
            sr1.Passed = filtereddata.Where(l => l.Outcome.Equals("Passed")).Count();
            sr1.Failed = filtereddata.Where(l => l.Outcome.Equals("Failed")).Count();
            sr1.Blocked = filtereddata.Where(l => l.Outcome.Equals("Blocked")).Count();
            if (sr1.Passed + sr1.Failed > 0)
            {
                sr1.FailRate = (float) Math.Round((float) sr1.Failed/(sr1.Passed + sr1.Failed)*100, 2);
                sr1.PassRate = (float) Math.Round((float) sr1.Passed/(sr1.Passed + sr1.Failed)*100, 2);
                sr1.ProgressRate = (float) Math.Round((float) (sr1.Passed + sr1.Failed)/sr1.Total*100, 2);
            }
            reportList.Add(sr1);

            return reportList;
        }
    }
}


//public static class StandardReporting
//{
//    public static List<resultsummary> GetResultSummaryList(int suiteId, string suiteName)
//    {
//        ConcurrentBag<resultsummary> resDetail = new ConcurrentBag<resultsummary>();
//        MTMInteraction.getsuiteresults(suiteId, resDetail, suiteName);
//        return resDetail.ToList<resultsummary>();
//    }
//}

//public class Demo
//{
//    public string A { get; set; }

//    public int B { get; set; }

//    public Demo(string a, int b)
//    {
//        this.A = a;
//        this.B = b;
//    }
//}

//public class Customer
//{

//    Demo myDemo = null;

//    public String FirstName { get; set; }
//    public String LastName { get; set; }

//    public Demo MyDemo {
//        get
//        {
//            return myDemo;
//        }
//        set
//        {
//            myDemo = new       
//        } 
//    }

//    public String Address { get; set; }
//    public Boolean IsNew { get; set; }

//    // A null value for IsSubscribed can indicate 
//    // "no preference" or "no response".
//    public Boolean? IsSubscribed { get; set; }

//    public Customer(String firstName, String lastName, Demo myDemo,
//        String address, Boolean isNew, Boolean? isSubscribed)
//    {
//        this.FirstName = firstName;
//        this.LastName = lastName;
//        this.MyDemo = MyDemo;
//        this.Address = address;
//        this.IsNew = isNew;
//        this.IsSubscribed = isSubscribed;
//    }

//    public static List<Customer> GetSampleCustomerList()
//    {
//        return new List<Customer>(new Customer[4] {
//        new Customer("A.", "Zero", new Demo("a", 1), 
//            "12 North Third Street, Apartment 45", 
//            false, true), 
//        new Customer("B.", "One", new Demo("B", 2), 
//            "34 West Fifth Street, Apartment 67", 
//            false, false),
//        new Customer("C.", "Two", new Demo("C", 3), 
//            "56 East Seventh Street, Apartment 89", 
//            true, null),
//        new Customer("D.", "Three", new Demo("D", 4), 
//            "78 South Ninth Street, Apartment 10", 
//            true, true)
//        });
//    }
//}