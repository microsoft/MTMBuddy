using System;
using System.Collections.Generic;
using MTMIntegration;

namespace ReportingLayer
{
    public class TestList
    {
        public string SuiteName { get; set; }
        public int TCId { get; set; }

        public string Title { get; set; }
        public int Priority { get; set; }
        public string Outcome { get; set; }
        public DateTime Date { get; set; }
        public string Tester { get; set; }
        public bool AutomationStatus { get; set; }


        public static List<TestList> Generate(List<resultsummary> rawData, string module = "",
            bool moduleinclusion = true, string tester = "", bool testerinclusion = true,
            string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();

            var reportList = new List<TestList>();
            var rd = new List<resultsummary>();
            var filtereddata = Utilities.filterdata(rawData, module, moduleinclusion, tester, testerinclusion,
                automationstatus);

            foreach (var res in filtereddata)
            {
                var tlitem = new TestList();
                tlitem.AutomationStatus = res.AutomationStatus;
                tlitem.Date = res.Date;
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