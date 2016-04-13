using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;
using MTMIntegration;

namespace ReportingLayer
{
    public class ReleaseLevelAppProgram
    {
        public string Application { get; set; }
        public string Program { get; set; }
        public int Active { get; set; }
        public int Blocked { get; set; }
        public int Passed { get; set; }
        public int Failed { get; set; }
        public int Total { get; set; }
        public float FailRate { get; set; }
        public float PassRate { get; set; }
        public float ProgressRate { get; set; }

        public int P1Total { get; set; }
        public int P2Total { get; set; }
        public int P3Total { get; set; }

        public int P1Passed { get; set; }
        public int P2Passed { get; set; }
        public int P3Passed { get; set; }

        public int P1Failed { get; set; }
        public int P2Failed { get; set; }
        public int P3Failed { get; set; }


        public int P1Blocked { get; set; }
        public int P2Blocked { get; set; }
        public int P3Blocked { get; set; }

        public int P1Active { get; set; }
        public int P2Active { get; set; }
        public int P3Active { get; set; }

        public float P1PassRate { get; set; }
        public float P2PassRate { get; set; }
        public float P3PassRate { get; set; }

        public float P1FailRate { get; set; }
        public float P2FailRate { get; set; }
        public float P3FailRate { get; set; }

        public float P1ProgressRate { get; set; }
        public float P2ProgressRate { get; set; }
        public float P3ProgressRate { get; set; }

        public int AutomationNew { get; set; }
        public int AutomatedRegression { get; set; }

        public int AutomationTotal { get; set; }
        public int ManualNew { get; set; }

        public int ManualRegression { get; set; }

        public int ManualTotal { get; set; }

        public float AutomationRatio { get; set; }

        public float RegressionAutoRatio { get; set; }

        public float NewAutoRatio { get; set; }

        public string LastRefreshDate { get; set; }

        public string Type { get; set; }

        public string Iteration { get; set; }

       public int planid { get; set; }


        public static List<ReleaseLevelAppProgram> Generate(int planid, List<string> Programs, string iteration,
            List<string> AppList, string module = "", bool moduleinclusion = true, string tester = "",
            bool testerinclusion = true, string automationstatus = "both")
        {
            //rawData = rawData.OrderBy(l => l.Priority).ThenBy(l => l.Outcome).ToList();


            List<ReleaseLevelAppProgram> reportList = new List<ReleaseLevelAppProgram>();
            List<resultsummary> rd = new List<resultsummary>();
            List<resultsummary> rd1 = new List<resultsummary>();
            Parallel.ForEach(Programs, program =>
                //                   foreach (string program in Programs)
            {
                

               // Parallel.ForEach(AppList, app =>
                        foreach (var app in AppList)
                {
                    ConcurrentBag<resultsummary> resDetail = new ConcurrentBag<resultsummary>();
                    string eligiblesuitelist = "-1'";
                    foreach (KeyValuePair<int,string> suite in MTMInteraction.suitedictionary)
                    {
                        //check if program matches
                        if (suite.Value.ToUpperInvariant().Contains("\\" + program.ToUpperInvariant()))
                        {
                            if (suite.Value.ToUpperInvariant().Contains("\\" + iteration.ToUpperInvariant() + "\\"))
                            {
                                if (suite.Value.ToUpperInvariant().Contains("\\" + app.ToUpperInvariant() + "\\"))
                                {
                                    eligiblesuitelist = eligiblesuitelist + ",'" + suite.Key.ToString() + "'";
                                }

                            }
                        }
                    }
                    if (!eligiblesuitelist.Equals("-1'"))
                    {
                        MTMInteraction.getsuiteresults(-1, resDetail, eligiblesuitelist);

                        List<resultsummary> rawData = resDetail.ToList();
                        List<resultsummary> filtereddata = Utilities.filterdata(rawData, module, moduleinclusion, tester,
                            testerinclusion,
                            automationstatus);

                        ReleaseLevelAppProgram modrep = new ReleaseLevelAppProgram();
                        modrep.Iteration = iteration;
                        modrep.planid = planid;
                        modrep.Program = program;
                        modrep.Application = app;
                        modrep.Type = "AppProgram";
                        modrep.LastRefreshDate = DateTime.Now.ToString(@"dMMyyyy", CultureInfo.InvariantCulture);
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

                        if (modrep.P1Passed + modrep.P1Failed > 0)
                        {
                            modrep.P1FailRate =
                                (float) Math.Round((float) modrep.P1Failed/(modrep.P1Passed + modrep.P1Failed)*100, 2);
                            modrep.P1PassRate =
                                (float) Math.Round((float) modrep.P1Passed/(modrep.P1Passed + modrep.P1Failed)*100, 2);
                            modrep.P1ProgressRate =
                                (float) Math.Round((float) (modrep.P1Passed + modrep.P1Failed)/modrep.P1Total*100, 2);
                        }

                        //P2
                        rd = filtereddata.Where(l => l.Priority.Equals(2)).ToList();
                        modrep.P2Total = rd.Count;
                        modrep.P2Active = rd.Where(l => l.Outcome.Equals("Active")).Count();
                        modrep.P2Passed = rd.Where(l => l.Outcome.Equals("Passed")).Count();
                        modrep.P2Failed = rd.Where(l => l.Outcome.Equals("Failed")).Count();
                        modrep.P2Blocked = rd.Where(l => l.Outcome.Equals("Blocked")).Count();

                        if (modrep.P2Passed + modrep.P2Failed > 0)
                        {
                            modrep.P2FailRate =
                                (float) Math.Round((float) modrep.P2Failed/(modrep.P2Passed + modrep.P2Failed)*100, 2);
                            modrep.P2PassRate =
                                (float) Math.Round((float) modrep.P2Passed/(modrep.P2Passed + modrep.P2Failed)*100, 2);
                            modrep.P2ProgressRate =
                                (float) Math.Round((float) (modrep.P2Passed + modrep.P2Failed)/modrep.P2Total*100, 2);
                        }

                        //P3
                        rd = filtereddata.Where(l => l.Priority.Equals(3)).ToList();
                        modrep.P3Total = rd.Count;
                        modrep.P3Active = rd.Where(l => l.Outcome.Equals("Active")).Count();
                        modrep.P3Passed = rd.Where(l => l.Outcome.Equals("Passed")).Count();
                        modrep.P3Failed = rd.Where(l => l.Outcome.Equals("Failed")).Count();
                        modrep.P3Blocked = rd.Where(l => l.Outcome.Equals("Blocked")).Count();

                        if (modrep.P3Passed + modrep.P3Failed > 0)
                        {
                            modrep.P3FailRate =
                                (float) Math.Round((float) modrep.P3Failed/(modrep.P3Passed + modrep.P3Failed)*100, 2);
                            modrep.P3PassRate =
                                (float) Math.Round((float) modrep.P3Passed/(modrep.P3Passed + modrep.P3Failed)*100, 2);
                            modrep.P3ProgressRate =
                                (float) Math.Round((float) (modrep.P3Passed + modrep.P3Failed)/modrep.P3Total*100, 2);
                        }

                        if (modrep.Passed + modrep.Failed > 0)
                        {
                            modrep.FailRate =
                                (float) Math.Round((float) modrep.Failed/(modrep.Passed + modrep.Failed)*100, 2);
                            modrep.PassRate =
                                (float) Math.Round((float) modrep.Passed/(modrep.Passed + modrep.Failed)*100, 2);
                            modrep.ProgressRate =
                                (float) Math.Round((float) (modrep.Passed + modrep.Failed)/modrep.Total*100, 2);
                        }

                        modrep.AutomationTotal = filtereddata.Where(l => l.AutomationStatus).Count();
                        modrep.ManualTotal = filtereddata.Where(l => !l.AutomationStatus).Count();

                        if (modrep.Total > 0)
                        {
                            modrep.AutomationRatio =
                                (float)
                                    Math.Round(
                                        (float) modrep.AutomationTotal/(modrep.AutomationTotal + modrep.ManualTotal)*100,
                                        2);
                            rd1 = Utilities.filterdata(filtereddata, "Regression", true, tester, testerinclusion,
                                automationstatus);
                            modrep.ManualRegression = rd1.Where(l => !l.AutomationStatus).Count();
                            modrep.AutomatedRegression = rd1.Where(l => l.AutomationStatus).Count();
                            if (rd1.Count > 0)
                                modrep.RegressionAutoRatio =
                                    (float)
                                        Math.Round(
                                            (float) modrep.AutomatedRegression/
                                            (modrep.AutomatedRegression + modrep.ManualRegression)*100, 2);
                            rd1 = Utilities.filterdata(filtereddata, "Regression", false, tester, testerinclusion,
                                automationstatus);
                            modrep.ManualNew = rd1.Where(l => !l.AutomationStatus).Count();
                            modrep.AutomationNew = rd1.Where(l => l.AutomationStatus).Count();
                            if (rd1.Count > 0)
                                modrep.NewAutoRatio =
                                    (float)
                                        Math.Round(
                                            (float) modrep.AutomationNew/(modrep.AutomationNew + modrep.ManualNew)*100,
                                            2);
                        }
                        reportList.Add(modrep);
                    }
                }//);
            });
                         


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