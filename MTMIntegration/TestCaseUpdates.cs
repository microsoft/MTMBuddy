using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;

namespace MTMIntegration
{
    public class AssociateTestCase
    {
        
       
      

        #region ResultUpdate
        
        private static readonly string TestPointQuery = "Select * from TestPoint where TestCaseId= '[#testcaseid#]'";

        private static readonly string TestPointSuiteQuery =
            "Select * from TestPoint where TestCaseId= '[#testcaseid#]' and SuiteId = '[#testsuiteid#]'";

     
        

        //private static ITestRun PlanRun;

        private static ITestCaseResult TestResult;

        private static string PlanBuildNumber=string.Empty ;


        public static string updateResult(string Tcid, string comments, string SuiteName, bool usebuildnumber,
            string result = "Passed", int bugid = 0, string Attachments = null)
        {
            var tsid = string.Empty;
            var returnstring = "success";
            var finalsuite = string.Empty;
            var debug = string.Empty;


            var he = MTMInteraction.TestPlan.QueryTestPointHierarchy(TestPointQuery.Replace("[#testcaseid#]", Tcid));
                // finds all occurences of this test case.
            //foreach (HierarchyEntry a in he.Children)
            //{
            //    debug = a.SuiteTitle;
            //    if (a.SuiteTitle.Equals(SuiteName))// to match the pass specified in config
            //    {
            //        suiteid = getFinalSuiteID(a);
            //    }
            //}
            if (he.Children.Count <= 0)
                return "Testcase not found";
            tsid = getFinalSuiteID(he, SuiteName, out finalsuite).ToString();
            if (tsid.Equals("-1"))
                return "Testcase does not belong to suite";

            var pointCollection =
                MTMInteraction.TestPlan.QueryTestPoints(
                    TestPointSuiteQuery.Replace("[#testcaseid#]", Tcid).Replace("[#testsuiteid#]", tsid));
                // like "SELECT * from TestPoint where TestCaseId='185716'
            var PlanRun = MTMInteraction.TestPlan.CreateTestRun(false);


            PlanRun.Title = "Test Run";
            if (usebuildnumber)
            {
                PlanRun.BuildNumber = PlanBuildNumber;
            }


            try
            {
                PlanRun.AddTestPoint(pointCollection[0], null);
                PlanRun.Save();
                var tp = pointCollection[0];
                var proj = tp.TestCaseWorkItem.Area;
                TestResult = PlanRun.QueryResults()[PlanRun.QueryResults().Count - 1];
                TestResult.DateStarted = DateTime.Now;
                //Console.WriteLine("Test Result count : {0}", PlanRun.QueryResults().Count);
                if (result.Equals("In progress", StringComparison.InvariantCultureIgnoreCase))
                {
                    TestResult.State = TestResultState.InProgress;
                    TestResult.Outcome = TestOutcome.None;
                }
                else if(result.Equals("Active",StringComparison.InvariantCultureIgnoreCase ))
                {
                    TestResult.State = TestResultState.Completed;
                    TestResult.Outcome = TestOutcome.None;
                }
                else
                {
                    TestResult.State = TestResultState.Completed;
                    //TestResult.Outcome = TestOutcome.Passed;
                    
                    TestResult.Outcome = (TestOutcome)Enum.Parse(typeof(TestOutcome), result);
                }

                TestResult.ComputerName = Dns.GetHostName();
                TestResult.RunBy = PlanRun.Owner;
                var tc = TestResult.GetTestCase();


                if (Attachments != null && !string.IsNullOrEmpty(Attachments))
                {
                    var attachmentlist = Attachments.Split(';');
                    foreach (var attachementpath in attachmentlist)
                    {
                        if (File.Exists(attachementpath))
                        {
                            TestResult.Attachments.Add(TestResult.CreateAttachment(attachementpath,
                                SourceFileAction.None));
                            PlanRun.Attachments.Add(PlanRun.CreateAttachment(attachementpath, SourceFileAction.None));
                        }
                        else
                        {
                            returnstring = "Success but Attachment-" + attachementpath + " not found";
                        }
                    }
                }

                TestResult.Comment = "Suite:" + SuiteName + " " + "Build:" + PlanBuildNumber + " " + comments;
                TestResult.DateCompleted = DateTime.Now;

                if (bugid != 0)
                {
                    try
                    {
                        var tmp = MTMInteraction.wstore.GetWorkItem(bugid);
                        TestResult.AssociateWorkItem(tmp);
                        TestResult.Save();
                        if (!string.IsNullOrEmpty(Attachments))
                        {
                            var attachmentlist = Attachments.Split(';');
                            foreach (var attachementpath in attachmentlist)
                            {
                                if (File.Exists(attachementpath))
                                {
                                    var att = new Attachment(attachementpath);
                                    tmp.Attachments.Add(att);
                                }
                            }
                        }

                        var linkTypeEnd = MTMInteraction.wstore.WorkItemLinkTypes.LinkTypeEnds["Test Case"];
                        var r = new RelatedLink(linkTypeEnd, tc.Id);
                        tmp.Links.Add(r);

                        tmp.Save();
                    }
                    catch (Exception e)
                    {
                        returnstring = "Result saved succesfully but bug could not be associated due to" + e.Message;
                    }
                }
                else
                    TestResult.Save();

                return returnstring;
            }

            catch (Exception)
            {
                PlanRun.Delete();
                //return ex.Message;
                throw;
            }
        }

        /// <summary>
        ///     This method returns the id of the first suite that contains the test case to be passed. This has been added to pass
        ///     test cases under specific sub suites like Pass 2 or Pass 1.
        /// </summary>
        /// <param name="testhe"> The first node. One of its sub suites contains the case to be passed.</param>
        /// <param name="suitename">Name of the suite can be a bracnh or leaf</param>
        /// <param name="finalsuitename">Name of the leaf suite </param>
        /// <returns>suite id</returns>
        private static int getFinalSuiteID(HierarchyEntry testhe, string suitename, out string finalsuitename)
        {
            finalsuitename = string.Empty;
            foreach (var branch in testhe.Children)
            {
                if (branch.SuiteTitle.Equals(suitename, StringComparison.InvariantCultureIgnoreCase))
                {
                    if (branch.Children.Count > 0)
                    {
                        var temp = branch;
                        while (temp.Children.Count > 0)
                        {
                            temp = temp.Children[0];
                            finalsuitename = finalsuitename + "\\" + temp.SuiteTitle;
                        }

                        return temp.SuiteId;
                    }
                        finalsuitename = branch.SuiteTitle;
                    return branch.SuiteId;
                }
                if (branch.Children.Count > 0)
                {
                    var suiteid = getFinalSuiteID(branch, suitename, out finalsuitename);
                    if (suiteid > 0)
                    {
                        return suiteid;
                    }
                }
            }
            finalsuitename = string.Empty;
            return -1;
        }

        #endregion ResultUpdate

       
    }
}
