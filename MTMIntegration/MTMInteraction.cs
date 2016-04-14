using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using TreeView = System.Windows.Forms.TreeView;

namespace MTMIntegration
{
    public static class MTMInteraction
    {
        private static Uri TempURI;

        private static List<string> tester = new List<string>();
        public static TfsTeamProjectCollection TfsProjColl;

        public static ITestManagementService Tms;

        public static ITestManagementTeamProject TeamProject;

        public static int TestPlanID;

        public static ITestPlan TestPlan;

        private static readonly Stopwatch stp = new Stopwatch();


        public static readonly Dictionary<int, string> suitedictionary = new Dictionary<int, string>();

        public static float OutcomeTime;
        public static float TesterTime;
        public static float PriorityTime;
        public static float AutomationTime;
        public static float Initialize;
        public static float tcidtime;
        public static float Titletime;
        public static float AutomationMethodTime;
        public static float AutomationMethodTime1;
        public static float AutomationTestNameFetchTime;

        //private static ITestRun PlanRun;


      

        private static string ProjectName;
        public static WorkItemStore wstore;

     

        private static readonly string TestSuiteQuery = "Select * from TestPoint where SuiteId in  ('[#testsuiteid#])";

        /// <summary>
        ///     planName
        /// </summary>
        public static string planName { get; set; }

        public static void clearperfcounters()
        {
            OutcomeTime = 0;
            TesterTime = 0;
            PriorityTime = 0;
            AutomationTime = 0;
            Initialize = 0;
            tcidtime = 0;
            Titletime = 0;
            AutomationMethodTime = 0;
            AutomationMethodTime1 = 0;
            AutomationTestNameFetchTime = 0;
        }

        /// <summary>
        /// </summary>
        /// <param name="VSTFconnectionstring"></param>
        /// <param name="VSTFProject"></param>
        /// <param name="PlanId"></param>
        /// <param name="flags"></param>
        /// <returns></returns>
        public static List<PlanDetails> GetAllCasesfromPlan(string VSTFconnectionstring, string VSTFProject, int PlanId,
            Dictionary<string, bool> flags)
        {
            var vstfconnection = new Uri(VSTFconnectionstring);
            initializeVSTFUpdate(vstfconnection, VSTFProject, PlanId, "");
            var allcases = new List<PlanDetails>();

            //Step 1: Get Application-Iteration-Suiteid mapping

            string[] iterations = {"DUT", "DIT", "SIT1", "SIT2", "UAT"};

            Parallel.ForEach(iterations, iteration =>
            {
                //Get Application-suiteid mapping
                var Applist = getsuitelist(PlanId, iteration, getSuite(PlanId));

                Parallel.ForEach(Applist.Keys, App =>
                {
                    var suiteid = Applist[App];
                    //get all subsuites
                    var type = getType(suiteid);

                    ITestPointCollection pointCollection = null;
                    if (type)
                    {
                        var suitelist = getsubsuites(suiteid);
                        suitelist = suitelist.Substring(0, suitelist.Length - 3);
                        pointCollection = TestPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suitelist));
                        // like "SELECT * from TestPoint where TestCaseId='185716'
                    }
                    else
                    {
                        pointCollection =
                            TestPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suiteid.ToString()));
                        // like "SELECT * from TestPoint where TestCaseId='185716'
                    }

                    var resultitem = new PlanDetails();


                    foreach (var tpoint in pointCollection)
                    {
                        resultitem = new PlanDetails();
                        resultitem.Iteration = iteration;
                        resultitem.Application = App;

                        var outcome = string.Empty;
                        outcome = tpoint.MostRecentResultOutcome.ToString();

                        if (outcome.Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (
                                !tpoint.MostRecentResult.Outcome.ToString()
                                    .Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                                resultitem.Outcome = "Active";
                            else
                                resultitem.Outcome = "Blocked";
                        }
                        else if (outcome.Equals("Unspecified", StringComparison.InvariantCultureIgnoreCase) ||
                                 outcome.Equals("None", StringComparison.InvariantCultureIgnoreCase))
                        {
                            resultitem.Outcome = "Active";
                        }
                        else
                        {
                            resultitem.Outcome = outcome;
                        }

                        resultitem.TcID = tpoint.TestCaseId;

                        resultitem.Priority = tpoint.TestCaseWorkItem.Priority;
                        allcases.Add(resultitem);
                    }
                });
            });

            return allcases;
        }


        /// <summary>
        ///     Establishes connection with TFS and select the project and the test plan
        /// </summary>
        /// <param name="ConnUri"></param>
        /// <param name="Project"></param>
        /// <param name="PlanID"></param>
        /// <param name="BuildNumber"></param>
        public static void initializeVSTFUpdate(Uri ConnUri, string Project, int PlanID, string BuildNumber)
        {
            TempURI = ConnUri;
            TestPlanID = PlanID;
            
            ProjectName = Project;
            TfsProjColl = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(TempURI);
            TfsProjColl.Authenticate();
            if (TfsProjColl.HasAuthenticated)
            {
                Tms = TfsProjColl.GetService<ITestManagementService>();

                wstore = (WorkItemStore) TfsProjColl.GetService(typeof (WorkItemStore));
                TeamProject = Tms.GetTeamProject(ProjectName);
                TestPlan = TeamProject.TestPlans.Find(TestPlanID);
                planName = TestPlan.Name;

                TeamProject.TfsIdentityStore.Refresh();

                foreach (var i in TeamProject.TfsIdentityStore.AllUserIdentities)
                {
                    tester.Add(i.DisplayName);
                }

            }
            else
            {
                var aexp = new ApplicationException("Unable to authenticate");
                throw aexp;
            }
        }


        /// <summary>
        ///     Recursively searches for a suite with given id and returns the suite
        /// </summary>
        /// <param name="rootsuite">The root suite to search</param>
        /// <param name="suiteid">Suite id to search for</param>
        /// <returns></returns>
        private static IStaticTestSuite getsubsuitebyid(IStaticTestSuite rootsuite, int suiteid)
        {
            for (var i = 0; i < rootsuite.SubSuites.Count; i++)
            {
                if (rootsuite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) rootsuite.SubSuites[i];
                    if (subsuite.Id.Equals(suiteid))
                        return subsuite;

                    if (subsuite.SubSuites.Count > 0)
                    {
                        var retsuite = getsubsuitebyid(subsuite, suiteid);
                        if (retsuite != null)
                            return retsuite;
                    }
                }
            }
            return null;
        }


        /// <summary>
        /// </summary>
        /// <param name="suiteid"></param>
        /// <returns></returns>
        private static string getsubsuites(int suiteid)
        {
            var suitelist = string.Empty;

            var tn = new TreeNode();

            IStaticTestSuite rootsuite = null;
            for (var i = 0; i < TestPlan.RootSuite.SubSuites.Count; i++)
            {
                if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) TestPlan.RootSuite.SubSuites[i];

                    if (subsuite.Id.Equals(suiteid))
                    {
                        rootsuite = subsuite;
                    }
                    if (subsuite.SubSuites.Count > 0)
                    {
                        var retsuite = getsubsuitebyid(subsuite, suiteid);
                        if (retsuite != null)
                            rootsuite = retsuite;
                    }
                }

                if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("DynamicTestSuite"))
                {
                    var suite = (IDynamicTestSuite) TestPlan.RootSuite.SubSuites[i];

                    if (suite.Id.Equals(suiteid))
                    {
                    }
                }
            }

            if (rootsuite.GetType().Name.Equals("StaticTestSuite"))
            {
                suitelist = getsubsuitelist(rootsuite, suitelist);
            }

            return suitelist;
        }

        /// <summary>
        ///     Construct MTM hierarchy as a tree
        /// </summary>
        /// <param name="planid">id of the plan</param>
        /// <param name="mtmtree">The treeview to be populated</param>
        /// <returns></returns>
        public static void getsuitetree(int planid, TreeView mtmtree)
        {
            TestPlan = TeamProject.TestPlans.Find(planid);
            foreach (IStaticTestSuite suite in TestPlan.RootSuite.SubSuites)
            {
                mtmtree.Nodes.Add(suite.Title, suite.Title, suite.Id);
                getsubsuitetree(suite, mtmtree.Nodes[suite.Title]);
            }
        }

        /// <summary>
        ///     Retrun subnode structure
        /// </summary>
        /// <param name="suite">Root suite</param>
        /// <param name="tn">Node to be populated</param>
        /// <returns></returns>
        public static void getsubsuitetree(IStaticTestSuite suite, TreeNode tn)
        {
            tn.Name = suite.Title;
            tn.ImageIndex = suite.Id;

            for (var i = 0; i < suite.SubSuites.Count; i++)
            {
                if (suite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) suite.SubSuites[i];

                    if (subsuite.SubSuites.Count > 0)
                    {
                        tn.Nodes.Add(subsuite.Title, subsuite.Title, subsuite.Id);

                        getsubsuitetree(subsuite, tn.Nodes[subsuite.Title]);
                    }
                    else
                    {
                        tn.Nodes.Add(subsuite.Title, subsuite.Title, subsuite.Id);
                    }
                }
            }
        }


        /// <summary>
        ///     Retrun subnode structure
        /// </summary>
        /// <param name="suite">Root suite</param>
        /// <param name="suitelist">List of suiteids under this suite</param>
        /// <returns></returns>
        public static string getsubsuitelist(IStaticTestSuite suite, string suitelist = "")
        {
            suitelist = suitelist + suite.Id + "','";
            for (var i = 0; i < suite.SubSuites.Count; i++)
            {
                if (suite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) suite.SubSuites[i];
                    if (subsuite.SubSuites.Count > 0)
                    {
                        suitelist = suitelist + subsuite.Id + "','";
                        suitelist = suitelist + getsubsuitelist(subsuite, suitelist);
                    }
                    else
                    {
                        suitelist = suitelist + subsuite.Id + "','";
                    }
                }
                else if (suite.SubSuites[i].GetType().Name.Equals("DynamicTestSuite"))
                {
                    var dynsuite = (IDynamicTestSuite) suite.SubSuites[i];

                    suitelist = suitelist + dynsuite.Id + "','";
                }

                else
                {
                    var reqsuite = (IRequirementTestSuite) suite.SubSuites[i];

                    suitelist = suitelist + reqsuite.Id + "','";
                }
            }
            return suitelist;
        }


        /// <summary>
        ///     Return results associated with suite
        /// </summary>
        /// <param name="suiteid">suiteid to fetch</param>
        /// <param name="resdetails">List of resultdetails which will be populated</param>
        /// <param name="suitename">Suitename will be appended to the result.</param>
        /// <returns></returns>
        public static bool getType(int suiteid)
        {
            for (var i = 0; i < TestPlan.RootSuite.SubSuites.Count; i++)
            {
                if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var suite = (IStaticTestSuite) TestPlan.RootSuite.SubSuites[i];
                    if (suite.Id.Equals(suiteid))
                    {
                        return true;
                    }


                    if (getsubsuitetype(suite, suiteid))
                        return true;
                }
            }

            return false;
        }

        /// <summary>
        ///     Get the type of Suite
        /// </summary>
        /// <param name="suite"></param>
        /// <param name="suiteid"></param>
        /// <returns>bool</returns>
        public static bool getsubsuitetype(IStaticTestSuite suite, int suiteid)
        {
            for (var i = 0; i < suite.SubSuites.Count; i++)
            {
                if (suite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) suite.SubSuites[i];

                    if (subsuite.Id.Equals(suiteid))
                    {
                        return true;
                    }

                    if (subsuite.SubSuites.Count > 0)
                    {
                        if (getsubsuitetype(subsuite, suiteid))
                            return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// </summary>
        /// <param name="suiteid"></param>
        /// <param name="resdetails"></param>
        /// <param name="suitename">List of suiteids</param>
        public static void getsuiteresults(int suiteid, ConcurrentBag<resultsummary> resdetails,
            string suitename = "None")
        {
            stp.Restart();
            var type = getType(suiteid);
            
            ITestPointCollection pointCollection = null;
            if (suiteid != -1)
            {
                if (type)
                {
                    var suitelist = getsubsuites(suiteid);
                    suitelist = suitelist.Substring(0, suitelist.Length - 2);
                    pointCollection = TestPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suitelist));
                    // like "SELECT * from TestPoint where TestCaseId='185716'
                }
                else
                {
                    
                    pointCollection =
                        TestPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suiteid.ToString()+"'"));
                    
                }
            }
            else
            {
                pointCollection =
                    TestPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suitename));
            }


            resdetails.AsParallel();
            var res = new resultsummary();


            stp.Stop();
            Initialize = Initialize + (float) stp.ElapsedMilliseconds/1000;
            foreach (var tpoint in pointCollection)
            {
                res = new resultsummary();
                stp.Restart();
                var outcome = string.Empty;
                outcome = tpoint.MostRecentResultOutcome.ToString();

                if (tpoint.MostRecentResult.State.Equals(TestResultState.InProgress))

                {
                    outcome = "In progress";
                }
                    if (outcome.Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (
                        !tpoint.MostRecentResult.Outcome.ToString()
                            .Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                        res.Outcome = "Active";
                    else
                        res.Outcome = "Blocked";
                }
                else if (outcome.Equals("Unspecified", StringComparison.InvariantCultureIgnoreCase) ||
                         outcome.Equals("None", StringComparison.InvariantCultureIgnoreCase))
                {
                    res.Outcome = "Active";
                }
                else
                {
                    res.Outcome = outcome;
                }
                stp.Stop();
                OutcomeTime = OutcomeTime + (float) stp.ElapsedMilliseconds/1000;
                stp.Restart();
                res.SuiteName = suitedictionary[tpoint.SuiteId];

                res.TcId = tpoint.TestCaseId;
                stp.Stop();
                tcidtime = tcidtime + (float) stp.ElapsedMilliseconds/1000;

                try
                {
                    stp.Restart();
                    res.AutomationStatus = tpoint.IsTestCaseAutomated;
                    stp.Stop();
                    AutomationTime = AutomationTime + (float) stp.ElapsedMilliseconds/1000;
                }
                catch (Exception)
                {
                    res.AutomationStatus = false;
                }
                try
                {
                    stp.Restart();
                    var tc = tpoint.TestCaseWorkItem;
                    res.Priority = tc.Priority;
                    stp.Stop();
                    PriorityTime = PriorityTime + (float) stp.ElapsedMilliseconds/1000;

                    stp.Restart();
                    //res.AutomationTestName = tc.Implementation.DisplayText;
                    stp.Stop();
                    AutomationTestNameFetchTime += (float) stp.ElapsedMilliseconds/1000;
                }
                catch (Exception)
                {
                    res.Priority = -1;
                }
                try
                {
                    stp.Restart();
                   
                        res.Tester = tpoint.AssignedTo.DisplayName;
                    
                   
                    stp.Stop();
                    TesterTime = TesterTime + (float) stp.ElapsedMilliseconds/1000;
                }


                catch (Exception)
                {
                   
                    res.Tester = "Nobody";
                }

                try
                {
                    stp.Restart();
                    res.Title = tpoint.TestCaseWorkItem.Title;
                    stp.Stop();
                    Titletime = Titletime + (float) stp.ElapsedMilliseconds/1000;
                }


                catch (Exception)
                {
                    res.Title = "No Title";
                }

              
                lock (resdetails)
                {
                    resdetails.Add(res);
                }
            }
        }

       


        /// <summary>
        ///     Gives the list of suites in a plan
        /// </summary>
        /// <param name="planId"></param>
        /// <returns>list of suite title</returns>
        public static List<string> getSuite(int planId)
        {
            TestPlan = TeamProject.TestPlans.Find(planId);
            var suiteList = new List<string>();
            for (var i = 0; i < TestPlan.RootSuite.SubSuites.Count; i++)
            {
                if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) TestPlan.RootSuite.SubSuites[i];
                    suiteList.Add(subsuite.Title);
                }
            }
            return suiteList;
        }


        /// <summary>
        ///     Get the dictionary with plan id as key and name as value.
        /// </summary>
        /// <returns> Dicitonary</returns>
        public static Dictionary<int, string> getPlanId()
        {
            var planName = new Dictionary<int, string>();

            var plan = TeamProject.TestPlans.Query("Select *from TestPlan");
            foreach (var i in plan)
            {
                planName.Add(i.Id, i.Name);
            }

            return planName;
        }

        /// <summary>
        ///     Look up function that searches user particular plan name using Id.
        /// </summary>
        /// <param name="planId"></param>
        /// <returns> string </returns>
        public static string getPlanName(int planId)
        {
            try {
                var planName = new Dictionary<int, string>();
                planName = getPlanId();
                var Name = string.Empty;
                if (planName.ContainsKey(planId))
                {
                    Name = planName[planId];
                    return Name;
                }

                return null;
            }catch(Exception)
            {
                return null;
            }
        }

        // get the query filter list

        public class filter
        {
            public string name { get; set; }
            public string op { get; set; }
            public string value { get; set; }
        }

        #region treelogic

        /// <summary>
        ///     Construct MTM hierarchy as a tree
        /// </summary>
        /// <param name="planid">id of the plan</param>
        /// <param name="mtmtree">The treeview to be populated</param>
        /// <returns></returns>
        public static Dictionary<string, int> getsuitelist(int planid, string iteration, List<string> selectedSuite)
        {
            var AppDictionary = new Dictionary<string, int>();
            AppDictionary.Clear();


            TestPlan = TeamProject.TestPlans.Find(planid);
            for (var j = 0; j < TestPlan.RootSuite.SubSuites.Count; j++)
            {
                if (TestPlan.RootSuite.SubSuites[j].GetType().Name.Equals("StaticTestSuite"))
                {
                    var suite = (IStaticTestSuite) TestPlan.RootSuite.SubSuites[j];
                    if (selectedSuite.Contains(suite.Title))
                    {
                        for (var i = 0; i < suite.SubSuites.Count; i++)
                        {
                            if (suite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                            {
                                var subsuite = (IStaticTestSuite) suite.SubSuites[i];
                                if (subsuite.SubSuites.Count > 0)
                                {
                                    if (subsuite.Title.Equals(iteration, StringComparison.InvariantCultureIgnoreCase))
                                    {
                                        if (AppDictionary.ContainsKey(suite.Title) == false)
                                        {
                                            AppDictionary.Add(suite.Title, subsuite.Id);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return AppDictionary;
        }

        /// <summary>
        ///     Construct MTM hierarchy as a tree
        /// </summary>
        /// <param name="planid">id of the plan</param>
        /// <param name="mtmtree">The treeview to be populated</param>
        /// <returns></returns>
        public static void buildsuitedictionary(int planid, bool includeparent = false)
        {
            suitedictionary.Clear();
            TestPlan = TeamProject.TestPlans.Find(planid);
            if (TestPlan == null)
            {
                var exp =
                    new ApplicationException(
                        "Invalid planid.You can get the planid from MTM in the Select Test Plan Window.");
                throw exp;
            }

            for (var i = 0; i < TestPlan.RootSuite.SubSuites.Count; i++)
            {
                if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var suite = (IStaticTestSuite) TestPlan.RootSuite.SubSuites[i];
                    if (includeparent)
                    {
                        suitedictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(suite.Id, suite.Title);
                    }
                    buildsubsuitedictionary(suite, TestPlan.RootSuite.Title + "\\", includeparent);
                }

                else if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("DynamicTestSuite"))
                {
                    var suite = (IDynamicTestSuite) TestPlan.RootSuite.SubSuites[i];
                    if (includeparent)
                    {
                        suitedictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(suite.Id, suite.Title);
                    }
                }
                else
                {
                    var suite = (IRequirementTestSuite) TestPlan.RootSuite.SubSuites[i];

                    if (includeparent)
                    {
                        suitedictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(suite.Id, suite.Title);
                    }
                }
            }
        }

        /// <summary>
        ///     Retrun subnode structure
        /// </summary>
        /// <param name="suite">Root suite</param>
        /// <param name="includeparent"></param>
        /// <param name="prefix"> prefix for title</param>
        /// <returns></returns>
        public static void buildsubsuitedictionary(IStaticTestSuite suite, string prefix, bool includeparent = false)
        {
            for (var i = 0; i < suite.SubSuites.Count; i++)
            {
                if (suite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) suite.SubSuites[i];
                    if (subsuite.SubSuites.Count > 0)
                    {
                        if (includeparent)
                        {
                            suitedictionary.Add(subsuite.Id, prefix + "\\" + suite.Title + "\\" + subsuite.Title);
                        }
                        else
                        {
                            suitedictionary.Add(subsuite.Id, subsuite.Title);
                        }
                        buildsubsuitedictionary(subsuite, prefix + "\\" + suite.Title, includeparent);
                    }
                    else
                    {
                        if (includeparent)
                        {
                            suitedictionary.Add(subsuite.Id, prefix + "\\" + suite.Title + "\\" + subsuite.Title);
                        }
                        else
                        {
                            suitedictionary.Add(subsuite.Id, subsuite.Title);
                        }
                    }
                }
                else if (suite.SubSuites[i].GetType().Name.Equals("DynamicTestSuite"))
                {
                    var dynsuite = (IDynamicTestSuite) suite.SubSuites[i];


                    if (includeparent)
                    {
                        suitedictionary.Add(dynsuite.Id, prefix + "\\" + suite.Title + "\\" + dynsuite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(dynsuite.Id, dynsuite.Title);
                    }
                }
                else
                {
                    var reqsuite = (IRequirementTestSuite) suite.SubSuites[i];
                    if (includeparent)
                    {
                        suitedictionary.Add(reqsuite.Id, prefix + "\\" + suite.Title + "\\" + reqsuite.Title);
                    }

                    else
                    {
                        suitedictionary.Add(reqsuite.Id, reqsuite.Title);
                    }
                }
            }
        }


        public static string getSuitePath(int suiteid)
        {
            if (suitedictionary.ContainsKey(suiteid))
            {
                return suitedictionary[suiteid];
            }
            return null;
        }

        /// <summary>
        ///     Get the name of Testers
        /// </summary>
        /// <returns>List of tester names</returns>
        public static List<string> getTester()
        {
            
            return tester;
        }


        /// <summary>
        ///     Query Interface Implementation.
        /// </summary>
        /// <param name="planId"></param>
        /// <param name="status"></param>
        public static List<querydetails> getQueryResults(int suiteid, List<filter> inFilter, string suitename)
        {
            var suitelist = getsubsuites(suiteid);
            suitelist = suitelist.Substring(0, suitelist.Length - 3);
            var pointCollection = TestPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suitelist));
            // like "SELECT * from TestPoint where TestCaseId='185716'
            var retrylist = new List<ITestPoint>();
            var resdetails = new List<querydetails>();
            var newresdetails = new List<querydetails>();

            var res = new querydetails();
            var testers = new List<string>();

            foreach (var fil in inFilter)
            {
                if (fil.name.Equals("Outcome", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (fil.op.Equals("="))
                    {
                        foreach (var tpoint in pointCollection)
                        {
                            res = new querydetails();

                            var outcome = string.Empty;
                            outcome = tpoint.MostRecentResultOutcome.ToString();
                            if (outcome.Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                            {
                                if (
                                    !tpoint.MostRecentResult.Outcome.ToString()
                                        .Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                                    outcome = "Unspecified";
                            }
                            if (outcome.Equals("Unspecified", StringComparison.InvariantCultureIgnoreCase))
                                outcome = "Active";
                            if (outcome.Equals(fil.value, StringComparison.InvariantCultureIgnoreCase))
                            {
                                try
                                {
                                    res.TcId = tpoint.TestCaseWorkItem.Id;
                                    res.Tester = tpoint.AssignedTo.DisplayName;
                                    res.Title = tpoint.TestCaseWorkItem.Title;
                                    res.Priority = tpoint.TestCaseWorkItem.Priority;
                                    res.AutomationStatus = tpoint.IsTestCaseAutomated;
                                    res.Outcome = outcome;
                                    //      res.Date = tpoint.MostRecentResult.DateCompleted;      
                                }

                                catch (Exception)
                                {
                                    res.AutomationStatus = false;
                                    res.Priority = -1;
                                    res.Tester = "Nobody";
                                }

                                resdetails.Add(res);
                            }
                        }
                    }

                    else if (fil.op.Equals("!="))
                    {
                        foreach (var tpoint in pointCollection)
                        {
                            res = new querydetails();

                            var outcome = string.Empty;
                            outcome = tpoint.MostRecentResultOutcome.ToString();
                            if (outcome.Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                            {
                                if (
                                    !tpoint.MostRecentResult.Outcome.ToString()
                                        .Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                                    outcome = "Unspecified";
                            }
                            if (outcome.Equals("Unspecified", StringComparison.InvariantCultureIgnoreCase))
                                outcome = "Active";
                            if (!outcome.Equals(fil.value, StringComparison.InvariantCultureIgnoreCase))
                            {
                                try
                                {
                                    res.TcId = tpoint.TestCaseWorkItem.Id;
                                    res.Tester = tpoint.AssignedTo.DisplayName;
                                    res.Title = tpoint.TestCaseWorkItem.Title;
                                    res.Priority = tpoint.TestCaseWorkItem.Priority;
                                    res.AutomationStatus = tpoint.IsTestCaseAutomated;
                                    res.Outcome = outcome;

                                    //        res.Date = tpoint.MostRecentResult.DateCompleted;
                                }

                                catch (Exception)
                                {
                                    res.AutomationStatus = false;
                                    res.Priority = -1;
                                    res.Tester = "Nobody";
                                }
                                resdetails.Add(res);
                            }
                        }
                    }
                }

                else
                {
                    foreach (var tpoint in pointCollection)
                    {
                        res = new querydetails();


                        var outcome = string.Empty;
                        outcome = tpoint.MostRecentResultOutcome.ToString();
                        if (outcome.Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                        {
                            if (
                                !tpoint.MostRecentResult.Outcome.ToString()
                                    .Equals("Blocked", StringComparison.InvariantCultureIgnoreCase))
                                outcome = "Unspecified";
                        }
                        if (outcome.Equals("Unspecified", StringComparison.InvariantCultureIgnoreCase))
                            outcome = "Active";
                        try
                        {
                            res.TcId = tpoint.TestCaseWorkItem.Id;
                            res.Tester = tpoint.AssignedTo.DisplayName;
                            res.Title = tpoint.TestCaseWorkItem.Title;
                            res.Priority = tpoint.TestCaseWorkItem.Priority;
                            res.AutomationStatus = tpoint.IsTestCaseAutomated;
                            res.Outcome = outcome;
                            //  res.Date = tpoint.MostRecentResult.DateCompleted;
                        }
                        catch (Exception)
                        {
                            res.AutomationStatus = false;
                            res.Priority = -1;
                            res.Tester = "Nobody";
                        }

                        resdetails.Add(res);
                    }
                }


                if (fil.name.Equals("TestCaseID", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (fil.op.Equals("="))
                    {
                        foreach (var item in resdetails)
                        {
                            if (item.TcId == int.Parse(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }


                    if (fil.op.Equals("!="))
                    {
                        foreach (var item in resdetails)
                        {
                            if (item.TcId != int.Parse(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }


                    if (fil.op.Equals(">"))
                    {
                        foreach (var item in resdetails)
                        {
                            if (item.TcId > int.Parse(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }

                    if (fil.op.Equals("<"))
                    {
                        foreach (var item in resdetails)
                        {
                            if (item.TcId < int.Parse(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }


                    resdetails = new List<querydetails>();

                    foreach (var i in newresdetails)
                    {
                        resdetails.Add(i);
                    }

                    newresdetails = new List<querydetails>();
                }

                if (fil.name.Equals("Priority", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (fil.op.Equals("="))
                    {
                        res = new querydetails();
                        foreach (var item in resdetails)
                        {
                            if (item.Priority == int.Parse(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }


                    if (fil.op.Equals("!="))
                    {
                        foreach (var item in resdetails)
                        {
                            if (item.Priority != int.Parse(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }


                    if (fil.op.Equals(">"))
                    {
                        foreach (var item in resdetails)
                        {
                            if (item.Priority > int.Parse(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }

                    if (fil.op.Equals("<"))
                    {
                        foreach (var item in resdetails)
                        {
                            if (item.Priority < int.Parse(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }


                    resdetails = new List<querydetails>();

                    foreach (var i in newresdetails)
                    {
                        resdetails.Add(i);
                    }

                    newresdetails = new List<querydetails>();
                }

                if (fil.name.Equals("Tester", StringComparison.InvariantCultureIgnoreCase))
                {
                    if (fil.op.Equals("="))
                    {
                        foreach (var item in resdetails)
                        {
                            if (item.Tester.Equals(fil.value, StringComparison.InvariantCultureIgnoreCase))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }

                    if (fil.op.Equals("!="))
                    {
                        foreach (var item in resdetails)
                        {
                            if (!item.Tester.Equals(fil.value, StringComparison.InvariantCultureIgnoreCase))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }

                    resdetails = new List<querydetails>();

                    foreach (var i in newresdetails)
                    {
                        resdetails.Add(i);
                    }

                    newresdetails = new List<querydetails>();
                }

                if (fil.name.Equals("Title", StringComparison.CurrentCultureIgnoreCase))
                {
                    if (fil.op.Equals("="))
                    {
                        foreach (var item in resdetails)
                        {
                            if (res.Title.Contains(fil.value))
                            {
                                newresdetails.Add(item);
                            }
                        }
                    }
                }
            }


            return resdetails;
        }

        /// <summary>
        ///     Assign Tester to select test cases in Query Filter
        /// </summary>
        /// <param name="suiteid"></param>
        /// <param name="selectedTcID"></param>
        /// <param name="name"></param>
        public static void assignTester(int suiteid, List<int> selectedTcID, string name)
        {
            var suitelist = getsubsuites(suiteid);
            suitelist = suitelist.Substring(0, suitelist.Length - 3)+"'";
            var pointCollection = TestPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suitelist));


            foreach (var tpoint in pointCollection)
            {
                if (selectedTcID.Contains(tpoint.TestCaseId))
                {
                    tpoint.AssignedTo = TeamProject.TfsIdentityStore.FindByDisplayName(name);
                 
                    
                    tpoint.Save();
                }
            }
        }

        /// <summary>
        ///     Create Palylist with selected test cases
        /// </summary>
        /// <param name="suiteid"></param>
        /// <param name="selectedTcID"></param>
        /// <param name="name"></param>
        public static void createPlaylist(int suiteid, List<int> selectedTcID, string fileLocation)
        {
            try
            {
                stp.Restart();
                string automatedTestName, playlistFileContent = "<Playlist Version=\"1.0\">";
                var automatedTestList = new ConcurrentBag<string>();
                automatedTestList.AsParallel();
                Parallel.ForEach(selectedTcID, testcaseID =>
                {
                    var pt = wstore.GetWorkItem(testcaseID);
                    automatedTestName = pt.Fields["Automated Test Name"].Value.ToString();
                    lock (automatedTestList)
                    {
                        automatedTestList.Add(automatedTestName);
                    }
                });
                var dedup = automatedTestList.Distinct().ToList();
                stp.Stop();
                AutomationMethodTime = (float) stp.ElapsedMilliseconds/1000;
                stp.Restart();
                foreach (var testName in dedup)
                {
                    playlistFileContent = playlistFileContent + "<Add Test=\"" + testName + "\" />";
                }

                playlistFileContent += "</Playlist>";
                var fs = new FileStream(fileLocation, FileMode.Create);
                using (var writer = new StreamWriter(fs))
                {
                    writer.WriteLine(playlistFileContent);
                }
                stp.Stop();
                AutomationMethodTime1 = (float) stp.ElapsedMilliseconds/1000;

               
                
              
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error generating playlist " + ex);
            }
        }

        #endregion

        #region wpftreelogic

        /// <summary>
        ///     Construct MTM hierarchy as a tree
        /// </summary>
        /// <param name="planid">id of the plan</param>
        /// <param name="mtmtree">The treeview to be populated</param>
        /// <returns></returns>
        public static void getwpfsuitetree(int planid, System.Windows.Controls.TreeView mtmtree,
            bool includeparent = false)
        {
            suitedictionary.Clear();
            TestPlan = TeamProject.TestPlans.Find(planid);
            if (TestPlan == null)
            {
                var exp =
                    new ApplicationException(
                        "Invalid planid.You can get the planid from MTM in the Select Test Plan Window.");
                throw exp;
            }

            for (var i = 0; i < TestPlan.RootSuite.SubSuites.Count; i++)
            {
                if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var suite = (IStaticTestSuite) TestPlan.RootSuite.SubSuites[i];
                    var tvitem = new TreeViewItem();
                    tvitem.Header = suite.Title;
                    tvitem.Tag = suite.Id;


                    if (includeparent)
                    {
                        suitedictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(suite.Id, suite.Title);
                    }
                    getwpfsubsuitetree(suite, tvitem, TestPlan.RootSuite.Title + "\\", includeparent);
                    mtmtree.Items.Add(tvitem);
                }
                else if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("DynamicTestSuite"))
                {
                    var suite = (IDynamicTestSuite) TestPlan.RootSuite.SubSuites[i];
                    var tvitem = new TreeViewItem();
                    tvitem.Header = suite.Title;
                    tvitem.Tag = suite.Id;


                    if (includeparent)
                    {
                        suitedictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(suite.Id, suite.Title);
                    }

                    mtmtree.Items.Add(tvitem);
                }

                else
                {
                    var suite = (IRequirementTestSuite) TestPlan.RootSuite.SubSuites[i];
                    var tvitem = new TreeViewItem();
                    tvitem.Header = suite.Title;
                    tvitem.Tag = suite.Id;


                    if (includeparent)
                    {
                        suitedictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(suite.Id, suite.Title);
                    }

                    mtmtree.Items.Add(tvitem);
                }
            }
        }

        /// <summary>
        ///     Retrun subnode structure
        /// </summary>
        /// <param name="suite">Root suite</param>
        /// <param name="tn">Node to be populated</param>
        /// <returns></returns>
        public static void getwpfsubsuitetree(IStaticTestSuite suite, TreeViewItem mtmtree, string prefix,
            bool includeparent = false)
        {
            for (var i = 0; i < suite.SubSuites.Count; i++)
            {
                if (suite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) suite.SubSuites[i];
                    if (subsuite.SubSuites.Count > 0)
                    {
                        var tvitem = new TreeViewItem();
                        tvitem.Header = subsuite.Title;
                        tvitem.Tag = subsuite.Id;
                        if (includeparent)
                        {
                            suitedictionary.Add(subsuite.Id, prefix + "\\" + suite.Title + "\\" + subsuite.Title);
                        }
                        else
                        {
                            suitedictionary.Add(subsuite.Id, subsuite.Title);
                        }
                        getwpfsubsuitetree(subsuite, tvitem, prefix + "\\" + suite.Title, includeparent);
                        mtmtree.Items.Add(tvitem);
                    }
                    else
                    {
                        var tvitem = new TreeViewItem();
                        tvitem.Header = subsuite.Title;
                        tvitem.Tag = subsuite.Id;
                        if (includeparent)
                        {
                            suitedictionary.Add(subsuite.Id, prefix + "\\" + suite.Title + "\\" + subsuite.Title);
                        }
                        else
                        {
                            suitedictionary.Add(subsuite.Id, subsuite.Title);
                        }
                        mtmtree.Items.Add(tvitem);
                    }
                }
                else if (suite.SubSuites[i].GetType().Name.Equals("DynamicTestSuite"))
                {
                    var dynsuite = (IDynamicTestSuite) suite.SubSuites[i];
                    var tvitem = new TreeViewItem();
                    tvitem.Header = dynsuite.Title;
                    tvitem.Tag = dynsuite.Id;
                    if (includeparent)
                    {
                        suitedictionary.Add(dynsuite.Id, prefix + "\\" + suite.Title + "\\" + dynsuite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(dynsuite.Id, dynsuite.Title);
                    }
                    mtmtree.Items.Add(tvitem);
                }

                else
                {
                    var dynsuite = (IRequirementTestSuite) suite.SubSuites[i];
                    var tvitem = new TreeViewItem();
                    tvitem.Header = dynsuite.Title;
                    tvitem.Tag = dynsuite.Id;
                    if (includeparent)
                    {
                        suitedictionary.Add(dynsuite.Id, prefix + "\\" + suite.Title + "\\" + dynsuite.Title);
                    }
                    else
                    {
                        suitedictionary.Add(dynsuite.Id, dynsuite.Title);
                    }
                    mtmtree.Items.Add(tvitem);
                }
            }
        }

        #endregion
    }
}