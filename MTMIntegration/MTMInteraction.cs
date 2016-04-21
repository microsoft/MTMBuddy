//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using TreeView = System.Windows.Forms.TreeView;

namespace MTMIntegration
{
    public static class MtmInteraction
    {

        #region Declarations
        private static Uri _tempUri;

        private static List<string> _tester = new List<string>();
        private static TfsTeamProjectCollection _tfsProjColl;

        private static ITestManagementService _tms;

        private static ITestManagementTeamProject _teamProject;

        /// <summary>
        /// The planid of the plan currently selected.
        /// </summary>
        private static int _testplanid;

        public static int TestPlanId
        {
            get { return _testplanid; }
        }


        /// <summary>
        /// Currently selected TestPlan
        /// </summary>
        private static ITestPlan TestPlan;

        private static readonly Stopwatch Stp = new Stopwatch();

        /// <summary>
        /// Disctionary of suiteids and suites in the current selected plan
        /// </summary>
        private static Dictionary<int, string> SuiteDictionary = new Dictionary<int, string>();

        /// <summary>
        /// Sum of time taken to retrieve all outcomes in the query
        /// </summary>
        public static float OutcomeTime;
        /// <summary>
        /// Sum of time taken to retrieve all Tester names in the query
        /// </summary>
        public static float TesterTime;
        /// <summary>
        /// Sum of time taken to retrieve all priority information in the query
        /// </summary>
        public static float PriorityTime;
        /// <summary>
        /// Sum of time taken to retrieve all Automation status information in the query
        /// </summary>
        public static float AutomationTime;
        /// <summary>
        /// Sum of time taken to retrieve all initializations in the query
        /// </summary>
        public static float Initialize;
        /// <summary>
        /// Sum of time taken to retrieve all testcase id information in the query
        /// </summary>
        public static float TCidtime;
        /// <summary>
        /// Sum of time taken to retrieve all the titles in the query
        /// </summary>
        public static float Titletime;
        /// <summary>
        /// Sum of time taken to retrieve names of Automation Methods in the query
        /// </summary>
        public static float AutomationMethodTime;
        /// <summary>
        /// Sum of time taken to add entries to play list
        /// </summary>
        public static float AutomationPlaylistAddition;

        public static float AutomationTestNameFetchTime;

           

        private static string _projectName;

        /// <summary>
        /// Work Item store of the currect TFS Project.
        /// </summary>
        public static WorkItemStore Wstore;

     

        private const string TestSuiteQuery = "Select * from TestPoint where SuiteId in  ('[#testsuiteid#])";

        /// <summary>
        ///     Name of the currently selected Test Plan
        /// </summary>
        public static string SelectedPlanName { get; set; }

        #endregion Declarations
        /// <summary>
        /// Clear the performance counters used to track the time taken for each field
        /// </summary>
        public static void Clear_performance_counters()
        {
            OutcomeTime = 0;
            TesterTime = 0;
            PriorityTime = 0;
            AutomationTime = 0;
            Initialize = 0;
            TCidtime = 0;
            Titletime = 0;
            AutomationMethodTime = 0;
            AutomationPlaylistAddition = 0;
       
        }

     


        /// <summary>
        ///     Establishes connection with TFS and select the project and the test plan
        /// </summary>
        /// <param name="connUri">Connection URI for the TFS Server</param>
        /// <param name="project">TFS Project Name</param>
        /// <param name="planId">Test Plan ID</param>
        
        public static void Initialize_VSTF(Uri connUri, string project, int planId)
        {
            _tempUri = connUri;
            _testplanid = planId;
            
            _projectName = project;
            _tfsProjColl = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(_tempUri);
            _tfsProjColl.Authenticate();
           
                _tms = _tfsProjColl.GetService<ITestManagementService>();

                Wstore = (WorkItemStore) _tfsProjColl.GetService(typeof (WorkItemStore));
                _teamProject = _tms.GetTeamProject(_projectName);
                TestPlan = _teamProject.TestPlans.Find(TestPlanId);
                SelectedPlanName = TestPlan.Name;

                _teamProject.TfsIdentityStore.Refresh();

                foreach (var i in _teamProject.TfsIdentityStore.AllUserIdentities)
                {
                    _tester.Add(i.DisplayName);
                }

            
            
           
        }

        #region SuiteOperation

        /// <summary>
        ///     Recursively searches for a suite with given id and returns the suite
        /// </summary>
        /// <param name="rootsuite">The root suite to search</param>
        /// <param name="suiteid">Suite id to search for</param>
        /// <returns></returns>
        private static IStaticTestSuite Getsubsuitebyid(IStaticTestSuite rootsuite, int suiteid)
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
                        var retsuite = Getsubsuitebyid(subsuite, suiteid);
                        if (retsuite != null)
                            return retsuite;
                    }
                }
            }
            return null;
        }


        /// <summary>
        /// Returns the list of all subsuites under the given suiteid
        /// </summary>
        /// <param name="suiteid"></param>
        /// <returns>List of all suiteids under the given suite</returns>
        private static string Getsubsuites(int suiteid)
        {
            var suitelist = string.Empty;

            #region GetSuite
            //Find the suite based on suiteid provides
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
                        var retsuite = Getsubsuitebyid(subsuite, suiteid);
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

            #endregion GetSuite
            if (rootsuite.GetType().Name.Equals("StaticTestSuite"))
            {
                suitelist = Getsubsuitelist(rootsuite, suitelist);
            }

            return suitelist;
        }

      /// <summary>
        ///     Retrun subnode structure
        /// </summary>
        /// <param name="suite">Root suite</param>
        /// <param name="tn">Node to be populated</param>
        /// <returns></returns>
        public static void GetSubSuiteTree(IStaticTestSuite suite, TreeNode tn)
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

                        GetSubSuiteTree(subsuite, tn.Nodes[subsuite.Title]);
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
        public static string Getsubsuitelist(IStaticTestSuite suite, string suitelist = "")
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
                        suitelist = suitelist + Getsubsuitelist(subsuite, suitelist);
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
                    var suite = (IStaticTestSuite)TestPlan.RootSuite.SubSuites[i];
                    if (suite.Id.Equals(suiteid))
                    {
                        return true;
                    }


                    if (Getsubsuitetype(suite, suiteid))
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
        public static bool Getsubsuitetype(IStaticTestSuite suite, int suiteid)
        {
            for (var i = 0; i < suite.SubSuites.Count; i++)
            {
                if (suite.SubSuites[i].GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite)suite.SubSuites[i];

                    if (subsuite.Id.Equals(suiteid))
                    {
                        return true;
                    }

                    if (subsuite.SubSuites.Count > 0)
                    {
                        if (Getsubsuitetype(subsuite, suiteid))
                            return true;
                    }
                }
            }

            return false;
        }

        #endregion SuiteOperation


        /// <summary>
        /// </summary>
        /// <param name="suiteid"></param>
        /// <param name="resdetails"></param>
        /// <param name="suitename">List of suiteids</param>
        public static void Getsuiteresults(int suiteid, ConcurrentBag<ResultSummary> resdetails,
            string suitename = "None")
        {
            Stp.Restart();
            var type = getType(suiteid);
            
            ITestPointCollection pointCollection = null;
            if (suiteid != -1)
            {
                if (type)
                {
                    var suitelist = Getsubsuites(suiteid);
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
            var res = new ResultSummary();


            Stp.Stop();
            Initialize = Initialize + (float) Stp.ElapsedMilliseconds/1000;
            foreach (var tpoint in pointCollection)
            {
                res = new ResultSummary();
                Stp.Restart();
                var outcome = string.Empty;
                outcome = tpoint.MostRecentResultOutcome.ToString();

                if (tpoint.MostRecentResult!=null && tpoint.MostRecentResult.State.Equals(TestResultState.InProgress))

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
                Stp.Stop();
                OutcomeTime = OutcomeTime + (float) Stp.ElapsedMilliseconds/1000;
                Stp.Restart();
                res.SuiteName = SuiteDictionary[tpoint.SuiteId];

                res.TCID = tpoint.TestCaseId;
                Stp.Stop();
                TCidtime = TCidtime + (float) Stp.ElapsedMilliseconds/1000;

                try
                {
                    Stp.Restart();
                    res.AutomationStatus = tpoint.IsTestCaseAutomated;
                    Stp.Stop();
                    AutomationTime = AutomationTime + (float) Stp.ElapsedMilliseconds/1000;
                }
                catch (Exception)
                {
                    res.AutomationStatus = false;
                }
               
                try
                {
                    Stp.Restart();
                    if (tpoint.AssignedTo != null)
                    {
                        res.Tester = tpoint.AssignedTo.DisplayName;
                    }
                    else
                    {
                        res.Tester = "Nobody";
                    }
                


                    Stp.Stop();
                    TesterTime = TesterTime + (float) Stp.ElapsedMilliseconds/1000;
                }


                catch (Exception)
                {
                   
                    res.Tester = "Nobody";
                }

                try
                {
                    Stp.Restart();
                    res.Title = tpoint.TestCaseWorkItem.Title;
                    Stp.Stop();
                    Titletime = Titletime + (float) Stp.ElapsedMilliseconds/1000;
                }


                catch (Exception)
                {
                    res.Title = "No Title";
                }
                try
                {
                    Stp.Restart();
                    var tc = tpoint.TestCaseWorkItem;
                    res.Priority = tc.Priority;
                    Stp.Stop();
                    PriorityTime = PriorityTime + (float)Stp.ElapsedMilliseconds / 1000;

                 
                }
                catch (Exception)
                {
                    res.Priority = -1;
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
        public static List<string> GetSuite(int planId)
        {
            TestPlan = _teamProject.TestPlans.Find(planId);
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
        public static Dictionary<int, string> GetPlanIdsDictionary()
        {
            var planName = new Dictionary<int, string>();

            var plan = _teamProject.TestPlans.Query("Select *from TestPlan");
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
        public static string GetPlanName(int planId)
        {
            try {
                var planName = new Dictionary<int, string>();
                planName = GetPlanIdsDictionary();
                var name = string.Empty;
                if (planName.ContainsKey(planId))
                {
                    name = planName[planId];
                    return name;
                }

                return null;
            }catch(Exception)
            {
                return null;
            }
        }

        // get the query filter list

        public class Filter
        {
            public string Name { get; set; }
            public string Op { get; set; }
            public string Value { get; set; }
        }

        #region treelogic

        /// <summary>
        ///     Construct MTM hierarchy as a tree
        /// </summary>
        /// <param name="planid">id of the plan</param>
        /// <param name="mtmtree">The treeview to be populated</param>
        /// <returns></returns>
        public static Dictionary<string, int> Getsuitelist(int planid, string iteration, List<string> selectedSuite)
        {
            var appDictionary = new Dictionary<string, int>();
            appDictionary.Clear();


            TestPlan = _teamProject.TestPlans.Find(planid);
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
                                        if (appDictionary.ContainsKey(suite.Title) == false)
                                        {
                                            appDictionary.Add(suite.Title, subsuite.Id);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }

            return appDictionary;
        }

        /// <summary>
        ///     Construct MTM hierarchy as a tree
        /// </summary>
        /// <param name="planid">id of the plan</param>
        /// <param name="mtmtree">The treeview to be populated</param>
        /// <returns></returns>
        public static void Buildsuitedictionary(int planid, bool includeparent = false)
        {
            SuiteDictionary.Clear();
            TestPlan = _teamProject.TestPlans.Find(planid);
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
                        SuiteDictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(suite.Id, suite.Title);
                    }
                    Buildsubsuitedictionary(suite, TestPlan.RootSuite.Title + "\\", includeparent);
                }

                else if (TestPlan.RootSuite.SubSuites[i].GetType().Name.Equals("DynamicTestSuite"))
                {
                    var suite = (IDynamicTestSuite) TestPlan.RootSuite.SubSuites[i];
                    if (includeparent)
                    {
                        SuiteDictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(suite.Id, suite.Title);
                    }
                }
                else
                {
                    var suite = (IRequirementTestSuite) TestPlan.RootSuite.SubSuites[i];

                    if (includeparent)
                    {
                        SuiteDictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(suite.Id, suite.Title);
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
        public static void Buildsubsuitedictionary(IStaticTestSuite suite, string prefix, bool includeparent = false)
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
                            SuiteDictionary.Add(subsuite.Id, prefix + "\\" + suite.Title + "\\" + subsuite.Title);
                        }
                        else
                        {
                            SuiteDictionary.Add(subsuite.Id, subsuite.Title);
                        }
                        Buildsubsuitedictionary(subsuite, prefix + "\\" + suite.Title, includeparent);
                    }
                    else
                    {
                        if (includeparent)
                        {
                            SuiteDictionary.Add(subsuite.Id, prefix + "\\" + suite.Title + "\\" + subsuite.Title);
                        }
                        else
                        {
                            SuiteDictionary.Add(subsuite.Id, subsuite.Title);
                        }
                    }
                }
                else if (suite.SubSuites[i].GetType().Name.Equals("DynamicTestSuite"))
                {
                    var dynsuite = (IDynamicTestSuite) suite.SubSuites[i];


                    if (includeparent)
                    {
                        SuiteDictionary.Add(dynsuite.Id, prefix + "\\" + suite.Title + "\\" + dynsuite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(dynsuite.Id, dynsuite.Title);
                    }
                }
                else
                {
                    var reqsuite = (IRequirementTestSuite) suite.SubSuites[i];
                    if (includeparent)
                    {
                        SuiteDictionary.Add(reqsuite.Id, prefix + "\\" + suite.Title + "\\" + reqsuite.Title);
                    }

                    else
                    {
                        SuiteDictionary.Add(reqsuite.Id, reqsuite.Title);
                    }
                }
            }
        }


        public static string GetSuitePath(int suiteid)
        {
            if (SuiteDictionary.ContainsKey(suiteid))
            {
                return SuiteDictionary[suiteid];
            }
            return null;
        }

        /// <summary>
        ///     Get the name of Testers
        /// </summary>
        /// <returns>List of tester names</returns>
        public static List<string> GetTester()
        {
            
            return _tester;
        }


     

        /// <summary>
        ///     Assign Tester to select test cases in Query Filter
        /// </summary>
        /// <param name="suiteid"></param>
        /// <param name="selectedTcId"></param>
        /// <param name="name"></param>
        public static void AssignTester(int suiteid, List<int> selectedTcId, string name)
        {
            var suitelist = Getsubsuites(suiteid);
            suitelist = suitelist.Substring(0, suitelist.Length - 3)+"'";
            var pointCollection = TestPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suitelist));


            foreach (var tpoint in pointCollection)
            {
                if (selectedTcId.Contains(tpoint.TestCaseId))
                {
                    tpoint.AssignedTo = _teamProject.TfsIdentityStore.FindByDisplayName(name);
                 
                    
                    tpoint.Save();
                }
            }
        }

        /// <summary>
        ///     Create Palylist with selected test cases
        /// </summary>
        /// <param name="suiteid"></param>
        /// <param name="selectedTcId"></param>
        /// <param name="name"></param>
        public static void CreatePlaylist(int suiteid, List<int> selectedTcId, string fileLocation)
        {
            try
            {
                Stp.Restart();
                string automatedTestName, playlistFileContent = "<Playlist Version=\"1.0\">";
                var automatedTestList = new ConcurrentBag<string>();
                automatedTestList.AsParallel();
                Parallel.ForEach(selectedTcId, testcaseId =>
                {
                    var pt = Wstore.GetWorkItem(testcaseId);
                    automatedTestName = pt.Fields["Automated Test Name"].Value.ToString();
                    lock (automatedTestList)
                    {
                        automatedTestList.Add(automatedTestName);
                    }
                });
                var dedup = automatedTestList.Distinct().ToList();
                Stp.Stop();
                AutomationMethodTime = (float) Stp.ElapsedMilliseconds/1000;
                Stp.Restart();
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
                Stp.Stop();
                AutomationPlaylistAddition = (float) Stp.ElapsedMilliseconds/1000;

               
                
              
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
        public static void Getwpfsuitetree(int planid, System.Windows.Controls.TreeView mtmtree,
            bool includeparent = false)
        {
            SuiteDictionary.Clear();
            TestPlan = _teamProject.TestPlans.Find(planid);
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
                        SuiteDictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(suite.Id, suite.Title);
                    }
                    Getwpfsubsuitetree(suite, tvitem, TestPlan.RootSuite.Title + "\\", includeparent);
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
                        SuiteDictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(suite.Id, suite.Title);
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
                        SuiteDictionary.Add(suite.Id, TestPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(suite.Id, suite.Title);
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
        public static void Getwpfsubsuitetree(IStaticTestSuite suite, TreeViewItem mtmtree, string prefix,
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
                            SuiteDictionary.Add(subsuite.Id, prefix + "\\" + suite.Title + "\\" + subsuite.Title);
                        }
                        else
                        {
                            SuiteDictionary.Add(subsuite.Id, subsuite.Title);
                        }
                        Getwpfsubsuitetree(subsuite, tvitem, prefix + "\\" + suite.Title, includeparent);
                        mtmtree.Items.Add(tvitem);
                    }
                    else
                    {
                        var tvitem = new TreeViewItem();
                        tvitem.Header = subsuite.Title;
                        tvitem.Tag = subsuite.Id;
                        if (includeparent)
                        {
                            SuiteDictionary.Add(subsuite.Id, prefix + "\\" + suite.Title + "\\" + subsuite.Title);
                        }
                        else
                        {
                            SuiteDictionary.Add(subsuite.Id, subsuite.Title);
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
                        SuiteDictionary.Add(dynsuite.Id, prefix + "\\" + suite.Title + "\\" + dynsuite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(dynsuite.Id, dynsuite.Title);
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
                        SuiteDictionary.Add(dynsuite.Id, prefix + "\\" + suite.Title + "\\" + dynsuite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(dynsuite.Id, dynsuite.Title);
                    }
                    mtmtree.Items.Add(tvitem);
                }
            }
        }

        #endregion



        #region ResultUpdate

        private static readonly string TestPointQuery = "Select * from TestPoint where TestCaseId= '[#testcaseid#]'";

        private static readonly string TestPointSuiteQuery =
            "Select * from TestPoint where TestCaseId= '[#testcaseid#]' and SuiteId = '[#testsuiteid#]'";


        //private static ITestRun PlanRun;

        private static ITestCaseResult _testResult;

        private static readonly string PlanBuildNumber = string.Empty;


        public static string UpdateResult(string tcid, string comments, string suiteName, bool usebuildnumber,
            string result = "Passed", int bugid = 0, string attachments = null)
        {
            var tsid = string.Empty;
            var returnstring = "success";
            var finalsuite = string.Empty;
            var debug = string.Empty;


            var he = MtmInteraction.TestPlan.QueryTestPointHierarchy(TestPointQuery.Replace("[#testcaseid#]", tcid));
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
            tsid = GetFinalSuiteId(he, suiteName, out finalsuite).ToString();
            if (tsid.Equals("-1"))
                return "Testcase does not belong to suite";

            var pointCollection =
                MtmInteraction.TestPlan.QueryTestPoints(
                    TestPointSuiteQuery.Replace("[#testcaseid#]", tcid).Replace("[#testsuiteid#]", tsid));
            // like "SELECT * from TestPoint where TestCaseId='185716'
            var planRun = MtmInteraction.TestPlan.CreateTestRun(false);


            planRun.Title = "Test Run";
            if (usebuildnumber)
            {
                planRun.BuildNumber = PlanBuildNumber;
            }


            try
            {
                planRun.AddTestPoint(pointCollection[0], null);
                planRun.Save();
                var tp = pointCollection[0];
                var proj = tp.TestCaseWorkItem.Area;
                _testResult = planRun.QueryResults()[planRun.QueryResults().Count - 1];
                _testResult.DateStarted = DateTime.Now;
                //Console.WriteLine("Test Result count : {0}", PlanRun.QueryResults().Count);
                if (result.Equals("In progress", StringComparison.InvariantCultureIgnoreCase))
                {
                    _testResult.State = TestResultState.InProgress;
                    _testResult.Outcome = TestOutcome.None;
                }
                else if (result.Equals("Active", StringComparison.InvariantCultureIgnoreCase))
                {
                    _testResult.State = TestResultState.Completed;
                    _testResult.Outcome = TestOutcome.None;
                }
                else
                {
                    _testResult.State = TestResultState.Completed;
                    //TestResult.Outcome = TestOutcome.Passed;

                    _testResult.Outcome = (TestOutcome)Enum.Parse(typeof(TestOutcome), result);
                }

                _testResult.ComputerName = Dns.GetHostName();
                _testResult.RunBy = planRun.Owner;
                var tc = _testResult.GetTestCase();


                if (attachments != null && !string.IsNullOrEmpty(attachments))
                {
                    var attachmentlist = attachments.Split(';');
                    foreach (var attachementpath in attachmentlist)
                    {
                        if (File.Exists(attachementpath))
                        {
                            _testResult.Attachments.Add(_testResult.CreateAttachment(attachementpath,
                                SourceFileAction.None));
                            planRun.Attachments.Add(planRun.CreateAttachment(attachementpath, SourceFileAction.None));
                        }
                        else
                        {
                            returnstring = "Success but Attachment-" + attachementpath + " not found";
                        }
                    }
                }

                _testResult.Comment = "Suite:" + suiteName + " " + "Build:" + PlanBuildNumber + " " + comments;
                _testResult.DateCompleted = DateTime.Now;

                if (bugid != 0)
                {
                    try
                    {
                        var tmp = MtmInteraction.Wstore.GetWorkItem(bugid);
                        _testResult.AssociateWorkItem(tmp);
                        _testResult.Save();
                        if (!string.IsNullOrEmpty(attachments))
                        {
                            var attachmentlist = attachments.Split(';');
                            foreach (var attachementpath in attachmentlist)
                            {
                                if (File.Exists(attachementpath))
                                {
                                    var att = new Attachment(attachementpath);
                                    tmp.Attachments.Add(att);
                                }
                            }
                        }

                        var linkTypeEnd = MtmInteraction.Wstore.WorkItemLinkTypes.LinkTypeEnds["Test Case"];
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
                    _testResult.Save();

                return returnstring;
            }

            catch (Exception)
            {
                planRun.Delete();
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
        private static int GetFinalSuiteId(HierarchyEntry testhe, string suitename, out string finalsuitename)
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
                    var suiteid = GetFinalSuiteId(branch, suitename, out finalsuitename);
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