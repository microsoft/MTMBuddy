//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Threading.Tasks;
using System.Windows.Controls;

using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;


[assembly: CLSCompliant(true)]
namespace MTMIntegration
{
    public static class MtmInteraction
    {

        #region Declarations
        private static Uri _tempUri;

        
        private static TfsTeamProjectCollection _tfsProjColl;

        private static ITestManagementService _tms;

        private static ITestManagementTeamProject _teamProject;


        private static string _projectName;

        /// <summary>
        /// Work Item store of the currect TFS Project.
        /// </summary>
        private static WorkItemStore _tfsStore;
      private const string TestSuiteQuery = "Select * from TestPoint where SuiteId in  ('[#testsuiteid#])";
        /// <summary>
        /// Currently selected _testPlan
        /// </summary>
        private static ITestPlan _testPlan;

        private static readonly Stopwatch Stp = new Stopwatch();

        /// <summary>
        /// Disctionary of suiteids and suites in the current selected plan
        /// </summary>
        private static readonly Dictionary<int, string> SuiteDictionary = new Dictionary<int, string>();


        /// <summary>
        /// ID of the current selected test plan
        /// </summary>
        public static int TestPlanId { get; private set; }
        /// <summary>
        /// Time taken to compute all the outcomes
        /// </summary>
        public static float OutcomeTime { get; private set; }

        /// <summary>
        /// Sum of time taken to retrieve all Tester names in the query
        /// </summary>
        public static float TesterTime { get; private set; }
        /// <summary>
        /// Sum of time taken to retrieve all priority information in the query
        /// </summary>
        public static float PriorityTime { get; private set; }
        /// <summary>
        /// Sum of time taken to retrieve all Automation status information in the query
        /// </summary>
        public static float AutomationTime { get; private set; }
        /// <summary>
        /// Sum of time taken to retrieve all initializations in the query
        /// </summary>
        public static float Initialize { get; private set; }
        /// <summary>
        /// Sum of time taken to retrieve all testcase id information in the query
        /// </summary>
        public static float TcidTime { get; private set; }
        /// <summary>
        /// Sum of time taken to retrieve all the titles in the query
        /// </summary>
        public static float TitleTime { get; private set; }
        /// <summary>
        /// Sum of time taken to retrieve names of Automation Methods in the query
        /// </summary>
        public static float AutomationMethodTime { get; private set; }
        /// <summary>
        /// Sum of time taken to add entries to play list
        /// </summary>
        public static float AutomationPlaylistAddition { get; private set; }

      

        /// <summary>
        ///     Name of the currently selected Test Plan
        /// </summary>
        public static string SelectedPlanName { get; set; }

        #endregion Declarations


        /// <summary>
        /// Clear the performance counters used to track the time taken for each field
        /// </summary>
        public static void ClearPerformanceCounters()
        {
            OutcomeTime = 0;
            TesterTime = 0;
            PriorityTime = 0;
            AutomationTime = 0;
            Initialize = 0;
            TcidTime = 0;
            TitleTime = 0;
            AutomationMethodTime = 0;
            AutomationPlaylistAddition = 0;
       
        }

     


        /// <summary>
        ///     Establishes connection with TFS and select the project and the test plan
        /// </summary>
        /// <param name="connUri">Connection URI for the TFS Server</param>
        /// <param name="project">TFS Project Name</param>
        /// <param name="planId">Test Plan ID</param>
        
        public static void InitializeVstfConnection(Uri connUri, string project, int planId)
        {
            _tempUri = connUri;
            TestPlanId = planId;
            
            _projectName = project;
            _tfsProjColl = TfsTeamProjectCollectionFactory.GetTeamProjectCollection(_tempUri);
            _tfsProjColl.Authenticate();
           
                _tms = _tfsProjColl.GetService<ITestManagementService>();

            _tfsStore = (WorkItemStore) _tfsProjColl.GetService(typeof (WorkItemStore));
                _teamProject = _tms.GetTeamProject(_projectName);
                _testPlan = _teamProject.TestPlans.Find(TestPlanId);
                SelectedPlanName = _testPlan.Name;
     
        }

        #region SuiteOperation

        /// <summary>
        ///     Recursively searches for a suite with given id and returns the suite
        /// </summary>
        /// <param name="rootsuite">The root suite to search</param>
        /// <param name="suiteid">Suite id to search for</param>
        /// <returns>Static suite which matches the suiteid within the rootsuite</returns>
        private static IStaticTestSuite FindSuitebySuiteId(IStaticTestSuite rootsuite, int suiteid)
        {
            foreach (ITestSuiteBase t in rootsuite.SubSuites)
            {
                if (!t.GetType().Name.Equals("StaticTestSuite")) continue;
                var subsuite = (IStaticTestSuite) t;
                if (subsuite.Id.Equals(suiteid))
                    return subsuite;

                if (subsuite.SubSuites.Count > 0)
                {
                    var retsuite = FindSuitebySuiteId(subsuite, suiteid);
                    if (retsuite != null)
                        return retsuite;
                }
            }
            return null;
        }


        /// <summary>
        /// Returns the list of all subsuites under the given suiteId
        /// </summary>
        /// <param name="suiteid"></param>
        /// <returns>List of all suiteids under the given suite</returns>
        private static string FindSuiteAndGetAllChildSuites(int suiteid)
        {
            var suitelist = suiteid.ToString()+ "','";
            IStaticTestSuite selectedSuite = null;

           
            //Traverse the test suite tree to get the suite referenced by suiteId.

            foreach (var subsuite in _testPlan.RootSuite.SubSuites.Where(t => t.GetType().Name.Equals("StaticTestSuite")).Cast<IStaticTestSuite>())
            {
                if (subsuite.Id.Equals(suiteid))
                {
                    selectedSuite = subsuite;
                    suitelist = GetChildSuiteIds(selectedSuite, suitelist);
                }
                if (subsuite.SubSuites.Count > 0)
                {
                    var retsuite = FindSuitebySuiteId(subsuite, suiteid);
                    if (retsuite != null)
                    {
                        selectedSuite = retsuite;
                        suitelist = GetChildSuiteIds(selectedSuite, suitelist);
                    }
                }
                
                
            }


            return suitelist;
        }

     


        /// <summary>
        ///     Returns all the child suite ids under the selected suite. It will recurse if the child suite has children.
        /// </summary>
        /// <param name="suite">Root suite</param>
        /// <param name="suiteList">List of suiteids under this suite</param>
        /// <returns>List of child suite ids</returns>
        public static string GetChildSuiteIds(IStaticTestSuite suite, string suiteList )
        {
            suiteList = suiteList + suite.Id + "','";
            foreach (ITestSuiteBase t in suite.SubSuites)
            {
//need to check for suite type as static suites have children and others dont
                if (t.GetType().Name.Equals("StaticTestSuite"))
                {
                    var subsuite = (IStaticTestSuite) t;
                    if (subsuite.SubSuites.Count > 0)
                    {
                        suiteList = suiteList + subsuite.Id + "','";
                        suiteList = suiteList + GetChildSuiteIds(subsuite, suiteList);
                    }
                    else
                    {
                        suiteList = suiteList + subsuite.Id + "','";
                    }
                }
                else if (t.GetType().Name.Equals("DynamicTestSuite"))
                {
                    var dynsuite = (IDynamicTestSuite) t;

                    suiteList = suiteList + dynsuite.Id + "','";
                }

                else if (t.GetType().Name.Equals("RequirementTestSuite"))
                {
                    var reqsuite = (IRequirementTestSuite) t;

                    suiteList = suiteList + reqsuite.Id + "','";
                }
            }
            return suiteList;
        }

    
       

        #endregion SuiteOperation

        #region GetResultsData

        /// <summary>
        /// Quries the MTM Repository for the testcases under the given suiteid and returns all results under it.
        /// </summary>
        /// <param name="suiteId"></param>
        /// <param name="testResults"></param>
        /// <param name="suiteName">List of suiteids</param>
        public static void GetResultDetails(int suiteId, ConcurrentBag<ResultSummary> testResults,
            string suiteName )
        {
            Stp.Restart();
           
            
            ITestPointCollection pointCollection;
            if (suiteId != -1)
            {
                
                    var suitelist = FindSuiteAndGetAllChildSuites(suiteId);
                    suitelist = suitelist.Substring(0, suitelist.Length - 2);
                    pointCollection = _testPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suitelist));
                    
               
            }
            else
            {
                pointCollection =
                    _testPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suiteName));
            }


         Stp.Stop();
            Initialize = Initialize + (float) Stp.ElapsedMilliseconds/1000;
            //This loop cannot be made parallel as the mtm methods to retrieve result details are not threadsafe
            foreach (var tpoint in pointCollection)
            {
                ResultSummary res = new ResultSummary();
                Stp.Restart();
                var outcome = tpoint.MostRecentResultOutcome.ToString();

                //Special logic for in-progress status. Though MTM shows the teststatus and outcome in similar way, it stores them separately.
                if (tpoint.MostRecentResult!=null && tpoint.MostRecentResult.State.Equals(TestResultState.InProgress))

                {
                    outcome = "In progress";
                }
                    if (outcome.Equals("Blocked", StringComparison.OrdinalIgnoreCase ))
                {
                    //If a result is unblocked, the most recent result still shows blocked hence we have to have special logic. The alternate option of always using tpoint.MostRecentResult.Outcome
                    //is not feasible as it is slower that tpoint.MostRecentResultOutcome hence should be used only when needed
                    res.Outcome = !tpoint.MostRecentResult.Outcome.ToString()
                        .Equals("Blocked", StringComparison.OrdinalIgnoreCase) ? "Active" : "Blocked";
                }
                    //There is no active status but MTM shows Active in UI
                else if (outcome.Equals("Unspecified", StringComparison.OrdinalIgnoreCase) ||
                         outcome.Equals("None", StringComparison.OrdinalIgnoreCase))
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

                res.TcId = tpoint.TestCaseId;
                Stp.Stop();
                TcidTime = TcidTime + (float) Stp.ElapsedMilliseconds/1000;

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
                    res.Tester = tpoint.AssignedTo != null ? tpoint.AssignedTo.DisplayName : "Nobody";
                


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
                    TitleTime = TitleTime + (float) Stp.ElapsedMilliseconds/1000;
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

                lock (testResults)
                {
                    testResults.Add(res);
                }
            }
        }


        #endregion GetResultsData

        #region PlanDetails

        /// <summary>
        ///     Get the dictionary with plan id as key and name as value.
        /// </summary>
        
        public static Dictionary<int, string> PlanIdsDictionary
        {
            get
            {
                var plan = _teamProject.TestPlans.Query("Select * from TestPlan");

                return plan.ToDictionary(i => i.Id, i => i.Name);
            }
        }

        /// <summary>
        ///     Look up function that searches user particular plan name using Id.
        /// </summary>
        /// <param name="planId"></param>
        /// <returns> Name of the plan </returns>
        public static string GetPlanName(int planId)
        {
            
               
                
                if (PlanIdsDictionary.ContainsKey(planId))
                {
                return PlanIdsDictionary[planId];
                    
                }

                return null;
            
        }


        #endregion PlanDetails








        #region TestCaseAssignment

        /// <summary>
        ///     Get the name of Testers
        /// </summary>
        /// <returns>List of tester names</returns>
        public static List<string> GetTester()
        {
            _teamProject.TfsIdentityStore.Refresh();

            return _teamProject.TfsIdentityStore.AllUserIdentities.Select(i => i.DisplayName).ToList();
            
        }


     

        /// <summary>
        ///     Assign Tester to select test cases in Query Filter
        /// </summary>
        /// <param name="suiteId"></param>
        /// <param name="selectedTcId"></param>
        /// <param name="name"></param>
        public static void AssignTester(int suiteId, List<int> selectedTcId, string name)
        {
            var suitelist = FindSuiteAndGetAllChildSuites(suiteId);
            suitelist = suitelist.Substring(0, suitelist.Length - 3)+"'";
            var pointCollection = _testPlan.QueryTestPoints(TestSuiteQuery.Replace("[#testsuiteid#]", suitelist));


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
      /// <param name="selectedTcId"></param>
        /// <param name="fileLocation"></param>
        public static void CreatePlaylist(List<int> selectedTcId, string fileLocation)
        {
            
                Stp.Restart();
                string automatedTestName, playlistFileContent = "<Playlist Version=\"1.0\">";
                var automatedTestList = new ConcurrentBag<string>();
                automatedTestList.AsParallel();
                Parallel.ForEach(selectedTcId, testcaseId =>
                {
                    var pt = _tfsStore.GetWorkItem(testcaseId);
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
            playlistFileContent = dedup.Aggregate(playlistFileContent, (current, testName) => current + "<Add Test=\"" + testName + "\" />");

            playlistFileContent += "</Playlist>";
                var fs = new FileStream(fileLocation, FileMode.Create);
                using (var writer = new StreamWriter(fs))
                {
                    writer.WriteLine(playlistFileContent);
                }
                fs.Dispose();
                Stp.Stop();
                AutomationPlaylistAddition = (float) Stp.ElapsedMilliseconds/1000;

               
                
              
           
        }

        #endregion TestCaseAssignment

        #region wpftreelogic

        /// <summary>
        ///     Construct MTM hierarchy as a tree
        /// </summary>
        /// <param name="planid">id of the plan</param>
        /// <param name="mtmtree">The treeview to be populated</param>
        /// <param name="includeparent"></param>
        /// <returns></returns>
        public static void Getwpfsuitetree(int planid, TreeView mtmtree,
            bool includeparent )
        {
            SuiteDictionary.Clear();
            _testPlan = _teamProject.TestPlans.Find(planid);
            if (_testPlan == null)
            {
                var exp =
                    new ApplicationException(
                        "Invalid planid.You can get the planid from MTM in the Select Test Plan Window.");
                throw exp;
            }

            foreach (ITestSuiteBase t in _testPlan.RootSuite.SubSuites)
            {
                if (t.GetType().Name.Equals("StaticTestSuite"))
                {
                    var suite = (IStaticTestSuite) t;
                    TreeViewItem tvitem = new TreeViewItem {Header = suite.Title, Tag = suite.Id};


                    if (includeparent)
                    {
                        SuiteDictionary.Add(suite.Id, _testPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(suite.Id, suite.Title);
                    }
                    Getwpfsubsuitetree(suite, tvitem, _testPlan.RootSuite.Title + "\\", includeparent);
                    mtmtree.Items.Add(tvitem);
                }
                else if (t.GetType().Name.Equals("DynamicTestSuite"))
                {
                    var suite = (IDynamicTestSuite) t;
                    var tvitem = new TreeViewItem {Header = suite.Title, Tag = suite.Id};


                    if (includeparent)
                    {
                        SuiteDictionary.Add(suite.Id, _testPlan.RootSuite.Title + "\\" + suite.Title);
                    }
                    else
                    {
                        SuiteDictionary.Add(suite.Id, suite.Title);
                    }

                    mtmtree.Items.Add(tvitem);
                }

                else
                {
                    var suite = (IRequirementTestSuite) t;
                    var tvitem = new TreeViewItem {Header = suite.Title, Tag = suite.Id};


                    if (includeparent)
                    {
                        SuiteDictionary.Add(suite.Id, _testPlan.RootSuite.Title + "\\" + suite.Title);
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
     /// Recursively build the WPF tree for a subsuite
     /// </summary>
     /// <param name="suite">The root suite to build the tree</param>
     /// <param name="mtmtree">Tree Object</param>
     /// <param name="prefix">Prefix to get the full path in suitedictionary</param>
     /// <param name="includeparent">Whether to include parent in suite name</param>
        public static void Getwpfsubsuitetree(IStaticTestSuite suite, TreeViewItem mtmtree, string prefix,
            bool includeparent )
     {
         foreach (ITestSuiteBase t in suite.SubSuites)
         {
             if (t.GetType().Name.Equals("StaticTestSuite"))
             {
                 var subsuite = (IStaticTestSuite) t;
                 if (subsuite.SubSuites.Count > 0)
                 {
                     var tvitem = new TreeViewItem {Header = subsuite.Title, Tag = subsuite.Id};
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
                     var tvitem = new TreeViewItem {Header = subsuite.Title, Tag = subsuite.Id};
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
             else if (t.GetType().Name.Equals("DynamicTestSuite"))
             {
                 var dynsuite = (IDynamicTestSuite) t;
                 TreeViewItem tvitem = new TreeViewItem {Header = dynsuite.Title, Tag = dynsuite.Id};
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

                else if (t.GetType().Name.Equals("RequirementTestSuite"))
                {
                 var dynsuite = (IRequirementTestSuite) t;
                 var tvitem = new TreeViewItem {Header = dynsuite.Title, Tag = dynsuite.Id};
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

        #endregion wpftreelogic



        #region ResultUpdate

       


        public static string UpdateResult(string tcid, string comments, string suiteName, bool usebuildnumber,
            string result = "Passed",string planBuildNumber="",float duration=0 )
        {

             const string testPointQuery = "Select * from TestPoint where TestCaseId= '[#testcaseid#]'";

        const string testPointSuiteQuery =
            "Select * from TestPoint where TestCaseId= '[#testcaseid#]' and SuiteId = '[#testsuiteid#]'";
       
       
   
            //Find all instances of the test case
            var he = _testPlan.QueryTestPointHierarchy(testPointQuery.Replace("[#testcaseid#]", tcid));
           
            if (he.Children.Count <= 0)
                return "Testcase not found";
            //Find the child testsuite where the testcase exists under the given suitename
           string tsid = GetFinalSuiteId(he, suiteName).ToString(CultureInfo.InvariantCulture);
            if (tsid.Equals("-1"))
                return "Testcase does not belong to suite";
            
            var pointCollection =
                _testPlan.QueryTestPoints(
                    testPointSuiteQuery.Replace("[#testcaseid#]", tcid).Replace("[#testsuiteid#]", tsid));
           
            var planRun = _testPlan.CreateTestRun(false);


            planRun.Title = "Test Run";
            if (usebuildnumber)
            {
                planRun.BuildNumber = planBuildNumber;
            }


            try
            {
                planRun.AddTestPoint(pointCollection[0], null);
                planRun.Save();
                ITestCaseResult testResult= planRun.QueryResults()[planRun.QueryResults().Count - 1];
                testResult.DateStarted = DateTime.Now;
               
                if (result.Equals("In progress", StringComparison.OrdinalIgnoreCase))
                {
                    testResult.State = TestResultState.InProgress;
                    testResult.Outcome = TestOutcome.None;
                }
                else if (result.Equals("Active", StringComparison.OrdinalIgnoreCase))
                {
                    testResult.State = TestResultState.Unspecified;
                    testResult.Outcome = TestOutcome.None;
                }
                else
                {
                    testResult.State = TestResultState.Completed;
                    //TestResult.Outcome = TestOutcome.Passed;

                    testResult.Outcome = (TestOutcome)Enum.Parse(typeof(TestOutcome), result);
                }

                testResult.ComputerName = Dns.GetHostName();
                testResult.RunBy = planRun.Owner;
               


                testResult.Duration =TimeSpan.FromSeconds(duration);

                testResult.Comment = "Suite:" + suiteName + " " + "Build:" + planBuildNumber + " " + comments;
                testResult.DateCompleted = DateTime.Now;

               testResult.Save();

                return "Success";
            }

            catch (Exception)
            {
                planRun.Delete();
              
                throw;
            }
        }

        /// <summary>
        ///     This method returns the id of the first suite that contains the test case to be passed. This has been added to pass
        ///     test cases under specific sub suites like Pass 2 or Pass 1.
        /// </summary>
        /// <param name="testhe"> The first node. One of its sub suites contains the case to be passed.</param>
        /// <param name="suitename">Name of the suite can be a bracnh or leaf</param>
     
        /// <returns>suite id</returns>
        private static int GetFinalSuiteId(HierarchyEntry testhe, string suitename)
        {
            
            foreach (var branch in testhe.Children)
            {
                if (branch.SuiteTitle.Equals(suitename, StringComparison.OrdinalIgnoreCase))
                {
                    if (branch.Children.Count > 0)
                    {
                        var temp = branch;
                        while (temp.Children.Count > 0)
                        {
                            temp = temp.Children[0];
                           
                        }

                        return temp.SuiteId;
                    }
                   
                    return branch.SuiteId;
                }
                if (branch.Children.Count > 0)
                {
                    var suiteid = GetFinalSuiteId(branch, suitename);
                    if (suiteid > 0)
                    {
                        return suiteid;
                    }
                }
            }
            
            return -1;
        }

        #endregion ResultUpdate
    }
}