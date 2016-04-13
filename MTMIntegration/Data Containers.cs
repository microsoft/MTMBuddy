using System;

namespace MTMIntegration
{
    /// <summary>
    ///     Result Summary object containing plan,suite, outcome, Runby and comments
    /// </summary>
    public class resultdetail
    {
        /// <summary>
        ///     Automation status of the test case
        /// </summary>
        public bool automationstatus;

        /// <summary>
        ///     Date of the result
        /// </summary>
        public DateTime date;

        /// <summary>
        ///     The Outcome of the result
        /// </summary>
        private string outcome;

        /// <summary>
        ///     Owner of the test case
        /// </summary>
        private string owner;

        /// <summary>
        ///     The priority for the testcase
        /// </summary>
        private int priority;

        /// <summary>
        ///     The identity of the person who updated the result
        /// </summary>
        private string runby;

        /// <summary>
        ///     The testsuite for the result
        /// </summary>
        private string suitename;

        /// <summary>
        ///     The testcaseid for the result
        /// </summary>
        private int tcid;

        /// <summary>
        ///     Tester as per MTM
        /// </summary>
        private string tester;

        /// <summary>
        ///     The testplan for the result
        /// </summary>
        private string testplan;

        public string TestPlan
        {
            get { return testplan; }
            set { testplan = value; }
        }

        public string SuiteName
        {
            get { return suitename; }
            set { suitename = value; }
        }

        public int TcId
        {
            get { return tcid; }
            set { tcid = value; }
        }

        /// <summary>
        ///     Title as per MTM
        /// </summary>
        public string Title { get; set; }

        public int Priority
        {
            get { return priority; }
            set { priority = value; }
        }

        public string Outcome
        {
            get { return outcome; }
            set { outcome = value; }
        }

        public string RunBy
        {
            get { return runby; }
            set { runby = value; }
        }

        public string Owner
        {
            get { return owner; }
            set { owner = value; }
        }

        public DateTime Date
        {
            get { return date; }
            set { date = value; }
        }

        public string Tester
        {
            get { return tester; }
            set { tester = value; }
        }

        public bool AutomationStatus
        {
            get { return automationstatus; }
            set { automationstatus = value; }
        }
    }

    /// <summary>
    ///     Result Summary object containing plan,suite, outcome, Runby and comments
    /// </summary>
    public class resultsummary
    {
        /// <summary>
        ///     Date of the result
        /// </summary>
        private DateTime date;

        /// <summary>
        ///     The Outcome of the result
        /// </summary>
        private string outcome;

        /// <summary>
        ///     The priority for the testcase
        /// </summary>
        private int priority;

        /// <summary>
        ///     Suitename
        /// </summary>
        private string suitename;

        /// <summary>
        ///     The testcaseid for the result
        /// </summary>
        private int tcid;

        /// <summary>
        ///     Tester as per MTM
        /// </summary>
        private string tester;

        public int TcId
        {
            get { return tcid; }
            set { tcid = value; }
        }

        public int Priority
        {
            get { return priority; }
            set { priority = value; }
        }

        public string Outcome
        {
            get { return outcome; }
            set { outcome = value; }
        }

        public DateTime Date
        {
            get { return date; }
            set { date = value; }
        }

        public string Tester
        {
            get { return tester; }
            set { tester = value; }
        }

        public string SuiteName
        {
            get { return suitename; }
            set { suitename = value; }
        }

        public bool AutomationStatus { get; set; }

        public string Title { get; set; }

        public string AutomationTestName { get; set; }
    }

    public class querydetails
    {
        /// <summary>
        ///     The Outcome of the result
        /// </summary>
        private string outcome;

        /// <summary>
        ///     The priority for the testcase
        /// </summary>
        private int priority;

        /// <summary>
        ///     Date of the result
        /// </summary>
        /// <summary>
        ///     Tester as per MTM
        /// </summary>
        private string tester;

        public bool Check { get; set; }

        public int TcId { get; set; }

        public int Priority
        {
            get { return priority; }
            set { priority = value; }
        }

        public string Outcome
        {
            get { return outcome; }
            set { outcome = value; }
        }

        public string Tester
        {
            get { return tester; }
            set { tester = value; }
        }

        /// <summary>
        ///     Suitename
        /// </summary>
        public string Title { get; set; }

        public bool AutomationStatus { get; set; }
    }

    public class PlanDetails
    {
        public string Application { get; set; }
        public string Iteration { get; set; }
        public int TcID { get; set; }
        public int Priority { get; set; }
        public string Outcome { get; set; }
        public bool Automation { get; set; }
        public string Tester { get; set; }
        public string FullPath { get; set; }
    }
}