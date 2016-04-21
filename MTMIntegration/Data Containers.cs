//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 

using System;

namespace MTMIntegration
{
 

    /// <summary>
    ///     Result Summary object containing plan,suite, outcome, Runby and comments
    /// </summary>
    public class resultsummary
    {
        /// <summary>
        ///     Date of the result.
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

        /// <summary>
        ///     The testcaseid for the result
        /// </summary>
        public int TcId
        {
            get { return tcid; }
            set { tcid = value; }
        }

        /// <summary>
        ///     The priority for the testcase
        /// </summary>
        public int Priority
        {
            get { return priority; }
            set { priority = value; }
        }

        /// <summary>
        ///     The Outcome of the result. This is fetched from MTM as is however in few cases we apply logic.
        ///     1. If the test result state is 'In Progress', we mark the outcome as 'In Progress'
        ///     2. In case of 'None' or 'Not Specified', we mark it as Active.
        /// </summary>
        public string Outcome
        {
            get { return outcome; }
            set { outcome = value; }
        }

        /// <summary>
        ///     Date of the result. This field is currently empty as fetching date results in performance issues.
        /// </summary>
        public DateTime Date
        {
            get { return date; }
            set { date = value; }
        }
        /// <summary>
        ///     Tester as per MTM. If no Tester is specified, we show 'Nobody'
        /// </summary>
        public string Tester
        {
            get { return tester; }
            set { tester = value; }
        }

        /// <summary>
        ///     Full path of the testsuite including plan name
        /// </summary>
        public string SuiteName
        {
            get { return suitename; }
            set { suitename = value; }
        }

        /// <summary>
        /// Automation status of the test case
        /// </summary>
        public bool AutomationStatus { get; set; }

        /// <summary>
        /// Title of the test case
        /// </summary>
        public string Title { get; set; }

    }

 
    
}