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
    public class ResultSummary
    {
        /// <summary>
        ///     The testcaseid for the result
        /// </summary>
        public int TcId { get; set; }

        /// <summary>
        ///     The priority for the testcase
        /// </summary>
        public int Priority { get; set; }

        /// <summary>
        ///     The Outcome of the result. This is fetched from MTM as is however in few cases we apply logic.
        ///     1. If the test result state is 'In Progress', we mark the outcome as 'In Progress'
        ///     2. In case of 'None' or 'Not Specified', we mark it as Active.
        /// </summary>
        public string Outcome { get; set; }

        /// <summary>
        ///     Date of the result. This field is currently empty as fetching date results in performance issues.
        /// </summary>
        public DateTime Date { get; set; }

        /// <summary>
        ///     Tester as per MTM. If no Tester is specified, we show 'Nobody'
        /// </summary>
        public string Tester { get; set; }

        /// <summary>
        ///     Full path of the testsuite including plan name
        /// </summary>
        public string SuiteName { get; set; }

        /// <summary>
        /// Automation status of the test case
        /// </summary>
        public bool AutomationStatus { get; set; }

        /// <summary>
        /// Title of the test case
        /// </summary>
        public string Title { get; set; }

    }

    public class ResultHistorySummary
    {
        
        /// <summary>
        /// Name of the plan for this instance of testcase
        /// </summary>
        public string PlanName { get; set; }

        /// <summary>
        ///     TestSuite Name
        /// </summary>
        public string SuiteName { get; set; }

        /// <summary>
        ///     The testcaseid for the result
        /// </summary>
        public int TcId { get; set; }

       

        /// <summary>
        ///     The Outcome of the result. This is fetched from MTM as is however in few cases we apply logic.
        ///     1. If the test result state is 'In Progress', we mark the outcome as 'In Progress'
        ///     2. In case of 'None' or 'Not Specified', we mark it as Active.
        /// </summary>
        public string Outcome { get; set; }

        /// <summary>
        ///     Date of the result. This field is currently empty as fetching date results in performance issues.
        /// </summary>
        public DateTime Date { get; set; }


 
        /// <summary>
        /// Identity of the person who actually executed the case
        /// </summary>
        public string RunBy { get; set; }


        /// <summary>
        ///     Tester as per MTM. If no Tester is specified, we show 'Nobody'
        /// </summary>
        public string Tester { get; set; }

        /// <summary>
        /// Automation status of the test case
        /// </summary>
        public bool AutomationStatus { get; set; }

        /// <summary>
        ///     The priority for the testcase
        /// </summary>
        public int Priority { get; set; }

        /// <summary>
        /// Title of the test case
        /// </summary>
        public string Title { get; set; }

        /// <summary>
        /// Comments and failure reasons captured in the test result
        /// </summary>
        public string Comment { get; set; }

        /// <summary>
        /// Bugs Associated with test result
        /// </summary>
        public string BugsAssociated { get; set; }

    }

    public class TestFailureDetail
    {
        public int TCId { get; set; }
        public int iterationid { get; set; }
        public int ActionId { get; set; }
        public string ActionOutcome { get; set; }
        public string StepComment { get; set; }
        public string StepTitle { get; set; }

    }

}