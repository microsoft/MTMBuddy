using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.TeamFoundation.Framework;
using Microsoft.TeamFoundation.Framework.Client;
using Microsoft.TeamFoundation.TestManagement.Client;
using Microsoft.TeamFoundation.Client;
using System.Collections.Specialized;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System.Collections;
using System.Configuration;
using System.Net;
using System.Security;
using System.Diagnostics;
using System.Windows.Forms;
using System.Threading;
using System.Collections.Concurrent;

namespace MTMIntegration
{
/// <summary>
	/// Result Summary object containing plan,suite, outcome, Runby and comments
	/// </summary>
	public class resultdetail
	{
		/// <summary>
		/// The testplan for the result
		/// </summary>
		private string testplan;
		public string TestPlan
		{
			get
			{
				return testplan;
			}
			set
			{
				testplan = value;
			}
		}
		/// <summary>
		/// The testsuite for the result
		/// </summary>
		private string suitename;

		public string SuiteName
		{
			get
			{
				return suitename;
			}
			set
			{
				suitename = value;
			}
		}
        /// <summary>
        /// The testcaseid for the result
        /// </summary>
        private int tcid;

        public int TcId
        {
            get
            {
                return tcid;
            }
            set
            {
                tcid = value;
            }
        }

        /// <summary>
        /// Title as per MTM
        /// </summary>
        public string Title { get; set; }
		
        /// <summary>
        /// The priority for the testcase
        /// </summary>
        private int priority;

        public int Priority
        {
            get
            {
                return priority;
            }
            set
            {
                priority = value;
            }
        }
        /// <summary>
        /// The Outcome of the result
        /// </summary>
		private string outcome;
		public string Outcome
		{
			get
			{
				return outcome;
			}
			set
			{
				outcome = value;
			}
		}
		/// <summary>
		/// The identity of the person who updated the result
		/// </summary>
		private string runby;
		public string RunBy
		{
			get
			{
				return runby;
			}
			set
			{
				runby = value;
			}
		}
		
		/// <summary>
		/// Owner of the test case
		/// </summary>
		private string owner;
		public string Owner
		{
			get
			{
				return owner;
			}
			set
			{
				owner = value;
			}
		}
		/// <summary>
		/// Date of the result
		/// </summary>
		public DateTime date;
		public DateTime Date
		{
			get
			{
				return date;
			}
			set
			{
				date = value;
			}
		}
		/// <summary>
		/// Tester as per MTM
		/// </summary>
		private string tester;

		public string Tester
		{
			get
			{
				return tester;
			}
			set
			{
				tester = value;
			}
		}
		/// <summary>
		/// Automation status of the test case
		/// </summary>
		public bool automationstatus;
		public bool AutomationStatus
		{
			get
			{
				return automationstatus;
			}
			set
			{
				automationstatus = value;
			}
		}
		
	   
		
		
	}

	  /// <summary>
	/// Result Summary object containing plan,suite, outcome, Runby and comments
	/// </summary>
	public class resultsummary
	{
		/// <summary>
		/// The testcaseid for the result
		/// </summary>
		private int tcid;

		public int TcId
		{
			get 
			{
				return tcid; 
			}
			set 
			{
				tcid = value; 
			}
		}
		 
		/// <summary>
		/// The priority for the testcase
		/// </summary>
		private int priority;

		public int Priority
		{
			get
			{
				return priority;
			}
			set
			{
				priority = value;
			}
		}
		/// <summary>
		/// The Outcome of the result
		/// </summary>
		private string outcome;

		public string Outcome
		{
			get
			{
				return outcome;
			}
			set
			{
				outcome = value;
			}
		}
		/// <summary>
		/// Date of the result
		/// </summary>
		private DateTime date;

		public DateTime Date
		{
			get
			{
				return date;
			}
			set
			{
				date = value;
			}
		}

		/// <summary>
		/// Tester as per MTM
		/// </summary>
		private string tester;

		public string Tester
		{
			get
			{
				return tester;
			}
			set
			{
				tester = value;
			}
		}
		/// <summary>
		/// Suitename
		/// </summary>
		private string suitename;

		public string SuiteName
		{
			get
			{
				return suitename;
			}
			set
			{
				suitename = value;
			}
		}

        public bool AutomationStatus { get; set; }
     
	}
}