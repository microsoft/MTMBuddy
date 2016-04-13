namespace ReportingLayer
{
    public class QueryFilter
    {
        public int TestCaseId { get; set; }
        public int Blocked { get; set; }
        public string Tester { get; set; }
        public string Title { get; set; }
        //     public string Priority { get; set; }

        public bool AutomationStatus { get; set; }
    }
}