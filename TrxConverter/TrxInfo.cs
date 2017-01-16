using System;

namespace TrxConverter
{
    public class TrxInfo
    {
        public string ExecutionId { get; set; }
        public string TestId { get; set; }
        public string TestName { get; set; }
        public string Duration { get; set; }
        public string Outcome { get; set; }
        public string StdOut { get; set; }
        public string ErrorMessage { get; set; }
        public string StackTrace { get; set; }
    }
}