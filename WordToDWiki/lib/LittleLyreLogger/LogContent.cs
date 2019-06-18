using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LittleLyreLogger
{
    public class LogContent
    {
        private LoggerParameters.LogSeverity prp_LogSeverity = LoggerParameters.LogSeverity.INFO;
        private string prp_LogMessage = string.Empty;
        private string prp_LogSubject = string.Empty;

        public LoggerParameters.LogSeverity LogSeverity { get => prp_LogSeverity; set => prp_LogSeverity = value; }
        public string LogMessage { get => prp_LogMessage; set => prp_LogMessage = value; }
        public string LogSubject { get => prp_LogSubject; set => prp_LogSubject = value; }

        public LogContent()
        {

        }

        public LogContent(string LSubject, string LMessage, LoggerParameters.LogSeverity LSeverity)
        {
            LogSubject = LSubject;
            LogSeverity = LSeverity;
            LogMessage = LMessage;
        }
    }
}
