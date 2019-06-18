using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LittleLyreLogger
{
    public class ListLogger : ILittleLyreLogger
    {
        private EventHandler<object> pevnt_OnLogAdded;

        private List<string> prp_LogList = null;
        private int prp_LogListLineCount = 1000;
        private int prp_LogMessageLength = 1000;
        private string prp_LogLineStart = "[+]";
        private int prp_LogSubjectLength = 30;
        
        private string LogElWhiteSpace = " ";
        
        public List<string> LogList { get => prp_LogList; }

        public int LogLineCount { get => prp_LogListLineCount; set => prp_LogListLineCount = value < 0 ? 0 : value > 100000 ? 100000 : value; }
        public int LogMessageLength { get => prp_LogMessageLength; set => prp_LogMessageLength = value < 0 ? 0 : value > 100000 ? 100000 : value; }
        public string LogLineStartString { get => prp_LogLineStart; set => prp_LogLineStart = value == null ? string.Empty : value.Length > 100 ? value.Substring(0, 100) : value; }
        public int LogSubjectLength { get => prp_LogSubjectLength; set => prp_LogSubjectLength = value < 0 ? 0 : value > 100000 ? 100000 : value; }

        public LoggerParameters.LogOutput GetOutput => LoggerParameters.LogOutput.ToList;

        public EventHandler<object> OnLogAdded { get => pevnt_OnLogAdded; set => pevnt_OnLogAdded = value; }

        

        public ListLogger()
        {
            prp_LogList = new List<string>();
        }

        public void AddLog(LogContent LContent)
        {
            if(LContent == null)
            {
                return;
            }

            if(LContent.LogMessage.Length > prp_LogMessageLength)
            {
                LContent.LogMessage = LContent.LogMessage.Substring(0, prp_LogMessageLength - 3);
                LContent.LogMessage += "...";
            }

            if (LContent.LogSubject.Length > prp_LogSubjectLength)
            {
                LContent.LogSubject = LContent.LogSubject.Substring(0, prp_LogSubjectLength - 3);
                LContent.LogSubject += "...";
            }

            prp_LogList.Add(LogLineStartString);
            prp_LogList.Add(LogElWhiteSpace);
            prp_LogList.Add(AddWhiteSpaces(LContent.LogSeverity.ToString(), LoggerParameters.SeverityLongStringLength));
            prp_LogList.Add(LogElWhiteSpace);
            prp_LogList.Add(AddWhiteSpaces(LContent.LogSubject, LogSubjectLength));
            prp_LogList.Add(LogElWhiteSpace);
            prp_LogList.Add(LContent.LogMessage);
            
            prp_LogList.Add(Environment.NewLine);

            if (prp_LogList.Count > prp_LogListLineCount)
            {
                while (prp_LogList.Count > this.prp_LogListLineCount)
                {
                    prp_LogList.RemoveAt(0);
                }
            }

            OnLogAdded?.Invoke(this, LogList);
        }

        private string AddWhiteSpaces(string in_Content, int MaxLength)
        {
            string out_Rslt = "[";

            for (int wc = 0; wc < MaxLength - in_Content.Length; wc++)
            {
                out_Rslt += " ";
            }

            out_Rslt += in_Content;
            out_Rslt += "]";
            return out_Rslt;
        }
    }
}
