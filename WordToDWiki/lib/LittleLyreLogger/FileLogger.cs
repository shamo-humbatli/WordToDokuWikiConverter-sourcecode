using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace LittleLyreLogger
{
    public class FileLogger : ILittleLyreLogger
    {
        private EventHandler<object> pevnt_OnLogAdded;

        private string prp_LogLine = null;
        private FileStream FStream = null;
        private StreamWriter SWriter = null;

        private int prp_LogListLineCount = 1000;
        private int prp_LogListLineLength = 1000;
        private string prp_LogLineStart = "[+]";
        private int prp_LogSubjectLength = 30;
        private int prp_LogMessageLength = 1000;
        private string prp_LogOutputFolder = "Logs";
        private string prp_LogFileName = "SHLogger";
        
        private string LogElWhiteSpace = " ";
        
        public string LogLine { get => prp_LogLine; }

        public int LogLineCount { get => prp_LogListLineCount; set => prp_LogListLineCount = value < 0 ? 0 : value > 100000 ? 100000 : value; }
        public int LogMessageLength { get => prp_LogMessageLength; set => prp_LogMessageLength = value < 0 ? 0 : value > 100000 ? 100000 : value; }
        public string LogLineStartString { get => prp_LogLineStart; set => prp_LogLineStart = value == null ? string.Empty : value.Length > 100 ? value.Substring(0, 100) : value; }
        public int LogSubjectLength { get => prp_LogSubjectLength; set => prp_LogSubjectLength = value < 0 ? 0 : value > 100000 ? 100000 : value; }
        public LoggerParameters.LogOutput GetOutput => LoggerParameters.LogOutput.ToList;

        public EventHandler<object> OnLogAdded { get => pevnt_OnLogAdded; set => pevnt_OnLogAdded = value; }

        
        public string LogOutputFolder { get => prp_LogOutputFolder; }
        public string LogFileName { get => prp_LogFileName; }

        public FileLogger(string in_OutputFolder, string in_LogFileName)
        {
            prp_LogLine = string.Empty;
            prp_LogOutputFolder = in_OutputFolder;
            prp_LogFileName = in_LogFileName;
            FStream = new FileStream(prp_LogOutputFolder + "\\" + in_LogFileName + "_" + DateTime.Now.ToLongTimeString() + ".log", FileMode.Create);
            SWriter = new StreamWriter(FStream);
        }

        public void AddLog(LogContent LContent)
        {
            if(LContent == null)
            {
                return;
            }

            if (LContent.LogMessage.Length > prp_LogMessageLength)
            {
                LContent.LogMessage = LContent.LogMessage.Substring(0, prp_LogMessageLength - 3);
                LContent.LogMessage += "...";
            }

            if (LContent.LogSubject.Length > prp_LogSubjectLength)
            {
                LContent.LogSubject = LContent.LogSubject.Substring(0, prp_LogSubjectLength - 3);
                LContent.LogSubject += "...";
            }

            string newLogLine = string.Empty;

            newLogLine = LogLineStartString;
            newLogLine += LogElWhiteSpace;
            newLogLine += AddWhiteSpaces(LContent.LogSeverity.ToString(), LoggerParameters.SeverityLongStringLength);
            newLogLine += LogElWhiteSpace;
            newLogLine += AddWhiteSpaces(LContent.LogSubject, LogSubjectLength);
            newLogLine += LogElWhiteSpace;
            newLogLine += LContent.LogMessage;

            SWriter.WriteLine(newLogLine);

            OnLogAdded?.Invoke(this, LogLine);
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
