using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LittleLyreLogger
{
    public interface ILittleLyreLogger
    {
        int LogLineCount { get; set; }
        int LogMessageLength { get; set; }
        int LogSubjectLength { get; set; }
        string LogLineStartString { get; set; }
        LoggerParameters.LogOutput GetOutput { get; }
        EventHandler<object> OnLogAdded { get; set; }
        void AddLog(LogContent LContent);
    }
}
