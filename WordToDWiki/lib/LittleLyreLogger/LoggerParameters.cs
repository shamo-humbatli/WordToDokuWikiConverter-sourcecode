using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace LittleLyreLogger
{
    public abstract class LoggerParameters
    {
        public enum LogSeverity
        {
            INFO = 0,
            DEBUG = 1,
            WARNING = 2,
            ERROR = 3
        }

        public enum LogOutput
        {
            ToList = 1,
            ToFile = 2
        }

        private static int prp_SeverityLongStringLength = LogSeverity.WARNING.ToString().Length;

        public static int SeverityLongStringLength { get => prp_SeverityLongStringLength; }
    }
}
