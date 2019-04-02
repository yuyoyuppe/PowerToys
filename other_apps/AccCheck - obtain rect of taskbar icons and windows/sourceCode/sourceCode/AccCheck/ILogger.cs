// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

namespace AccCheck.Logging
{
    // Signifies the severity of a log event
    [ComVisible(true)]
    public enum EventLevel { Trace, Information, Warning, Error, Pass };

    [ComVisible(true), GuidAttribute("137AD71F-4657-4363-C9E4-D6D734F1F530")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ILogger
    {
        /// <summary>Method that will log an event to the derived logger (ie. ConsoleLogger)</summary>
        void Log(LogEvent logEvent);

        EventLevel LogLevel
        {
            get;
            set;
        }
    }

    [ComVisible(true), GuidAttribute("A088D8BC-DAB3-4165-A3C2-CEC7EDB98287")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IAccumulatingLogger : ILogger, ILoggerStatistics
    {
        void Clear();
        void DumpToLogger(ILogger l);

        ILogEvent GetEvent(int i);
    }

    [ComVisible(true), GuidAttribute("D7922009-A860-492c-8E55-1C44A184E217")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface ILoggerStatistics
    {
        int InformationalCount
        {
            get;
        }

        int TraceCount
        {
            get;
        }

        int ErrorCount
        {
            get;
        }

        int WarningCount
        {
            get;
        }
        
    }

    
    [ComVisible(true), GuidAttribute("53E59F12-B117-4965-98A0-C7E29D1ADF5E")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public interface ILoggerCallback
    {
        void LogEntryNotification(object object1, object object2, ILogEvent logEvent);
    }
        
}

