// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;

namespace AccCheck.Logging
{
    /*
     * This logger allows chaining of loggers, so that the output 
     * of a log can go several places.
     */
    internal class LogList : ILogger, ILoggerStatistics
    {
        private List<ILogger> _loggers = new List<ILogger>();
        private EventLevel _logLevel = EventLevel.Warning;
        private ILoggerStatistics _stats = null;
        private ILoggerCallback _loggerCallback = null;
        
        public void AddLogger(ILogger l)
        {
            _loggers.Add(l);
        }

        #region ILogger Members

        public void Log(LogEvent e)
        {
            // we don't check the log level, since some loggers
            // ignore that, and logs everything (e.g. AccumulatingLogger)
            foreach (ILogger l in _loggers)
            {
                l.Log(e);
            }

            
            if (_loggerCallback != null)
            {
                if (e.AutomationElement != null)
                {
                    _loggerCallback.LogEntryNotification(e.AutomationElement, null, e);
                }
                else
                {
                    if (e.Accessible != null)
                    {
                        _loggerCallback.LogEntryNotification(e.Accessible.IAccessible, e.Accessible.ChildId, e);
                    }
                }
            }
        }

        public ILogger GetLogger(Type loggerType)
        {
            // we just return the first logger that comes up with the input type
            // as we don’t expect a scenario where the same type of logger added to VM 
            // would have different entries.  If there are two they would be identical. 
            foreach (ILogger log in _loggers)
            {
                if (log.GetType().Equals(loggerType))
                {
                    return log;
                }
            }

            return null; 
        }
        #endregion

        public EventLevel LogLevel
        {
            get 
            {
                return _logLevel;
            }
            set 
            { 
                foreach (ILogger logger in _loggers)
                {
                    logger.LogLevel = value;
                }
            }
        }
        
        public int Priority
        {
            set 
            { 
                foreach (ILogger log in _loggers)
                {
                    BaseLogger baseLog = log as BaseLogger;
                    if (baseLog != null)
                    {
                        baseLog.Priority = value;
                    }
                }
            }
        }
        
#region ILoggerStatistics Members
        public int InformationalCount
        {
            get 
            { 
                GetILoggerStatistics();
                return _stats.InformationalCount; 
            }
        }

        public int TraceCount
        {
            get 
            { 
                GetILoggerStatistics();
                return _stats.TraceCount; 
            }
        }

        public int ErrorCount
        {
            get 
            { 
                GetILoggerStatistics();
                return _stats.ErrorCount; 
            }
        }

        public int WarningCount
        {
            get 
            { 
                GetILoggerStatistics();
                return _stats.WarningCount; 
            }
        }

#endregion

        internal ILoggerCallback LoggerCallback
        { 
            get 
            { 
                return _loggerCallback; 
            }
            
            set 
            { 
                _loggerCallback = value; 
            } 
        }

        private void GetILoggerStatistics()
        {
            if (_stats == null)
            {
                foreach (ILogger l in _loggers)
                {
                    ILoggerStatistics stats = l as ILoggerStatistics;
                    if (stats != null)
                    {
                        // the first logger in the list will give us the right results
                        _stats = stats;
                        break;
                    }
                }
            }
        }
        

    }
}
