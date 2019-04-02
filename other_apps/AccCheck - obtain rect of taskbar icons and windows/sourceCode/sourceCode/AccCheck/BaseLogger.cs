// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;

namespace AccCheck.Logging
{
    public abstract class BaseLogger : ILogger, ILoggerStatistics
    {
        EventLevel _loglevel = EventLevel.Information;
        int _cErrors = 0;
        int _cWarnings = 0;
        int _cInformation = 0;
        int _cTrace = 0;
        int _priority = 3;

        public EventLevel LogLevel
        {
            get { return _loglevel; }
            set { _loglevel = value; }
        }

        public int InformationalCount
        {
            get { return _cInformation; }
            set { _cInformation = value; }
        }

        public int TraceCount
        {
            get { return _cTrace; }
            set { _cTrace = value; }
        }

        public int ErrorCount
        {
            get { return _cErrors; }
            set { _cErrors = value; }
        }

        public int WarningCount
        {
            get { return _cWarnings; }
            set { _cErrors = value; }
        }

        public int Priority
        {
            get { return _priority; }
            set { _priority = value; }
        }

        /// <summary>Objects that derive from this need to override this and do their custom logging</summary>
        protected abstract void WriteToLog(LogEvent logEvent);

        #region ILogger Members

        /// <summary>Manages the different EventLevel counts; Calls method to do the actual logging of the logEvent.</summary>
        public virtual void Log(LogEvent logEvent)
        {
            // only log events that meet the bar
            if (logEvent.Level >= _loglevel && logEvent.Priority <= _priority)
            {
                WriteToLog(logEvent);

                switch (logEvent.Level)
                {
                    case EventLevel.Error:
                        {
                            _cErrors++;
                            break;
                        }
                    case EventLevel.Warning:
                        {
                            _cWarnings++;
                            break;
                        }
                    case EventLevel.Information:
                        {
                            _cInformation++;
                            break;
                        }
                    case EventLevel.Trace:
                        {
                            _cTrace++;
                            break;
                        }
                }
            }
        }

        #endregion
    }
}
