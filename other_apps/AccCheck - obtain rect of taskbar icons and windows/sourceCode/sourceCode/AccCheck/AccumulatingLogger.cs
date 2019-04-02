// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;
using System.Runtime.InteropServices;

namespace AccCheck.Logging
{
    /*
     * This logger stores everything it is told to log in a list, and 
     * can play the log events back to another logger at any time.
     */
    [ComVisible(true), GuidAttribute("05A38449-4815-401b-95E4-CB586367FB2C")]
    [ClassInterface(ClassInterfaceType.None)]
    public class AccumulatingLogger : BaseLogger, IAccumulatingLogger
    {
        protected List<LogEvent> _events = new List<LogEvent>();

        public List<LogEvent> Events
        {
            get 
            { 
                return _events; 
            }
            set 
            { 
                _events = value; 
            }
        }

        private int _sequenceNumber = 0;

        public void Clear()
        {
            lock (_events)
            {
                _events.Clear();
                _sequenceNumber = 0;
                ErrorCount = 0;
                WarningCount= 0;
                InformationalCount = 0;
                TraceCount = 0;
            }
        }

        public void DumpToLogger(ILogger l)
        {
            lock (_events)
            {
                foreach (LogEvent ev in _events)
                {
                    l.Log(ev);
                }
            }
        }

        // Only dumps the events whose EventLevel is present in levels
        public void DumpToLogger(ILogger l, List<EventLevel> levels)
        {
            lock (_events)
            {
                foreach (LogEvent ev in _events)
                {
                    if (levels.Contains(ev.Level))
                    {
                        l.Log(ev);
                    }
                }
            }
        }

        public LogEvent Get(int i)
        {
            lock (_events)
            {
                if (_events.Count > i)
                {
                    return _events[i];
                }
                else
                {
                    return null;
                }
            }
        }
        
        public ILogEvent GetEvent(int i)
        {
            return Get(i);
        }

        /// <summary>This is called by the BaseLogger if the information should be logged.</summary>
        protected override void WriteToLog(LogEvent logEvent)
        {
            lock (_events)
            {
                logEvent.SequenceNumber = _sequenceNumber++;
                _events.Add(logEvent);
            }
        }
    }
}
