// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Xml.Serialization;

namespace AccCheck.Logging
{
    /*
     * This class filters the log events based on an XML file
     * of previous log events.
     */
    public class SuppressionFilterLog : ILogger
    {
        private ILogger _logger;
        private List<LogEvent> _suppressionList = new List<LogEvent>();
        private const string _globalSuppressionFile = "Global.xml";
        
        public SuppressionFilterLog(ILogger logger)
        {
            _logger = logger;
            if (File.Exists(_globalSuppressionFile))
            {
                SuppressionFilterLog.AddSuppressionEntries(_globalSuppressionFile, _suppressionList);
            }
        }

        public void AddSuppressionFile(string suppressionFile)
        {
            if (String.IsNullOrEmpty(suppressionFile))
            {
                throw new ArgumentException("Suppression file is null or empty");
            }
            else if (!File.Exists(suppressionFile))
            {
                throw new FileNotFoundException(string.Format("Could not find the suppression file \"{0}\"", suppressionFile));
            }
            else 
            {
                bool suppressionFileOpened = false;
                try
                {
                    SuppressionFilterLog.AddSuppressionEntries(suppressionFile, _suppressionList);
                    suppressionFileOpened = true;
                }
                catch (InvalidOperationException )
                {
                    // If file does not contain a LogFile object then it might be the old format.
                }

                if (!suppressionFileOpened)
                {
                    try
                    {
                        using (FileStream fs = new FileStream(suppressionFile, FileMode.Open))
                        {
                            XmlSerializer xml = new XmlSerializer(typeof(List<LogEvent>));
                            _suppressionList.AddRange((List<LogEvent>)xml.Deserialize(fs));
                            fs.Flush();
                            fs.Close();
                        }
                    }
                    catch (InvalidOperationException )
                    {
                        // If file does not an array of LogEvents then its not a suppression file
                    }
                }
            }
        }

        public void Clear()
        {
            _suppressionList.Clear();
            if (File.Exists(_globalSuppressionFile))
            {
                SuppressionFilterLog.AddSuppressionEntries(_globalSuppressionFile, _suppressionList);
            }
        }

        static private void AddSuppressionEntries(string suppressionFile, List<LogEvent> suppressionList)
        {
            using (FileStream fs = new FileStream(suppressionFile, FileMode.Open))
            {
                XmlSerializer xml = new XmlSerializer(typeof(LogFile));
                LogFile logFile = (LogFile)xml.Deserialize(fs);
                if (!logFile.IsVersionOk())
                {
                    throw new ArgumentException("Suppression file format is too old");
                }
                
                foreach (LogEvent logEvent in logFile.LogEvents)
                {
                    if (logEvent.Suppressed)
                    {
                        suppressionList.Add(logEvent);
                    }
                }
                fs.Flush();
                fs.Close();
            }
        }
        
        #region ILogger Members

        public void Log(LogEvent e)
        {
            if (_suppressionList != null && _suppressionList.Contains(e))
            {
                e.Suppressed = true;
            }
            _logger.Log(e);
        }

        // forward this to the real logger
        public EventLevel LogLevel
        {
            get { return _logger.LogLevel; }
            set { _logger.LogLevel = value; }
        }

        #endregion
    }
}
