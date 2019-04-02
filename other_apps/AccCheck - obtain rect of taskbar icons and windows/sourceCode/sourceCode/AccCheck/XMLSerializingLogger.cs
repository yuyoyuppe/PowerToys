// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;
using System.IO;
using System.Collections;

namespace AccCheck.Logging
{
    public class XMLSerializingLogger : AccumulatingLogger
    {

        public void SerializeSuppressionsToFile(String filename)
        {
            List<LogEvent> events = new List<LogEvent>();
            foreach (LogEvent e in _events)
            {
                if (e.Suppressed)
                {
                    events.Add(e);
                }
            }
            
            XmlSerializer xml = new XmlSerializer(typeof(List<LogEvent>));

            // make sure we release the file so its not in use when we try to read it in again.
            using (TextWriter writer = new StreamWriter(filename))
            {
                xml.Serialize(writer, events);
            }
        }

        
        public void SerializeToFile(String filename, LogFile logFile)
        {
            logFile.PrepareToSerialize(_events);
            
            XmlSerializer xml = new XmlSerializer(typeof(LogFile));
            
            // make sure we release the file so its not in use when we try to read it in again.
            using (TextWriter writer = new StreamWriter(filename))
            {
                xml.Serialize(writer, logFile);
            }
            
        }
        
    }
}
