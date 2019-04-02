// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace AccCheck.Logging
{
    /*
     * This logger logs everything to a text file. The file is
     * kept open for the duration of the instance.
     */
    public class TextFileLogger : BaseLogger
    {
        private StreamWriter _streamwriter;

        public TextFileLogger(String filename)
        {
            _streamwriter = new StreamWriter(filename);
            _streamwriter.AutoFlush = true;
        }

        /// <summary>This is called by the BaseLogger if the information should be logged.</summary>
        protected override void WriteToLog(LogEvent logEvent)
        {
            // skipped suppressed entries
            if (!logEvent.Suppressed)
            {
                _streamwriter.WriteLine(logEvent.ToString());
            }
        }
    }
}
