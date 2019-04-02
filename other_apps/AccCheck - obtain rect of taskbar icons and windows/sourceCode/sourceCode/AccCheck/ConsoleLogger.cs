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
    /*
     * Simple console logger
     */
    [ComVisible(true), GuidAttribute("1178A631-C55F-4823-BE46-54FFED6251B8")]
    [ClassInterface(ClassInterfaceType.None)]
    public class ConsoleLogger : BaseLogger
    {
        /// <summary>This is called by the BaseLogger if the information should be logged.</summary>
        protected override void WriteToLog(LogEvent logEvent)
        {
            if (logEvent.Suppressed)
            {
                // skipped suppressed entries
                return;
            }
            
            ConsoleColor curColor = Console.ForegroundColor;

            switch (logEvent.Level)
            {
                case EventLevel.Error:
                    Console.ForegroundColor = ConsoleColor.Red;
                    break;
                case EventLevel.Warning:
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    break;

            }

            Console.WriteLine(logEvent.ToString());

            Console.ForegroundColor = curColor;
        }
    }
}
