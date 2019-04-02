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
    /// <summary>
    /// UsageReporter is the entry point for reporting statistics from
    /// AccChecker execution.  We can report on errors/warnings found,
    /// which binaries were scanned, and whether UI was shown.  Other
    /// statistics considered interesting could also be added to this interface.
    /// </summary>
    public class UsageReporter
    {
        // Create the singleton instance of UsageReporter
        // This can be changed for outside-of-Windows environments
        // to simply create an instance of this class.
        private static readonly UsageReporter instance = CreateReporter();

        private static UsageReporter CreateReporter()
        {
            return new UsageReporter();
        }

        protected UsageReporter()
        {
        }

        /// <summary>
        /// Singleton instance accessor, which can be used anywhere
        /// in AccChecker.
        /// </summary>
        public static UsageReporter Instance
        {
            get
            {
                return instance;
            }
        }

        /// <summary>
        /// Are we showing UI for this usage?
        /// </summary>
        public virtual bool ShowingUI
        {
            get
            {
                return _showingUI;
            }
            set
            {
                _showingUI = value;
            }
        }

        /// <summary>
        /// Notify reporter of the start of a window scan
        /// </summary>
        /// <param name="hwnd"></param>
        public virtual void BeginWindowScan(
            Verification.VerificationManager vm, 
            IntPtr hwnd)
        {
        }

        /// <summary>
        /// Notify reporter of the end of a window scan
        /// </summary>
        /// <param name="hwnd"></param>
        public virtual void EndWindowScan(
            Verification.VerificationManager vm, 
            IntPtr hwnd)
        {
        }

        /// <summary>
        /// Report UI complexity statistics for a scan 
        /// </summary>
        /// <param name="objectCount">Number of objects scanned</param>
        /// <param name="treeDepth">Tree depth of scanned UI</param>
        public virtual void ReportUIComplexity(uint objectCount, uint treeDepth)
        {
        }

        protected bool _showingUI = false;
    }
}
