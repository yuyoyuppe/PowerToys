// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Threading;
using System.Collections.Generic;
using System.Windows.Forms;
using AccCheck;

namespace AccCheckUI
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            try
            {
                Win32API.SetProcessDPIAware();
            }
            catch (EntryPointNotFoundException)
            {
                // On pre-Vista OS's, this function does not exist,
                // and is not necessary, so failure is acceptable.
            }
            
            try
            {
                // Let's see if an instance is already running.
                const bool acquireAtCreation = true; // Do acquire the mutex.
                bool mutexCreated;
                using (Mutex mutex = new Mutex(acquireAtCreation, "UIAccessibilityCheckerRunning", out mutexCreated))
                {
                    // if we did not create the mutex then someone else is already running in this session.
                    if (!mutexCreated)
                    {
                        return;
                    }
                    
                    MainForm mainForm = new MainForm();

                    if (args.Length > 0)
                    {
                        mainForm.AddVerificationDlls(args);
                    }
                    
                    Application.Run(mainForm);
                }
            }
            catch (UnauthorizedAccessException)
            { 
                // In case of access denied on the mutex just return
                return;
            }
        }
    }
}
