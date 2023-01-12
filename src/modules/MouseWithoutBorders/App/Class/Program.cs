// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

// <summary>
//     Application entry and pre-process/intialization.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using System;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;
using ManagedCommon;

[module: SuppressMessage("Microsoft.MSInternal", "CA904:DeclareTypesInMicrosoftOrSystemNamespace", Scope = "namespace", Target = "MouseWithoutBorders", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1014:MarkAssembliesWithClsCompliant", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Globalization", "CA1304:SpecifyCultureInfo", Scope = "member", Target = "MouseWithoutBorders.Program.#Main()", MessageId = "System.String.ToLower", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "MouseWithoutBorders.Program.#Main()", Justification = "Dotnet port with style preservation")]

namespace MouseWithoutBorders.Class
{
    internal static class Program
    {
        private static FormHelper formHelper;

        internal static FormHelper FormHelper => formHelper != null && !formHelper.IsDisposed ? formHelper : (formHelper = new FormHelper());

        private static FormDot dotForm;

        internal static FormDot DotForm => dotForm != null && !dotForm.IsDisposed ? dotForm : (dotForm = new FormDot());

        [STAThread]
        private static void Main()
        {
            Common.Log(Application.ProductName + " Started!");
            Thread.CurrentThread.Name = Application.ProductName + " main thread";
            Common.BinaryName = Path.GetFileNameWithoutExtension(Application.ExecutablePath);

            try
            {
                string[] args = Environment.GetCommandLineArgs();

                var parentPid = 0;

                User = WindowsIdentity.GetCurrent().Name;
                Common.Log("*** Started as " + User);

                // NOTE(yuyoyuppe): unified logic between elevated/non-elevated scenarios
                Common.RunWithNoAdminRight = true;

                Common.Log(Environment.CommandLine);

                if (args.Length > 1 && args[1] != null)
                {
                    if (Common.CheckSecondInstance(Common.RunWithNoAdminRight))
                    {
                        Common.Log("*** Second instance, exiting...");
                        return;
                    }

                    string myDesktop = Common.GetMyDesktop();
                    string arg = args[1].Trim();
                    if (arg.Equals("winlogon", StringComparison.OrdinalIgnoreCase))
                    {
                        // Executed by service, running on logon desktop
                        Common.Log("*** Running on " + arg + " desktop");
                        Common.RunOnLogonDesktop = true;
                    }
                    else if (args[1].Trim().Equals("default", StringComparison.OrdinalIgnoreCase))
                    {
                        Common.Log("*** Running on " + arg + " desktop");
                    }
                    else if (args[1].Equals(myDesktop, StringComparison.OrdinalIgnoreCase))
                    {
                        Common.Log("*** Running on " + myDesktop);
                        if (myDesktop.Equals("Screen-saver", StringComparison.OrdinalIgnoreCase))
                        {
                            Common.RunOnScrSaverDesktop = true;
                            Setting.Values.LastX = Common.JUST_GOT_BACK_FROM_SCREENSAVER;
                        }
                    }
                    else
                    {
                        int.TryParse(args[1], out parentPid);
                    }
                }
                else
                {
                    if (Common.CheckSecondInstance(true))
                    {
                        Common.Log("*** Second instance, exiting...");
                        return;
                    }

                    if (!Common.RunWithNoAdminRight)
                    {
                        Common.StartMouseWithoutBordersService();
                        return;
                    }
                }

                Common.TelemetryForEvent?.TrackEvent("Start");

                try
                {
                    Common.CurrentProcess = Process.GetCurrentProcess();
                    Common.CurrentProcess.PriorityClass = ProcessPriorityClass.RealTime;
                }
                catch (Exception e)
                {
                    Common.Log(e);
                }

                Common.Log(Environment.OSVersion.ToString());

                // Environment.OSVersion is unreliable from 6.2 and up, so just forcefully call the APIs and log the exception unsupported by Windows:
                int setProcessDpiAwarenessResult = -1;

                try
                {
                    setProcessDpiAwarenessResult = NativeMethods.SetProcessDpiAwareness(2);
                    Common.Log(string.Format(CultureInfo.InvariantCulture, "SetProcessDpiAwareness: {0}.", setProcessDpiAwarenessResult));
                }
                catch (DllNotFoundException)
                {
                    Common.Log("SetProcessDpiAwareness is unsupported in Windows 7 and lower.");
                }
                catch (EntryPointNotFoundException)
                {
                    Common.Log("SetProcessDpiAwareness is unsupported in Windows 7 and lower.");
                }
                catch (Exception e)
                {
                    Common.Log(e);
                }

                try
                {
                    if (setProcessDpiAwarenessResult != 0)
                    {
                        Common.Log(string.Format(CultureInfo.InvariantCulture, "SetProcessDPIAware: {0}.", NativeMethods.SetProcessDPIAware()));
                    }
                }
                catch (Exception e)
                {
                    Common.Log(e);
                }

                System.Threading.Thread mainUIThread = Thread.CurrentThread;
                Common.UIThreadID = mainUIThread.ManagedThreadId;
                Thread.UpdateThreads(mainUIThread);

                StartInputCallbackThread();

                Application.EnableVisualStyles();
                _ = Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
                Application.SetCompatibleTextRenderingDefault(false);

                Common.Init();
                Common.WndProcCounter++;

                var formScreen = new FrmScreen();
                if (parentPid != 0)
                {
                    RunnerHelper.WaitForPowerToysRunner(parentPid, () =>
                    {
                        formScreen.Quit(true, false);
                        Application.Exit();
                    });
                }

                Application.Run(formScreen);
            }
            catch (Exception e)
            {
                Common.Log(e);
            }
        }

        internal static void StartInputCallbackThread()
        {
            System.Collections.Hashtable dummy = Setting.Values.VKMap; // Reading from registry to memory.
            Thread inputCallback = new(new ThreadStart(InputCallbackThread), "InputCallback Thread");
            inputCallback.SetApartmentState(ApartmentState.STA);
            inputCallback.Priority = ThreadPriority.Highest;
            inputCallback.Start();
        }

        private static void InputCallbackThread()
        {
            Common.InputCallbackThreadID = Thread.CurrentThread.ManagedThreadId;
            while (!Common.InitDone)
            {
                Thread.Sleep(100);
            }

            Application.Run(new FrmInputCallback());
        }

        internal static void StartService()
        {
            if (Common.RunWithNoAdminRight)
            {
                return;
            }

            try
            {
                // Kill all but me
                Process me = Process.GetCurrentProcess();
                Process[] ps = Process.GetProcessesByName(Common.BinaryName);
                foreach (Process pp in ps)
                {
                    if (pp.Id != me.Id)
                    {
                        Common.Log(string.Format(CultureInfo.InvariantCulture, "Killing process {0}.", pp.Id));
                        pp.KillProcess();
                    }
                }
            }
            catch (Exception e)
            {
                Common.Log(e);
            }

            Common.StartMouseWithoutBordersService();
        }

        internal static string User { get; set; }
    }
}
