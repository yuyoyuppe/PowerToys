// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;

namespace MouseWithoutBorders
{
    internal static class Program
    {
        internal static FormHelper FormHelper;

        private static FormDot dotForm;

        internal static FormDot DotForm
        {
            get
            {
                return dotForm != null && !dotForm.IsDisposed ? dotForm : (dotForm = new FormDot());
            }
        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        private static void Main()
        {
            string[] args = Environment.GetCommandLineArgs();

            if (args.Length > 1 && !string.IsNullOrEmpty(args[1]))
            {
                string command = args[1];
                string arg = args.Length > 2 && !string.IsNullOrEmpty(args[2]) ? args[2] : string.Empty;

                if (command.Equals("SvcExec", StringComparison.OrdinalIgnoreCase))
                {
                    Process.Start(Path.GetDirectoryName(Application.ExecutablePath) + "\\MouseWithoutBorders.exe", "\"" + arg + "\"");
                }
                else if (command.Equals("install", StringComparison.OrdinalIgnoreCase))
                {
                    Process.Start(Path.GetDirectoryName(Application.ExecutablePath) + "\\MouseWithoutBorders.exe");
                }
                else if (command.Equals("help-in", StringComparison.OrdinalIgnoreCase))
                {
                    Process.Start(
@"mailto:truongdo?subject=Mouse Without Borders&body=%0D%0ACHECKLIST:
%0D%0A(1) Run the same, latest version in all machines.
%0D%0A(2) Review all options on the Settings form.
%0D%0A(3) Check firewall/network connection.
%0D%0A(4) Ensure all machines/devices can see each other (pinging by name resolves to right IP Addresses).
%0D%0A
%0D%0AHelp page: http://www.aka.ms/mm, SAW: http://www.aka.ms/mmsaw
");
                }
                else if (command.Equals("help-ex", StringComparison.OrdinalIgnoreCase))
                {
                    Process.Start(@"http://www.aka.ms/mm");
                }
                else if (command.Equals("InternalError", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show(arg, Application.ProductName);
                }
                else if (command.Equals("Terminate", StringComparison.OrdinalIgnoreCase))
                {
                    CleanupProcesses();
                }

                return;
            }

            Application.EnableVisualStyles();
            Application.SetHighDpiMode(HighDpiMode.SystemAware);
            Application.SetCompatibleTextRenderingDefault(false);

            dotForm = new FormDot();
            Application.Run(FormHelper = new FormHelper());
        }

        private static void CleanupProcesses()
        {
            Process[] ps = Process.GetProcessesByName("MouseWithoutBorders");
            foreach (Process p in ps)
            {
                p.Kill();
            }

            ps = Process.GetProcessesByName("MouseWithoutBordersHelper");
            foreach (Process p in ps)
            {
                if (p.Id != Environment.ProcessId)
                {
                    p.Kill();
                }
            }
        }
    }
}
