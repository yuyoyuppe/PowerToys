// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Windows.Forms;

// <summary>
//     Startup/main form + helper methods.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using MouseWithoutBorders.Class;
using MouseWithoutBorders.Form;
using Timer = System.Windows.Forms.Timer;

[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#ShowMouseWithoutBordersUIOnWinLogonDesktop(System.Boolean)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#ChangeIcon(System.Int32)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#UpdateMenu(MouseWithoutBorders.MachineInf[])", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#frmScreen_Load(System.Object,System.EventArgs)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#MenuNewVersion(System.Object,System.EventArgs)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#.ctor()", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#helperTimer_Tick(System.Object,System.EventArgs)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#LoadPlugins()", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Design", "CA1031:DoNotCatchGeneralExceptionTypes", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#mnuPersonalizeLogonScrPluginClick(System.Object,System.EventArgs)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Globalization", "CA1304:SpecifyCultureInfo", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#helperTimer_Tick(System.Object,System.EventArgs)", MessageId = "System.String.ToLower", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#helperTimer_Tick(System.Object,System.EventArgs)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#LogonLogo", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Performance", "CA1814:PreferJaggedArraysOverMultidimensional", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#frmScreen_Load(System.Object,System.EventArgs)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#LoadNewLogonBackground()", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Performance", "CA1800:DoNotCastUnnecessarily", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#UncheckAllMenus()", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#WndProc(System.Windows.Forms.Message&)", MessageId = "MouseWithoutBorders.NativeMethods.SendMessage(System.IntPtr,System.Int32,System.IntPtr,System.IntPtr)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#Destroy()", MessageId = "MouseWithoutBorders.NativeMethods.SendMessage(System.IntPtr,System.Int32,System.IntPtr,System.IntPtr)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Usage", "CA1806:DoNotIgnoreMethodResults", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#helperTimer_Tick(System.Object,System.EventArgs)", MessageId = "MouseWithoutBorders.NativeMethods.SendMessage(System.IntPtr,System.Int32,System.IntPtr,System.IntPtr)", Justification = "Dotnet port with style preservation")]
[module: SuppressMessage("Microsoft.Globalization", "CA1300:SpecifyMessageBoxOptions", Scope = "member", Target = "MouseWithoutBorders.frmScreen.#mnuPersonalizeLogonScrPluginClick(System.Object,System.EventArgs)", Justification = "Dotnet port with style preservation")]

namespace MouseWithoutBorders
{
    internal partial class FrmScreen : System.Windows.Forms.Form
    {
#pragma warning disable CA2213 // Disposing is done by ComponentResourceManager
        private Cursor dotCur;
        private Cursor dropCur;
        private int[,] logonLogo;
        private Timer helperTimer;
#pragma warning restore CA2213

        internal int CurIcon { get; set; }

        internal NotifyIcon NotifyIcon { get; set; }

        internal System.Windows.Forms.ToolStripMenuItem MnuALLPC { get; set; }

        internal System.Windows.Forms.ContextMenuStrip MainMenu { get; set; }

        internal FrmScreen()
        {
            InitializeComponent();
            Text = Setting.Values.MyID;
            NotifyIcon.BalloonTipText = Application.ProductName;
            NotifyIcon.BalloonTipTitle = Application.ProductName;
            mnuGenDumpFile.Visible = true;

            Common.WndProcCounter++;

            try
            {
                mnuWindowsPhone.Visible = false; // No longer supported.
            }
            catch (Exception e)
            {
                Common.Log(e);
            }
        }

        private void FrmScreen_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (e.CloseReason != CloseReason.WindowsShutDown)
            {
                if (Tag == null)
                {
                    e.Cancel = true;
                }
            }
            else
            {
                Common.StartServiceAndSendLogoffSignal();
                Quit(true, true);
            }
        }

        internal void Destroy()
        {
            if (helperTimer != null)
            {
                helperTimer.Stop();
                helperTimer.Dispose();
                helperTimer = null;
            }

            NotifyIcon.Visible = false;
            NotifyIcon.Dispose();
            if (!Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop)
            {
                Common.RunDDHelper(true);
            }

            // Common.UnhookClipboard();
            Tag = "myself";
            Close();
        }

        internal void Quit(bool cleanup, bool isFormClosing)
        {
            Tag = "Quitting...";

            Common.TelemetryForEvent?.TrackMetric(new MetricTelemetry(nameof(Common.SwitchCount), Common.SwitchCount));
            Setting.Values.SwitchCount += Common.SwitchCount;
            Common.TelemetryForEvent?.TrackEvent("Stop.");
            Process me = Process.GetCurrentProcess();
            Common.WndProcCounter++;

            try
            {
                if (cleanup)
                {
                    Common.Cleanup();
                }

                Common.WndProcCounter++;
                Common.TelemetryForEvent?.Flush();
                if (!Common.RunOnScrSaverDesktop)
                {
                    Common.ReleaseAllKeys();
                }

                Process[] ps = Process.GetProcessesByName(Common.BinaryName);
                foreach (Process p in ps)
                {
                    if (p.Id != me.Id && p.SessionId == me.SessionId)
                    {
                        p.KillProcess();
                    }
                }

                if (!isFormClosing)
                {
                    ps = Process.GetProcessesByName("MouseWithoutBordersSvc");
                    foreach (Process p in ps)
                    {
                        p.KillProcess();
                    }
                }

                Common.RunDDHelper(true);
            }
            catch (Exception e)
            {
                _ = MessageBox.Show(e.ToString());
            }

            Common.MMSleep(1);

            me.KillProcess();
        }

        private void MnuExit_Click(object sender, EventArgs e)
        {
            Quit(true, false);
        }

        internal void MenuOnClick(object sender, EventArgs e)
        {
            string name = (sender as ToolStripMenuItem).Text;
            Common.SwitchToMachine(name);
        }

        internal void UpdateMenu()
        {
            try
            {
                ChangeIcon(-1);
                while (MainMenu.Items[MainMenu.Items.Count - 1].Tag != null &&
                    ((string)MainMenu.Items[MainMenu.Items.Count - 1].Tag).StartsWith("MACHINE: ", StringComparison.CurrentCultureIgnoreCase))
                {
                    MainMenu.Items.Remove(MainMenu.Items[MainMenu.Items.Count - 1]);
                }

                while (!mnuSendScreenCapture.DropDown.Items[
                    mnuSendScreenCapture.DropDown.Items.Count - 1].Text.Equals("Myself", StringComparison.OrdinalIgnoreCase))
                {
                    mnuSendScreenCapture.DropDown.Items.Remove(mnuSendScreenCapture.DropDown.Items[
                        mnuSendScreenCapture.DropDown.Items.Count - 1]);
                }

                while (!mnuGetScreenCapture.DropDown.Items[
                    mnuGetScreenCapture.DropDown.Items.Count - 1].Text.Equals("All", StringComparison.OrdinalIgnoreCase))
                {
                    mnuGetScreenCapture.DropDown.Items.Remove(mnuGetScreenCapture.DropDown.Items[
                        mnuGetScreenCapture.DropDown.Items.Count - 1]);
                }

                for (int i = 0; i < Common.MAX_MACHINE; i++)
                {
                    string newMachine = Common.MachineMatrix[i].Trim();

                    if (Common.MachinePool.TryFindMachineByName(newMachine, out MachineInf inf) && MachinePool.IsAlive(inf))
                    {
                        ToolStripMenuItem newItem = new(
                            newMachine,
                            null,
                            new EventHandler(MenuOnClick));
                        newItem.Tag = "MACHINE: " + inf.Name;
                        newItem.ToolTipText = "Switch Mouse/keyboard to " + newMachine;
                        newItem.Visible = true;
                        _ = MainMenu.Items.Add(newItem);

                        if (!newMachine.Equals(Common.MachineName.Trim(), StringComparison.OrdinalIgnoreCase))
                        {
                            ToolStripMenuItem newItem2 = new(
                                inf.Name.Trim(),
                                null,
                                new EventHandler(MnuSendScreenCaptureClick));
                            newItem2.Visible = true;
                            _ = mnuSendScreenCapture.DropDown.Items.Add(newItem2);

                            ToolStripMenuItem newItem3 = new(
                                inf.Name.Trim(),
                                null,
                                new EventHandler(MnuGetScreenCaptureClick));
                            newItem3.Visible = true;
                            _ = mnuGetScreenCapture.DropDown.Items.Add(newItem3);
                        }
                    }
                }

                mnuGetScreenCapture.DropDown.Items[0].Enabled = mnuGetScreenCapture.DropDown.Items.Count > 1;
            }
            catch (Exception e)
            {
                Common.Log(e);
            }
        }

        // We dont do something heavy in our timer
        private void FrmScreen_Load(object sender, EventArgs e)
        {
            try
            {
                // Common.UpdateMenu();
                // Common.HookClipboard();
                CurIcon = Common.ICON_ONE;
                ChangeIcon(-1);
                dotCur = CustomCursor.CreateDotCursor();
                Cursor = dotCur;
                dropCur = CustomCursor.CreateCursor(Icon.ToBitmap(), 0, 0);
                BackColor = Color.White;
                Opacity = 0.15;
                UpdateNotifyIcon();

                if (picLogonLogo.Image != null)
                {
                    Bitmap b = new(picLogonLogo.Image);
                    logonLogo = new int[b.Width, b.Height];

                    for (int i = 0; i < b.Width; i++)
                    {
                        for (int j = 0; j < b.Height; j++)
                        {
                            Color c = b.GetPixel(i, j);
                            logonLogo[i, j] = !(c.G == 255 && c.R == 0 && c.B == 0) ? (c.B << 16) | (c.G << 8) | c.R : -1;
                        }
                    }

                    b.Dispose();
                }
            }
            catch (Exception ex)
            {
                BackColor = Color.White;
                Opacity = 0.15;
                Common.Log(ex);
            }

            helperTimer = new System.Windows.Forms.Timer();
            helperTimer.Interval = 100;
            helperTimer.Tick += new EventHandler(HelperTimer_Tick);
            helperTimer.Start();
            Common.WndProcCounter++;

            if (Environment.OSVersion.Version.Major > 6 ||
                   (Environment.OSVersion.Version.Major == 6 && Environment.OSVersion.Version.Minor >= 1))
            {
                // Win7 up
                _ = NativeMethods.ChangeWindowMessageFilterEx(Handle, NativeMethods.WM_SHOW_SETTINGS_FORM, 1, IntPtr.Zero);
            }
        }

        private long count = 21;

        // private bool checkForNewVersion;
#if USING_FORM
        private static frmLogon fLogon = null;
#endif
        private bool busy;
        private bool shownSetupFormOneTime;
        private bool myDesktopNotActive;

        private void HelperTimer_Tick(object sender, EventArgs e)
        {
            Common.WndProcCounter++;

            if (busy)
            {
                return;
            }

            busy = true;

            try
            {
                if (!Common.IsMyDesktopActive() || Common.CurrentProcess.SessionId != NativeMethods.WTSGetActiveConsoleSessionId())
                {
                    myDesktopNotActive = true;

                    SocketStuff sk = Common.Sk;

                    // We are not on application desktop
                    if (sk != null)
                    {
                        turnedOff = true;
                        Common.SecondOpenSocketTry = false;
                        sk.Close(false);
                        Common.Sk = null;
                        count = 21;
#if USING_FORM
                        if (fLogon != null) fLogon.Hide();
#endif

                        if (Common.MainFormVisible)
                        {
                            Common.MainFormDot();
                        }

                        InputSimu.ResetSystemKeyFlags();

                        Common.DoSomethingInTheInputCallbackThread(() =>
                        {
                            Common.Hook?.ResetLastSwitchKeys();
                        });

                        Common.CheckForDesktopSwitchEvent(true);
                    }
                }
                else
                {
                    if (Common.Sk == null)
                    {
                        if (!Common.SecondOpenSocketTry)
                        {
                            if (!Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop && !Common.GetUserName())
                            {
                                // While Windows 8 is hybrid-shutting down, user name would be empty (as returned from the .Net API), we should not do anything in this case.
                                Common.Log("No active user.");
                                Thread.Sleep(1000);
                                busy = false;
                                return;
                            }

                            if (myDesktopNotActive)
                            {
                                myDesktopNotActive = false;
                                Common.MyKey = Setting.Values.MyKey;
                            }

                            Common.UpdateMachinePoolStringSetting();

                            if (!Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop && (Setting.Values.FirstRun || Common.KeyCorrupted))
                            {
                                if (!shownSetupFormOneTime)
                                {
                                    shownSetupFormOneTime = true;
                                    Common.ShowMachineMatrix();

                                    if (Common.KeyCorrupted && !Setting.Values.FirstRun)
                                    {
                                        Common.KeyCorrupted = false;
                                        string msg = "The security key is corrupted for some reason, please resetup.";
                                        MessageBox.Show(msg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                    }
                                }
                            }
                        }

                        Common.ReopenSockets(false);
                        Common.SecondOpenSocketTry = true;

                        if (Common.Sk != null)
                        {
                            Common.SendHello();
                            Common.SendHeartBeat();

                            if (!Common.RunOnScrSaverDesktop && !Common.RunOnLogonDesktop)
                            {
                                NotifyIcon.Visible = false;
                                NotifyIcon.Visible = true;

                                // Common.ReHookClipboard();
                            }

                            Common.RunDDHelper();
                        }

                        count = 0;

                        Common.InitDone = true;
#if SHOW_ON_WINLOGON
                        if (Common.RunOnLogonDesktop)
                        {
                            ShowMouseWithoutBordersUIOnWinLogonDesktop(true);
                        }
#endif
                    }

                    if ((count % 2) == 0)
                    {
                        if (Common.PleaseReopenSocket == 10 || (Common.PleaseReopenSocket > 0 && count > 0 && count % 300 == 0))
                        {
                            if (!Common.AtLeastOneSocketEstablished() || Common.PleaseReopenSocket == 10)
                            {
                                Thread.Sleep(1000);
                                if (Common.PleaseReopenSocket > 0)
                                {
                                    Common.PleaseReopenSocket--;
                                }

                                // Double check.
                                if (!Common.AtLeastOneSocketEstablished())
                                {
                                    Common.GetMachineName();
                                    Common.Log("Common.pleaseReopenSocket: " + Common.PleaseReopenSocket.ToString(CultureInfo.InvariantCulture));
                                    Common.ReopenSockets(false);
                                    Common.NewDesMachineID = Common.DesMachineID = Common.MachineID;
                                }
                            }
                            else
                            {
                                Common.PleaseReopenSocket = 0;
                            }
                        }

                        if (Common.PleaseReopenSocket == Common.REOPEN_WHEN_HOTKEY)
                        {
                            Common.PleaseReopenSocket = 0;
                            Common.ReopenSockets(true);
                        }
                        else if (Common.PleaseReopenSocket == Common.REOPEN_WHEN_WSAECONNRESET)
                        {
                            Common.PleaseReopenSocket = 0;
                            Thread.Sleep(1000);
                            Common.UpdateClientSockets("REOPEN_WHEN_WSAECONNRESET");
                        }

                        if (Common.RunOnLogonDesktop)
                        {
                            PaintMynameOnDesktop();
                        }

                        /*
                        if (checkClipboard)
                        {
                            Common.CtrlC = -1;
                            checkClipboard = false;
                            Common.CheckClipboard();
                        }
                        else if (--Common.CtrlC == 0)//Giving some timeout from the time Ctrl+C
                        {
                            if ((Common.GetTick() - Common.LastClipboardEventTime > 5000) && Common.CheckClipboard())
                            {
                                Common.ReHookClipboard();
                            }
                        }
                         * */
                    }

                    // One more time after 1/3 minutes (Sometimes XP has explorer started late)
                    if (count == 600 || count == 1800)
                    {
                        Common.RunDDHelper();
                    }

                    if (count == 600)
                    {
                        if (!Common.GeneratedKey)
                        {
                            Common.MyKey = Setting.Values.MyKey;

                            if (!Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop)
                            {
                                Common.ShowMachineMatrix();

                                Common.MatrixForm?.UpdateKeyTextBox();

                                Common.Sk?.Close(false);

                                MessageBox.Show("The security key must be auto generated in one of the machines.", Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else if (!Common.KeyCorrupted && !Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop && !Setting.Values.FirstRun && Common.AtLeastOneSocketConnected())
                        {
                            int myKeyDaysToExpire = Setting.Values.MyKeyDaysToExpire;

                            if (myKeyDaysToExpire <= 0)
                            {
                                Common.ShowMachineMatrix();

                                Common.Sk?.Close(false);

                                string msg = "The security key has expired, please generate a new key.";
                                MessageBox.Show(msg, Application.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                            else if (myKeyDaysToExpire <= 15)
                            {
                                Common.ShowToolTip($"The security key will expire in {myKeyDaysToExpire} days! Please regenerate a new key.", 10000);
                            }
                        }
                    }
                }

                if (count == 20)
                {
#if SHOW_ON_WINLOGON
                    // if (Common.RunOnLogonDesktop) ShowMouseWithoutBordersUIOnWinLogonDesktop(false);
#endif
                    Common.CheckForDesktopSwitchEvent(true);
                    Common.UpdateClientSockets("helperTimer_Tick"); // Sockets may be closed by the remote host when both machines switch desktop at the same time.
                }

                count++;
                if (count % 5 == 0)
                {
                    if (Common.RunOnLogonDesktop)
                    {
                        ShowMessageOnLogonDesktop(Common.DesMachineID != ID.ALL);
                    }

                    if (Common.ToggleIcons != null)
                    {
                        Common.ToggleIcon();
                    }

                    if (count % 20 == 0)
                    {
                        Common.LogAll();

                        // Need to review this code on why it is needed (moved from MoveToMyNeighbourIfNeeded(...))
                        for (int i = 0; i < Common.MachineMatrix.Length; i++)
                        {
                            if (string.IsNullOrEmpty(Common.MachineMatrix[i]) && !Common.InMachineMatrix(Common.MachineName))
                            {
                                Common.MachineMatrix[i] = Common.MachineName;
                            }
                        }

                        if (count % 600 == 0 && Common.Sk != null)
                        {
                            // Send out Heartbeat every 1 or 5 mins
                            if (Setting.Values.BlockScreenSaver || count % 3000 == 0)
                            {
                                Common.SendAwakeBeat();
                                Common.RemoveDeadMachines();

                                // GC.Collect();
                                // GC.WaitForPendingFinalizers();
                            }

                            Common.CloseAnUnusedSocket();
                        }
                    }
                    else if ((count % 36005) == 0)
                    {// One hour
                        Common.SaveSwitchCount();

                        int rv = 0;

                        if (!Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop && Common.IsMyDesktopActive() && (rv = Common.SendMessageToHelper(0x400, IntPtr.Zero, IntPtr.Zero)) <= 0)
                        {
                            Common.TelemetryLogTrace($"{Common.HELPER_FORM_TEXT} not found: {rv}", SeverityLevel.Warning);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Common.Log(ex);
            }
            finally
            {
                busy = false;
            }
        }

        private void MnuMachineMatrix_Click(object sender, EventArgs e)
        {
            Common.ShowMachineMatrix();
        }

        private void FrmScreen_MouseMove(object sender, MouseEventArgs e)
        {
            if (!Common.IsDroping)
            {
                if (Cursor != dotCur)
                {
                    Cursor = dotCur;
                    Refresh();

                    // Application.DoEvents();
                }
            }
        }

        internal void ChangeIcon(int iconCode)
        {
            try
            {
                Graphics g;
                Pen p;
                Bitmap bm = new(imgListIcon.Images[0]);

                /*
                if (curIcon == Common.ICON_ONE)
                {
                    string[] x = Common.MachineMatrix;
                    Graphics g = Graphics.FromImage(bm);
                    if (x != null)
                    {
                        for (int i = 0; i < Common.MAX_MACHINE; i++)
                        {
                            if (x[i].Trim().Length > 0) g.FillEllipse(
                                #if DEBUG
                                Brushes.White,
                                #else
                                Brushes.Salmon,
                                #endif
                                3 + (i % 2) * 16, 3 + (i / 2) * 16, 10, 10);
                        }
                    }
                    g.Dispose();
                }
                */

                if (CurIcon != Common.ICON_ONE)
                {
                    p = new Pen(Color.Red, 2);
                    g = Graphics.FromImage(bm);
                    g.DrawRectangle(p, 1, 1, bm.Width - 2, bm.Height - 2);
                    g.Dispose();
                }

                if (iconCode == Common.ICON_SMALL_CLIPBOARD)
                {
                    g = Graphics.FromImage(bm);
                    g.FillEllipse(Brushes.Blue, 8, 8, 16, 16);
                    g.Dispose();
                }
                else if (iconCode == Common.ICON_BIG_CLIPBOARD)
                {
                    g = Graphics.FromImage(bm);
                    g.FillEllipse(Brushes.Yellow, 8, 8, 16, 16);
                    g.Dispose();
                }
                else if (iconCode == Common.ICON_ERROR)
                {
                    g = Graphics.FromImage(bm);
                    g.FillEllipse(Brushes.Red, 8, 8, 16, 16);
                    g.Dispose();
                }

#if DEBUG
                p = new Pen(Color.White, 4);
                g = Graphics.FromImage(bm);
                g.DrawRectangle(p, 6, 6, bm.Width - 12, bm.Height - 12);
                g.Dispose();
                p.Dispose();
#endif

                Common.Log($"Changing icon to {iconCode}.");

                if (NotifyIcon.Icon != null)
                {
                    NativeMethods.DestroyIcon(NotifyIcon.Icon.Handle);
                }

                NotifyIcon.Icon = Icon.FromHandle(bm.GetHicon());
                bm.Dispose();
            }
            catch (Exception e)
            {
                Common.Log(e);
            }
        }

        internal void MnuAllPC_Click(object sender, EventArgs e)
        {
            Common.Log("mnuAllPC_Click");
            Common.SwitchToMultipleMode(MnuALLPC.Checked, true);
            CurIcon = MnuALLPC.Checked ? Common.ICON_ALL : Common.ICON_ONE;
            ChangeIcon(CurIcon);
        }

        internal void UpdateMultipleModeIconAndMenu()
        {
            Common.DoSomethingInUIThread(() =>
            {
                MnuALLPC.Checked = Common.NewDesMachineID == ID.ALL;
                CurIcon = MnuALLPC.Checked ? Common.ICON_ALL : Common.ICON_ONE;
                ChangeIcon(CurIcon);
            });
        }

        // private bool checkClipboard = false;
        protected override void WndProc(ref Message m)
        {
            const int WM_ENDSESSION = 0x0016;
            const int WM_QUERYENDSESSION = 0x0011;

            // const int WM_DRAWCLIPBOARD = 0x0308;
            // const int WM_CHANGECBCHAIN = 0x030D;
            switch (m.Msg)
            {
                /*
                case WM_DRAWCLIPBOARD:
                    if (Common.NextClipboardViewer != IntPtr.Zero && Common.NextClipboardViewer != this.Handle)
                    {
                        int rv = NativeMethods.SendMessage(Common.NextClipboardViewer, m.Msg, m.WParam, m.LParam);
                        Common.Log("SendMessage returned " + rv.ToString(CultureInfo.CurrentCulture));
                    }

                    if (Common.GetTick() - Common.LastClipboardEventTime < 1000)
                    {
                        Common.Log("GetTick() - lastClipboardEventTime < 1000");
                        Common.LastClipboardEventTime = Common.GetTick();
                        return;
                    }
                    Common.LastClipboardEventTime = Common.GetTick();

                    checkClipboard = true;//Do the task in a next message loop.
                    break;

                case WM_CHANGECBCHAIN:
                    if (m.WParam == Common.NextClipboardViewer)
                        Common.NextClipboardViewer = m.LParam;
                    else
                        if (Common.NextClipboardViewer != IntPtr.Zero && Common.NextClipboardViewer != this.Handle)
                        {
                            int rv = NativeMethods.SendMessage(Common.NextClipboardViewer, m.Msg, m.WParam, m.LParam);
                            Common.Log("SendMessage returned " + rv.ToString(CultureInfo.CurrentCulture));
                        }
                    break;
                */

                case NativeMethods.WM_SHOW_DRAG_DROP:
                    Point p = default;
                    _ = NativeMethods.GetCursorPos(ref p);
                    Width = 70;
                    Height = 70;
                    Left = p.X - (Width / 3);
                    Top = p.Y - (Height / 3);
                    BackColor = Color.White;
                    Opacity = 0.15;
                    if (Cursor != dropCur)
                    {
                        Cursor = dropCur;
                    }

                    Show();
                    break;

                case NativeMethods.WM_HIDE_DRAG_DROP:
                    Common.MainFormDot();

                    /*
                    this.Width = 1;
                    this.Height = 1;
                    this.Left = 0;
                    this.Top = 0;
                    //this.Hide();
                    //Common.MainFormVisible = false;
                     * */

                    if (Cursor != dotCur)
                    {
                        Cursor = dotCur;
                    }

                    break;

                case NativeMethods.WM_HIDE_DD_HELPER:
                    Common.MainForm3Pixels();
                    Common.MMSleep(0.2);
                    if (m.WParam.ToInt32() == 1)
                    {
                        InputSimu.MouseUp(); // A file is being draged
                    }

                    IntPtr h = (IntPtr)NativeMethods.FindWindow(null, Common.HELPER_FORM_TEXT);

                    if (h.ToInt32() > 0)
                    {
                        Common.Log("Hide Mouse Without Borders Helper.");

                        // Common.ShowWindow(h, 1);
                        _ = NativeMethods.ShowWindow(h, 0);

                        // Common.SetWindowPos(h, Common.HWND_NOTOPMOST, 10, 10, 50, 50, 0);
                        // Common.SetWindowPos(h, Common.HWND_TOPMOST, 0, 0, 800, 60, 0);// For debug
                    }

                    break;

                case NativeMethods.WM_CHECK_EXPLORER_DRAG_DROP:
                    Common.Log("Got WM_CHECK_EXPLORER_DRAG_DROP!");
                    Common.DragDropStep04();
                    break;

                case NativeMethods.WM_QUIT:
                    // Quit(true);
                    break;

                case NativeMethods.WM_SWITCH:
                    break;

                case WM_QUERYENDSESSION:
                    Common.Log("WM_QUERYENDSESSION...");
                    Common.StartServiceAndSendLogoffSignal();
                    break;

                case WM_ENDSESSION:
                    Quit(true, true);
                    break;

                case NativeMethods.WM_SHOW_SETTINGS_FORM:
                    if (!Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop)
                    {
                        Common.ShowMachineMatrix();
                    }

                    break;

                default:
                    base.WndProc(ref m);
                    break;
            }
        }

        private void NotifyIcon_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                Common.ShowMachineMatrix();
            }
        }

        internal void SetTrayIconText(string value)
        {
            if (value.Length > 63)
            {
                value = value[..63];
            }

            NotifyIcon.Text = value;
        }

        internal void UpdateNotifyIcon()
        {
            string[] x = Common.MachineMatrix;
            string iconText;
            if (x != null && (x[0].Length > 0 || x[1].Length > 0 || x[2].Length > 0 || x[3].Length > 0))
            {
                iconText = Setting.Values.MatrixOneRow
                    ? string.Format(
                        CultureInfo.CurrentCulture,
                        "[{0}][{1}][{2}][{3}]",
                        x[0].Trim(),
                        x[1].Trim(),
                        x[2].Trim(),
                        x[3].Trim())
                    : string.Format(
                        CultureInfo.CurrentCulture,
                        "[{0}][{1}]\r\n[{2}][{3}]",
                        x[0].Trim(),
                        x[1].Trim(),
                        x[2].Trim(),
                        x[3].Trim());

                SetTrayIconText(iconText);
            }
            else
            {
                SetTrayIconText(Application.ProductName);
            }
        }

        internal void ShowToolTip(string txt, int timeOutInMilliseconds, ToolTipIcon icon = ToolTipIcon.Info)
        {
            if (!Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop)
            {
                NotifyIcon.ShowBalloonTip(timeOutInMilliseconds, Application.ProductName, txt, icon);
            }
        }

        private void MnuSendScreenCaptureClick(object sender, EventArgs e)
        {
            string mnuCaption = (sender as ToolStripMenuItem).Text;
            string captureFile = Common.CaptureScreen();

            if (captureFile != null && File.Exists(captureFile))
            {
                if (mnuCaption.Equals("Myself", StringComparison.OrdinalIgnoreCase))
                {
                    Common.OpenImage(captureFile);
                }
                else
                {
                    Common.SendImage(mnuCaption, captureFile);
                }
            }
        }

        private void MnuGetScreenCaptureClick(object sender, EventArgs e)
        {
            string mnuCaption = (sender as ToolStripMenuItem).Text;

            // Send Capturescreencmd
            ID des = mnuCaption.Equals("All", StringComparison.OrdinalIgnoreCase) ? ID.ALL : Common.IDfromName(mnuCaption);
            Common.SendPackage(des, PackageType.Capturescreencmd);
        }

        private void FrmScreen_Shown(object sender, EventArgs e)
        {
            Common.AssertOneInstancePerDesktopSession();

            Common.MainForm = this;
            Hide();
            if (!Common.RunOnLogonDesktop && !Common.RunOnScrSaverDesktop)
            {
                NotifyIcon.Visible = false;
                NotifyIcon.Visible = true;
            }
        }

        private void MnuAbout_Click(object sender, EventArgs e)
        {
            if (Common.AboutForm == null)
            {
                _ = (Common.AboutForm = new FrmAbout()).ShowDialog();
            }
            else
            {
                Common.AboutForm.Activate();
            }
        }

#if SHOW_ON_WINLOGON
        private void ShowMouseWithoutBordersUIOnWinLogonDesktop(bool initMenuState)
        {
            Common.PaintCount = 0;
            try
            {
                if (initMenuState)
                {
                    HideMenuWhenRunOnLogonDesktop();
                }

#if !USING_FORM
                PaintMynameOnDesktop();
#else
                if (fLogon == null) fLogon = new frmLogon();
                fLogon.Show();
                fLogon.LabelDesktop.ContextMenuStrip = fLogon.ContextMenuStrip = MainMenu;
#endif
            }
            catch (Exception e)
            {
                Common.Log(e);
            }
        }

        internal void HideMenuWhenRunOnLogonDesktop()
        {
            mnuHelp.Visible = false;
            mnuGetScreenCapture.Visible = false;
            mnuSendScreenCapture.Visible = false;
            toolStripSeparator1.Visible = false;
            toolStripSeparator2.Visible = false;
            mnuAbout.Visible = false;
            mnuMachineMatrix.Visible = false;
            mnuReinstallKeyboardAndMouseHook.Visible = false;
            toolStripMenuItem1.Visible = false;
            toolStripMenuItem2.Visible = false;
            mnuExit.Visible = false;
        }
#endif

        private void MainMenu_MouseLeave(object sender, EventArgs e)
        {
#if SHOW_ON_WINLOGON
            if (Common.RunOnLogonDesktop)
            {
                MainMenu.Hide();
                Thread.Sleep(50);
                PaintMynameOnDesktop();
            }
#endif
        }

        private int colorB = 240;
        private int colorBDelta = 5;
        private bool turnedOff = true;

        internal void ShowMessageOnLogonDesktop(bool turnOff)
        {
            if (turnOff && turnedOff)
            {
                colorB = 240;
                colorBDelta = 5;
                return;
            }

            turnedOff = false;
            IntPtr hdsktop = NativeMethods.GetDesktopWindow();
            IntPtr hdc = NativeMethods.GetWindowDC(hdsktop);
            int textLengthInPixels;

            if (hdc != IntPtr.Zero)
            {
                NativeMethods.RECT r;
                string machineMatrix = Application.ProductName + " is in multiple mode, keyboard may repeat in all machines: ";
                r.Top = 0;
                r.Left = Common.ScreenWidth / 5;
                r.Right = Common.ScreenWidth - (Common.ScreenWidth / 5);
                r.Bottom = 20;

                for (int i = 0; i < Common.MAX_MACHINE; i++)
                {
                    string newMachine = Common.MachineMatrix[i].Trim();

                    if (Common.MachinePool.TryFindMachineByName(newMachine, out MachineInf inf) && MachinePool.IsAlive(inf))
                    {
                        machineMatrix += "[" + inf.Name.Trim() + "]";
                    }
                }

                if (turnOff)
                {
                    turnedOff = true;
                    colorBDelta = -40;
                    colorB = 0;
                }
                else
                {
                    colorBDelta = colorB == 255 ? -colorBDelta : colorB == 240 ? colorBDelta > 0 ? 5 : -40 : colorB == 0 ? -colorBDelta : colorBDelta;
                    colorB += colorBDelta;
                }

                _ = NativeMethods.SetBkColor(hdc, 0);
                _ = NativeMethods.SetTextColor(hdc, colorB << 8);
                _ = NativeMethods.DrawText(hdc, machineMatrix, machineMatrix.Length, ref r, 0x1 | 0x4 | 0x100 | 0x400);
                textLengthInPixels = r.Right - r.Left;
                r.Left = (Common.ScreenWidth - textLengthInPixels) / 2;
                r.Right = r.Left + textLengthInPixels;
                _ = NativeMethods.DrawText(hdc, machineMatrix, machineMatrix.Length, ref r, 0x1 | 0x4 | 0x100);
                _ = NativeMethods.ReleaseDC(hdsktop, hdc);
            }
        }

#if SHOW_ON_WINLOGON
        private const byte MINC = 185;
        private const byte MAXC = 255;
        private const byte MINCD = 10;
        private byte paintColorR = MINC;
        private byte paintColorG = MINC;
        private byte paintColorB = MINC;
        private bool paintColorDown;

        internal void PaintMynameOnDesktop()
        {
            if (Setting.Values.HideLogonLogo)
            {
                return;
            }

            if (Common.PaintCount > 1500)
            {
                return; // running for 5 mins only
            }

            Common.PaintCount++;
            IntPtr hdsktop = NativeMethods.GetDesktopWindow();
            IntPtr hdc = NativeMethods.GetWindowDC(hdsktop);
            if (hdc != IntPtr.Zero)
            {
                int c;

                // int rv = NativeMethods.DrawText(hdc, Common.BinaryName, 10, ref r, 0);
                for (int i = 0; i < logonLogo.GetUpperBound(0); i++)
                {
                    for (int j = 0; j < logonLogo.GetUpperBound(1); j++)
                    {
                        if (logonLogo[i, j] != -1)
                        {
                            c = logonLogo[i, j];
                            if (c == 0xFFFFFF)
                            {
                                c = (paintColorB << 16) | (paintColorG << 8) | paintColorR;
                            }

                            int rv = (int)NativeMethods.SetPixel(hdc, i, j, (uint)c);
                        }
                    }
                }

                // Common.Log("PaintMynameOnDesktop: last rv = " + rv.ToString(CultureInfo.CurrentCulture));
                _ = NativeMethods.ReleaseDC(hdsktop, hdc);
            }

            if (!paintColorDown)
            {
                if (paintColorR < MAXC)
                {
                    paintColorR += MINCD;
                }
                else
                {
                    if (paintColorG < MAXC)
                    {
                        paintColorG += MINCD;
                    }
                    else
                    {
                        if (paintColorB < MAXC)
                        {
                            paintColorB += MINCD;
                        }
                        else
                        {
                            paintColorDown = true;
                        }
                    }
                }
            }
            else
            {
                if (paintColorB > MINC)
                {
                    paintColorB -= MINCD;
                }
                else
                {
                    if (paintColorR > MINC)
                    {
                        paintColorR -= MINCD;
                    }
                    else
                    {
                        if (paintColorG > MINC)
                        {
                            paintColorG -= MINCD;
                        }
                        else
                        {
                            paintColorDown = false;
                        }
                    }
                }
            }
        }
#endif

        private void MnuHelp_Click(object sender, EventArgs e)
        {
            Common.OpenHelpLink();
        }

        internal void MnuReinstallKeyboardAndMouseHook_Click(object sender, EventArgs e)
        {
            Common.DoSomethingInTheInputCallbackThread(() =>
            {
                if (Common.Hook != null)
                {
                    Common.Hook.Stop();
                    Common.Hook = null;
                }

                Common.InputCallbackForm.Close();
                Common.InputCallbackForm = null;
                Program.StartInputCallbackThread();
            });
        }

        private void MnuGenDumpFile_Click(object sender, EventArgs e)
        {
            Common.GenerateLog();
        }

        private void MainMenu_Opening(object sender, CancelEventArgs e)
        {
            UpdateMenu();
        }
    }
}
