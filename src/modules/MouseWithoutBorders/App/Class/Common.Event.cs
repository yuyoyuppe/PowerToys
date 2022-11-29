// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

// <summary>
//     Keyboard/Mouse hook callback implementation.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using MouseWithoutBorders.Class;
using MouseWithoutBorders.Form;

namespace MouseWithoutBorders
{
    internal partial class Common
    {
        internal static TelemetryClient TelemetryForEvent { get; private set; }

        internal static TelemetryClient TelemetryForLog => Setting.Values.SendErrorLogV2 ? TelemetryForEvent : null;

        private static readonly DATA KeybdDackage = new();
        private static readonly DATA MouseDackage = new();
        private static ulong inputEventCount;
        private static ulong invalidPackageCount;
        internal static int MOVE_MOUSE_RELATIVE = 100000;
        internal static int XY_BY_PIXEL = 300000;

        static Common()
        {
            TelemetryForEvent = CreateTelemetryClient();
        }

        private static TelemetryClient CreateTelemetryClient()
        {
            /* During GDPR review (2018/05/14) we already deleted our Application Insights account thus the client would not send any telemetry data to Application Insights.
             * In the future if we decide to enable telemetry back, a full review must be performed.
             * */

            return null;
        }

        internal static ulong InvalidPackageCount
        {
            get => Common.invalidPackageCount;
            set => Common.invalidPackageCount = value;
        }

        internal static ulong InputEventCount
        {
            get => Common.inputEventCount;
            set => Common.inputEventCount = value;
        }

        internal static ulong RealInputEventCount
        {
            get;
            set;
        }

        private static Point actualLastPos;
        private static int myLastX;
        private static int myLastY;

        [SuppressMessage("Microsoft.Maintainability", "CA1502:AvoidExcessiveComplexity", Justification = "Dotnet port with style preservation")]
        internal static void MouseEvent(MOUSEDATA e, int dx, int dy)
        {
            try
            {
                PaintCount = 0;
                bool switchByMouseEnabled = IsSwitchingByMouseEnabled();

                if (switchByMouseEnabled && Sk != null && (DesMachineID == MachineID || !Setting.Values.MoveMouseRelatively) && e.dwFlags == WM_MOUSEMOVE)
                {
                    Point p = MoveToMyNeighbourIfNeeded(e.X, e.Y, desMachineID);

                    if (!p.IsEmpty)
                    {
                        HasSwitchedMachineSinceLastCopy = true;

                        Common.Log(string.Format(
                            CultureInfo.CurrentCulture,
                            "***** Host Machine: newDesMachineIDEx set = [{0}]. Mouse is now at ({1},{2})",
                            newDesMachineIDEx,
                            e.X,
                            e.Y));

                        myLastX = e.X;
                        myLastY = e.Y;

                        PrepareToSwitchToMachine(newDesMachineIDEx, p);
                    }
                }

                if (desMachineID != MachineID && SwitchLocation.Count <= 0)
                {
                    MouseDackage.Des = desMachineID;
                    MouseDackage.Type = PackageType.Mouse;
                    MouseDackage.Md.dwFlags = e.dwFlags;
                    MouseDackage.Md.WDelta = e.WDelta;

                    // Relative move
                    if (Setting.Values.MoveMouseRelatively && Math.Abs(dx) >= MOVE_MOUSE_RELATIVE && Math.Abs(dy) >= MOVE_MOUSE_RELATIVE)
                    {
                        MouseDackage.Md.X = dx;
                        MouseDackage.Md.Y = dy;
                    }
                    else
                    {
                        MouseDackage.Md.X = (e.X - primaryScreenBounds.Left) * 65535 / screenWidth;
                        MouseDackage.Md.Y = (e.Y - primaryScreenBounds.Top) * 65535 / screenHeight;
                    }

                    SkSend(MouseDackage, null, false);

                    if (MouseDackage.Md.dwFlags is WM_LBUTTONUP or WM_RBUTTONUP)
                    {
                        Thread.Sleep(10);
                    }

                    NativeMethods.GetCursorPos(ref actualLastPos);

                    if (actualLastPos != Common.LastPos)
                    {
                        Common.Log($"Mouse cursor has moved unexpectedly: Expected: {Common.LastPos}, actual: {actualLastPos}.");
                        Common.LastPos = actualLastPos;
                    }
                }

#if SHOW_ON_WINLOGON_EX
                if (RunOnLogonDesktop && e.dwFlags == WM_RBUTTONUP &&
                    desMachineID == machineID &&
                    e.x > 2 && e.x < 100 && e.y > 2 && e.y < 20)
                {
                    DoSomthingInUIThread(delegate()
                    {
                        MainForm.HideMenuWhenRunOnLogonDesktop();
                        MainForm.MainMenu.Hide();
                        MainForm.MainMenu.Show(e.x - 5, e.y - 3);
                    });
                }
#endif
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }

        private static bool IsSwitchingByMouseEnabled()
        {
            return (EasyMouseOption)Setting.Values.EasyMouse == EasyMouseOption.Enable || InputHook.EasyMouseKeyDown;
        }

        internal static void PrepareToSwitchToMachine(ID newDesMachineID, Point desMachineXY)
        {
            Log($"PrepareToSwitchToMachine: newDesMachineID = {newDesMachineID}, desMachineXY = {desMachineXY}");

            if (((GetTick() - lastJump < 100) && (GetTick() - lastJump > 0)) || desMachineID == ID.ALL)
            {
                Log("PrepareToSwitchToMachine: lastJump");
                return;
            }

            lastJump = GetTick();

            string newDesMachineName = NamefromID(newDesMachineID);

            if (!IsConnectedTo(newDesMachineID))
            {// Connection lost, cancel switching
                Log("No active connection found for " + newDesMachineName);

                // ShowToolTip("No active connection found for [" + newDesMachineName + "]!", 500);
            }
            else
            {
                Common.newDesMachineID = newDesMachineID;
                SwitchLocation.X = desMachineXY.X;
                SwitchLocation.Y = desMachineXY.Y;
                SwitchLocation.ResetCount();
                _ = EvSwitch.Set();

                // PostMessage(mainForm.Handle, WM_SWITCH, IntPtr.Zero, IntPtr.Zero);
                if (newDesMachineID != DragMachine)
                {
                    if (!IsDragging && !IsDroping)
                    {
                        if (MouseDown && !RunOnLogonDesktop && !RunOnScrSaverDesktop)
                        {
                            DragDropStep02();
                        }
                    }
                    else if (DragMachine != (ID)1)
                    {
                        ChangeDropMachine();
                    }
                }
                else
                {
                    DragDropStep11();
                }

                // Change des machine
                if (desMachineID != newDesMachineID)
                {
                    Log("MouseEvent: Switching to new machine:" + newDesMachineName);

                    // Ask current machine to hide the Mouse cursor
                    if (newDesMachineID != ID.ALL && desMachineID != MachineID)
                    {
                        SendPackage(desMachineID, PackageType.Hidemouse);
                    }

                    DesMachineID = newDesMachineID;

                    if (desMachineID == MachineID)
                    {
                        if (GetTick() - clipboardCopiedTime < BIG_CLIPBOARD_DATA_TIMEOUT)
                        {
                            clipboardCopiedTime = 0;
                            Common.GetRemoteClipboard("PrepareToSwitchToMachine");
                        }
                    }
                    else
                    {
                        // Ask the new active machine to get clipboard data (if the data is too big)
                        SendPackage(desMachineID, PackageType.Machineswitched);
                    }

                    _ = Interlocked.Increment(ref switchCount);
                }
            }
        }

        internal static void SaveSwitchCount()
        {
            if (SwitchCount > 0)
            {
                _ = Task.Run(() =>
                {
                    Common.TelemetryForEvent?.TrackMetric(new MetricTelemetry(nameof(SwitchCount), SwitchCount));
                    Common.TelemetryForEvent?.Flush();
                    Setting.Values.SwitchCount += SwitchCount;
                    _ = Interlocked.Exchange(ref switchCount, 0);
                });
            }
        }

        internal static void KeybdEvent(KEYBDDATA e)
        {
            try
            {
                PaintCount = 0;
                if (desMachineID != newDesMachineID)
                {
                    Log("KeybdEvent: Switching to new machine...");
                    DesMachineID = newDesMachineID;
                }

                if (desMachineID != MachineID)
                {
                    KeybdDackage.Des = desMachineID;
                    KeybdDackage.Type = PackageType.Keybd;
                    KeybdDackage.Kd = e;
                    KeybdDackage.DateTime = GetTick();
                    SkSend(KeybdDackage, null, false);
                    if (KeybdDackage.Kd.dwFlags is WM_KEYUP or WM_SYSKEYUP)
                    {
                        Thread.Sleep(10);
                    }
                }
            }
            catch (Exception ex)
            {
                Log(ex);
            }
        }
    }
}
