// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading;
using System.Windows.Forms;

// <summary>
//     Drag/Drop implementation.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using MouseWithoutBorders.Class;

namespace MouseWithoutBorders
{
    /* Common.DragDrop.cs
     * Drag&Drop is one complicated implementation of the tool with some tricks.
     *
     * SEQUENCE OF EVENTS:
     * DragDropStep01: MachineX: Remember mouse down state since it could be a start of a draging
     * DragDropStep02: MachineY: Send an message to the MachineX to ask it to check if it is
     *                           doing drag/drop
     * DragDropStep03: MachineX: Got explorerdd, send WM_CHECK_EXPLORER_DRAG_DROP to its mainForm
     * DragDropStep04: MachineX: Show Mouse Without Borders Helper form at mouse cursor to get DragEnter event.
     * DragDropStepXX: MachineX: Mouse Without Borders Helper: Called by DragEnter, check if draging a single file,
     *                           remember the file (set as its window caption)
     * DragDropStep05: MachineX: Get the file name from Mouse Without Borders Helper, hide Mouse Without Borders Helper window
     * DragDropStep06: MachineX: Broadcast a message saying that it has some drag file.
     * DragDropStep08: MachineY: Got clipboarddd, isDropping set, get the MachineX name from the package.
     * DragDropStep09: MachineY: Since isDropping is true, show up the drop form (looks like an icon).
     * DragDropStep10: MachineY: MouseUp, set isDropping to false, hide the drop "icon" and get data.
     * DragDropStep11: MachineX: Mouse move back without drop event, cancelling drag/dop
     *                           SendClipboardBeatDragDropEnd
     * DragDropStep12: MachineY: Hide the drop "icon" when received clipboardddend.
     *
     * FROM VERSION 1.6.3: Drag/Drop is temporary removed, Drop action cannot be done from a lower intergrity app to a higher one.
     * We have to run a helper process...
     * http://forums.microsoft.com/MSDN/ShowPost.aspx?PageIndex=1&SiteID=1&PageID=1&PostID=736086
     *
     * 2008.10.28: Trying to restore the Drag/Drop feature by adding the drag/drop helper process. Coming in version
     * 1.6.5
     * */

    internal partial class Common
    {
        private static bool isDragging;

        internal static bool IsDragging
        {
            get => Common.isDragging;
            set => Common.isDragging = value;
        }

        internal static void DragDropStep01(int wParam)
        {
            if (!Setting.Values.TransferFile)
            {
                return;
            }

            if (wParam == WM_LBUTTONDOWN)
            {
                MouseDown = true;
                DragMachine = desMachineID;
                dropMachineID = ID.NONE;
                Log("DragDropStep01: MouseDown");
            }
            else if (wParam == WM_LBUTTONUP)
            {
                MouseDown = false;
                Log("DragDropStep01: MouseUp");
            }

            if (wParam == WM_RBUTTONUP && IsDroping)
            {
                IsDroping = false;
                LastIDWithClipboardData = ID.NONE;
            }
        }

        internal static void DragDropStep02()
        {
            if (desMachineID == MachineID)
            {
                Log("DragDropStep02: SendCheckExplorerDragDrop sent to myself");
                DoSomethingInUIThread(() =>
                {
                    _ = NativeMethods.PostMessage(MainForm.Handle, NativeMethods.WM_CHECK_EXPLORER_DRAG_DROP, (IntPtr)0, (IntPtr)0);
                });
            }
            else
            {
                SendCheckExplorerDragDrop();
                Log("DragDropStep02: SendCheckExplorerDragDrop sent");
            }
        }

        internal static void DragDropStep03(DATA package)
        {
            if (RunOnLogonDesktop || RunOnScrSaverDesktop)
            {
                return;
            }

            if (package.Des == MachineID || package.Des == ID.ALL)
            {
                Log("DragDropStep03: Explorerdd Received.");
                dropMachineID = package.Src; // Drop machine is the machine that sent Explorerdd
                if (MouseDown || IsDroping)
                {
                    Log("DragDropStep03: Mouse is down, check if draging...sending WM_CHECK_EXPLORER_DRAG_DROP to myself...");
                    DoSomethingInUIThread(() =>
                    {
                        _ = NativeMethods.PostMessage(MainForm.Handle, NativeMethods.WM_CHECK_EXPLORER_DRAG_DROP, (IntPtr)0, (IntPtr)0);
                    });
                }
            }
        }

        private static int dragDropStep05ExCalledByIpc;

        internal static void DragDropStep04()
        {
            if (!IsDroping)
            {
                IntPtr h = (IntPtr)NativeMethods.FindWindow(null, Common.HELPER_FORM_TEXT);
                if (h.ToInt32() > 0)
                {
                    _ = Interlocked.Exchange(ref dragDropStep05ExCalledByIpc, 0);

                    MainForm.Hide();
                    MainFormVisible = false;

                    Point p = default;

                    // NativeMethods.SetWindowText(h, "");
                    _ = NativeMethods.SetWindowPos(h, NativeMethods.HWND_TOPMOST, 0, 0, 0, 0, NativeMethods.SWP_SHOWWINDOW);

                    for (int i = -10; i < 10; i++)
                    {
                        if (dragDropStep05ExCalledByIpc > 0)
                        {
                            Log("DragDropStep04: DragDropStep05ExCalledByIpc.");
                            break;
                        }

                        _ = NativeMethods.GetCursorPos(ref p);
                        Log("DragDropStep04: Moving Mouse Without Borders Helper to (" + p.X.ToString(CultureInfo.CurrentCulture) + ", " + p.Y.ToString(CultureInfo.CurrentCulture) + ")");
                        _ = NativeMethods.SetWindowPos(h, NativeMethods.HWND_TOPMOST, p.X - 100 + i, p.Y - 100 + i, 200, 200, 0);
                        _ = NativeMethods.SendMessage(h, 0x000F, IntPtr.Zero, IntPtr.Zero); // WM_PAINT
                        Thread.Sleep(20);
                        Application.DoEvents();

                        // if (GetText(h).Length > 1) break;
                    }
                }
                else
                {
                    Log("DragDropStep04: Mouse without Borders Helper not found!");
                }
            }
            else
            {
                Log("DragDropStep04: IsDropping == true, skip chekking");
            }

            Log("DragDropStep04: Got WM_CHECK_EXPLORER_DRAG_DROP, done with processing jump to DragDropStep05...");
        }

        internal static void DragDropStep05Ex(string dragFileName)
        {
            Log("DragDropStep05 called.");

            _ = Interlocked.Exchange(ref dragDropStep05ExCalledByIpc, 1);

            if (RunOnLogonDesktop || RunOnScrSaverDesktop)
            {
                return;
            }

            if (!IsDroping)
            {
                _ = Common.ImpersonateLoggedOnUserAndDoSomething(() =>
                {
                    if (!string.IsNullOrEmpty(dragFileName) && (File.Exists(dragFileName) || Directory.Exists(dragFileName)))
                    {
                        Common.LastDragDropFile = dragFileName;
                        /*
                         * possibleDropMachineID is used as desID sent in DragDropStep06();
                         * */
                        if (dropMachineID == ID.NONE)
                        {
                            dropMachineID = newDesMachineID;
                        }

                        DragDropStep06();
                        Log("DragDropStep05: File dragging: " + dragFileName);
                        _ = NativeMethods.PostMessage(MainForm.Handle, NativeMethods.WM_HIDE_DD_HELPER, (IntPtr)1, (IntPtr)0);
                    }
                    else
                    {
                        Log("DragDropStep05: File not found: [" + dragFileName + "]");
                        _ = NativeMethods.PostMessage(MainForm.Handle, NativeMethods.WM_HIDE_DD_HELPER, (IntPtr)0, (IntPtr)0);
                    }

                    Log("DragDropStep05: WM_HIDE_DDHelper sent");
                });
            }
            else
            {
                Log("DragDropStep05: IsDropping == true, change drop machine...");
                IsDroping = false;
                MainFormVisible = true; // WM_HIDE_DRAG_DROP
                SendDropBegin(); // To dropMachineID set in DragDropStep03
            }

            MouseDown = false;
        }

        internal static void DragDropStep06()
        {
            IsDragging = true;
            Log("DragDropStep06: SendClipboardBeatDragDrop");
            SendClipboardBeatDragDrop();
            SendDropBegin();
        }

        internal static void DragDropStep08(DATA package)
        {
            GetNameOfMachineWithClipboardData(package);
            Log("DragDropStep08: clipboarddd Received. machine with drag file was set");
        }

        internal static void DragDropStep08_2(DATA package)
        {
            if (package.Des == MachineID && !RunOnLogonDesktop && !RunOnScrSaverDesktop)
            {
                IsDroping = true;
                dropMachineID = MachineID;
                Log("DragDropStep08_2: Clipboarddop Received. IsDroping set");
            }
        }

        internal static void DragDropStep09(int wParam)
        {
            if (wParam == WM_MOUSEMOVE && IsDroping)
            {
                // Show/Move form
                DoSomethingInUIThread(() =>
                {
                    _ = NativeMethods.PostMessage(MainForm.Handle, NativeMethods.WM_SHOW_DRAG_DROP, (IntPtr)0, (IntPtr)0);
                });
            }
            else if (wParam == WM_LBUTTONUP && (IsDroping || IsDragging))
            {
                if (IsDroping)
                {
                    // Hide form, get data
                    DragDropStep10();
                }
                else
                {
                    IsDragging = false;
                    LastIDWithClipboardData = ID.NONE;
                }
            }
        }

        internal static void DragDropStep10()
        {
            Log("DragDropStep10: Hide the form and get data...");
            IsDroping = false;
            IsDragging = false;
            LastIDWithClipboardData = ID.NONE;

            DoSomethingInUIThread(() =>
            {
                _ = NativeMethods.PostMessage(MainForm.Handle, NativeMethods.WM_HIDE_DRAG_DROP, (IntPtr)0, (IntPtr)0);
            });

            Common.TelemetryForEvent?.TrackEvent("DraDro");
            GetRemoteClipboard("desktop");
        }

        internal static void DragDropStep11()
        {
            Log("DragDropStep11: Mouse drag comming back, candeling drag/drop");
            SendClipboardBeatDragDropEnd();
            IsDroping = false;
            IsDragging = false;
            DragMachine = (ID)1;
            LastIDWithClipboardData = ID.NONE;
            LastDragDropFile = null;
            MouseDown = false;
        }

        internal static void DragDropStep12()
        {
            Log("DragDropStep12: clipboardddend received");
            IsDroping = false;
            LastIDWithClipboardData = ID.NONE;

            DoSomethingInUIThread(() =>
            {
                _ = NativeMethods.PostMessage(MainForm.Handle, NativeMethods.WM_HIDE_DRAG_DROP, (IntPtr)0, (IntPtr)0);
            });
        }

        internal static void SendCheckExplorerDragDrop()
        {
            DATA package = new();
            package.Type = PackageType.Explorerdd;

            /*
             * package.src = newDesMachineID:
             * sent from the master machine but the src must be the
             * new des machine since the previous des machine will get this and set
             * to possibleDropMachineID in DragDropStep3()
             * */
            package.Src = newDesMachineID;

            package.Des = desMachineID;
            package.MachineName = MachineName;

            SkSend(package, null, false);
        }

        private static void ChangeDropMachine()
        {
            // desMachineID = current drop machine
            // newDesMachineID = new drop machine

            // 1. Cancelling dropping in current drop machine
            if (dropMachineID == MachineID)
            {
                // Drag/Drop coming through me
                IsDroping = false;
            }
            else
            {
                // Drag/Drop coming back
                SendClipboardBeatDragDropEnd();
            }

            // 2. SendClipboardBeatDragDrop to new drop machine
            // new drop machine is not me
            if (newDesMachineID != MachineID)
            {
                dropMachineID = newDesMachineID;
                SendDropBegin();
            }

            // New drop machine is me
            else
            {
                IsDroping = true;
            }
        }

        internal static void SendClipboardBeatDragDrop()
        {
            SendPackage(ID.ALL, PackageType.Clipboarddd);
        }

        internal static void SendDropBegin()
        {
            Log("SendDropBegin...");
            SendPackage(dropMachineID, PackageType.Clipboarddop);
        }

        internal static void SendClipboardBeatDragDropEnd()
        {
            if (desMachineID != MachineID)
            {
                SendPackage(desMachineID, PackageType.Clipboardddend);
            }
        }

        private static bool isDroping;
        private static ID dragMachine;

        internal static ID DragMachine
        {
            get => Common.dragMachine;
            set => Common.dragMachine = value;
        }

        internal static bool IsDroping
        {
            get => Common.isDroping;
            set => Common.isDroping = value;
        }

        internal static bool MouseDown { get; set; }
    }
}
