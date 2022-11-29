// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;

// <summary>
//     Back-end thread for the socket.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using MouseWithoutBorders.Class;

[module: SuppressMessage("Microsoft.Reliability", "CA2002:DoNotLockOnObjectsWithWeakIdentity", Scope = "member", Target = "MouseWithoutBorders.Common.#PreProcess(MouseWithoutBorders.DATA)", Justification = "Dotnet port with style preservation")]

namespace MouseWithoutBorders
{
    internal partial class Common
    {
        private static readonly uint QUEUESIZE = 50;
        private static readonly int[] RecentProcessedPackageIDs = new int[QUEUESIZE];
        private static int recentProcessedPackageIndex;
        private static long processedPackageCount;
        private static long skippedPackageCount;

        internal static long IJustGotAKey { get; set; }

        private static bool PreProcess(DATA package)
        {
            if (package.Type == PackageType.Invalid)
            {
                if ((Common.InvalidPackageCount % 100) == 0)
                {
                    ShowToolTip("Invalid packages received!", 1000, ToolTipIcon.Warning, false);
                }

                Common.InvalidPackageCount++;
                Common.Log("Invalid packages received!");
                return false;
            }
            else if (package.Type == 0)
            {
                Common.Log("Got an unknown package!");
                return false;
            }
            else if (package.Type is not PackageType.Clipboardtxt and not PackageType.Clipboardimg

                // BEGIN: These package types are sent by TcpSend which is single direction.
                and not PackageType.Handshake and not PackageType.Handshakeack)
            {
                // END
                lock (RecentProcessedPackageIDs)
                {
                    for (int i = 0; i < QUEUESIZE; i++)
                    {
                        if (RecentProcessedPackageIDs[i] == package.Id)
                        {
                            skippedPackageCount++;
                            return false;
                        }
                    }

                    processedPackageCount++;
                    recentProcessedPackageIndex = (int)((recentProcessedPackageIndex + 1) % QUEUESIZE);
                    RecentProcessedPackageIDs[recentProcessedPackageIndex] = package.Id;
                }
            }

            return true;
        }

        private static System.Drawing.Point lastXY;

        internal static void ProcessPackage(DATA package, TcpSk tcp)
        {
            if (!PreProcess(package))
            {
                return;
            }

            switch (package.Type)
            {
                case PackageType.Keybd:
                    PackageRecv.Keybd++;
                    if (package.Des == MachineID || package.Des == ID.ALL)
                    {
                        IJustGotAKey = GetTick();

                        // TODO(@yuyoyuppe): disabled to drop elevation requirement
                        bool nonElevated = Common.RunWithNoAdminRight && false;
                        if (nonElevated && Setting.Values.OneWayControlMode)
                        {
                            if ((package.Kd.dwFlags & (int)Common.LLKHF.UP) == (int)Common.LLKHF.UP)
                            {
                                Common.ShowOneWayModeMessage();
                            }

                            return;
                        }

                        InputSimu.SendKey(package.Kd);
                    }

                    break;

                case PackageType.Mouse:
                    PackageRecv.Mouse++;

                    if (package.Des == MachineID || package.Des == ID.ALL)
                    {
                        if (desMachineID != MachineID)
                        {
                            NewDesMachineID = DesMachineID = MachineID;
                        }

                        // TODO(@yuyoyuppe): disabled to drop elevation requirement
                        bool nonElevated = Common.RunWithNoAdminRight && false;
                        if (nonElevated && Setting.Values.OneWayControlMode && package.Md.dwFlags != Common.WM_MOUSEMOVE)
                        {
                            if (!IsDroping)
                            {
                                if (package.Md.dwFlags is WM_LBUTTONDOWN or WM_RBUTTONDOWN)
                                {
                                    Common.ShowOneWayModeMessage();
                                }
                            }
                            else if (package.Md.dwFlags is WM_LBUTTONUP or WM_RBUTTONUP)
                            {
                                IsDroping = false;
                            }

                            return;
                        }

                        if (Math.Abs(package.Md.X) >= MOVE_MOUSE_RELATIVE && Math.Abs(package.Md.Y) >= MOVE_MOUSE_RELATIVE)
                        {
                            if (package.Md.dwFlags == Common.WM_MOUSEMOVE)
                            {
                                InputSimu.MoveMouseRelative(
                                    package.Md.X < 0 ? package.Md.X + MOVE_MOUSE_RELATIVE : package.Md.X - MOVE_MOUSE_RELATIVE,
                                    package.Md.Y < 0 ? package.Md.Y + MOVE_MOUSE_RELATIVE : package.Md.Y - MOVE_MOUSE_RELATIVE);
                                _ = NativeMethods.GetCursorPos(ref lastXY);

                                Point p = MoveToMyNeighbourIfNeeded(lastXY.X, lastXY.Y, MachineID);

                                if (!p.IsEmpty)
                                {
                                    HasSwitchedMachineSinceLastCopy = true;

                                    Common.Log(string.Format(
                                        CultureInfo.CurrentCulture,
                                        "***** Controlled Machine: newDesMachineIDEx set = [{0}]. Mouse is now at ({1},{2})",
                                        newDesMachineIDEx,
                                        lastXY.X,
                                        lastXY.Y));

                                    SendNextMachine(package.Src, newDesMachineIDEx, p);
                                }
                            }
                            else
                            {
                                _ = NativeMethods.GetCursorPos(ref lastXY);
                                package.Md.X = lastXY.X * 65535 / screenWidth;
                                package.Md.Y = lastXY.Y * 65535 / screenHeight;
                                _ = InputSimu.SendMouse(package.Md);
                            }
                        }
                        else
                        {
                            _ = InputSimu.SendMouse(package.Md);
                            _ = NativeMethods.GetCursorPos(ref lastXY);
                        }

                        LastX = lastXY.X;
                        LastY = lastXY.Y;
                        CustomCursor.ShowFakeMouseCursor(LastX, LastY);
                    }

                    DragDropStep01(package.Md.dwFlags);
                    DragDropStep09(package.Md.dwFlags);
                    break;

                case PackageType.NextMachine:
                    Log("PackageType.NextMachine received!");

                    if (IsSwitchingByMouseEnabled())
                    {
                        PrepareToSwitchToMachine((ID)package.Md.WDelta, new Point(package.Md.X, package.Md.Y));
                    }

                    break;

                case PackageType.Explorerdd:
                    PackageRecv.Explorerdd++;
                    DragDropStep03(package);
                    break;

                case PackageType.Heartbeat:
                case PackageType.Heartbeatex:
                    PackageRecv.Heartbeat++;

                    Common.GeneratedKey = Common.GeneratedKey || package.Type == PackageType.Heartbeatex;

                    if (Common.GeneratedKey)
                    {
                        Setting.Values.MyKey = Common.MyKey;
                        SendPackage(ID.ALL, PackageType.Heartbeatexl2);
                    }

                    string desMachine = Common.AddToMachinePool(package);

                    if (Setting.Values.FirstRun && !string.IsNullOrEmpty(desMachine))
                    {
                        Common.UpdateSetupMachineMatrix(desMachine);
                        Common.UpdateClientSockets("UpdateSetupMachineMatrix");
                    }

                    break;

                case PackageType.Heartbeatexl2:
                    Common.GeneratedKey = true;
                    Setting.Values.MyKey = Common.MyKey;
                    SendPackage(ID.ALL, PackageType.Heartbeatexl3);

                    break;

                case PackageType.Heartbeatexl3:
                    Common.GeneratedKey = true;
                    Setting.Values.MyKey = Common.MyKey;

                    break;

                case PackageType.Awake:
                    PackageRecv.Heartbeat++;
                    _ = Common.AddToMachinePool(package);
                    Common.HumanBeingDetected();
                    break;

                case PackageType.Hello:
                    PackageRecv.Hello++;
                    SendHeartBeat();
                    string newMachine = Common.AddToMachinePool(package);
                    if (Setting.Values.MachineMatrixString == null)
                    {
                        string tip = newMachine + " saying Hello!";
                        tip += "\r\n Right Click to setup your machine Matrix";
                        ShowToolTip(tip);
                    }

                    break;

                case PackageType.Hi:
                    PackageRecv.Hello++;
                    break;

                case PackageType.Byebye:
                    PackageRecv.Byebye++;
                    ProcessByeByeMessage(package);
                    break;

                case PackageType.Clipboard:
                    PackageRecv.Clipboard++;
                    if (!RunOnLogonDesktop && !RunOnScrSaverDesktop)
                    {
                        clipboardCopiedTime = GetTick();
                        GetNameOfMachineWithClipboardData(package);
                        SignalBigClipboardData();
                    }

                    break;

                case PackageType.Machineswitched:
                    if (GetTick() - clipboardCopiedTime < BIG_CLIPBOARD_DATA_TIMEOUT && (package.Des == MachineID))
                    {
                        clipboardCopiedTime = 0;
                        Common.GetRemoteClipboard("PackageType.Machineswitched");
                    }

                    break;

                case PackageType.Clipboardcapture:
                    PackageRecv.Clipboard++;
                    if (!RunOnLogonDesktop && !RunOnScrSaverDesktop)
                    {
                        if (package.Des == MachineID || package.Des == ID.ALL)
                        {
                            GetNameOfMachineWithClipboardData(package);
                            GetRemoteClipboard("mspaint," + LastMachineWithClipboardData);
                        }
                    }

                    break;

                case PackageType.Capturescreencmd:
                    PackageRecv.Clipboard++;
                    if (package.Des == MachineID || package.Des == ID.ALL)
                    {
                        Common.SendImage(package.Src, Common.CaptureScreen());
                    }

                    break;

                case PackageType.Clipboardask:
                    PackageRecv.Clipboardask++;

                    if (package.Des == MachineID)
                    {
                        _ = Task.Run(() =>
                        {
                            try
                            {
                                System.Threading.Thread thread = Thread.CurrentThread;
                                thread.Name = $"{nameof(PackageType.Clipboardask)}.{thread.ManagedThreadId}";
                                Thread.UpdateThreads(thread);

                                string remoteMachine = package.MachineName;
                                System.Net.Sockets.TcpClient client = ConnectToRemoteClipboardSocket(remoteMachine);
                                bool clientPushData = true;

                                if (ShakeHand(ref remoteMachine, client.Client, out Stream enStream, out Stream deStream, ref clientPushData, ref package.PostAction))
                                {
                                    SocketStuff.SendClipboardData(client.Client, enStream);
                                }
                            }
                            catch (Exception e)
                            {
                                Log(e);
                            }
                        });
                    }

                    break;

                case PackageType.Clipboarddd:
                    PackageRecv.Clipboarddd++;
                    DragDropStep08(package);
                    break;

                case PackageType.Clipboarddop:
                    PackageRecv.Clipboarddd++;
                    DragDropStep08_2(package);
                    break;

                case PackageType.Clipboardddend:
                    PackageRecv.Clipboardddend++;
                    DragDropStep12();
                    break;

                case PackageType.Clipboardtxt:
                case PackageType.Clipboardimg:
                    clipboardCopiedTime = 0;
                    if (package.Type == PackageType.Clipboardimg)
                    {
                        PackageRecv.Clipboardimg++;
                    }
                    else
                    {
                        PackageRecv.Clipboardtxt++;
                    }

                    if (tcp != null)
                    {
                        Common.ReceiveClipboardDataUsingTCP(
                            package,
                            package.Type == PackageType.Clipboardimg,
                            tcp);
                    }

                    break;

                case PackageType.Hidemouse:
                    HasSwitchedMachineSinceLastCopy = true;
                    HideMouseCursor(true);
                    MainFormDotEx(false);
                    ReleaseAllKeys();
                    break;

                default:
                    if ((package.Type & PackageType.Matrix) == PackageType.Matrix)
                    {
                        PackageRecv.Matrix++;
                        UpdateMachineMatrix(package);
                        break;
                    }
                    else
                    {
                        // We should never get to this point!
                        Common.Log("Invalid package received!");
                        return;
                    }
            }
        }

        private static void GetNameOfMachineWithClipboardData(DATA package)
        {
            LastIDWithClipboardData = package.Src;
            List<MachineInf> matchingMachines = Common.MachinePool.TryFindMachineByID(LastIDWithClipboardData);
            if (matchingMachines.Count >= 1)
            {
                LastMachineWithClipboardData = matchingMachines[0].Name.Trim();
            }

            /*
            lastMachineWithClipboardData =
                Common.GetString(BitConverter.GetBytes(package.machineNameHead));
            lastMachineWithClipboardData +=
                Common.GetString(BitConverter.GetBytes(package.machineNameTail));
            lastMachineWithClipboardData = lastMachineWithClipboardData.Trim();
             * */
        }

        private static void SignalBigClipboardData()
        {
            Log("SignalBigClipboardData");
            SetToggleIcon(new int[TOGGLE_ICONS_SIZE] { ICON_BIG_CLIPBOARD, -1, ICON_BIG_CLIPBOARD, -1 });
        }
    }
}
