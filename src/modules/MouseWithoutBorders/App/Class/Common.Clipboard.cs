// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Sockets;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

// <summary>
//     Clipboard related routines.
// </summary>
// <history>
//     2008 created by Truong Do (ductdo).
//     2009-... modified by Truong Do (TruongDo).
// </history>
using MouseWithoutBorders.Class;
using MouseWithoutBorders.Exceptions;
using Clipb = System.Windows.Forms.Clipboard;

namespace MouseWithoutBorders
{
    internal partial class Common
    {
        private const uint BIG_CLIPBOARD_DATA_TIMEOUT = 30000;
        private const uint MAX_CLIPBOARD_DATA_SIZE_CAN_BE_SENT_INSTANTLY_TCP = 1024 * 1024; // 1MB
        private const uint MAX_CLIPBOARD_FILE_SIZE_CAN_BE_SENT = 100 * 1024 * 1024; // 100MB
        private const int TEXT_HEADER_SIZE = 12;
        private const int DATA_SIZE = 48;
        private const string TEXTTYPE_SEP = "{4CFF57F7-BEDD-43d5-AE8F-27A61E886F2F}";
        private static long lastClipboardEventTime;
        private static string lastMachineWithClipboardData;
        private static string lastDragDropFile;
        private static long clipboardCopiedTime;

        internal static ID LastIDWithClipboardData { get; set; }

        internal static string LastDragDropFile
        {
            get => Common.lastDragDropFile;
            set => Common.lastDragDropFile = value;
        }

        internal static string LastMachineWithClipboardData
        {
            get => Common.lastMachineWithClipboardData;
            set => Common.lastMachineWithClipboardData = value;
        }

        internal static long LastClipboardEventTime
        {
            get => Common.lastClipboardEventTime;
            set => Common.lastClipboardEventTime = value;
        }

        internal static IntPtr NextClipboardViewer { get; set; }

        internal static bool IsClipboardDataImage { get; private set; }

        internal static byte[] LastClipboardData { get; private set; }

        private static object lastClipboardObject = string.Empty;

        internal static bool HasSwitchedMachineSinceLastCopy { get; set; }

        internal static bool CheckClipboardEx(object data, bool isFilePath)
        {
            Log($"{nameof(CheckClipboardEx)}: ShareClipboard = {Setting.Values.ShareClipboard}, TransferFile = {Setting.Values.TransferFile}, data = {data?.GetType().Name}.");
            Log($"{nameof(CheckClipboardEx)}: {nameof(Setting.Values.OneWayClipboardMode)} = {Setting.Values.OneWayClipboardMode}.");

            if (!Setting.Values.ShareClipboard)
            {
                return false;
            }

            if (Common.RunWithNoAdminRight && Setting.Values.OneWayClipboardMode)
            {
                return false;
            }

            if (GetTick() - LastClipboardEventTime < 1000)
            {
                Log("GetTick() - lastClipboardEventTime < 1000");
                LastClipboardEventTime = GetTick();
                return false;
            }

            LastClipboardEventTime = GetTick();

            try
            {
                IsClipboardDataImage = false;
                LastClipboardData = null;
                LastDragDropFile = null;
                GC.Collect();

                string stringData = data as string;
                byte[] byteData = stringData != null ? null : data as byte[];

                if (stringData != null)
                {
                    if (!HasSwitchedMachineSinceLastCopy)
                    {
                        if (lastClipboardObject is string lastStringData && lastStringData.Equals(stringData, StringComparison.OrdinalIgnoreCase))
                        {
                            Log("CheckClipboardEx: Same string data.");
                            return false;
                        }
                    }

                    HasSwitchedMachineSinceLastCopy = false;

                    if (isFilePath)
                    {
                        Common.Log("Clipboard cotains FileDropList");

                        if (!Setting.Values.TransferFile)
                        {
                            Common.Log("TransferFile option is unchecked.");
                            return false;
                        }

                        string filePath = stringData;

                        _ = Common.ImpersonateLoggedOnUserAndDoSomething(() =>
                        {
                            if (File.Exists(filePath) || Directory.Exists(filePath))
                            {
                                if (File.Exists(filePath) && new FileInfo(filePath).Length <= MAX_CLIPBOARD_FILE_SIZE_CAN_BE_SENT)
                                {
                                    Log("Clipboard contains: " + filePath);
                                    LastDragDropFile = filePath;
                                    SendClipboardBeat();
                                    SetToggleIcon(new int[TOGGLE_ICONS_SIZE] { ICON_BIG_CLIPBOARD, -1, ICON_BIG_CLIPBOARD, -1 });
                                }
                                else
                                {
                                    if (Directory.Exists(filePath))
                                    {
                                        Log("Clipboard contains a directory: " + filePath);
                                        LastDragDropFile = filePath;
                                        SendClipboardBeat();
                                    }
                                    else
                                    {
                                        LastDragDropFile = filePath + " - File too big (greater than 100MB), please drag and drop the file instead!";
                                        SendClipboardBeat();
                                        Log("Clipboard: File too big: " + filePath);
                                    }

                                    SetToggleIcon(new int[TOGGLE_ICONS_SIZE] { ICON_ERROR, -1, ICON_ERROR, -1 });
                                }
                            }
                            else
                            {
                                Log("CheckClipboardEx: File not found: " + filePath);
                            }
                        });
                    }
                    else
                    {
                        byte[] texts = Common.GetBytesU(stringData);

                        using MemoryStream ms = new();
                        using (DeflateStream s = new(ms, CompressionMode.Compress, true))
                        {
                            s.Write(texts, 0, texts.Length);
                        }

                        Common.Log("Plain/Zip = " + texts.Length.ToString(CultureInfo.CurrentCulture) + "/" +
                            ms.Length.ToString(CultureInfo.CurrentCulture));

                        LastClipboardData = ms.GetBuffer();
                    }
                }
                else if (byteData != null)
                {
                    if (!HasSwitchedMachineSinceLastCopy)
                    {
                        if (lastClipboardObject is byte[] lastByteData && Enumerable.SequenceEqual(lastByteData, byteData))
                        {
                            Log("CheckClipboardEx: Same byte[] data.");
                            return false;
                        }
                    }

                    HasSwitchedMachineSinceLastCopy = false;

                    Common.Log("Clipboard cotains image");
                    IsClipboardDataImage = true;
                    LastClipboardData = byteData;
                }
                else
                {
                    Log("*** Clipboard contains something else!");
                    return false;
                }

                lastClipboardObject = data;

                if (LastClipboardData != null && LastClipboardData.Length > 0)
                {
                    if (LastClipboardData.Length > MAX_CLIPBOARD_DATA_SIZE_CAN_BE_SENT_INSTANTLY_TCP)
                    {
                        SendClipboardBeat();
                        SetToggleIcon(new int[TOGGLE_ICONS_SIZE] { ICON_BIG_CLIPBOARD, -1, ICON_BIG_CLIPBOARD, -1 });
                    }
                    else
                    {
                        SetToggleIcon(new int[TOGGLE_ICONS_SIZE] { ICON_SMALL_CLIPBOARD, -1, -1, -1 });
                        SendClipboardDataUsingTCP(LastClipboardData, IsClipboardDataImage);
                    }

                    return true;
                }
            }
            catch (Exception e)
            {
                Log(e);
            }

            return false;
        }

        private static void SendClipboardDataUsingTCP(byte[] bytes, bool image)
        {
            if (Sk == null)
            {
                return;
            }

            new Task(() =>
            {
                System.Threading.Thread thread = Thread.CurrentThread;
                thread.Name = $"{nameof(SendClipboardDataUsingTCP)}.{thread.ManagedThreadId}";
                Thread.UpdateThreads(thread);
                int l = bytes.Length;
                int index = 0;
                int len;
                DATA package = new();
                byte[] buf = new byte[PACKAGESIZEEX];
                int dataStart = PACKAGESIZEEX - DATA_SIZE;

                while (true)
                {
                    if ((index + DATA_SIZE) > l)
                    {
                        len = l - index;
                        Array.Clear(buf, 0, PACKAGESIZEEX);
                    }
                    else
                    {
                        len = DATA_SIZE;
                    }

                    Array.Copy(bytes, index, buf, dataStart, len);
                    package.Bytes = buf;

                    package.Type = image ? PackageType.Clipboardimg : PackageType.Clipboardtxt;
                    package.Des = ID.ALL;
                    SkSend(package, (uint)MachineID, false);

                    index += DATA_SIZE;
                    if (index >= l)
                    {
                        break;
                    }
                }

                package.Type = PackageType.Clipboarddataend;
                package.Des = ID.ALL;
                SkSend(package, (uint)MachineID, false);
            }).Start();
        }

        internal static void ReceiveClipboardDataUsingTCP(DATA data, bool image, TcpSk tcp)
        {
            try
            {
                if (Sk == null || RunOnLogonDesktop || RunOnScrSaverDesktop)
                {
                    return;
                }

                MemoryStream m = new();
                int dataStart = PACKAGESIZEEX - DATA_SIZE;
                m.Write(data.Bytes, dataStart, DATA_SIZE);
                int unexpectedCount = 0;

                bool done = false;
                do
                {
                    data = SocketStuff.TcpRecvData(tcp, out int err);

                    switch (data.Type)
                    {
                        case PackageType.Clipboardimg:
                        case PackageType.Clipboardtxt:
                            m.Write(data.Bytes, dataStart, DATA_SIZE);
                            break;

                        case PackageType.Clipboarddataend:
                            done = true;
                            break;

                        default:
                            ProcessPackage(data, tcp);
                            if (++unexpectedCount > 100)
                            {
                                Log("ReceiveClipboardDataUsingTCP: unexpectedCount > 100!");
                                done = true;
                            }

                            break;
                    }
                }
                while (!done);

                LastClipboardEventTime = GetTick();

                if (image)
                {
                    Image im = Image.FromStream(m);
                    Clipboard.SetImage(im);
                    LastClipboardEventTime = GetTick();
                }
                else
                {
                    Common.SetClipboardData(m.GetBuffer());
                    LastClipboardEventTime = GetTick();
                }

                m.Dispose();

                SetToggleIcon(new int[TOGGLE_ICONS_SIZE] { ICON_SMALL_CLIPBOARD, -1, ICON_SMALL_CLIPBOARD, -1 });
            }
            catch (Exception e)
            {
                Log("ReceiveClipboardDataUsingTCP: " + e.Message);
            }
        }

        private static readonly object ClipboardThreadOldLock = new();
        private static System.Threading.Thread clipboardThreadOld;

        internal static void GetRemoteClipboard(string postAction)
        {
            if (!RunOnLogonDesktop && !RunOnScrSaverDesktop)
            {
                if (Common.LastMachineWithClipboardData == null ||
                    Common.LastMachineWithClipboardData.Length < 1)
                {
                    return;
                }

                new Task(() =>
                {
                    System.Threading.Thread thread = Thread.CurrentThread;
                    thread.Name = $"{nameof(ConnectAndGetData)}.{thread.ManagedThreadId}";
                    Thread.UpdateThreads(thread);
                    ConnectAndGetData(postAction);
                }).Start();
            }
        }

        private static Stream m;

        private static void ConnectAndGetData(object postAction)
        {
            if (Sk == null)
            {
                Log("ConnectAndGetData: Sk == null!");
                return;
            }

            string remoteMachine;
            TcpClient clipboardTcpClient = null;
            string postAct = (string)postAction;

            Log("ConnectAndGetData.postAction: " + postAct);

            ClipboardPostAction clipboardPostAct = postAct.Contains("mspaint,") ? ClipboardPostAction.Mspaint
                : postAct.Equals("desktop", StringComparison.OrdinalIgnoreCase) ? ClipboardPostAction.Desktop
                : ClipboardPostAction.Other;

            try
            {
                remoteMachine = postAct.Contains("mspaint,") ? postAct.Split(new char[] { ',' })[1] : Common.LastMachineWithClipboardData;

                remoteMachine = remoteMachine.Trim();

                if (!IsConnectedByAClientSocketTo(remoteMachine))
                {
                    Log($"No potential inbound connection from {MachineName} to {remoteMachine}, ask for a push back instead.");
                    ID ipn = MachinePool.ResolveID(remoteMachine);

                    if (ipn != ID.NONE)
                    {
                        SkSend(
                            new DATA()
                            {
                                Type = PackageType.Clipboardask,
                                Des = ipn,
                                MachineName = MachineName,
                                PostAction = clipboardPostAct,
                            },
                            null,
                            false);
                    }
                    else
                    {
                        Log($"Unable to resolve {remoteMachine} to its long IP.");
                    }

                    return;
                }

                ShowToolTip("Connecting to " + remoteMachine, 2000, ToolTipIcon.Info, Setting.Values.ShowClipNetStatus);

                clipboardTcpClient = ConnectToRemoteClipboardSocket(remoteMachine);
            }
            catch (ThreadAbortException)
            {
                Common.Log("The current thread is being aborted (1).");
                if (clipboardTcpClient != null && clipboardTcpClient.Connected)
                {
                    clipboardTcpClient.Client.Close();
                }

                return;
            }
            catch (Exception e)
            {
                Common.Log(e);
                Common.SetToggleIcon(new int[Common.TOGGLE_ICONS_SIZE]
                {
                    Common.ICON_BIG_CLIPBOARD,
                    -1, Common.ICON_BIG_CLIPBOARD, -1,
                });
                ShowToolTip(e.Message, 1000, ToolTipIcon.Warning, Setting.Values.ShowClipNetStatus);
                return;
            }

            bool clientPushData = false;

            if (!ShakeHand(ref remoteMachine, clipboardTcpClient.Client, out Stream enStream, out Stream deStream, ref clientPushData, ref clipboardPostAct))
            {
                return;
            }

            ReceiveAndProcessClipboardData(remoteMachine, clipboardTcpClient.Client, enStream, deStream, postAct);
        }

        internal static void ReceiveAndProcessClipboardData(string remoteMachine, Socket s, Stream enStream, Stream deStream, string postAct)
        {
            lock (ClipboardThreadOldLock)
            {
                // Do not enable two connections at the same time.
                if (clipboardThreadOld != null && clipboardThreadOld.ThreadState != System.Threading.ThreadState.AbortRequested
                    && clipboardThreadOld.ThreadState != System.Threading.ThreadState.Aborted && clipboardThreadOld.IsAlive
                    && clipboardThreadOld.ManagedThreadId != Thread.CurrentThread.ManagedThreadId)
                {
                    // TODO(@yuyoyuppe): thread cancellation token
                    // clipboardThreadOld.Abort();
                    if (clipboardThreadOld.Join(3000))
                    {
                        if (m != null)
                        {
                            m.Flush();
                            m.Close();
                            m = null;
                        }
                    }
                }

                clipboardThreadOld = Thread.CurrentThread;
            }

            try
            {
                byte[] header = new byte[1024];
                byte[] buf = new byte[NETWORK_STREAM_BUF_SIZE];
                string fileName = null;
                string tempFile = "data", savingFolder = string.Empty;
                Common.ToggleIconsIndex = 0;
                int rv;
                long receivedCount = 0;

                if ((rv = deStream.ReadEx(header, 0, header.Length)) < header.Length)
                {
                    Common.Log("Reading header failed: " + rv.ToString(CultureInfo.CurrentCulture));
                    Common.SetToggleIcon(new int[Common.TOGGLE_ICONS_SIZE]
                    {
                        Common.ICON_BIG_CLIPBOARD,
                        -1, -1, -1,
                    });
                    return;
                }

                fileName = Common.GetStringU(header).Replace("\0", string.Empty);
                Common.Log("Header: " + fileName);
                string[] headers = fileName.Split(new char[] { '*' });

                if (headers.Length < 2 || !long.TryParse(headers[0], out long dataSize))
                {
                    Common.Log(string.Format(
                        CultureInfo.CurrentCulture,
                        "Reading header failed: {0}:{1}",
                        headers.Length,
                        fileName));
                    Common.SetToggleIcon(new int[Common.TOGGLE_ICONS_SIZE]
                    {
                        Common.ICON_BIG_CLIPBOARD,
                        -1, -1, -1,
                    });
                    return;
                }

                fileName = headers[1];

                Common.Log(string.Format(
                    CultureInfo.CurrentCulture,
                    "Receiving {0}:{1} from {2}...",
                    Path.GetFileName(fileName),
                    dataSize,
                    remoteMachine));
                ShowToolTip(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        "Receiving {0} from {1}...",
                        Path.GetFileName(fileName),
                        remoteMachine),
                    5000,
                    ToolTipIcon.Info,
                    Setting.Values.ShowClipNetStatus);
                if (fileName.StartsWith("image", StringComparison.CurrentCultureIgnoreCase) ||
                    fileName.StartsWith("text", StringComparison.CurrentCultureIgnoreCase))
                {
                    m = new MemoryStream();
                }
                else
                {
                    if (postAct.Equals("desktop", StringComparison.OrdinalIgnoreCase))
                    {
                        _ = ImpersonateLoggedOnUserAndDoSomething(() =>
                        {
                            savingFolder = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\MouseWithoutBorders\\";

                            if (!Directory.Exists(savingFolder))
                            {
                                _ = Directory.CreateDirectory(savingFolder);
                            }
                        });

                        tempFile = savingFolder + Path.GetFileName(fileName);
                        m = new FileStream(tempFile, FileMode.Create);
                    }
                    else if (postAct.Contains("mspaint"))
                    {
                        tempFile = GetMyStorageDir() + @"ScreenCapture-" +
                            remoteMachine + ".png";
                        m = new FileStream(tempFile, FileMode.Create);
                    }
                    else
                    {
                        tempFile = GetMyStorageDir();
                        tempFile += Path.GetFileName(fileName);
                        m = new FileStream(tempFile, FileMode.Create);
                    }

                    Common.Log("==> " + tempFile);
                }

                ShowToolTip(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        "Receiving {0} from {1}...",
                        Path.GetFileName(fileName),
                        remoteMachine),
                    5000,
                    ToolTipIcon.Info,
                    Setting.Values.ShowClipNetStatus);

                do
                {
                    rv = deStream.ReadEx(buf, 0, buf.Length);

                    if (rv > 0)
                    {
                        receivedCount += rv;

                        if (receivedCount > dataSize)
                        {
                            rv -= (int)(receivedCount - dataSize);
                        }

                        m.Write(buf, 0, rv);
                    }

                    if (Common.ToggleIcons == null)
                    {
                        Common.SetToggleIcon(new int[Common.TOGGLE_ICONS_SIZE]
                        {
                                    Common.ICON_SMALL_CLIPBOARD,
                                    -1, Common.ICON_SMALL_CLIPBOARD, -1,
                        });
                    }

                    string text = string.Format(CultureInfo.CurrentCulture, "{0}KB received: {1}", m.Length / 1024, Path.GetFileName(fileName));

                    DoSomethingInUIThread(() =>
                    {
                        MainForm.SetTrayIconText(text);
                    });
                }
                while (rv > 0);

                if (m != null && fileName != null)
                {
                    m.Flush();
                    Common.Log(m.Length.ToString(CultureInfo.CurrentCulture) + " bytes received.");
                    Common.LastClipboardEventTime = Common.GetTick();
                    string toolTipText = null;
                    string sizeText = m.Length >= 1024
                        ? (m.Length / 1024).ToString(CultureInfo.CurrentCulture) + "KB"
                        : m.Length.ToString(CultureInfo.CurrentCulture) + "Bytes";
                    TelemetryForEvent?.TrackEvent($"BigC", new Dictionary<string, string> { { "Size", sizeText } });

                    if (fileName.StartsWith("image", StringComparison.CurrentCultureIgnoreCase))
                    {
                        Clipboard.SetImage(Image.FromStream(m));
                        toolTipText = string.Format(
                            CultureInfo.CurrentCulture,
                            "{0} {1} from {2} is in Clipboard.",
                            sizeText,
                            "image",
                            remoteMachine);
                    }
                    else if (fileName.StartsWith("text", StringComparison.CurrentCultureIgnoreCase))
                    {
                        byte[] data = (m as MemoryStream).GetBuffer();
                        toolTipText = string.Format(
                            CultureInfo.CurrentCulture,
                            "{0} {1} from {2} is in Clipboard.",
                            sizeText,
                            "text",
                            remoteMachine);
                        Common.SetClipboardData(data);
                    }
                    else if (tempFile != null)
                    {
                        if (postAct.Equals("desktop", StringComparison.OrdinalIgnoreCase))
                        {
                            toolTipText = string.Format(
                                CultureInfo.CurrentCulture,
                                "{0} {1} received from {2}!",
                                sizeText,
                                Path.GetFileName(fileName),
                                remoteMachine);

                            _ = ImpersonateLoggedOnUserAndDoSomething(() =>
                            {
                                ProcessStartInfo startInfo = new();
                                startInfo.UseShellExecute = true;
                                startInfo.WorkingDirectory = savingFolder;
                                startInfo.FileName = savingFolder;
                                startInfo.Verb = "open";
                                _ = Process.Start(startInfo);
                            });
                        }
                        else if (postAct.Contains("mspaint"))
                        {
                            m.Close();
                            m = null;
                            OpenImage(tempFile);
                            toolTipText = string.Format(
                                CultureInfo.CurrentCulture,
                                "{0} {1} from {2} is in Mspaint.",
                                sizeText,
                                Path.GetFileName(tempFile),
                                remoteMachine);
                        }
                        else
                        {
                            StringCollection filePaths = new()
                            {
                                tempFile,
                            };
                            Clipboard.SetFileDropList(filePaths);
                            toolTipText = string.Format(
                                CultureInfo.CurrentCulture,
                                "{0} {1} from {2} is in Clipboard.",
                                sizeText,
                                Path.GetFileName(fileName),
                                remoteMachine);
                        }
                    }

                    if (!string.IsNullOrWhiteSpace(toolTipText))
                    {
                        Common.ShowToolTip(toolTipText, 5000, ToolTipIcon.Info, Setting.Values.ShowClipNetStatus);
                    }

                    DoSomethingInUIThread(() =>
                    {
                        MainForm.UpdateNotifyIcon();
                    });

                    m?.Close();
                    m = null;
                }
            }
            catch (ThreadAbortException)
            {
                Common.Log("The current thread is being aborted (3).");
                s.Close();

                if (m != null)
                {
                    m.Close();
                    m = null;
                }

                return;
            }
            catch (Exception e)
            {
                if (e is IOException)
                {
                    string log = $"{nameof(ReceiveAndProcessClipboardData)}: Exception accessing the socket: {e.InnerException?.GetType()}/{e.Message}. (This is expected when the remote machine closes the connection during desktop switch or reconnection.)";
                    Common.Log(log);
                }
                else
                {
                    Common.Log(e);
                }

                Common.SetToggleIcon(new int[Common.TOGGLE_ICONS_SIZE]
                {
                    Common.ICON_BIG_CLIPBOARD,
                    -1, Common.ICON_BIG_CLIPBOARD, -1,
                });
                ShowToolTip(e.Message, 1000, ToolTipIcon.Info, Setting.Values.ShowClipNetStatus);

                if (m != null)
                {
                    m.Close();
                    m = null;
                }

                return;
            }

            s.Close();
        }

        internal static bool ShakeHand(ref string remoteName, Socket s, out Stream enStream, out Stream deStream, ref bool clientPushData, ref ClipboardPostAction postAction)
        {
            const int CLIPBOARD_HANDSHAKE_TIMEOUT = 30;
            s.ReceiveTimeout = CLIPBOARD_HANDSHAKE_TIMEOUT * 1000;
            s.NoDelay = true;
            s.SendBufferSize = s.ReceiveBufferSize = 1024000;

            bool handShaken = false;
            enStream = deStream = null;

            try
            {
                DATA package = new()
                {
                    Type = clientPushData ? PackageType.Clipboarpush : PackageType.Clipboard,
                    PostAction = postAction,
                    Src = MachineID,
                    MachineName = MachineName,
                };

                byte[] buf = new byte[PACKAGESIZEEX];

                NetworkStream ns = new(s);
                enStream = Common.GetEncryptedStream(ns);
                Common.SendOrReceiveARandomDataBlockPerInitialIV(enStream);
                Log($"{nameof(ShakeHand)}: Writing header package.");
                enStream.Write(package.Bytes, 0, PACKAGESIZEEX);

                Log($"{nameof(ShakeHand)}: Sent: clientPush={clientPushData}, postAction={postAction}.");

                deStream = Common.GetDecryptedStream(ns);
                Common.SendOrReceiveARandomDataBlockPerInitialIV(deStream, false);

                Log($"{nameof(ShakeHand)}: Reading header package.");

                int bytesReceived = deStream.ReadEx(buf, 0, Common.PACKAGESIZEEX);
                package.Bytes = buf;

                string name = "Unknown";

                if (bytesReceived == Common.PACKAGESIZEEX)
                {
                    if (package.Type is PackageType.Clipboard or PackageType.Clipboarpush)
                    {
                        name = remoteName = package.MachineName;

                        Common.Log($"{nameof(ShakeHand)}: Connection from {name}:{package.Src}");

                        if (Common.MachinePool.ResolveID(name) == package.Src && Common.IsConnectedTo(package.Src))
                        {
                            clientPushData = package.Type == PackageType.Clipboarpush;
                            postAction = package.PostAction;
                            handShaken = true;
                            Log($"{nameof(ShakeHand)}: Received: clientPush={clientPushData}, postAction={postAction}.");
                        }
                        else
                        {
                            Common.Log($"{nameof(ShakeHand)}: No active connection to the machine: {name}.");
                        }
                    }
                    else
                    {
                        Common.Log($"{nameof(ShakeHand)}: Unexpected package type: {package.Type}.");
                    }
                }
                else
                {
                    Common.Log($"{nameof(ShakeHand)}: BytesTransferred != PACKAGESIZEEX: {bytesReceived}");
                }

                if (!handShaken)
                {
                    string msg = $"Clipboard connection rejected: {name}:{remoteName}/{package.Src}\r\n\r\nMake sure you run the same version in all machines.";
                    Common.Log(msg);
                    Common.ShowToolTip(msg, 3000, ToolTipIcon.Warning);
                    Common.SetToggleIcon(new int[Common.TOGGLE_ICONS_SIZE] { Common.ICON_BIG_CLIPBOARD, -1, -1, -1 });
                }
            }
            catch (ThreadAbortException)
            {
                Common.Log($"{nameof(ShakeHand)}: The current thread is being aborted.");
                s.Close();
            }
            catch (Exception e)
            {
                if (e is IOException)
                {
                    string log = $"{nameof(ShakeHand)}: Exception accessing the socket: {e.InnerException?.GetType()}/{e.Message}. (This is expected when the remote machine closes the connection during desktop switch or reconnection.)";
                    Common.Log(log);
                }
                else
                {
                    Common.Log(e);
                }

                Common.SetToggleIcon(new int[Common.TOGGLE_ICONS_SIZE]
                {
                    Common.ICON_BIG_CLIPBOARD,
                    -1, Common.ICON_BIG_CLIPBOARD, -1,
                });
                MainForm.UpdateNotifyIcon();
                ShowToolTip(e.Message + "\r\n\r\nMake sure you run the same version in all machines.", 1000, ToolTipIcon.Warning, Setting.Values.ShowClipNetStatus);

                if (m != null)
                {
                    m.Close();
                    m = null;
                }
            }

            return handShaken;
        }

        internal static TcpClient ConnectToRemoteClipboardSocket(string remoteMachine)
        {
            TcpClient clipboardTcpClient;
            clipboardTcpClient = new TcpClient(AddressFamily.InterNetworkV6);
            clipboardTcpClient.Client.DualMode = true;

            SocketStuff sk = Common.Sk;

            if (sk != null)
            {
                Common.DoSomethingInUIThread(() => Common.MainForm.ChangeIcon(Common.ICON_SMALL_CLIPBOARD));

                System.Net.IPAddress ip = GetConnectedClientSocketIPAddressFor(remoteMachine);
                Common.Log($"{nameof(ConnectToRemoteClipboardSocket)}Connecting to {remoteMachine}:{ip}:{sk.TcpPort}...");

                if (ip != null)
                {
                    clipboardTcpClient.Connect(ip, sk.TcpPort);
                }
                else
                {
                    clipboardTcpClient.Connect(remoteMachine, sk.TcpPort);
                }

                Common.Log($"Connected from {clipboardTcpClient.Client.LocalEndPoint}. Getting data...");
                return clipboardTcpClient;
            }
            else
            {
                throw new ExpectedSocketException($"{nameof(ConnectToRemoteClipboardSocket)}: No longer connected.");
            }
        }

        internal static void SetClipboardData(byte[] data)
        {
            if (data == null || data.Length <= 0)
            {
                Common.Log("data is null or empty!");
                return;
            }

            if (data.Length > 1024000)
            {
                ShowToolTip(
                    string.Format(
                        CultureInfo.CurrentCulture,
                        "Decompressing {0} clipboard data ...",
                        (data.Length / 1024).ToString(CultureInfo.CurrentCulture) + "KB"),
                    5000,
                    ToolTipIcon.Info,
                    Setting.Values.ShowClipNetStatus);
            }

            string st = string.Empty;

            using (MemoryStream ms = new(data))
            {
                using DeflateStream s = new(ms, CompressionMode.Decompress, true);
                const int BufferSize = 1024000; // Buffer size should be big enough, this is critical to performance!

                int rv = 0;

                do
                {
                    byte[] buffer = new byte[BufferSize];
                    rv = s.ReadEx(buffer, 0, BufferSize);

                    if (rv > 0)
                    {
                        st += Common.GetStringU(buffer);
                    }
                    else
                    {
                        break;
                    }
                }
                while (true);
            }

            int textTypeCount = 0;
            string[] txts = st.Split(new string[] { TEXTTYPE_SEP }, StringSplitOptions.RemoveEmptyEntries);
            string tmp;
            DataObject data1 = new();

            foreach (string txt in txts)
            {
                if (string.IsNullOrEmpty(txt.Trim(new char[] { '\0' })))
                {
                    continue;
                }

                tmp = txt[3..];

                if (txt.StartsWith("RTF", StringComparison.CurrentCultureIgnoreCase))
                {
                    Common.Log(((double)tmp.Length / 1024).ToString("0.00", CultureInfo.InvariantCulture) + "KB of RTF <-");
                    data1.SetData(DataFormats.Rtf, tmp);
                }
                else if (txt.StartsWith("HTM", StringComparison.CurrentCultureIgnoreCase))
                {
                    Common.Log(((double)tmp.Length / 1024).ToString("0.00", CultureInfo.InvariantCulture) + "KB of HTM <-");
                    data1.SetData(DataFormats.Html, tmp);
                }
                else if (txt.StartsWith("TXT", StringComparison.CurrentCultureIgnoreCase))
                {
                    Common.Log(((double)tmp.Length / 1024).ToString("0.00", CultureInfo.InvariantCulture) + "KB of TXT <-");
                    data1.SetData(DataFormats.UnicodeText, tmp);
                }
                else
                {
                    if (textTypeCount == 0)
                    {
                        Common.Log(((double)txt.Length / 1024).ToString("0.00", CultureInfo.InvariantCulture) + "KB of UNI <-");
                        data1.SetData(DataFormats.UnicodeText, txt);
                    }

                    Common.Log("Invalid clipboard format received!");
                }

                textTypeCount++;
            }

            if (textTypeCount > 0)
            {
                Clipboard.SetDataObject(data1);
            }
        }
    }

    internal static class Clipboard
    {
        public static void SetFileDropList(StringCollection filePaths)
        {
            Common.DoSomethingInUIThread(() =>
            {
                try
                {
                    _ = Common.Retry(
                        nameof(Clipb.SetFileDropList),
                        () =>
                        {
                            Clipb.SetFileDropList(filePaths);
                            return true;
                        },
                        (log) => Common.TelemetryLogTrace(
                            log,
                            SeverityLevel.Information),
                        () => Common.LastClipboardEventTime = Common.GetTick());
                }
                catch (ExternalException e)
                {
                    Common.Log(e);
                }
                catch (ThreadStateException e)
                {
                    Common.Log(e);
                }
                catch (ArgumentNullException e)
                {
                    Common.Log(e);
                }
                catch (ArgumentException e)
                {
                    Common.Log(e);
                }
            });
        }

        public static void SetImage(Image image)
        {
            Common.DoSomethingInUIThread(() =>
            {
                try
                {
                    _ = Common.Retry(
                        nameof(Clipb.SetImage),
                        () =>
                    {
                        Clipb.SetImage(image);
                        return true;
                    },
                        (log) => Common.TelemetryLogTrace(log, SeverityLevel.Information),
                        () => Common.LastClipboardEventTime = Common.GetTick());
                }
                catch (ExternalException e)
                {
                    Common.Log(e);
                }
                catch (ThreadStateException e)
                {
                    Common.Log(e);
                }
                catch (ArgumentNullException e)
                {
                    Common.Log(e);
                }
            });
        }

        public static void SetText(string text)
        {
            Common.DoSomethingInUIThread(() =>
            {
                try
                {
                    _ = Common.Retry(
                        nameof(Clipb.SetText),
                        () =>
                    {
                        Clipb.SetText(text);
                        return true;
                    },
                        (log) => Common.TelemetryLogTrace(log, SeverityLevel.Information),
                        () => Common.LastClipboardEventTime = Common.GetTick());
                }
                catch (ExternalException e)
                {
                    Common.Log(e);
                }
                catch (ThreadStateException e)
                {
                    Common.Log(e);
                }
                catch (ArgumentNullException e)
                {
                    Common.Log(e);
                }
            });
        }

        public static void SetDataObject(DataObject dataObject)
        {
            Common.DoSomethingInUIThread(() =>
            {
                try
                {
                    Clipb.SetDataObject(dataObject, true, 10, 200);
                }
                catch (ExternalException e)
                {
                    string dataFormats = string.Join(",", dataObject.GetFormats());
                    Common.Log($"{e.Message}: {dataFormats}");
                }
                catch (ThreadStateException e)
                {
                    Common.Log(e);
                }
                catch (ArgumentNullException e)
                {
                    Common.Log(e);
                }
            });
        }
    }
}
