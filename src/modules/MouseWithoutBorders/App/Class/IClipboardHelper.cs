// Copyright (c) Microsoft Corporation
// The Microsoft Corporation licenses this file to you under the MIT license.
// See the LICENSE file in the project root for more information.

using System;
using System.Collections;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Drawing;
using System.IO.Pipes;
using System.Runtime.InteropServices;
using System.Security.Principal;
using System.Threading;
using System.Windows.Forms;
#if !MM_HELPER
using MouseWithoutBorders.Class;
#endif
using Clipb = System.Windows.Forms.Clipboard;

namespace MouseWithoutBorders
{
    public interface IClipboardHelper
    {
        void SendLog(string log);

        void SendDragFile(string fileName);

        void SendClipboardData(object data, bool isFilePath);
    }

#if !MM_HELPER
    public sealed class ClipboardHelper : MarshalByRefObject, IClipboardHelper
    {
        void MouseWithoutBorders.IClipboardHelper.SendLog(string log)
        {
            Common.Log("FROM HELPER: " + log);

            if (!string.IsNullOrEmpty(log))
            {
                if (log.StartsWith("Screen capture ended", StringComparison.InvariantCulture))
                {
                    Common.TelemetryForEvent?.TrackEvent("SSCap");

                    if (Setting.Values.FisrtCtrlShiftS)
                    {
                        Setting.Values.FisrtCtrlShiftS = false;
                        Common.ShowToolTip("Selective screen capture has been triggered, you can change the hotkey on the Settings form.", 10000);
                    }
                }
                else if (log.StartsWith("Trace:", StringComparison.InvariantCulture))
                {
                    Common.TelemetryLogTrace(log, SeverityLevel.Information);
                }
            }
        }

        void MouseWithoutBorders.IClipboardHelper.SendDragFile(string fileName)
        {
            Common.DragDropStep05Ex(fileName);
        }

        void MouseWithoutBorders.IClipboardHelper.SendClipboardData(object data, bool isFilePath)
        {
            _ = Common.CheckClipboardEx(data, isFilePath);
        }
    }
#endif

    internal sealed class IpcHelper
    {
        private const string ChannelName = "MouseWithoutBorders";
        private const string RemoteObjectName = "ClipboardHelper";

#if !MM_HELPER
        private static void CleanupStream(PipeStream s)
        {
            s.Close();
        }

        private static NamedPipeServerStream serverChannel;
        private static IDictionary channelProperties;

        internal static void CreateIpcServer(bool cleanup)
        {
            if (cleanup && serverChannel != null)
            {
                CleanupStream(serverChannel);
                serverChannel = null;
                return;
            }

            if (channelProperties == null)
            {
                channelProperties = new Hashtable
                {
                    ["portName"] = ChannelName,
                    ["exclusiveAddressUse"] = false,
                    ["rejectRemoteRequests"] = true,
                };

                if (Common.RunWithNoAdminRight)
                {
                    channelProperties["authorizedGroup"] = WindowsIdentity.GetCurrent().Name;
                }
                else
                {
                    _ = Common.ImpersonateLoggedOnUserAndDoSomething(() =>
                    {
                        channelProperties["authorizedGroup"] = WindowsIdentity.GetCurrent(true).Name;
                    });
                }
            }

            try
            {
                if (serverChannel != null)
                {
                    CleanupStream(serverChannel);
                    serverChannel = null;
                }

                serverChannel = new NamedPipeServerStream(ChannelName, PipeDirection.InOut);
                Common.IpcChannelCreated = true;
            }
            catch (Exception e)
            {
                Common.IpcChannelCreated = false;
                Common.ShowToolTip("Error setting up clipboard sharing, clipboard sharing will not work!", 5000, ToolTipIcon.Error);
                Common.Log(e);
            }
        }
#else
        internal static IClipboardHelper CreateIpcClient()
        {
            try
            {
                // TODO(@yuyoyuppe): IPC feature
                /*
                IpcChannel channel = new IpcChannel();
                ChannelServices.RegisterChannel(channel, true);
                return (IClipboardHelper)Activator.GetObject(typeof(IClipboardHelper), "ipc://" + ChannelName + "/" + RemoteObjectName);
                */
            }
            catch (Exception e)
            {
                Logger.LogEvent(e.Message, EventLogEntryType.Error);
            }

            return null;
        }
#endif

    }

    internal static class Logger
    {
#if MM_HELPER
        private const string EventSourceName = "MouseWithoutBordersHelper";
#else
        private const string EventSourceName = "MouseWithoutBorders";
#endif

        internal static void LogEvent(string message, EventLogEntryType logType = EventLogEntryType.Information)
        {
            try
            {
                if (!EventLog.SourceExists(EventSourceName))
                {
                    EventLog.CreateEventSource(EventSourceName, "Application");
                }

                EventLog.WriteEntry(EventSourceName, message, logType);
            }
            catch (Exception e)
            {
                Debug.WriteLine(message + ": " + e.Message);
            }
        }
    }
#if MM_HELPER

    internal static class ClipboardMMHelper
    {
        internal static IntPtr NextClipboardViewer = IntPtr.Zero;
        private static FormHelper helperForm;
        private static bool addClipboardFormatListenerResult;

        private static void Log(string log)
        {
            helperForm.SendLog(log);
        }

        private static void Log(Exception e)
        {
            Log($"Trace: {e}");
        }

        internal static void HookClipboard(FormHelper f)
        {
            helperForm = f;

            try
            {
                addClipboardFormatListenerResult = NativeMethods.AddClipboardFormatListener(f.Handle);

                int err = addClipboardFormatListenerResult ? 0 : Marshal.GetLastWin32Error();

                if (err != 0)
                {
                    Log($"Trace: {nameof(NativeMethods.AddClipboardFormatListener)}: GetLastError = {err}");
                }
            }
            catch (EntryPointNotFoundException e)
            {
                Log($"{nameof(NativeMethods.AddClipboardFormatListener)} is unavailable in this version of Windows.");
                Log(e);
            }
            catch (Exception e)
            {
                Log(e);
            }

            // Fallback
            if (!addClipboardFormatListenerResult)
            {
                NextClipboardViewer = NativeMethods.SetClipboardViewer(f.Handle);
                int err = NextClipboardViewer == IntPtr.Zero ? Marshal.GetLastWin32Error() : 0;

                if (err != 0)
                {
                    Log($"Trace: {nameof(NativeMethods.SetClipboardViewer)}: GetLastError = {err}");
                }
            }

            Log($"Trace: Clipboard monitor method {(addClipboardFormatListenerResult ? nameof(NativeMethods.AddClipboardFormatListener) : NextClipboardViewer != IntPtr.Zero ? nameof(NativeMethods.SetClipboardViewer) : "(none)")} is used.");
        }

        internal static void UnhookClipboard()
        {
            if (addClipboardFormatListenerResult)
            {
                addClipboardFormatListenerResult = false;
                _ = NativeMethods.RemoveClipboardFormatListener(helperForm.Handle);
            }
            else
            {
                _ = NativeMethods.ChangeClipboardChain(helperForm.Handle, NextClipboardViewer);
                NextClipboardViewer = IntPtr.Zero;
            }
        }

        private static void ReHookClipboard()
        {
            UnhookClipboard();
            HookClipboard(helperForm);
        }

        internal static bool UpdateNextClipboardViewer(Message m)
        {
            if (m.WParam == NextClipboardViewer)
            {
                NextClipboardViewer = m.LParam;
                return true;
            }

            return false;
        }

        internal static void PassMessageToTheNextViewer(Message m)
        {
            if (NextClipboardViewer != IntPtr.Zero && NextClipboardViewer != helperForm.Handle)
            {
                _ = NativeMethods.SendMessage(NextClipboardViewer, m.Msg, m.WParam, m.LParam);
            }
        }

        public static bool ContainsFileDropList()
        {
            bool rv = false;

            try
            {
                rv = Common.Retry(nameof(Clipb.ContainsFileDropList), () => { return Clipb.ContainsFileDropList(); }, (log) => Log(log));
            }
            catch (ExternalException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ThreadStateException e)
            {
                Log(e);
                ReHookClipboard();
            }

            return rv;
        }

        public static bool ContainsImage()
        {
            bool rv = false;

            try
            {
                rv = Common.Retry(nameof(Clipb.ContainsImage), () => { return Clipb.ContainsImage(); }, (log) => Log(log));
            }
            catch (ExternalException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ThreadStateException e)
            {
                Log(e);
                ReHookClipboard();
            }

            return rv;
        }

        public static bool ContainsText()
        {
            bool rv = false;

            try
            {
                rv = Common.Retry(nameof(Clipb.ContainsText), () => { return Clipb.ContainsText(); }, (log) => Log(log));
            }
            catch (ExternalException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ThreadStateException e)
            {
                Log(e);
                ReHookClipboard();
            }

            return rv;
        }

        public static StringCollection GetFileDropList()
        {
            StringCollection rv = null;

            try
            {
                rv = Common.Retry(nameof(Clipb.GetFileDropList), () => { return Clipb.GetFileDropList(); }, (log) => Log(log));
            }
            catch (ExternalException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ThreadStateException e)
            {
                Log(e);
                ReHookClipboard();
            }

            return rv;
        }

        public static Image GetImage()
        {
            Image rv = null;

            try
            {
                rv = Common.Retry(nameof(Clipb.GetImage), () => { return Clipb.GetImage(); }, (log) => Log(log));
            }
            catch (ExternalException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ThreadStateException e)
            {
                Log(e);
                ReHookClipboard();
            }

            return rv;
        }

        public static string GetText(TextDataFormat format)
        {
            string rv = null;

            try
            {
                rv = Common.Retry(nameof(Clipb.GetText), () => { return Clipb.GetText(format); }, (log) => Log(log));
            }
            catch (ExternalException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ThreadStateException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (System.ComponentModel.InvalidEnumArgumentException e)
            {
                Log(e);
            }

            return rv;
        }

        public static void SetImage(Image image)
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
                    (log) => Log(log));
            }
            catch (ExternalException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ThreadStateException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ArgumentNullException e)
            {
                Log(e);
            }
        }

        public static void SetText(string text)
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
                    (log) => Log(log));
            }
            catch (ExternalException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ThreadStateException e)
            {
                Log(e);
                ReHookClipboard();
            }
            catch (ArgumentNullException e)
            {
                Log(e);
            }
        }
    }

#endif

    internal sealed class SharedConst
    {
        internal const int QUIT_CMD = 0x409;
    }

    internal sealed partial class Common
    {
        internal static bool IpcChannelCreated { get; set; }

        internal static T Retry<T>(string name, Func<T> func, Action<string> log, Action preRetry = null)
        {
            int count = 0;

            do
            {
                try
                {
                    T rv = func();

                    if (count > 0)
                    {
                        log($"Trace: {name} has been successful after {count} retry.");
                    }

                    return rv;
                }
                catch (Exception)
                {
                    count++;

                    preRetry?.Invoke();

                    if (count > 10)
                    {
                        throw;
                    }

                    Application.DoEvents();
                    Thread.Sleep(200);
                }
            }
            while (true);
        }
    }
}
