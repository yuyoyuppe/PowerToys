// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using Accessibility;
using AccCheck.Logging;
using AccCheck;
using System.Windows.Forms;
using System.Reflection;
using VerificationRoutines;

namespace VerificationRoutines
{

    // indicates what an event represents
    public enum AccEventIdentifier { SystemSound, SystemAlert, SystemForeground, SystemMenuStart, 
        SystemMenuEnd, SystemMenuPopupStart, SystemMenuPopupEnd, SystemCaptureStart, SystemCaptureEnd, 
        SystemMovesizeStart, SystemMovesizeEnd, SystemContexthelpStart, SystemContexthelpEnd, 
        SystemDragdropStart, SystemDragdropEnd, SystemDialogStart, SystemDialogEnd, SystemScrollingStart, 
        SystemScrollingEnd, SystemSwitchStart, SystemSwitchEnd, SystemMinimizeStart, SystemMinimizeEnd, 
        SystemDesktopSwitch, ConsoleCaret, ConsoleUpdateRegion, ConsoleUpdateSimple, ConsoleUpdateScroll, 
        ConsoleLayout, ConsoleStartApplication, ConsoleEndApplication, ObjectCreate, ObjectDestroy, 
        ObjectShow, ObjectHide, ObjectReorder, ObjectFocus, ObjectSelection, ObjectSelectionAdd, 
        ObjectSelectionRemove, ObjectSelectionWithin, ObjectStateChange, ObjectLocationChange, 
        ObjectNameChange, ObjectDescriptionChange, ObjectValueChange, ObjectParentChange, ObjectHelpChange, 
        ObjectDefactionChange, ObjectAcceleratorChange, ObjectInvoked, ObjectTextselectionChanged, 
        ObjectContentScrolled };

    // represents an accessible event
    public struct AccEvent
    {
        public AccEventIdentifier evnt;
        public MsaaElement msaaElement;
        public Accessible accessible;
        public int dtTimestamp;
    }

    public class Events : IDisposable
    {
        #region Win32 Conversion
        private static Hashtable EventMap = new Hashtable();
        private static void InitializeEventMap()
        {
            if (EventMap.Count == 0)
            {
                EventMap.Add(Win32API.EVENT_SYSTEM_SOUND, AccEventIdentifier.SystemSound);
                EventMap.Add(Win32API.EVENT_SYSTEM_ALERT, AccEventIdentifier.SystemAlert);
                EventMap.Add(Win32API.EVENT_SYSTEM_FOREGROUND, AccEventIdentifier.SystemForeground);
                EventMap.Add(Win32API.EVENT_SYSTEM_MENUSTART, AccEventIdentifier.SystemMenuStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_MENUEND, AccEventIdentifier.SystemMenuEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_MENUPOPUPSTART, AccEventIdentifier.SystemMenuPopupStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_MENUPOPUPEND, AccEventIdentifier.SystemMenuPopupEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_CAPTURESTART, AccEventIdentifier.SystemCaptureStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_CAPTUREEND, AccEventIdentifier.SystemCaptureEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_MOVESIZESTART, AccEventIdentifier.SystemMovesizeStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_MOVESIZEEND, AccEventIdentifier.SystemMovesizeEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_CONTEXTHELPSTART, AccEventIdentifier.SystemContexthelpStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_CONTEXTHELPEND, AccEventIdentifier.SystemContexthelpEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_DRAGDROPSTART, AccEventIdentifier.SystemDragdropStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_DRAGDROPEND, AccEventIdentifier.SystemDragdropEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_DIALOGSTART, AccEventIdentifier.SystemDialogStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_DIALOGEND, AccEventIdentifier.SystemDialogEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_SCROLLINGSTART, AccEventIdentifier.SystemScrollingStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_SCROLLINGEND, AccEventIdentifier.SystemScrollingEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_SWITCHSTART, AccEventIdentifier.SystemSwitchStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_SWITCHEND, AccEventIdentifier.SystemSwitchEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_MINIMIZESTART, AccEventIdentifier.SystemMinimizeStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_MINIMIZEEND, AccEventIdentifier.SystemMinimizeEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_DESKTOPSWITCH, AccEventIdentifier.SystemDesktopSwitch);
                EventMap.Add(Win32API.EVENT_CONSOLE_CARET, AccEventIdentifier.ConsoleCaret);
                EventMap.Add(Win32API.EVENT_CONSOLE_UPDATE_REGION, AccEventIdentifier.ConsoleUpdateRegion);
                EventMap.Add(Win32API.EVENT_CONSOLE_UPDATE_SIMPLE, AccEventIdentifier.ConsoleUpdateSimple);
                EventMap.Add(Win32API.EVENT_CONSOLE_UPDATE_SCROLL, AccEventIdentifier.ConsoleUpdateScroll);
                EventMap.Add(Win32API.EVENT_CONSOLE_LAYOUT, AccEventIdentifier.ConsoleLayout);
                EventMap.Add(Win32API.EVENT_CONSOLE_START_APPLICATION, AccEventIdentifier.ConsoleStartApplication);
                EventMap.Add(Win32API.EVENT_CONSOLE_END_APPLICATION, AccEventIdentifier.ConsoleEndApplication);
                EventMap.Add(Win32API.EVENT_OBJECT_CREATE, AccEventIdentifier.ObjectCreate);
                EventMap.Add(Win32API.EVENT_OBJECT_DESTROY, AccEventIdentifier.ObjectDestroy);
                EventMap.Add(Win32API.EVENT_OBJECT_SHOW, AccEventIdentifier.ObjectShow);
                EventMap.Add(Win32API.EVENT_OBJECT_HIDE, AccEventIdentifier.ObjectHide);
                EventMap.Add(Win32API.EVENT_OBJECT_REORDER, AccEventIdentifier.ObjectReorder);
                EventMap.Add(Win32API.EVENT_OBJECT_FOCUS, AccEventIdentifier.ObjectFocus);
                EventMap.Add(Win32API.EVENT_OBJECT_SELECTION, AccEventIdentifier.ObjectSelection);
                EventMap.Add(Win32API.EVENT_OBJECT_SELECTIONADD, AccEventIdentifier.ObjectSelectionAdd);
                EventMap.Add(Win32API.EVENT_OBJECT_SELECTIONREMOVE, AccEventIdentifier.ObjectSelectionRemove);
                EventMap.Add(Win32API.EVENT_OBJECT_SELECTIONWITHIN, AccEventIdentifier.ObjectSelectionWithin);
                EventMap.Add(Win32API.EVENT_OBJECT_STATECHANGE, AccEventIdentifier.ObjectStateChange);
                EventMap.Add(Win32API.EVENT_OBJECT_LOCATIONCHANGE, AccEventIdentifier.ObjectLocationChange);
                EventMap.Add(Win32API.EVENT_OBJECT_NAMECHANGE, AccEventIdentifier.ObjectNameChange);
                EventMap.Add(Win32API.EVENT_OBJECT_DESCRIPTIONCHANGE, AccEventIdentifier.ObjectDescriptionChange);
                EventMap.Add(Win32API.EVENT_OBJECT_VALUECHANGE, AccEventIdentifier.ObjectValueChange);
                EventMap.Add(Win32API.EVENT_OBJECT_PARENTCHANGE, AccEventIdentifier.ObjectParentChange);
                EventMap.Add(Win32API.EVENT_OBJECT_HELPCHANGE, AccEventIdentifier.ObjectHelpChange);
                EventMap.Add(Win32API.EVENT_OBJECT_DEFACTIONCHANGE, AccEventIdentifier.ObjectDefactionChange);
                EventMap.Add(Win32API.EVENT_OBJECT_ACCELERATORCHANGE, AccEventIdentifier.ObjectAcceleratorChange);
                EventMap.Add(Win32API.EVENT_OBJECT_INVOKED, AccEventIdentifier.ObjectInvoked);
                EventMap.Add(Win32API.EVENT_OBJECT_TEXTSELECTIONCHANGED, AccEventIdentifier.ObjectTextselectionChanged);
                EventMap.Add(Win32API.EVENT_OBJECT_CONTENTSCROLLED, AccEventIdentifier.ObjectContentScrolled);
            }
        }
        #endregion

        #region Private member variables
        // wineventhook identifier
        private List<IntPtr> _hooks = new List<IntPtr>();
        // always lock this list before accessing it
        private List<AccEvent> _events = new List<AccEvent>();
        
        // event-thread callbacks that we need to hold on to
        private Win32API.WinEventProcDef _wineventproc;
        private GCHandle _wineventprocHandle;
        private Thread _eventThread;
        private IntPtr _hwnd;
        private List<int> _eventsToListenFor = new List<int>();
        
        #endregion

        static Events()
        {
            InitializeEventMap();
        }
        
        public Events(IntPtr hwnd, params int[] events)
        {
            _hwnd = hwnd;
            if (events != null)
            {
                _eventsToListenFor.AddRange(events);
            }
            SetupThread();
        }
        #region Private methods

        /*
         * This method runs on a separate thread, and sets a window event hook. 
         * It populates the _events list when new events arrive.
         */
        private void RegisterForAndPumpMessages(object hwndObject)
        {
            uint process;

            try
            {
                _wineventproc = new Win32API.WinEventProcDef(WinEventProc);
                _wineventprocHandle = GCHandle.Alloc(_wineventproc);

                IntPtr hwnd = (IntPtr)hwndObject;

                foreach (int eventId in _eventsToListenFor)
                {
                    Win32API.GetWindowThreadProcessId(hwnd, out process);
                    IntPtr hook = Win32API.SetWinEventHook(eventId, eventId, 0, _wineventproc, process, 0, Win32API.WINEVENT_OUTOFCONTEXT | Win32API.WINEVENT_SKIPOWNPROCESS);
                    _hooks.Add(hook);
                }
                
                Win32API.MSG msg = new Win32API.MSG();

                // The timer will then post a message (WM_TIMER) that we can decide to ignore or 
                // process. We wait for any external messages to kill the thread.
                Win32API.SetTimer(IntPtr.Zero, UIntPtr.Zero, 200, null);

                while (Thread.CurrentThread.ThreadState != System.Threading.ThreadState.Aborted &&
                    Win32API.GetMessage(ref msg, IntPtr.Zero, 0, 0) != 0)
                {
                    Win32API.TranslateMessage(ref msg);
                    Win32API.DispatchMessage(ref msg);
                }
            }
            catch (ThreadAbortException)
            {
                foreach (IntPtr hook in _hooks)
                {
                    Win32API.UnhookWinEvent(hook);
                }
                _hooks.Clear();
                _eventThread = null;
            }
        }

        /*
         * This is the callback we get from our event hook.
         */
        private unsafe void WinEventProc(IntPtr winEventHook, int eventId, IntPtr hwnd, int idObject, int idChild, int eventThread, int eventTime)
        {
            Accessible acc = null;
            if (Accessible.FromEvent(hwnd, idObject, idChild, out acc) == Win32API.S_OK)
            {
                AccEvent e = new AccEvent();
                e.accessible = acc;
                e.dtTimestamp = eventTime;
                if (EventMap.Contains(eventId))
                {
                    e.evnt = (AccEventIdentifier)EventMap[eventId];
                }
                lock (((ICollection)_events).SyncRoot)
                {
                    _events.Add(e);
                }
            }
        }

        #endregion

        ~Events()
        {
            foreach (IntPtr hook in _hooks)
            {
                Win32API.UnhookWinEvent(hook);
            }
            _hooks.Clear();
            _wineventproc = null;
        }

        /*
         * Creates the thread that listens for winevents.
         */
        private void SetupThread()
        {
            if (_eventThread == null)
            {
                ParameterizedThreadStart pts = new ParameterizedThreadStart(RegisterForAndPumpMessages);
                _eventThread = new Thread(pts);
                _eventThread.Start(_hwnd);
            }
        }

        public AccEvent GetNextEvent()
        {
            if (_events.Count == 0)
            {
                throw new ElementNotFoundException("No elements left");
            }
            
            AccEvent e;
            lock (((ICollection)_events).SyncRoot)
            {
                e = _events[0];
                _events.RemoveAt(0);
            }

            try
            {
                // this is fairly expensive so only do it when the event is asked for and outside the lock
                // so we don't hold up processing of new events.
                e.msaaElement = MsaaElement.FindMsaaElement(e.accessible);
            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception)) 
                {
                    throw; 
                }

                e.msaaElement = null;
            }

            return e;
        }

        public AccEvent GetNextEventType(AccEventIdentifier evnt)
        {
            AccEvent ev = new AccEvent();
            bool eventFound = false;
            lock (((ICollection)_events).SyncRoot)
            {
                foreach (AccEvent e in _events)
                {
                    if (e.evnt == evnt)
                    {
                        _events.Remove(e);
                        ev = e;
                        eventFound = true;
                        break;
                    }
                }
            }

            if (!eventFound)
            {
                throw new ElementNotFoundException("By event type " + Enum.GetName(typeof(AccEventIdentifier), evnt));
            }

            try
            {
                // this is fairly expensive so only do it when the event is asked for and outside the lock
                // so we don't hold up processing of new events.
                ev.msaaElement = MsaaElement.FindMsaaElement(ev.accessible);
            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception)) 
                {
                    throw; 
                }

                ev.msaaElement = null;
            }
            
            return ev;
        }

        public void Dispose()
        {
            if (_eventThread != null)
            {
                foreach (IntPtr hook in _hooks)
                {
                    Win32API.UnhookWinEvent(hook);
                }
                _hooks.Clear();
                _events.Clear();
                _eventsToListenFor.Clear();
                _wineventprocHandle.Free();
                _eventThread.Abort();
                _eventThread = null;
            }
        }
    }

    public class ElementNotFoundException : Exception
    {
        private string m_Element = "";

        override public string Message
        {
            get
            {
                return "Element not found: " + m_Element + ".";
            }
        }

        public ElementNotFoundException(string sElementName)
        {
            m_Element = sElementName;
        }

    } // ElementNotFoundException

}
