// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using Accessibility;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;
using System.Drawing;
using System.Reflection;

namespace VerificationRoutines
{

    [Verification(
        "CheckTabbing",
        "Checks the tabbing of a window",
        Group = ClassificationsGroup.Navigation,
        Priority = 1,
        IsIntrusive = true)]
    public class CheckTabbing : IVerificationRoutine
    {
        private const int _sanityCheckLimit = 100;

        private IntPtr _hwnd = IntPtr.Zero;
        private uint _process;

        public List<MsaaElement> _tabbableElements = new List<MsaaElement>();
        public List<MsaaElement> _skippedTabbedElements = new List<MsaaElement>();

        #region LogMessages


        private static LogMessage AppearsToNotSupportTabbing = new LogMessage("The application appears to not support tabbing.", "AppearsToNotSupportTabbing", 2);
        private static LogMessage TabbingNotCyclic = new LogMessage("The tabbing in this application is not cyclic.  Tabbing forwards or backwards {0} times doesn't come back to the same control", "TabbingNotCyclic", 2);
        private static LogMessage TabbingNotSymmetric = new LogMessage("Tabbing backwards and forwards does not yield the same element.", "TabbingNotSymmetric", 2);
        private static LogMessage MissingItemInTabOrder = new LogMessage("This element appears to be tabbable but is not in the tab order", "MissingItemInTabOrder", 1);
        private static LogMessage TabbedToUnknownElement = new LogMessage("Tabbing has gone to an element outside the target hwnd.", "TabbedToUnknownElement", 3);
        private static LogMessage TabOrderInQuestion = new LogMessage("Tab order may be not following standard Windows Guidelines", "TabOrderInQuestion", 2);
        private static LogMessage StartingTab = new LogMessage("The starting tab is {0}", "StartingTab");
        private static LogMessage TabbedBackwardTo = new LogMessage("Tabbed back to {0}", "TabbedBackwardTo");
        private static LogMessage TabbedForwardTo = new LogMessage("Tabbed forward to {0} expecting {1}", "TabbedForwardTo");
        private static LogMessage TabbedToInfo = new LogMessage("{0}", "TabbedToInfo");

        #endregion LogMessages

        /// -------------------------------------------------------------------
        /// <summary>Entry point to this test class</summary>
        /// -------------------------------------------------------------------
        public void Execute(IntPtr hwnd, ILogger logger, bool allowUI, GraphicsHelper graphicsHelper)
        {
            TestLogger.Logger = logger;
            TestLogger.RootHwnd = hwnd;

            // if we are going to check tabbing make sure we are at the root so we 
            // don't tab outside the window we are looking at
            _hwnd = Win32API.GetAncestor(hwnd, Win32API.GA_ROOT);

            MsaaElement.CacheMsaaTree(_hwnd);

            Win32API.GetWindowThreadProcessId(_hwnd, out _process);

            // get a chance for the event thread to start listening 
            System.Threading.Thread.Sleep(2000);
            if (!Win32API.SetForegroundWindow(_hwnd))
            {
                Accessible accessible = null;
                int ret = Accessible.FromWindow(_hwnd, out accessible);
                if (ret != Win32API.S_OK)
                {
                    throw new ArgumentException("AccessibleObjectFromWindow failed for toplevel hwnd");
                }
                try 
                {
                    accessible.Select(Win32API.SELFLAG_TAKEFOCUS);
                }
                catch (Exception exception)
                {
                    if (!Accessible.IsExpectedException(exception))
                    {
                        throw;
                    }
                }
            }

            // let the app get focus then empty the queue
            System.Threading.Thread.Sleep(1000);

            if (TabBackward())
            {
                if (TabForward())
                {
                    CheckForMissingTabbedElements(MsaaElement.RootElement);
                    CheckTabOrder();
                }
            }

            _tabbableElements.Clear();
        }

        /// -------------------------------------------------------------------
        /// <summary>Tab backward thru the UI build a list of the tabbed to elements.  
        /// Make sure we get back to the same place we started.  The list that is built from
        /// this is used in the other methods</summary>
        /// -------------------------------------------------------------------
        private bool TabBackward()
        {
            int sanityCheck = 0;

            MsaaElement focusedElement = null;
            MsaaElement lastFocusedElement = focusedElement;
            while (sanityCheck < _sanityCheckLimit)
            {
                // tab to the next control
                HitTabBackwardKey();

                if (!GetFocusedElement(out focusedElement))
                {
                    return false;
                }

                // If it is a role client with no name just pretend it did not happen
                if ((focusedElement.GetRole() == AccRole.Client || focusedElement.GetRole() == AccRole.Window) && focusedElement.GetName() == "")
                {
                    sanityCheck++;
                    continue;
                }
                
                if (_tabbableElements.Count == 0)
                {
                    // Put the first element in the list.  This is the one we want to circle around to.
                    _tabbableElements.Add(focusedElement);
                    lastFocusedElement = focusedElement;
                    TestLogger.Log(EventLevel.Information, StartingTab, this.GetType(), focusedElement, focusedElement.GetName());
                    continue;
                }

                if (!TabbedOutsideTheApp())
                {
                    if (focusedElement.GetName() == lastFocusedElement.GetName() &&
                        focusedElement.GetRole() == lastFocusedElement.GetRole() &&
                        focusedElement.GetBoundingBox().Equals(lastFocusedElement.GetBoundingBox()))
                    {
                        // if we hit the tab key and we are still in the same app then tabbing is not doing anything
                        TestLogger.Log(EventLevel.Warning, AppearsToNotSupportTabbing, this.GetType(), focusedElement);
                        return false;
                    }
                }
                
                lastFocusedElement = focusedElement;
                if (focusedElement == null)
                {
                    // need to cast the null to MsaaElement to disambiguate the overloads for log.
                    TestLogger.Log(EventLevel.Error, TabbedToUnknownElement, this.GetType(), (MsaaElement)null);
                    return false;
                }

                TestLogger.Log(EventLevel.Information, TabbedBackwardTo, this.GetType(), focusedElement, focusedElement.GetName());

                // If we have arrived at the same element then we have come full circle
                if (focusedElement.Compare(_tabbableElements[0]))
                {
                    break;
                }

                _tabbableElements.Add(focusedElement);
                sanityCheck++;
            }

            //We never got back to where we started 
            if (sanityCheck >= _sanityCheckLimit && _tabbableElements.Count > 0)
            {
                TestLogger.Log(EventLevel.Error, TabbingNotCyclic, this.GetType(), focusedElement, _sanityCheckLimit);
            }

            return true;
        }

        /// -------------------------------------------------------------------
        /// <summary>Tab forward thru the UI checking the list we got from the backward tabbing  
        /// make sure the same elements are tabbed to.  Make sure we get back to the same 
        /// place we started</summary>
        /// -------------------------------------------------------------------
        private bool TabForward()
        {
            MsaaElement focusedElement = null;

            int sanityCheck = 0;
            int tabIndex = _tabbableElements.Count - 1;

            MsaaElement lastFocusedElement = null;
            while (sanityCheck < _sanityCheckLimit)
            {
                // tab to the next control
                HitTabForwardKey();

                if (!GetFocusedElement(out focusedElement))
                {
                    return false;
                }

                // If it is a role client with no name just pretend it did not happen
                if ((focusedElement.GetRole() == AccRole.Client || focusedElement.GetRole() == AccRole.Window) && focusedElement.GetName() == "")
                {
                    sanityCheck++;
                    continue;
                }

                if (!TabbedOutsideTheApp() && lastFocusedElement != null)
                {
                    if (focusedElement.GetName() == lastFocusedElement.GetName() &&
                        focusedElement.GetRole() == lastFocusedElement.GetRole() &&
                        focusedElement.GetBoundingBox().Equals(lastFocusedElement.GetBoundingBox()))
                    {
                        // if we hit the tab key and we are still in the same app then tabbing is not doing anything
                        TestLogger.Log(EventLevel.Error, AppearsToNotSupportTabbing, this.GetType(), focusedElement);
                        return false;
                    }
                }
                
                lastFocusedElement = focusedElement;
                if (focusedElement == null)
                {
                    TestLogger.Log(EventLevel.Error, TabbedToUnknownElement, this.GetType(), focusedElement);
                    return false;
                }

                TestLogger.Log(EventLevel.Information, TabbedForwardTo, this.GetType(), focusedElement, focusedElement.GetName(), _tabbableElements[tabIndex].GetName());

                if (!focusedElement.Compare(_tabbableElements[tabIndex--]))
                {
                    MsaaElement expected = _tabbableElements[tabIndex + 1];
                    TestLogger.Log(EventLevel.Error, TabbingNotSymmetric, this.GetType(), focusedElement);
                    break;
                }

                // If we have arrived at the same element then we have come full circle
                if (focusedElement.Compare(_tabbableElements[0]))
                {
                    TestLogger.Log(EventLevel.Information, TabbedToInfo, this.GetType(), focusedElement, "Back to the beginning");
                    break;
                }

                sanityCheck++;
            }

            //We never got back to where we started 
            if (sanityCheck >= _sanityCheckLimit)
            {
                TestLogger.Log(EventLevel.Error, TabbingNotCyclic, this.GetType(), focusedElement, _sanityCheckLimit);
            }

            return true;
        }

        /// -------------------------------------------------------------------
        /// <summary>Examine the tree and make sure there are not elements that we should
        /// have tabbed to but did not.  This uses the list of elements from TabBackward</summary>
        /// -------------------------------------------------------------------
        private bool CheckForMissingTabbedElements(MsaaElement element)
        {
            bool tabbedToAChild = false;
            List<MsaaElement> children = element.Children;
            if (children != null)
            {
                foreach (MsaaElement child in children)
                {
                    List<AccState> states = new List<AccState>(element.GetStates());

                    // if a parent  is invisable or disabled no need to look at the children 
                    // they may appear visable for example in a tabbed dialog.
                    if (!states.Contains(AccState.Invisible) && !states.Contains(AccState.Unavailable))
                    {
                        if (CheckForMissingTabbedElements(child))
                        {
                            tabbedToAChild = true;
                        }
                    }
                }
            }

            bool weTabbedToThis = false;
            foreach (MsaaElement tabbedToItem in _tabbableElements)
            {
                if (tabbedToItem.Compare(element))
                {
                    weTabbedToThis = true;
                    break;
                }
            }


            if (!weTabbedToThis && !tabbedToAChild && ShouldBeTabbable(element))
            {
                TestLogger.Log(EventLevel.Error, MissingItemInTabOrder, this.GetType(), element);
            }

            return weTabbedToThis || tabbedToAChild;
        }


        /// -------------------------------------------------------------------
        /// <summary>Check that the tab order follows some pattern of top right 
        /// to bottom left movement.  Since we cannot really determine if things 
        /// are correct, or in error, the routine will only flag issues if there 
        /// is a percentage threshold of controls that do not follow this 
        /// standard windows design of top-left to bottom-right order.  This relys 
        /// the list from TabBackward.</summary>
        /// -------------------------------------------------------------------
        private void CheckTabOrder()
        {
            List<Rectangle> rectList = new List<Rectangle>();

            int numControls = _tabbableElements.Count;

            // Place the rects in list for easy manipulation below.
            foreach (MsaaElement oi in _tabbableElements)
            {
                rectList.Add(oi.GetBoundingBox());
            }

            if (numControls > 2)
            {
                // Since we cycle the tab list, we will always find at minimum one error when we cycle back
                // to the 1st control, which is really a valid situation. Set this to -1 so we allow one 
                // 'error' and take this into account.
                int badCount = -1;

                // The tab list is built starting with the current focused control when 
                // the 1st tab verification is ran, so the 1st control in the list is not guaranteed to 
                // be the topmost, leftmost control of the client.  Because of this, we need to cycle the entire list.  
                // The tab list is also in reverse order so need to walk the list backwards.
                // Walk the list backwards and verify that the current control's location is to the right or 
                // below the previous control in the tab order.
                for (int curTabPos = numControls - 1, prevTabPos = 0; curTabPos > -1; curTabPos--)
                {
                    if ((rectList[curTabPos].Left < rectList[prevTabPos].Left) && (rectList[curTabPos].Top < rectList[prevTabPos].Top))
                    {
                        badCount++;
                    }
                    prevTabPos = curTabPos;
                }

                // There are two different ways to determine the threshold for reporting an error:
                // Error threshold = 6 bad, or bad > 30% of controls which ever is greater
                // Warning threshold = 3 bad, or bad > 15% of controls which ever is greater
                const double percentBadError = .30;
                const double percentBadWarning = .15;
                const int numControlsBadError = 6;
                const int numControlsBadWarning = 3;

                int threadsholdError = (int)Math.Max(percentBadError * numControls, numControlsBadError);
                int threadsholdWarning = (int)Math.Max(percentBadWarning * numControls, numControlsBadWarning);

                // Set this to an initial state and use as a flag also.  If this is set to anything 
                // else besides Information, then we know we need to report a problem.
                EventLevel eventLevel = EventLevel.Information;

                if (badCount > threadsholdError)
                {
                    eventLevel = EventLevel.Error;
                }
                else if (badCount > threadsholdWarning)
                {
                    eventLevel = EventLevel.Warning;
                }

                // If we have hit something that needs reported, then log it
                if (eventLevel != EventLevel.Information)
                {
                    TestLogger.Log(eventLevel, TabOrderInQuestion, this.GetType(), MsaaElement.RootElement);
                }
            }
        }

        private bool ShouldBeTabbable(MsaaElement element)
        {
            bool tabbable = false;

            List<AccState> states = new List<AccState>(element.GetStates());
            if (states.Contains(AccState.Focusable))
            {
                if (!states.Contains(AccState.Invisible) &&
                    !states.Contains(AccState.Unavailable) &&
                    !states.Contains(AccState.ReadOnly))
                {
                    AccRole role = element.GetRole();
                    if (role != AccRole.Window &&
                        role != AccRole.RadioButton &&
                        role != AccRole.Menubar &&
                        role != AccRole.MenuItem &&
                        role != AccRole.Titlebar &&
                        role != AccRole.PropertyPage &&
                        role != AccRole.PageTab &&
                        role != AccRole.Scrollbar &&
                        role != AccRole.Grip &&
                        role != AccRole.Client &&
                        role != AccRole.Pane &&
                        role != AccRole.Separator &&
                        role != AccRole.Whitespace &&
                        role != AccRole.Invalid &&
                        role != AccRole.Grouping &&
                        role != AccRole.Statusbar &&
                        role != AccRole.StaticText &&
                        role != AccRole.ListItem &&
                        role != AccRole.TreeItem &&
                        role != AccRole.Dialog &&
                        role != AccRole.SpinButton &&
                        role != AccRole.Toolbar &&
                        role != AccRole.ProgressBar)
                    {
                        tabbable = true;
                    }

                    // check boxes in trees and list are not tabbable
                    if (role == AccRole.CheckButton)
                    {
                        MsaaElement parent = element.Parent;
                        if (parent.GetRole() == AccRole.Tree ||
                            parent.GetRole() == AccRole.List)
                        {
                            tabbable = false;
                        }
                    }

                    if (role == AccRole.Text)
                    {
                        MsaaElement parent = element.Parent;
                        if (parent.GetRole() == AccRole.ListItem)
                        {
                            tabbable = false;
                        }
                    }

                    
                }
            }

            return tabbable;
        }


        private bool GetFocusedElement(out MsaaElement focusElement)
        {
            bool ret = false;
            focusElement = null;
            
            Win32API.GUITHREADINFO guiThreadInfo = new Win32API.GUITHREADINFO();
            guiThreadInfo.cbSize = Marshal.SizeOf(guiThreadInfo.GetType());
            Win32API.GetGUIThreadInfo(0, ref guiThreadInfo);
            
            IntPtr hwnd = guiThreadInfo.hwndFocus;
            if (hwnd != IntPtr.Zero)
            {
                Accessible accessible = null;
                int status = Accessible.FromWindow(hwnd, out accessible);
                if (status == Win32API.S_OK)
                {                    
                    focusElement = MsaaElement.FindMsaaElement(accessible);
                    if (focusElement != null)
                    {
                        ret = true;
                        List<MsaaElement> children = focusElement.Children;
                        if (children != null)
                        {
                            foreach (MsaaElement child in children)
                            {
                                MsaaElement.UpdateMsaaCachedDynamicProperties(child);
                                List<AccState> states = new List<AccState>(child.GetStates());
                                if (states.Contains(AccState.Focused))
                                {
                                    focusElement = child;
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            return ret;
        }

        private void HitTabForwardKey()
        {
            // true is press key down false is release
            VerificationHelper.SendKeyboardInput(Win32API.VK_TAB, true);
            VerificationHelper.SendKeyboardInput(Win32API.VK_TAB, false);
            System.Threading.Thread.Sleep(200);
        }

        private void HitTabBackwardKey()
        {
            // true is press key down false is release
            VerificationHelper.SendKeyboardInput(Win32API.VK_RSHIFT, true);
            VerificationHelper.SendKeyboardInput(Win32API.VK_TAB, true);
            VerificationHelper.SendKeyboardInput(Win32API.VK_TAB, false);
            VerificationHelper.SendKeyboardInput(Win32API.VK_RSHIFT, false);
            System.Threading.Thread.Sleep(200);
        }

        private bool TabbedOutsideTheApp()
        {
            // make sure our app still has focus
            uint focusedProcess;

            Win32API.GUITHREADINFO guiThreadInfo = new Win32API.GUITHREADINFO();
            guiThreadInfo.cbSize = Marshal.SizeOf(guiThreadInfo.GetType());
            Win32API.GetGUIThreadInfo(0, ref guiThreadInfo);

            Win32API.GetWindowThreadProcessId(guiThreadInfo.hwndFocus, out focusedProcess);
            if (focusedProcess != _process)
            {
                Accessible focusedAccessible = null;

                Accessible.FromWindow(guiThreadInfo.hwndFocus, out focusedAccessible);

                // Trying to report this error when focus go to AccChecker itself causes problems.
                // This hapens when the user clicks the cancel button.
                if (Win32API.GetCurrentProcessId() == focusedProcess)
                {
                    Accessible.FromWindow(_hwnd, out focusedAccessible);
                }

                TestLogger.Log(EventLevel.Error, TabbedToUnknownElement, this.GetType(), focusedAccessible);
                return true;
            }

            return false;
        }
    }
}
