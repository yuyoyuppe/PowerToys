// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;
using UIA=UIAutomationClient;
using System.Runtime.InteropServices;

namespace UIAVerifications
{
    [Verification(
        "CheckTabbingUia",
        "Verfies tabbing through a window hits all/most important elements",
        Group = "UIA Tree",
        Priority = 1,
        IsIntrusive = true)]
    public class CheckTabbingUIA : IVerificationRoutine
    {
        private const int _sanityCheckLimit = 100;

        private IntPtr _hwnd = IntPtr.Zero;
        private uint _process;

        public List<UIA.IUIAutomationElement> _tabbableElements = new List<UIA.IUIAutomationElement>();

        #region LogMessages

        private static LogMessage AppearsToNotSupportTabbing = new LogMessage("The application appears to not support tabbing.", "AppearsToNotSupportTabbing", 2);
        private static LogMessage TabbingNotCyclic = new LogMessage("The tabbing in this application is not cyclic.  Tabbing forwards or backwards {0} times doesn't come back to the same control", "TabbingNotCyclic", 2);
        private static LogMessage TabbingGotStuck = new LogMessage("The tabbing appears to get stuck at this element", "TabbingGotStuck", 2);
        private static LogMessage TabbingNotSymmetric = new LogMessage("Tabbing backwards and forwards does not yield the same element.", "TabbingNotSymmetric", 3);
        private static LogMessage MissingItemInTabOrder = new LogMessage("This element appears to be tabbable but is not in the tab order", "MissingItemInTabOrder", 1);
        private static LogMessage TabbedToUnknownElement = new LogMessage("Tabbing has gone to an element outside the target hwnd.", "TabbedToUnknownElement", 3);
        private static LogMessage TabOrderInQuestion = new LogMessage("Tab order may be not following standard Windows Guidelines", "TabOrderInQuestion", 2);
        private static LogMessage StartingTab = new LogMessage("The starting tab is {0}", "StartingTab");
        private static LogMessage TabbedBackwardTo = new LogMessage("Tabbed back to {0}", "TabbedBackwardTo");
        private static LogMessage TabbedForwardTo = new LogMessage("Tabbed forward to {0} expecting {1}", "TabbedForwardTo");
        private static LogMessage TabbedToInfo = new LogMessage("{0}", "TabbedToInfo");

        #endregion LogMessages


        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            LogHelper.Logger = logger;
            LogHelper.RootHwnd = hwnd;

            // if we are going to check tabbing make sure we are at the root so we 
            // don't tab outside the window we are looking at
            _hwnd = Win32API.GetAncestor(hwnd, Win32API.GA_ROOT);
             UIA.IUIAutomationElement rootElement = UIAGlobalContext.CacheUIATree(_hwnd);

             Win32API.GetWindowThreadProcessId(_hwnd, out _process);
             if (!Win32API.SetForegroundWindow(_hwnd))
             {
                 rootElement.SetFocus();
             }

             // let the app get focus then empty the queue
             System.Threading.Thread.Sleep(1000);
             if (TabBackward())
             {
                 if (TabForward())
                 {
                     CheckForMissingTabbedElements(rootElement);
                     CheckTabOrder();
                 }
             }

             _tabbableElements.Clear();
        }

        private bool TabBackward()
        {
            UIA.IUIAutomationElement initElement = UIAGlobalContext.GetFocusedElement();

            // tab to the previous control
            HitTabBackwardKey();

            UIA.IUIAutomationElement focusedElement = UIAGlobalContext.GetFocusedElement();
            if (initElement == null || focusedElement == null)
            {
                return false;
            }

            if (UIAGlobalContext.Automation.CompareElements(initElement, focusedElement) != 0)
            {
                if (!TabbedOutsideTheApp())
                {
                    // if we hit the tab key and we are still in the same app then tabbing is not doing anything
                    LogHelper.Log(EventLevel.Warning, AppearsToNotSupportTabbing, GetType(), UIAGlobalContext.LocalRootElement);
                }
                return false;
            }

            _tabbableElements.Add(focusedElement);
            LogHelper.Log(EventLevel.Information, StartingTab, this.GetType(), focusedElement, focusedElement.CachedName);

            bool reachedEnd = false;
            for (int i = 0; i < _sanityCheckLimit; i++)
            {
                HitTabBackwardKey();
                UIA.IUIAutomationElement nextTabElement = UIAGlobalContext.GetFocusedElement();
                if (nextTabElement == null)
                {
                    LogHelper.Log(EventLevel.Error, TabbedToUnknownElement, this.GetType(), (UIA.IUIAutomationElement)null);
                    return false;
                }

                if (UIAGlobalContext.Automation.CompareElements(nextTabElement, focusedElement) != 0)
                {
                    if (!TabbedOutsideTheApp())
                    {
                        // if we hit the tab key and we are still in the same app then tabbing is not doing anything
                        LogHelper.Log(EventLevel.Warning, TabbingGotStuck, GetType(), nextTabElement);
                    }
                    return false;
                }

                focusedElement = nextTabElement;
                LogHelper.Log(EventLevel.Information, TabbedBackwardTo, this.GetType(), focusedElement, focusedElement.CachedName);

                // If we have arrived at the same element then we have come full circle
                if (UIAGlobalContext.Automation.CompareElements(focusedElement, _tabbableElements[0]) != 0)
                {
                    reachedEnd = true;
                    break;
                }

                _tabbableElements.Add(focusedElement);
            }

            if (!reachedEnd)
            {
                LogHelper.Log(EventLevel.Error, TabbingNotCyclic, this.GetType(), focusedElement, _sanityCheckLimit);
            }
            return true;
        }

        private bool TabForward()
        {
            UIA.IUIAutomationElement focusedElement = UIAGlobalContext.GetFocusedElement();
            if (focusedElement == null)
            {
                return false;
            }

            bool reachedEnd = false;
            int tabIndex = _tabbableElements.Count - 1;
            for (int i = 0; i < _sanityCheckLimit; i++)
            {
                HitTabForwardKey();
                UIA.IUIAutomationElement nextTabElement = UIAGlobalContext.GetFocusedElement();
                if (nextTabElement == null)
                {
                    LogHelper.Log(EventLevel.Error, TabbedToUnknownElement, this.GetType(), (UIA.IUIAutomationElement)null);
                    return false;
                }

                if (UIAGlobalContext.Automation.CompareElements(nextTabElement, focusedElement) != 0)
                {
                    if (!TabbedOutsideTheApp())
                    {
                        // if we hit the tab key and we are still in the same app then tabbing is not doing anything
                        LogHelper.Log(EventLevel.Warning, TabbingGotStuck, GetType(), nextTabElement);
                    }
                    return false;
                }

                LogHelper.Log(EventLevel.Information, TabbedForwardTo, this.GetType(), focusedElement, focusedElement.CachedName, _tabbableElements[tabIndex].CachedName);

                if (UIAGlobalContext.Automation.CompareElements(nextTabElement, _tabbableElements[tabIndex]) == 0)
                {
                    UIA.IUIAutomationElement  expected = _tabbableElements[tabIndex];
                    LogHelper.Log(EventLevel.Error, TabbingNotSymmetric, this.GetType(), nextTabElement);
                    break;
                }

                tabIndex--;
                focusedElement = nextTabElement;


                // If we have arrived at the same element then we have come full circle
                if (UIAGlobalContext.Automation.CompareElements(focusedElement, _tabbableElements[0]) != 0)
                {
                    LogHelper.Log(EventLevel.Information, TabbedToInfo, this.GetType(), focusedElement, "Back to the beginning");
                    reachedEnd = true;
                    break;
                }
            }

            if (!reachedEnd)
            {
                LogHelper.Log(EventLevel.Error, TabbingNotCyclic, this.GetType(), focusedElement, _sanityCheckLimit);
            }
            return true;
        }

        private bool CheckForMissingTabbedElements(UIA.IUIAutomationElement rootElement)
        {
            bool tabbedToAChild = false;
            UIA.IUIAutomationElementArray children = rootElement.GetCachedChildren();
            if (children != null)
            {
                for (int i = 0; i < children.Length; i++)
                {
                    if (CheckForMissingTabbedElements(children.GetElement(i)))
                    {
                        tabbedToAChild = true;
                    }
                }
            }

            bool weTabbedToThis = false;
            foreach (UIA.IUIAutomationElement tabbedToItem in _tabbableElements)
            {
                if (UIAGlobalContext.Automation.CompareElements(rootElement, tabbedToItem) != 0)
                {
                    weTabbedToThis = true;
                    break;
                }
            }

            if (!weTabbedToThis && !tabbedToAChild && (ShouldBeTabbable(rootElement)))
            {
                LogHelper.Log(EventLevel.Error, MissingItemInTabOrder, this.GetType(), rootElement);
            }

            return weTabbedToThis || tabbedToAChild;
        }

        private bool ShouldBeTabbable(UIA.IUIAutomationElement element)
        {
            if (element.CachedIsKeyboardFocusable == 0)
            {
                return false;
            }

            if (element.CachedIsEnabled == 0)
            {
                return false;
            }

            int controlType = element.CachedControlType;
            if (controlType == UIA.UIA_ControlTypeIds.UIA_WindowControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_RadioButtonControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_MenuBarControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_MenuItemControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_TitleBarControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_TabItemControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_ScrollBarControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_ThumbControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_PaneControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_SeparatorControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_CustomControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_GroupControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_StatusBarControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_TextControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_ToolBarControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_HeaderControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_HeaderItemControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_ListItemControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_TreeItemControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_SpinnerControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_ToolBarControlTypeId ||
                controlType == UIA.UIA_ControlTypeIds.UIA_ProgressBarControlTypeId)
            {
                return false;
            }

            if (controlType == UIA.UIA_ControlTypeIds.UIA_CheckBoxControlTypeId)
            {
                UIA.IUIAutomationElement parent = element.GetCachedParent();
                if (parent.CachedControlType == UIA.UIA_ControlTypeIds.UIA_ListControlTypeId ||
                    parent.CachedControlType == UIA.UIA_ControlTypeIds.UIA_TreeControlTypeId)
                {
                    return false;
                }
            }

            if (controlType == UIA.UIA_ControlTypeIds.UIA_EditControlTypeId)
            {
                UIA.IUIAutomationElement parent = element.GetCachedParent();
                if (parent.CachedControlType == UIA.UIA_ControlTypeIds.UIA_ListItemControlTypeId)
                {
                    return false;
                }
            }

            return true;
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
            foreach (UIA.IUIAutomationElement ae in _tabbableElements)
            {
                UIAutomationClient.tagRECT tempRect = ae.CachedBoundingRectangle;
                Rectangle rect = new Rectangle(tempRect.left, tempRect.top, tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
                rectList.Add(rect);
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
                    LogHelper.Log(eventLevel, TabOrderInQuestion, this.GetType(), UIAGlobalContext.LocalRootElement);
                }
            }
        }

        private void HitTabForwardKey()
        {
            // true is press key down false is release
            Win32API.SendKeyboardInput(Win32API.VK_TAB, true);
            Win32API.SendKeyboardInput(Win32API.VK_TAB, false);
            System.Threading.Thread.Sleep(200);
        }

        private void HitTabBackwardKey()
        {
            // true is press key down false is release
            Win32API.SendKeyboardInput(Win32API.VK_RSHIFT, true);
            Win32API.SendKeyboardInput(Win32API.VK_TAB, true);
            Win32API.SendKeyboardInput(Win32API.VK_TAB, false);
            Win32API.SendKeyboardInput(Win32API.VK_RSHIFT, false);
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
                UIA.IUIAutomationElement focusedElement = UIAGlobalContext.ElementFromHandle(guiThreadInfo.hwndFocus);

                // Trying to report this error when focus go to AccChecker itself causes problems.
                // This hapens when the user clicks the cancel button.
                if (Win32API.GetCurrentProcessId() == focusedProcess)
                {
                    focusedElement = UIAGlobalContext.ElementFromHandle(_hwnd);
                }

                LogHelper.Log(EventLevel.Error, TabbedToUnknownElement, this.GetType(), focusedElement);
                return true;
            }

            return false;
        }

    }
}
