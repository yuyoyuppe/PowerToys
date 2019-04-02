// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;
using UIA=UIAutomationClient;

namespace UIAVerifications
{
    [Verification(
        "CheckOrphanChildrenUia",
        "Hit test around the application in an attempt to find elements not in the proper tree",
        Group = "UIA Tree",
        Priority = 1)]
    public class CheckOrphanedChildrenUIA : IVerificationRoutine
    {

        #region LogMessages

        static LogMessage UiaElementIsOrphaned = new LogMessage("Element is an orphaned AutomationElement in the tree", "UiaElementIsOrphaned", 1);
        static LogMessage UiaElementDoesNotParentToRoot = new LogMessage("Element's parent chain does not go to the Root Element", "UiaElementDoesNotParentToRoot", 1);
        static LogMessage ElementFromPointReturnedNull = new LogMessage("IUIAutomation.ElementFromPoint({0}, {1}) returned null", "ElementFromPointReturnedNull", 2);

        #endregion LogMessages        
       
        public List<UIA.IUIAutomationElement> _orphanedElements = new List<UIA.IUIAutomationElement>();
        public List<UIA.IUIAutomationElement> _unparentedElements = new List<UIA.IUIAutomationElement>();

        public class UiaElementMatcher
        {
            internal UIA.IUIAutomationElement _element;
            public UiaElementMatcher(UIA.IUIAutomationElement el)
            {
                _element = el;
            }

            public bool Matches(UIA.IUIAutomationElement el)
            {
                return UIAGlobalContext.Automation.CompareElements(_element, el) != 0;
            }

            public bool IsContainedIn(List<UIA.IUIAutomationElement> list)
            {
                UIA.IUIAutomationElement found = list.Find(Matches);
                return found != null;
            }
        }

        ///  ------------------------------------------------------------------
        /// <summary>Search within the cordinates of the element to see if 
        /// the AutomationElements's found from ElementFromPoint are within 
        /// the tree of the Cached Elements.  If we can't find the Element within 
        /// the cached tree, then check its parent chain to see if it thinks it
        /// should be a child of this Automation Element.</summary>
        ///  ------------------------------------------------------------------
        public void ElementFromPoint_OrphanChildren(UIA.IUIAutomationElement automationElement)
        {
            _orphanedElements.Clear();
            _unparentedElements.Clear();

            const int STEP_X = 75;  // Controls are usually longer than higher so optimize here.
            const int STEP_Y = 25;

            // Check to see if the element is onscreen.  If it is not onscreen, then don't test
            if (automationElement.CurrentIsOffscreen != 0)
            {
                return;
            }

            // Find the HWND for the element, and get the Window Rect for the Hwnd
            IntPtr hwnd = automationElement.CurrentNativeWindowHandle;

            if (hwnd == IntPtr.Zero)
            {
                return;
            }

            Win32API.RECT win32Rect = Win32API.RECT.Empty;
            Win32API.GetWindowRect(hwnd, ref win32Rect);

            for (int x = win32Rect.left + 1; x < win32Rect.right; x += STEP_X)
            {
                for (int y = win32Rect.top + 1; y < win32Rect.bottom; y += STEP_Y)
                {
                    UIA.IUIAutomationElement fromPoint;

                    // The non-client area bits are not important to check.
                    try
                    {
                        fromPoint = UIAGlobalContext.ElementFromPoint(x, y);
                    }
                    catch (Exception exception)
                    {
                        if (VerificationHelper.IsCriticalException(exception))
                        {
                            throw;
                        }

                        LogHelper.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, null, automationElement, string.Format("IUIAutomation.ElementFromPoint({0}, {1})", x, y));
                        continue;
                    }

                    if (fromPoint == null)
                    {
                        //"AccessibleObjectFromPoint({0}, {1}) returned a null IAccessible"),
                        LogHelper.Log(EventLevel.Error, ElementFromPointReturnedNull, this.GetType(), automationElement, x, y);
                    }
                    else
                    {
                        // If we find it in the cache, then we pass, if not, then walk it's parent chain
                        // and find out if it thinks it's a descendant of our root element.
                        if (!FindInCachedTree(automationElement, fromPoint))
                        {
                            UIA.IUIAutomationTreeWalker rawWalker = UIAGlobalContext.Automation.RawViewWalker;

                            UIA.IUIAutomationElement current = fromPoint;
                            UIA.IUIAutomationElement rootElement = UIAGlobalContext.GetRootElement();
                            while (current != null)
                            {
                                if (UIAGlobalContext.Automation.CompareElements(automationElement, current) != 0)
                                {
                                    // Don't Report repeats
                                    UiaElementMatcher matcher = new UiaElementMatcher(fromPoint);
                                    if (matcher.IsContainedIn(_orphanedElements))
                                    {
                                        break;
                                    }

                                    _orphanedElements.Add(fromPoint);

                                    LogHelper.Log(EventLevel.Warning, UiaElementIsOrphaned, this.GetType(), fromPoint);
                                    break;
                                }

                                // While traversing upwards, if we end our parent chain, without hitting the
                                // desktop root element, something is wrong.
                                UIA.IUIAutomationElement nextParent = rawWalker.GetParentElement(current);
                                if (nextParent == null)
                                {
                                    if (UIAGlobalContext.Automation.CompareElements(rootElement, current) == 0)
                                    {
                                        // Don't Report repeats
                                        UiaElementMatcher matcher = new UiaElementMatcher(fromPoint);
                                        if (matcher.IsContainedIn(_unparentedElements))
                                        {
                                            break;
                                        } 
                                        
                                        _unparentedElements.Add(fromPoint);

                                        LogHelper.Log(EventLevel.Warning, UiaElementDoesNotParentToRoot, this.GetType(), fromPoint);
                                        break;
                                    }
                                }
                                current = nextParent;
                            }
                        }
                    }
                }
            }
        }

        private bool FindInCachedTree(UIA.IUIAutomationElement treeRoot, UIA.IUIAutomationElement elementToFind)
        {
            if (UIAGlobalContext.Automation.CompareElements(treeRoot, elementToFind) != 0)
            {
                return true;
            }

            UIA.IUIAutomationElementArray children = treeRoot.GetCachedChildren();
            if (children != null)
            {
                for (int i = 0; i < children.Length; i++)
                {
                    if (FindInCachedTree(children.GetElement(i), elementToFind))
                    {
                        return true;
                    }
                }
            }
            return false;
        }
        
    
        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            LogHelper.Logger = logger;
            LogHelper.RootHwnd = hwnd;

            ElementFromPoint_OrphanChildren(UIAGlobalContext.CacheUIATree(hwnd));
        }
    }
}
