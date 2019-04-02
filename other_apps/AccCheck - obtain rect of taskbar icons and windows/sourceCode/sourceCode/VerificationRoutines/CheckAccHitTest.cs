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
using Accessibility;
using System.Drawing;
using System.Diagnostics;

namespace VerificationRoutines
{
    [Verification(
        "CheckOrphanChildren",
        "TBD",
        Group = ClassificationsGroup.TreeStructure,
        Priority = 1)]
    public class CheckOrphanChildren : IVerificationRoutine
    {
        static IntPtr _intPtrFromWindowFromAccessibleObject = IntPtr.Zero;

        #region LogMessages

        static LogMessage ElementIsOrphaned = new LogMessage("Element is an orphaned IAccessible in the tree", "ElementIsOrphaned", 1);
        static LogMessage AccessibleObjectFromPointReturnedNot_S_OK = new LogMessage("AccessibleObjectFromPoint({0}, {1}) returned {2} and expected S_OK", "AccessibleObjectFromPointReturnedNot_S_OK", 2);
        static LogMessage AccessibleObjectFromPointReturnedNullChildId = new LogMessage("AccessibleObjectFromPoint({0}, {1}) returned a valid IAccessible, but a null childId", "AccessibleObjectFromPointReturnedNullChildId", 2);
        static LogMessage AccessibleObjectFromPointReturnedNullIAccessible = new LogMessage("AccessibleObjectFromPoint({0}, {1}) returned a null IAccessible", "AccessibleObjectFromPointReturnedNullIAccessible", 1);

        #endregion LogMessages

        public void Execute(IntPtr hwnd, ILogger logger, bool allowUI, GraphicsHelper graphicsHelper)
        {
            TestLogger.Logger = logger;
            TestLogger.RootHwnd = hwnd;
            AccHitTest_OrphanChildren(MsaaElement.CacheMsaaTree(hwnd));
        }

        /// -------------------------------------------------------------------
        /// <summary>Instantiate Win32WinEnumProc delegate</summary>
        /// -------------------------------------------------------------------
        private static Win32API.WinEnumProc EnumerateChildWindows
        {
            get
            {
                return new Win32API.WinEnumProc(FindChildWindow);
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Compares each child window to see if 
        /// _intPtrFromWindowFromAccessibleObject is in the list.</summary>
        /// <param name="hwnd"> "Next" hwnd passed from EnumWindows </param>
        /// <param name="lParam">Used to determine if the window was found.  Set 
        /// to _intPtrFromWindowFromAccessibleObject IntPtr if found</param>
        /// <returns>true - continue enum, false enum should be stopped</returns>
        /// -------------------------------------------------------------------
        private static bool FindChildWindow(IntPtr hwnd, ref IntPtr lParam)
        {
            if (hwnd == _intPtrFromWindowFromAccessibleObject)
                lParam = _intPtrFromWindowFromAccessibleObject;
            // Continue until the two are the same
            return (!(hwnd == lParam));
        }

        ///  ------------------------------------------------------------------
        /// <summary>Seach within the cordinates of the element to see if 
        /// the IAccessible's found from AccessibleObjectFromPoint are within 
        /// the tree of the MsaaElements.  If we can't find the IAccessible within 
        /// the cached tree, then look to see if the IAccessible's hwnd is a child
        /// of the element's hwnd...if it is not,then it's an orphanded IAccessible
        /// which cannot be navigated to from the element.</summary>
        ///  ------------------------------------------------------------------
        [TestMethod(Description = "Check for orphan children")]
        public void AccHitTest_OrphanChildren(MsaaElement element)
        {
            const int STEP_X = 75;  // Controls are usually longer than higher so optimize here.
            const int STEP_Y = 25;

            // Check to see if the element is visible.  If it is not visible, then don't test
            // TODO$: Pull this when Paul has his visible override implemented.
            if (!element.IsVisible)
            {
                return;
            }

            // Find the rectangle that we are going to search within based on the rectangle 
            // if the element.
            Win32API.RECT win32Rect = Win32API.RECT.Empty;
            Win32API.GetWindowRect(element.GetHwnd(), ref win32Rect);

            Accessible accFromPoint = null;
            int hResult;

            for (int x = win32Rect.left + 1; x < win32Rect.right; x += STEP_X)
            {
                for (int y = win32Rect.top + 1; y < win32Rect.bottom; y += STEP_Y)
                {
                    // The non-client area bits are not important to check.
                    uint lresult = 0;
                    Win32API.SendMessageTimeout(element.GetHwnd(), Win32API.WM_NCHITTEST, IntPtr.Zero, Win32API.MAKELPARAM(x, y), 0, 10000, out lresult);
                    if (lresult == Win32API.HTTOPRIGHT || lresult == Win32API.HTTOPLEFT || lresult == Win32API.HTBOTTOMLEFT || lresult == Win32API.HTBOTTOMRIGHT ||
                        lresult == Win32API.HTLEFT || lresult == Win32API.HTTOP || lresult == Win32API.HTRIGHT || lresult == Win32API.HTBOTTOM ||
                        lresult == Win32API.HTBORDER || lresult == Win32API.HTCAPTION)
                    {
                        continue;
                    }
                    
                    try
                    {
                        hResult = Accessible.FromPoint(new Point(x, y), out accFromPoint);
                    }
                    catch (VariantNotIntException)
                    {
                        //"AccessibleObjectFromPoint({0}, {1}) returned a valid IAccessible, but a null childId"),
                        TestLogger.Log(EventLevel.Error, AccessibleObjectFromPointReturnedNullChildId, this.GetType(), element, x, y);
                        continue;
                    }
                    catch (Exception exception)
                    {
                        if (VerificationHelper.IsCriticalException(exception)) 
                        {
                            throw; 
                        }
                        
                        TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, this.GetType(), exception, null, element, string.Format("AccessibleObjectFromPoint({0}, {1})", x, y));
                        continue;
                    }

                    if (hResult != Win32API.S_OK)
                    {
                        //"AccessibleObjectFromPoint({0}, {1}) returned {2} and expected S_OK"
                        TestLogger.Log(EventLevel.Warning, AccessibleObjectFromPointReturnedNot_S_OK, this.GetType(), element, x, y, hResult);
                    }
                    else if (accFromPoint == null)
                    {
                        //"AccessibleObjectFromPoint({0}, {1}) returned a null IAccessible"),
                        TestLogger.Log(EventLevel.Error, AccessibleObjectFromPointReturnedNullIAccessible, this.GetType(), element, x, y);
                    }
                    else
                    {
                        MsaaElement msaaElementFromPoint = MsaaElement.FindMsaaElement(accFromPoint);

                        // If we find it in the cache, then we pass, if not, then go see if it's hwnd is a child of the 
                        // element's tree.  There is times where the point will be pass through to the background, or
                        // we have an overlapping window.
                        if (msaaElementFromPoint == null)
                        {
                            _intPtrFromWindowFromAccessibleObject = accFromPoint.Hwnd;
                            // We found an IAccessible from WindowFromAccessibleObject that not in the cached tree.
                            // If the hwnd of this IAccessible is not part of the element's hwnd tree we don't care.  
                            // An example of this is a window that is overlapping (ie. notepad.exe) the window 
                            // we are testing (calc.exe).  When we get a point from the bounding rect of 
                            // the element (calc.exe), and then get wet window from that point (notepad.exe), 
                            // we would find the overlapping window which we don't care about.

                            IntPtr rootHwnd = element.GetHwnd();
                            IntPtr lParam = IntPtr.Zero;

                            // don't report errors for the root
                            if (rootHwnd != _intPtrFromWindowFromAccessibleObject)
                            {
                                Win32API.EnumChildWindows(rootHwnd, EnumerateChildWindows, ref lParam);

                                // If EnumerateChildWindows returns a non IntPtr.Zero value, that means that 
                                // the a window found in EnumerateChildWindows is the same as _intPtrFromWindowFromAccessibleObject, 
                                // which means that the window is a child if rootHwnd and we have an orphaned child since it's 
                                // not in the cache.
                                if (lParam != IntPtr.Zero)
                                {
                                    TestLogger.Log(EventLevel.Warning, ElementIsOrphaned, this.GetType(), accFromPoint);
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}
