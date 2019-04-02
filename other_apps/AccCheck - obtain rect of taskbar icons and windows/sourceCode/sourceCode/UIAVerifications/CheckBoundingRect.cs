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
using System.Drawing;
using UIA=UIAutomationClient;

namespace UIAVerifications
{
    [Verification(
        "CheckBoundingRectUIA",
        "Checks the controls bounding rectangle via hittesting",
        Group = "UIA Properties",
        Priority = 1)]
    public class CheckBoundingRectUIA : IVerificationRoutine
    {

        #region LogMessages

        private static LogMessage BoundingRectIsIncorrect = new LogMessage("The {0} edge of the reported rectangle is inconsistent with hit testing {1} the rect at point {2}.", "BoundingRectIsIncorrect", 1);
        private static LogMessage BoundingRectNotHittestable = new LogMessage("The reported rectangle does not appear to contain a point where this element can be retrieved with hit testing.", "BoundingRectNotHittestable", 1);

        #endregion LogMessages
        
        private bool _errorLoggedOutsideTop;
        private bool _errorLoggedOutsideBottom;
        private bool _errorLoggedOutsideRight;
        private bool _errorLoggedOutsideleft;

        // Hittest points outside the bounding rect expecting to get back an element 
        // other that the element being verified.  During vista the titlebar buttons 
        // had problems that this verification would have caught.
        public void VerifyPointsOutsideRect(UIA.IUIAutomationElement automationElement)
        {
            // We only want to do this verification on leaf nodes
            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null && children.Length > 0)
            {
                return;
            }
            

            _errorLoggedOutsideTop = false;
            _errorLoggedOutsideBottom = false;
            _errorLoggedOutsideRight = false;
            _errorLoggedOutsideleft = false;
            
            // Check the rects twice, once for really bad error that would be pri 1
            // and once for not so bad errors pri 3.
            CheckTheRect(automationElement, 10, 1);
            CheckTheRect(automationElement, 1, 3);
            
        }

        // Hittest points inside the bounding rect expecting to get back the element 
        // verified.  Since we cannot expect all points inside to return the element 
        // because of transparency try a few and if we don't find any return log an error.
        public void VerifyPointsInsideRect(UIA.IUIAutomationElement automationElement)
        {
            // We only want to do this verification on leaf nodes
            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null && children.Length > 0)
            {
                return;
            }

            UIAutomationClient.tagRECT tempRect = automationElement.CachedBoundingRectangle;
            Rectangle rect = new Rectangle(tempRect.left, tempRect.top, tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
            if (rect == Rectangle.Empty || rect.Width == 0 || rect.Height == 0)
            {
                return;
            }

            // Make sure the rect is not clipped by its parent
            UIA.IUIAutomationElement parent = automationElement.GetCachedParent();
            int lastControlType = automationElement.CachedControlType;
            while (parent != null)
            {
                bool fShouldClip = true;

                // Avoid clipping the system menu
                // The system menu is a child of the titlebar in UIA, but the titlebar's bounding rect
                // does not contain the system menu.  Thus, we end up clipping to an unusably small region.
                // The pattern of TitleBar containing MenuBar is rare enough that it uniquely identifies this situation.
                if (lastControlType == UIA.UIA_ControlTypeIds.UIA_MenuBarControlTypeId &&
                    parent.CachedControlType == UIA.UIA_ControlTypeIds.UIA_TitleBarControlTypeId)
                {
                    fShouldClip = false;
                }

                if (fShouldClip)
                {
                    tempRect = parent.CachedBoundingRectangle;
                    Rectangle rectParent = new Rectangle(tempRect.left, tempRect.top, tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
                    rect.Intersect(rectParent);
                }

                lastControlType = parent.CachedControlType;
                parent = parent.GetCachedParent();
            }

            // if the rect is not visible because of clipping there is not point in going on
            if (rect.IsEmpty)
            {
                return;
            }

            Point center = new Point(rect.Left + (rect.Width / 2), rect.Top + (rect.Height / 2));
            if (PointIsInElement(center, automationElement))
            {
                return;
            }

            Point topMiddle = new Point(rect.Left + (rect.Width / 2), rect.Top);
            if (PointIsInElement(topMiddle, automationElement))
            {
                return;
            }

            Point bottomMiddle = new Point(rect.Left + (rect.Width / 2), rect.Bottom - 1);
            if (PointIsInElement(bottomMiddle, automationElement))
            {
                return;
            }

            Point leftMiddle = new Point(rect.Left, rect.Top + (rect.Height / 2));
            if (PointIsInElement(leftMiddle, automationElement))
            {
                return;
            }

            Point rightMiddle = new Point(rect.Right - 1, rect.Top + (rect.Height / 2));
            if (PointIsInElement(rightMiddle, automationElement))
            {
                return;
            }

            LogHelper.Log(EventLevel.Warning, BoundingRectNotHittestable, this.GetType(), automationElement);

        }

        private bool PointIsInElement(Point screenPoint, UIA.IUIAutomationElement automationElement)
        {
            bool returnVal = false;

            try
            {
                UIA.IUIAutomationElement fromPoint = UIAGlobalContext.ElementFromPoint(screenPoint.X, screenPoint.Y);

                if (fromPoint != null)
                {
                    if (UIAGlobalContext.Automation.CompareElements(automationElement, fromPoint) != 0)
                    {
                        // The point we got hit test to be in the element 
                        returnVal = true;
                    }
                }

            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception))
                {
                    throw;
                }

                // we want to ingore all exceptions here because the error in this case
                // is if we get back the same element in this case "element" if there is 
                // any exception then we cannot say for sure its an error so just ingore it.
            }


            return returnVal;
        }

        private void CheckTheRect(UIA.IUIAutomationElement automationElement, int fudgeFactor, int priority)
        {
            UIAutomationClient.tagRECT tempRect = automationElement.CachedBoundingRectangle;
            Rectangle rect = new Rectangle(tempRect.left, tempRect.top, tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
            if (rect == Rectangle.Empty || rect.Width == 0 || rect.Height == 0)
            {
                return;
            }
            
            LogMessage logMessage = BoundingRectIsIncorrect;
            logMessage.Priority = priority;
            
            Point topMiddle = new Point(rect.Left + (rect.Width / 2), rect.Top - 1 - fudgeFactor);
            if (!_errorLoggedOutsideTop && PointIsInElement(topMiddle, automationElement))
            {
                LogHelper.Log(EventLevel.Error, logMessage, this.GetType(), automationElement, "top", "outside", topMiddle);
                _errorLoggedOutsideTop = true;
            }
            
            Point bottomMiddle = new Point(rect.Left + (rect.Width / 2), rect.Bottom + fudgeFactor);
            if (!_errorLoggedOutsideBottom && PointIsInElement(bottomMiddle, automationElement))
            {
                LogHelper.Log(EventLevel.Error, logMessage, this.GetType(), automationElement, "bottom", "outside", bottomMiddle);
                _errorLoggedOutsideBottom = true;
            }
            
            Point leftMiddle = new Point(rect.Left - 1 - fudgeFactor, rect.Top + (rect.Height / 2));
            if (!_errorLoggedOutsideleft && PointIsInElement(leftMiddle, automationElement))
            {
                LogHelper.Log(EventLevel.Error, logMessage, this.GetType(), automationElement, "left", "outside", leftMiddle);
                _errorLoggedOutsideleft = true;
            }
            
            Point rightMiddle = new Point(rect.Right + fudgeFactor, rect.Top + (rect.Height / 2));
            if (!_errorLoggedOutsideRight && PointIsInElement(rightMiddle, automationElement))
            {
                LogHelper.Log(EventLevel.Error, logMessage, this.GetType(), automationElement, "right", "outside", rightMiddle);
                _errorLoggedOutsideRight = true;
            }
        }

        private void TraverseFromNode(UIA.IUIAutomationElement automationElement)
        {
            VerifyPointsOutsideRect(automationElement);
            VerifyPointsInsideRect(automationElement);
            
            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null)
            {
                for (int i = 0; i < children.Length; i++)
                {
                    TraverseFromNode(children.GetElement(i));
                }
            }
        }
        
    
        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            LogHelper.Logger = logger;
            LogHelper.RootHwnd = hwnd;

            TraverseFromNode(UIAGlobalContext.CacheUIATree(hwnd));
        }
    }
}
