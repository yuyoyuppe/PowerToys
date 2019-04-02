// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using Accessibility;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;

namespace VerificationRoutines
{
    [Verification(
        "CheckBoundingRect",
        "TBD",
        Group = ClassificationsGroup.Properties,
        Priority = 3)]
    /// -------------------------------------------------------------------
    /// <summary>Check accLocation Tests</summary>
    /// <TESTS>
    /// Test : Verify that points outside the bounding rect hit test to something other than that element.
    /// </TESTS>
    /// -------------------------------------------------------------------
    public class CheckBoundingRect : MsaaTraversingVerificationBase, IVerificationRoutine
    {
        #region LogMessages

        private static LogMessage BoundingRectIsIncorrect = new LogMessage("The {0} edge of the reported rectangle is inconsistent with hit testing {1} the rect at point {2}.", "BoundingRectIsIncorrect", 1);
        private static LogMessage BoundingRectNotHittestable = new LogMessage("The reported rectangle does not appear to contain a point where this element can be retrieved with hit testing.", "BoundingRectNotHittestable", 1);

        private bool _errorLoggedOutsideTop;
        private bool _errorLoggedOutsideBottom;
        private bool _errorLoggedOutsideRight;
        private bool _errorLoggedOutsideleft;
        
        #endregion LogMessages

        protected override bool IgnoreInvisibleElements
        {
            get
            {
                return true;
            }
        }

        // Hittest points outside the bounding rect expecting to get back an element 
        // other that the element being verified.  During vista the titlebar buttons 
        // had problems that this verification would have caught.
        [TestMethod("Verify points outside the bounding rect")]
        public void VerifyPointsOutsideRect(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }

            // We only wnat to do this verification on leaf nodes
            if (element.Children.Count > 0)
            {
                return;
            }

            // In some situations, properties of the element have changed as a side
            // effect of another test (tabbing will add /t to the elements accValue.  
            // Need to update the properties that can change so that when we compare 
            // the elements, the cached value is up todate.
            MsaaElement.UpdateMsaaCachedDynamicProperties(element);

            _errorLoggedOutsideTop = false;
            _errorLoggedOutsideBottom = false;
            _errorLoggedOutsideRight = false;
            _errorLoggedOutsideleft = false;
            
            // Check the rects twice, once for really bad error that would be pri 1
            // and once for not so bad errors pri 3.
            CheckTheRect(element, 10, 1);
            CheckTheRect(element, 1, 3);

        }
            
        // Hittest points inside the bounding rect expecting to get back the element 
        // verified.  Since we cannot expect all points inside to return the element 
        // because of transparency try a few and if we don't find any return log an error.
        [TestMethod("Verify points Inside the bounding rect")]
        public void VerifyPointsInsideRect(MsaaElement element)
        {
            if (VerificationHelper.ShouldIgnoreElement(element, this.GetType()))
            {
                return;
            }

            // We only wnat to do this verification on leaf nodes
            if (element.Children.Count > 0)
            {
                return;
            }

            // In some situations, properties of the element have changed as a side
            // effect of another test (tabbing will add /t to the elements accValue.  
            // Need to update the properties that can change so that when we compare 
            // the elements, the cached value is up todate.
            MsaaElement.UpdateMsaaCachedDynamicProperties(element);

            Rectangle rect = element.GetBoundingBox();
            if (rect == Rectangle.Empty || rect.Width == 0 || rect.Height == 0)
            {
                return;
            }

            // Make sure the rect is not clipped by its parent
            MsaaElement parent = element.Parent;
            while (parent != null)
            {
                Rectangle rectParent = parent.GetBoundingBox();
                rect.Intersect(rectParent);
                parent = parent.Parent;
            }

            // if the rect is not visible because of clipping there is not point in going on
            if (rect.IsEmpty)
            {
                return;
            }
            
            Point center = new Point(rect.Left + (rect.Width / 2), rect.Top + (rect.Height / 2));
            if (PointIsInElement(center, element))
            {
                return;
            }

            Point topMiddle = new Point(rect.Left + (rect.Width / 2), rect.Top);
            if (PointIsInElement(topMiddle, element))
            {
                return;
            }
            
            Point bottomMiddle = new Point(rect.Left + (rect.Width / 2), rect.Bottom - 1);
            if (PointIsInElement(bottomMiddle, element))
            {
                return;
            }
            
            Point leftMiddle = new Point(rect.Left, rect.Top + (rect.Height / 2));
            if (PointIsInElement(leftMiddle, element))
            {
                return;
            }
            
            Point rightMiddle = new Point(rect.Right - 1, rect.Top + (rect.Height / 2));
            if (PointIsInElement(rightMiddle, element))
            {
                return;
            }

            TestLogger.Log(EventLevel.Warning, BoundingRectNotHittestable, this.GetType(), element);

        }
            
        private bool PointIsInElement(Point screenPoint, MsaaElement element)
        {
                bool returnVal = false;

                try
                {
                    Accessible acc = null;
                    int result = Accessible.FromPoint(screenPoint, out acc);

                    // debug code to move the mouse so you can see what points are being tested
                    // MoveMouse(screenPoint);
                    // System.Threading.Thread.Sleep(500);
                    
                    if (result == Win32API.S_OK && acc != null)
                    {
                        if (element.Compare(acc))
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


        private void CheckTheRect(MsaaElement element, int fudgeFactor, int priority)
        {
            Rectangle rect = element.GetBoundingBox();
            if (rect == Rectangle.Empty || rect.Width == 0 || rect.Height == 0)
            {
                return;
            }

            LogMessage logMessage = BoundingRectIsIncorrect;
            logMessage.Priority = priority;
            
            Point topMiddle = new Point(rect.Left + (rect.Width / 2), rect.Top - 1 - fudgeFactor);
            if (!_errorLoggedOutsideTop && PointIsInElement(topMiddle, element))
            {
                TestLogger.Log(EventLevel.Error, logMessage, this.GetType(), element, "top", "outside", topMiddle);
                _errorLoggedOutsideTop = true;
            }
            
            Point bottomMiddle = new Point(rect.Left + (rect.Width / 2), rect.Bottom + fudgeFactor);
            if (!_errorLoggedOutsideBottom && PointIsInElement(bottomMiddle, element))
            {
                TestLogger.Log(EventLevel.Error, logMessage, this.GetType(), element, "bottom", "outside", bottomMiddle);
                _errorLoggedOutsideBottom = true;
            }
            
            Point leftMiddle = new Point(rect.Left - 1 - fudgeFactor, rect.Top + (rect.Height / 2));
            if (!_errorLoggedOutsideleft && PointIsInElement(leftMiddle, element))
            {
                TestLogger.Log(EventLevel.Error, logMessage, this.GetType(), element, "left", "outside", leftMiddle);
                _errorLoggedOutsideleft = true;
            }
            
            Point rightMiddle = new Point(rect.Right + fudgeFactor, rect.Top + (rect.Height / 2));
            if (!_errorLoggedOutsideRight && PointIsInElement(rightMiddle, element))
            {
                TestLogger.Log(EventLevel.Error, logMessage, this.GetType(), element, "right", "outside", rightMiddle);
                _errorLoggedOutsideRight = true;
            }
        }
/*
        // this is used for debuging to see what points we are testing
        private void MoveMouse(Win32API.POINT point)
        {
            Win32API.INPUT mouseInput = new Win32API.INPUT();

            mouseInput.type = Win32API.INPUT_MOUSE;
            mouseInput.union.mouseInput.dwFlags = Win32API.MOUSEEVENTF_MOVE | Win32API.MOUSEEVENTF_ABSOLUTE | Win32API.MOUSEEVENTF_VIRTUALDESK;
            mouseInput.union.mouseInput.mouseData = 0;
            mouseInput.union.mouseInput.time = 0;
            mouseInput.union.mouseInput.dwExtraInfo = new IntPtr( 0 );

            int screenWidth = Win32API.GetSystemMetrics(Win32API.SM_CXVIRTUALSCREEN);
            int screenHeight = Win32API.GetSystemMetrics(Win32API.SM_CYVIRTUALSCREEN);
            int screenLeft = Win32API.GetSystemMetrics(Win32API.SM_XVIRTUALSCREEN);
            int screenTop = Win32API.GetSystemMetrics(Win32API.SM_YVIRTUALSCREEN);
            
            double x = ( ( point.x - screenLeft ) * 65536 ) / screenWidth + 65536 / ( screenWidth * 2 );
            double y = ( ( point.y - screenTop ) * 65536 ) / screenHeight + 65536 / ( screenHeight * 2 );

            mouseInput.union.mouseInput.dx = (int)x;
            mouseInput.union.mouseInput.dy = (int)y;

            Win32API.SendInput(1, ref mouseInput, Marshal.SizeOf(mouseInput));
        }
*/        
    }
}
