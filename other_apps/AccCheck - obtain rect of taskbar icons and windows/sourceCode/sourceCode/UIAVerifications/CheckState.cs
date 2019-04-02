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
        "CheckStateUIA",
        "Checks several various \"state\" properties of UIA elements for consistency",
        Group = "UIA Properties",
        Priority = 2)]
    public class CheckStateUIA : IVerificationRoutine
    {

        #region LogMessages

        private static LogMessage ElementShouldBeOffScreen = new LogMessage("Element should be offscreen", "ElementShouldBeOffScreen", 2);
        private static LogMessage InconsistentProperties = new LogMessage("A control has properties that are inconsistent {0} and {1}", "InconsistentProperties", 2);

        #endregion LogMessages
        private void UiaState_Inconsistencies(UIA.IUIAutomationElement automationElement)
        {
            // We only want to do this verification on leaf nodes, since an element may not itself be focusable, but it's children are
            UIA.IUIAutomationElementArray children = automationElement.GetCachedChildren();
            if (children != null && children.Length > 0)
            {
                return;
            }
            
            if (automationElement.CurrentHasKeyboardFocus != 0 && automationElement.CachedIsKeyboardFocusable == 0)
            {
                LogHelper.Log(EventLevel.Error, InconsistentProperties, this.GetType(), automationElement, "HasKeyboardFocus", "IsKeyboardFocusable");
            }
        }

        public void UiaState_ShouldBeOffScreen(UIA.IUIAutomationElement automationElement)
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

            // If the rect is clipped by the parent there is nothing left it should be offscreen
            UIA.IUIAutomationElement parent = automationElement.GetCachedParent();

            if (parent == null)
            {
                return;
            }

            tempRect = parent.CachedBoundingRectangle;
            Rectangle rectParent = new Rectangle(tempRect.left, tempRect.top, tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
            rect.Intersect(rectParent);

            if (rect.IsEmpty)
            {
                if( automationElement.CachedIsOffscreen == 0 )
                {
                    LogHelper.Log(EventLevel.Warning, ElementShouldBeOffScreen, this.GetType(), automationElement);
                }

                return;
            }


        }


        private void TraverseFromNode(UIA.IUIAutomationElement automationElement)
        {
            UiaState_Inconsistencies(automationElement);
            UiaState_ShouldBeOffScreen(automationElement);
            
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
