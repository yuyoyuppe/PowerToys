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

namespace UIAVerifications
{
    [Verification(
        "CheckNameUIA",
        "TBD",
        Group = "UIA Properties",
        Priority = 1)]
    public class CheckName : IVerificationRoutine
    {

        #region LogMessages

        static LogMessage ElementHasNoName = new LogMessage("Element has no name", "ElementHasNoName", 1);
        static LogMessage UiaNameLengthTooLong = new LogMessage("Element's Name should not return a string longer than {0} characters", "UiaNameLengthTooLong", 2);
        static LogMessage UiaNameShouldNotContainControlType = new LogMessage("Element's Name \"{0}\" should not contain the text of the Element's ControlType \"{1}\"", "UiaNameShouldNotContainControlType", 2);

        #endregion LogMessages

        private UIA.IUIAutomationElement [] _firstNoNameErrors = new UIA.IUIAutomationElement[3];
        private int [] _suppressFirstErrorControlTypes = new int [] {UIA.UIA_ControlTypeIds.UIA_ListControlTypeId, UIA.UIA_ControlTypeIds.UIA_TreeControlTypeId, UIA.UIA_ControlTypeIds.UIA_ToolBarControlTypeId};
        
        private void UiaName_FocusControlsHaveNames(UIA.IUIAutomationElement automationElement)
        {
            if (automationElement.CachedIsKeyboardFocusable != 0 &&
                automationElement.CachedIsEnabled != 0)
            {
                UIAutomationClient.tagRECT tempRect = automationElement.CachedBoundingRectangle;
                Rectangle rect = new Rectangle(tempRect.left, tempRect.top, tempRect.right - tempRect.left, tempRect.bottom - tempRect.top);
                if (rect == Rectangle.Empty || rect.Width == 0 || rect.Height == 0)
                {
                    // ignore stuff that is affectively invisible
                    return;
                }

                // Title bar does not have a name but does have a value
                // Tab Control is the same
                // Panes frequently lack a name, and they report as focusable, since their children are focusable
                int controlType = automationElement.CachedControlType;
                if (controlType != UIA.UIA_ControlTypeIds.UIA_TitleBarControlTypeId && 
                    controlType != UIA.UIA_ControlTypeIds.UIA_TabControlTypeId &&
                    controlType != UIA.UIA_ControlTypeIds.UIA_PaneControlTypeId)
                {
                    string name = automationElement.CachedName;
                    if (string.IsNullOrEmpty(name))
                    {
                        bool errorSuppressed = false;
                        
                        for (int i = 0; i < _firstNoNameErrors.Length; i++)
                        {
                            if  (_suppressFirstErrorControlTypes[i] == controlType) 
                            {   
                                if (_firstNoNameErrors[i] == null)
                                {
                                    errorSuppressed = true;
                                    _firstNoNameErrors[i] = automationElement;
                                }
                                else
                                {
                                    LogHelper.Log(EventLevel.Error, ElementHasNoName, this.GetType(), _firstNoNameErrors[i]);
                                    // put a bad value in here so from now we just log all no names for this control type
                                    _suppressFirstErrorControlTypes[i] = -1;
                                }
                                break;
                            }
                        }
                        
                        if (!errorSuppressed)
                        {
                            //"There is no name for the element");
                            LogHelper.Log(EventLevel.Error, ElementHasNoName, this.GetType(), automationElement);
                        }
                    }
                }
            }
        }

        private void UiaName_FocusableControlsNameDoesNotIncludeRoleText(UIA.IUIAutomationElement automationElement)
        {
            string name = automationElement.CachedName;

            if (automationElement.CachedIsKeyboardFocusable != 0 &&
                automationElement.CachedIsEnabled != 0)
            {
                string lname = name.ToLower();
                string controlType = automationElement.CachedLocalizedControlType.ToLower();

                // People often use the role in the name correctly like "Explorer pane"
                // but starting with the role is usually bad like "pane1"
                // If the name has a space in it it is probably not something that defaulted like "button_1"
                if (!string.IsNullOrEmpty(controlType) && lname.StartsWith(controlType) && !lname.Contains(" "))
                {
                    // "accName ({0}) should not contain the text of the Role ({1})"
                    LogHelper.Log(EventLevel.Error, UiaNameShouldNotContainControlType, this.GetType(), automationElement, name, controlType);
                }
            }
        }

        private void UiaName_NameStringNotTooLong(UIA.IUIAutomationElement automationElement)
        {
            const int MAX_NAME_LENGTH = 32000;

            string name = automationElement.CachedName;

            if (name != null)
            {
                if (name.Length > MAX_NAME_LENGTH)
                {
                    LogHelper.Log(EventLevel.Error, UiaNameLengthTooLong, this.GetType(), automationElement, MAX_NAME_LENGTH);
                }
            }
        }

        private void TraverseFromNode(UIA.IUIAutomationElement automationElement)
        {
            UiaName_FocusControlsHaveNames(automationElement);
            UiaName_FocusableControlsNameDoesNotIncludeRoleText(automationElement);
            UiaName_NameStringNotTooLong(automationElement);
            
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

            for (int i = 0; i < _firstNoNameErrors.Length; i++)
            {
                _firstNoNameErrors[i] = null;
            }

            TraverseFromNode(UIAGlobalContext.CacheUIATree(hwnd));
        }
    }
}
