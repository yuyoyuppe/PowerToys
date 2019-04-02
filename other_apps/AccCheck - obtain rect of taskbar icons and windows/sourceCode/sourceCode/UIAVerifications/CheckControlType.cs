//-----------------------------------------------------------------------------
// Copyright (c) Microsoft Corporation. All rights reserved.
// Description: This is the set of tests that are used within the AccChecker application
//-----------------------------------------------------------------------------
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
        "CheckControlTypeUIA",
        "Verifies that the element has a control type, and supports the required patterns for that control type",
        Group = "UIA Properties",
        Priority = 1)]
    public class CheckControlType : IVerificationRoutine
    {

        #region LogMessages

        static LogMessage ElementHasInvalidControlType = new LogMessage("Element either failed to return a control type or returned an unknown Control Type", "ElementHasInvalidControlType", 1);
        static LogMessage CustomControlTypeWithoutLocalizedControlType = new LogMessage("Element has the CustomControlType but has an empty LocalizedControlType", "CustomControlTypeWithoutLocalizedControlType", 1);
        static LogMessage ControlTypeRequiredPatternNotSupported = new LogMessage("Element with a ControlType of {0} should support {1} but this element does not", "ControlTypeRequiredPatternNotSupported", 1);
        static LogMessage ControlTypeNoExpectedPatternsSupported = new LogMessage("Element with a ControlType of {0} should support at least one of these patterns:{1} but this element does not", "ControlTypeNoExpectedPatternsSupported", 1);

         #endregion LogMessages

        class RequiredPatterns
        {
            public int controlTypeId;
            public int[] patternIds;
            public bool allAreRequired;

            public RequiredPatterns(int cid, int[] pids, bool allreq)
            {
                controlTypeId = cid;
                patternIds = pids;
                allAreRequired = allreq;
            }
        }

        // This table takes the necessary parts from:
        // http://msdn.microsoft.com/en-us/library/dd319586(VS.85).aspx

        // Entries that would be empty are simply ommitted

        // Several entries here, according to the rules of the table have optional patterns, at least 1 of which should be supported
        // but all of the listed patterns are truly optional, so they have been intentionally commented out, with the keyword OMITTING

        // Several other entries seem to be correct but current Proxy implementations don't have the full support, these are
        // commented out with the keyword SUPPRESSED
        static RequiredPatterns[] controlTypeToPatternMap = {
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ButtonControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_InvokePatternId, UIA.UIA_PatternIds.UIA_ExpandCollapsePatternId, UIA.UIA_PatternIds.UIA_TogglePatternId}, false),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_CalendarControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_GridPatternId, UIA.UIA_PatternIds.UIA_TablePatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_CheckBoxControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_TogglePatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ComboBoxControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_ExpandCollapsePatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_DataGridControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_GridPatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_DataItemControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_SelectionItemPatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_DocumentControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_TextPatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_EditControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_RangeValuePatternId, UIA.UIA_PatternIds.UIA_TextPatternId, UIA.UIA_PatternIds.UIA_ValuePatternId}, false),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_GroupControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_ExpandCollapsePatternId}, false),
            /* OMITTING: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_HeaderControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_TransformPatternId}, false), */
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_HeaderItemControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_InvokePatternId, UIA.UIA_PatternIds.UIA_TransformPatternId}, false),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_HyperlinkControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_InvokePatternId}, true),
            /* OMITTING: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ImageControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_GridItemPatternId, UIA.UIA_PatternIds.UIA_TableItemPatternId}, false), */
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ListControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_GridPatternId, UIA.UIA_PatternIds.UIA_MultipleViewPatternId, UIA.UIA_PatternIds.UIA_ScrollPatternId, UIA.UIA_PatternIds.UIA_SelectionPatternId}, false),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ListItemControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_SelectionItemPatternId}, true),
            /* OMITTING: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_MenuBarControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_DockPatternId, UIA.UIA_PatternIds.UIA_ExpandCollapsePatternId, UIA.UIA_PatternIds.UIA_TransformPatternId}, false), */
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_MenuItemControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_InvokePatternId, UIA.UIA_PatternIds.UIA_ExpandCollapsePatternId, UIA.UIA_PatternIds.UIA_TogglePatternId, UIA.UIA_PatternIds.UIA_SelectionItemPatternId}, false),
            /* OMITTING: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_PaneControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_DockPatternId, UIA.UIA_PatternIds.UIA_ScrollPatternId, UIA.UIA_PatternIds.UIA_TransformPatternId}, false), */
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ProgressBarControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_RangeValuePatternId, UIA.UIA_PatternIds.UIA_ValuePatternId}, false),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_RadioButtonControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_SelectionItemPatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ScrollBarControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_RangeValuePatternId}, false),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_SliderControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_RangeValuePatternId, UIA.UIA_PatternIds.UIA_SelectionPatternId, UIA.UIA_PatternIds.UIA_ValuePatternId}, false),
            /* SUPPRESSED: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_SpinnerControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_RangeValuePatternId, UIA.UIA_PatternIds.UIA_SelectionPatternId, UIA.UIA_PatternIds.UIA_ValuePatternId}, false), */
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_SplitButtonControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_InvokePatternId, UIA.UIA_PatternIds.UIA_ExpandCollapsePatternId}, true),
            /* OMITTING: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_StatusBarControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_GridPatternId}, false), */
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_TabControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_SelectionPatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_TabItemControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_SelectionItemPatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_TableControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_GridPatternId, UIA.UIA_PatternIds.UIA_GridItemPatternId, UIA.UIA_PatternIds.UIA_TablePatternId, UIA.UIA_PatternIds.UIA_TableItemPatternId}, true),
            /* OMITTING: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_TextControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_GridItemPatternId, UIA.UIA_PatternIds.UIA_TableItemPatternId, UIA.UIA_PatternIds.UIA_TextPatternId}, false), */
            /* SUPPRESSED: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ThumbControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_TransformPatternId}, true), */
            /* OMITTING: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ToolBarControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_DockPatternId, UIA.UIA_PatternIds.UIA_ExpandCollapsePatternId, UIA.UIA_PatternIds.UIA_TransformPatternId}, false), */
            /* OMITTING: new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_ToolTipControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_TextPatternId, UIA.UIA_PatternIds.UIA_WindowPatternId}, false), */
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_TreeControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_ScrollPatternId, UIA.UIA_PatternIds.UIA_SelectionPatternId}, false),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_TreeItemControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_ExpandCollapsePatternId}, true),
            new RequiredPatterns(UIA.UIA_ControlTypeIds.UIA_WindowControlTypeId, new int[] {UIA.UIA_PatternIds.UIA_TransformPatternId, UIA.UIA_PatternIds.UIA_WindowPatternId}, true),
        };

        private void UiaControlType_CheckPatterns(UIA.IUIAutomationElement automationElement)
        {
            int controlType = automationElement.CachedControlType;

            if (controlType < UIA.UIA_ControlTypeIds.UIA_ButtonControlTypeId /* First standard control type */ ||
                controlType > UIA.UIA_ControlTypeIds.UIA_SeparatorControlTypeId /* Last standard control type */)
            {
                LogHelper.Log(EventLevel.Error, ElementHasInvalidControlType, this.GetType(), automationElement);
            }

            // If it's custom, verify that it has a localized Control Type, if it's not verify it implements the required patterns
            if (controlType == UIA.UIA_ControlTypeIds.UIA_CustomControlTypeId)
            {
                if (string.IsNullOrEmpty(automationElement.CachedLocalizedControlType))
                {
                    LogHelper.Log(EventLevel.Error, CustomControlTypeWithoutLocalizedControlType, this.GetType(), automationElement);
                }
            }
            else
            {
                // Find the RequiredPatterns in the pattern map
                RequiredPatterns reqPat = null;
                foreach (RequiredPatterns rp in controlTypeToPatternMap)
                {
                    if (rp.controlTypeId == controlType)
                    {
                        reqPat = rp;
                        break;
                    }
                }
                // If the pattern isn't in the table, nothing needs checking
                if (reqPat != null)
                {
                    // if All Patterns are required, verify every pattern listed is supported
                    // if All Patterns aren't required, verify at least one pattern is supported
                    if (reqPat.allAreRequired)
                    {
                        foreach (int patternId in reqPat.patternIds)
                        {
                            object patternObject = automationElement.GetCurrentPattern(patternId);
                            if (patternObject == null)
                            {
                                string patternName = UIAGlobalContext.Automation.GetPatternProgrammaticName(patternId);
                                LogHelper.Log(EventLevel.Error, ControlTypeRequiredPatternNotSupported, this.GetType(), automationElement, automationElement.CachedLocalizedControlType, patternName);
                            }
                        }
                    }
                    else
                    {
                        bool passed = false;
                        StringBuilder patternList = new StringBuilder("{") ;
                        foreach (int patternId in reqPat.patternIds)
                        {
                            object patternObject = automationElement.GetCurrentPattern(patternId);
                            if (patternObject != null)
                            {
                                passed = true;
                                break;
                            }
                            string patternName = UIAGlobalContext.Automation.GetPatternProgrammaticName(patternId);
                            patternList.Append(patternName + ", ");
                        }
                        if (!passed)
                        {
                            // Ditch the ending ", "
                            patternList.Remove(patternList.Length - 2, 2);
                            patternList.Append("}");
                            LogHelper.Log(EventLevel.Error, ControlTypeNoExpectedPatternsSupported, this.GetType(), automationElement, automationElement.CachedLocalizedControlType, patternList.ToString());
                        }
                    }
                }
            }
        }


        private void TraverseFromNode(UIA.IUIAutomationElement automationElement)
        {
            UiaControlType_CheckPatterns(automationElement);
             
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
