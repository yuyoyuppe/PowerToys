// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Data;
using System.Text;
using System.Runtime.InteropServices;
using System.ComponentModel;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;
using UIA=UIAutomationClient;


namespace UIAVerifications
{
    public class UIAGlobalCache : ICache
    {
        public void Clear()
        {
            try
            {
                UIAGlobalContext.ClearCache();   
            }
            catch (Exception)
            {
                // If uia fails to load we get this.  They already got a message that something is wrong when it failed to load.
            }
        }
    }

    public static class UIAGlobalContext 
    {
        private static UIA.IUIAutomation _automation = null;
        private static UIA.IUIAutomationElement _automationRootElement = null;
        private static UIA.IUIAutomationCacheRequest _cacheRequest = null;
    
        public static void ClearCache()
        {
            _automationRootElement = null;
        }

        public static UIA.IUIAutomation Automation
        {
            get
            {
                if (_automation == null)
                {
                    _automation = new UIA.CUIAutomation();
                }
                return _automation;
            }
        }

        public static UIA.IUIAutomationElement CacheUIATree(IntPtr hwnd)
        {
            BuildCacheRequest();
            _automationRootElement = Automation.ElementFromHandleBuildCache(hwnd, _cacheRequest);

            return _automationRootElement;
        }

        public static UIA.IUIAutomationElement LocalRootElement
        {
            get
            {
                return _automationRootElement;
            }
        }

        public static UIA.IUIAutomationElement GetRootElement()
        {
            BuildCacheRequest();
            UIA.IUIAutomationCacheRequest rootCacheRequest = _cacheRequest.Clone();
            rootCacheRequest.TreeScope = UIA.TreeScope.TreeScope_Element;
            return Automation.GetRootElementBuildCache(rootCacheRequest);
        }

        public static UIA.IUIAutomationElement GetFocusedElement()
        {
            BuildCacheRequest();
            UIA.IUIAutomationCacheRequest rootCacheRequest = _cacheRequest.Clone();
            rootCacheRequest.TreeScope = UIA.TreeScope.TreeScope_Element;
            return Automation.GetFocusedElementBuildCache(rootCacheRequest);
        }

        public static UIA.IUIAutomationElement ElementFromPoint(int x, int y)
        {
            BuildCacheRequest();
            UIA.IUIAutomationCacheRequest rootCacheRequest = _cacheRequest.Clone();
            rootCacheRequest.TreeScope = UIA.TreeScope.TreeScope_Element;

            UIA.tagPOINT pt = new UIA.tagPOINT();
            pt.x = x;
            pt.y = y;
            return Automation.ElementFromPointBuildCache(pt, rootCacheRequest);
        }

        public static UIA.IUIAutomationElement ElementFromHandle(IntPtr hwnd)
        {
            BuildCacheRequest();
            UIA.IUIAutomationCacheRequest rootCacheRequest = _cacheRequest.Clone();
            rootCacheRequest.TreeScope = UIA.TreeScope.TreeScope_Element;

            UIA.IUIAutomationElement element = null;
            try
            {
                element = Automation.ElementFromHandleBuildCache(hwnd, rootCacheRequest);
            }
            catch (COMException )
            {
                // ignore comexception and just return null                
            }

            return element;
        }


        private static void BuildCacheRequest()
        {
            if (_cacheRequest == null)
            {
                _cacheRequest = Automation.CreateCacheRequest();
                
                _cacheRequest.TreeScope = UIA.TreeScope.TreeScope_Subtree;
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_NamePropertyId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_LocalizedControlTypePropertyId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_BoundingRectanglePropertyId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_ControlTypePropertyId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_IsOffscreenPropertyId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_IsEnabledPropertyId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_IsKeyboardFocusablePropertyId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_AccessKeyPropertyId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_RuntimeIdPropertyId);
                
                _cacheRequest.AddPattern(UIA.UIA_PatternIds.UIA_ValuePatternId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_ValueValuePropertyId);

                _cacheRequest.AddPattern(UIA.UIA_PatternIds.UIA_RangeValuePatternId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_RangeValueValuePropertyId);

                _cacheRequest.AddPattern(UIA.UIA_PatternIds.UIA_TogglePatternId);
                _cacheRequest.AddProperty(UIA.UIA_PropertyIds.UIA_ToggleToggleStatePropertyId);

                _cacheRequest.TreeFilter = Automation.RawViewCondition;
            }
        }

        public static string GetElementState(UIA.IUIAutomationElement element)
        {
            string enabled = element.CachedIsEnabled != 0 ? "Enabled" : "Disabled";
            string onscreen = element.CachedIsOffscreen != 0 ? "Offscreen" : "Onscreen";
            string focusable = element.CachedIsKeyboardFocusable != 0 ? "Focusable" : "Unfocusable";
            return String.Format("{0}, {1}, {2}", enabled, onscreen, focusable);
        }

        public static string GetElementValue(UIA.IUIAutomationElement element)
        {
            string rValue = "";
            UIA.IUIAutomationValuePattern valuePattern = element.GetCachedPattern(UIA.UIA_PatternIds.UIA_ValuePatternId) as UIA.IUIAutomationValuePattern;
            if (valuePattern != null)
            {
                rValue = valuePattern.CachedValue;
            }
            else
            {
                UIA.IUIAutomationRangeValuePattern rvaluePattern = element.GetCachedPattern(UIA.UIA_PatternIds.UIA_RangeValuePatternId) as UIA.IUIAutomationRangeValuePattern;
                if (rvaluePattern != null)
                {
                    rValue = rvaluePattern.CachedValue.ToString();
                }
            }
            return rValue;
        }
    }
}

