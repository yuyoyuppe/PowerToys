// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.IO;
using System.Reflection;
using System.Collections.Generic;
using System.Text;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;
using System.Drawing;
using System.Diagnostics;
using System.Runtime.InteropServices;
using UIA=UIAutomationClient;

namespace UIAVerifications
{


    /// -----------------------------------------------------------------------
    /// <summary>Class the wraps a lot of the logging functionality for the 
    /// verification classes</summary>
    /// -----------------------------------------------------------------------
    static class LogHelper
    {
        static ILogger _logger = null;
        static IntPtr _rootHwnd = IntPtr.Zero;
        
        /// <summary>Exception :({0})</summary>
        static LogMessage UNEXPECTED_TEST_ERROR = new LogMessage("Exception :({0})", "UNEXPECTEDTESTERROR");

        /// -------------------------------------------------------------------
        /// <summary>Global logger</summary>
        /// -------------------------------------------------------------------
        public static ILogger Logger
        {
            set { _logger = value; }
        }

        public static IntPtr RootHwnd
        {
            set { _rootHwnd = value; }
        }

        internal static void Log(EventLevel eventLevel, LogMessage message, Type routine, UIA.IUIAutomationElement automationElement, params object[] formatArgs)
        {
            Debug.Assert(_logger != null);

            LogHelper.CommonLog(eventLevel, message, "", routine, automationElement, formatArgs);

        }

        /// -------------------------------------------------------------------
        /// <summary>LogException</summary>
        /// -------------------------------------------------------------------
        internal static void LogException(EventLevel eventLevel, LogMessage message, Type routine, Exception exception, object[] filterErrorCodes, UIA.IUIAutomationElement automationElement, params object[] formatArgs)
        {
            Debug.Assert(_logger != null);

            if (FilterThisException(exception, filterErrorCodes))
            {
                return;
            }

            Exception messageException = exception;
            message.Description = string.Format("{0} : {1}", message.Description, messageException.Message);

            if (messageException is COMException)
            {
                message.Description += string.Format(":ErrorCode = {0}", ((COMException)messageException).ErrorCode);
            }
            
            LogHelper.CommonLog(eventLevel, message, exception.StackTrace, routine, automationElement, formatArgs);
        }

        /// -------------------------------------------------------------------
        /// <summary>LogException</summary>
        /// -------------------------------------------------------------------
        internal static void LogTestException(EventLevel eventLevel, Type routine, Exception exception)
        {
            LogHelper.CommonLog(eventLevel, LogHelper.UNEXPECTED_TEST_ERROR, exception.StackTrace, routine, null, exception.Message);
        }

        /// -------------------------------------------------------------------
        /// <summary>FilterThisException</summary>
        /// -------------------------------------------------------------------
        static bool FilterThisException(Exception exception, object[] filterErrorCodes)
        {
            bool filter = false;

            if (filterErrorCodes == null)
            {
                filter = false;
            }
            else if (exception is COMException)
            {
                COMException ce = (COMException)exception;
                filter = Array.IndexOf(filterErrorCodes, ce.ErrorCode) != -1;
            }
            else
            {
                filter = Array.IndexOf(filterErrorCodes, exception.GetType()) != -1;
            }
            return filter;
        }

        /// -------------------------------------------------------------------
        /// <summary>CommonLog</summary>
        /// -------------------------------------------------------------------
        private static void CommonLog(EventLevel eventLevel, LogMessage message, string extraText, Type routine, UIA.IUIAutomationElement automationElement, params object[] formatArgs)
        {
            Debug.Assert(_logger != null);

            string text = message.Description;

            if (formatArgs != null)
            {
                try
                {
                    text = string.Format(text, formatArgs);
                }
                catch (FormatException e)
                {
                    // Convert the object array to a string array needed for Join
                    string[] array = Array.ConvertAll<object, string>(formatArgs, delegate(object o) { return "\"" + o.ToString() + "\""; });

                    Debug.Assert(false, string.Format("Error with string.Format(\"{0}\", and formatArgs = {1}) : ",
                        text, string.Join(", ", array), e.Message));

                }
            }

            if (automationElement == null)
            {
                automationElement = UIAGlobalContext.GetRootElement();
            }

            LogEvent le = null;
            if (automationElement == null)
            {
                // Something drastic happened, can't use the element.
                le = new LogEvent(eventLevel, message, text, extraText, Rectangle.Empty, routine);
            }
            else
            {
                Win32API.RECT boundingRectangle = automationElement.CachedBoundingRectangle;
                
                le = new LogEvent(eventLevel, message, text, extraText, boundingRectangle, routine);
                le.ParentChain = CreateParentChain(automationElement);
                le.SetFromWindow((IntPtr)automationElement.CurrentNativeWindowHandle, _rootHwnd);
                string value = UIAGlobalContext.GetElementValue(automationElement);

                le.SetUIAProps(automationElement, 
                               automationElement.CachedName, 
                               value, automationElement.CachedLocalizedControlType, 
                               UIAGlobalContext.GetElementState(automationElement),
                               automationElement.CurrentFrameworkId,
                               automationElement.CurrentProviderDescription);

                IQueryString queryString = VerificationManager.GetQueryStringObject();
                if (queryString != null)
                {
                    le.QueryString = queryString.Generate(automationElement, null);
                    le.ParentQueryString = le.QueryString.Substring(0, le.QueryString.LastIndexOf(le.QueryString[0]));
                }
            }
            _logger.Log(le);
        }

        public static string CreateParentChain(UIA.IUIAutomationElement automationElement)
        {
            List<string> parentChain = new List<string>();

            UIA.IUIAutomationTreeWalker treeWalker = UIAGlobalContext.Automation.ControlViewWalker;
            
            UIA.IUIAutomationElement element = automationElement;
            while (element != null && UIAGlobalContext.Automation.CompareElements(element, UIAGlobalContext.GetRootElement()) == 0)
            {
                String name = element.CurrentName;
                
                if (!String.IsNullOrEmpty(name))
                {
                    parentChain.Add(name);
                }
        
                element = treeWalker.GetParentElement(element);
            }
            
            parentChain.Reverse();
            
            return string.Join(".", parentChain.ToArray());
        }
        
    }
}
