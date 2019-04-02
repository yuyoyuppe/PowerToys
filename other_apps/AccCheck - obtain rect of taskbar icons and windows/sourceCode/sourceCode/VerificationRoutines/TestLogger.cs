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
using System.Reflection;
using Accessibility;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace VerificationRoutines
{

    /// -----------------------------------------------------------------------
    /// <summary>Attribute that will describe the test methods</summary>
    /// -----------------------------------------------------------------------
    public class TestMethodAttribute : Attribute
    {
        string _description = string.Empty;
        string _testMethod = string.Empty;

        public string TestMethod
        {
            get { return _testMethod; }
            set { _testMethod = value; }
        }

        public string Description
        {
            get { return _description; }
            set { _description = value; }
        }

        public TestMethodAttribute() { }

        public TestMethodAttribute(string description)
        {
            _description = description;
        }
    }

    /// -----------------------------------------------------------------------
    /// <summary>Class the wraps a lot of the logging functionality for the 
    /// verification classes</summary>
    /// -----------------------------------------------------------------------
    static class TestLogger
    {
        static ILogger _logger = null;
        static IntPtr _rootHwnd = IntPtr.Zero;
        static object locker = new object();  // Use to stop reentry into the logging code

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

        internal static void Log(EventLevel eventLevel, LogMessage message, Type routine, Accessible accessible, params object[] formatArgs)
        {
            Debug.Assert(_logger != null);

            MsaaElement msaaElement = MsaaElement.FindMsaaElement(accessible);
            if (msaaElement != null)
            {
                TestLogger.CommonLog(eventLevel, message, "", routine, msaaElement, formatArgs);
            }
            else
            {
                TestLogger.CommonLog(eventLevel, message, "", routine, accessible, formatArgs);
            }

        }

        /// -------------------------------------------------------------------
        /// <summary>Log</summary>
        /// -------------------------------------------------------------------
        internal static void Log(EventLevel eventLevel, LogMessage message, Type routine, MsaaElement element, params object[] formatArgs)
        {
            TestLogger.CommonLog(eventLevel, message, "", routine, element, formatArgs);
        }

        /// -------------------------------------------------------------------
        /// <summary>LogException</summary>
        /// -------------------------------------------------------------------
        internal static void LogException(EventLevel eventLevel, LogMessage message, Type routine, Exception exception, object[] filterErrorCodes, Accessible accessible, params object[] formatArgs)
        {
            Debug.Assert(_logger != null);

            if (FilterThisException(exception, filterErrorCodes))
            {
                return;
            }
            message.Description = string.Format("{0} : {1}",  message.Description, exception.GetType().ToString());

            if (exception is COMException)
            {
                message.Description += string.Format(":ErrorCode = {0}", ((COMException)exception).ErrorCode);
            }

            TestLogger.CommonLog(eventLevel, message, exception.StackTrace, routine, accessible, formatArgs);
        }

        /// -------------------------------------------------------------------
        /// <summary>LogException</summary>
        /// -------------------------------------------------------------------
        internal static void LogException(EventLevel eventLevel, LogMessage message, Type routine, Exception exception, object[] filterErrorCodes, MsaaElement element, params object[] formatArgs)
        {
            Debug.Assert(_logger != null);

            if (FilterThisException(exception, filterErrorCodes))
            {
                return;
            }

            // If 'PREVIOUSTHROW' then this exception actually occurred earlier while building the tree and caching the properties and not
            // in the verification routine. When reporting the exception in the verification routine, the stack trace to the original throw 
            // is confusing when debugging the issue.  Beacause of this, the verification routine is throwing a new exception (w/PREVIOUSTHROW) 
            // to get the verification routine stack trace, and setting its inner exception to the original exception. Need to use the 
            // correct message of the appropriate exception.
            Exception messageException; // = exception.Message.StartsWith("PREVIOUSTHROW") ? exception.InnerException : exception;

            if (exception.Message.StartsWith("PREVIOUSTHROW"))
            {
                messageException = exception.InnerException;
                message.Description = string.Format("Unexpected exception was thrown while calling {0}. {1} : {2}", 
                    exception.Message.Substring("PREVIOUSTHROW".Length), message.Description, messageException.Message);
            }
            else
            {
                messageException = exception;
                message.Description = string.Format("{0} : {1}", message.Description, messageException.Message);
            }

            if (messageException is COMException)
            {
                message.Description += string.Format(":ErrorCode = {0}", ((COMException)messageException).ErrorCode);
            }
            TestLogger.CommonLog(eventLevel, message, exception.StackTrace, routine, element, formatArgs);
        }

        /// -------------------------------------------------------------------
        /// <summary>LogException</summary>
        /// -------------------------------------------------------------------
        internal static void LogTestException(EventLevel eventLevel, Type routine, Exception exception)
        {
            TestLogger.CommonLog(eventLevel, TestLogger.UNEXPECTED_TEST_ERROR, exception.StackTrace, routine, (MsaaElement)null, exception.Message);
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
        private static void CommonLog(EventLevel eventLevel, LogMessage message, string extraText, Type routine, MsaaElement element, params object[] formatArgs)
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

            if (element == null)
            {
                element = MsaaElement.RootElement;
            }

            LogEvent le = null;
            if (element == null)
            {
                // Something drastic happened, can't use the element.
                le = new LogEvent(eventLevel, message, text, extraText, Rectangle.Empty, routine);
            }
            else
            {
                le = new LogEvent(eventLevel, message, text, extraText, element.GetBoundingBox(), routine);
                le.ParentChain = element.CreateParentChain();
                le.SetFromWindow(element.GetHwnd(), _rootHwnd);
                le.SetIAccessibleProps(element.Accessible, element.GetName(), element.GetValue(), element.GetRole().ToString(), element.GetStatesString());
                
                IQueryString queryString = VerificationManager.GetQueryStringObject();
                if (queryString != null)
                {
                    le.QueryString = queryString.Generate(element.Accessible.IAccessible, element.Accessible.ChildId);
                    le.ParentQueryString = le.QueryString.Substring(0, le.QueryString.LastIndexOf(le.QueryString[0]));
                }
            }
            _logger.Log(le);
        }
        
        private static void CommonLog(EventLevel eventLevel, LogMessage message, string extraText, Type routine, Accessible accessible, params object[] formatArgs)
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

                    Debug.Assert(false, string.Format("Error with string.Format(\"{0}\", and formatArgs = {1}) : ", text, string.Join(", ", array), e.Message));
                }
            }

            if (accessible == null)
            {
                // if we don't have an accessible then just log the info we have.
                LogEvent logEvent = new LogEvent(eventLevel, message, text, extraText, Rectangle.Empty, routine);
                _logger.Log(logEvent);
                return;
            }
            

            Rectangle loc;
            try 
            {
                loc = accessible.Location;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    loc = Rectangle.Empty;
                }
                else
                {
                    throw;
                }
            }
            
            LogEvent le = new LogEvent(eventLevel, message, text, extraText, loc, routine);

            try
            {
                le.ParentChain = accessible.CreateParentChain();
            }
            catch (Exception e)
            {
                le.ParentChain = "Call failed to CreateParentChain with " + e.Message;
            }

            le.SetFromWindow(accessible.Hwnd, _rootHwnd);

            string name = "";
            try 
            {
                name = accessible.Name;
                if (name == null)
                    name = "";
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    name = "";
                }
                else
                {
                    throw;
                }
            }

            int role = 0;
            try 
            {
                role = accessible.Role;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    role = 0;
                }
                else
                {
                    throw;
                }
            }
            
            int state = 0;
            try 
            {
                state = accessible.State;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    state = 0;
                }
                else
                {
                    throw;
                }
            }
            
            string value = "";
            try 
            {
                value = accessible.Value;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    value = "";
                }
                else
                {
                    throw;
                }
            }

            le.SetIAccessibleProps(accessible, name, value, MsaaElement.LookupRole(role).ToString(), MsaaElement.GetStatesString(state));

            _logger.Log(le);
        }
    }
}
