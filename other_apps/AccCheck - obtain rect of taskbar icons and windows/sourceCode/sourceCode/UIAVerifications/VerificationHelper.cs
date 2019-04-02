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
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace UIAVerifications
{
    public static class VerificationHelper
    {
        static private object[] _acceptableCommonErrorCodes = new object[] { Win32API.E_FAIL, Win32API.UIA_E_ELEMENTNOTAVAILABLE, typeof(ArgumentException) };

        /// -------------------------------------------------------------------
        /// <summary>Get the array of valid error codes returned that we can 
        /// ignore at times.</summary>
        /// -------------------------------------------------------------------
        public static object[] AcceptableCommonErrorCodes
        {
            get { return _acceptableCommonErrorCodes; }
        }
        
 
        // Call IsCriticalException w/in a catch-all-exception handler to allow critical exceptions
        // to be thrown (this is copied from exception handling code in WinForms but feel free to
        // add new critical exceptions).  Usage:
        //      try 
        //      {
        //          Somecode(); 
        //      }
        //      catch (Exception e) 
        //      { 
        //          if (IsCriticalException(e)) 
        //              throw; 
        //          // ignore non-critical errors from external code
        //      }
        public static bool IsCriticalException(Exception e)
        {
            return e is SEHException || e is NullReferenceException || e is StackOverflowException || e is OutOfMemoryException || e is System.Threading.ThreadAbortException;
        }
  
    }
}

