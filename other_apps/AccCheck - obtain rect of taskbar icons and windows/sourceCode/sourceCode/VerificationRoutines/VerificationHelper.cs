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
    static class ClassificationsGroup
    {
        public const string Others = "Other";
        public const string Properties = "Properties";
        public const string TreeStructure = "Tree";
        public const string Navigation = "Navigation";
        public const string Focus = "Focus";
    }
    
    public static class VerificationHelper
    {
        internal const string TYPE_EXCEPTION = "exception";
        internal const string TYPE_CHILDID = "int";
        internal const string TYPE_IACCESSIBLE = "__comobject";
        internal const string TYPE_HRESULT = "int";
        internal const string TYPE_STRING = "string";
        internal const string TYPE_NULL = "null";

        static private object[] _acceptableCommonErrorCodes = new object[] { Win32API.DISP_E_MEMBERNOTFOUND, Win32API.E_NOTIMPL, Win32API.E_FAIL, typeof(ArgumentException) };

        /// -------------------------------------------------------------------
        /// <summary>Get the array of valid error codes returned that we can 
        /// ignore at times.</summary>
        /// -------------------------------------------------------------------
        public static object[] AcceptableCommonErrorCodes
        {
            get { return _acceptableCommonErrorCodes; }
        }
        
        public static void SendKeyboardInput( int key, bool press )
        {
            Win32API.INPUT ki = new Win32API.INPUT();
            ki.type = Win32API.INPUT_KEYBOARD;
            ki.union.keyboardInput.wVk = (short) key;
            ki.union.keyboardInput.wScan = (short) Win32API.MapVirtualKey( ki.union.keyboardInput.wVk, 0 );
            int dwFlags = 0;
            if( ki.union.keyboardInput.wScan > 0 )
                dwFlags |= Win32API.KEYEVENTF_SCANCODE;
            if( !press )
                dwFlags |= Win32API.KEYEVENTF_KEYUP;
            ki.union.keyboardInput.dwFlags = dwFlags;
            if( IsExtendedKey( key ) )
            {
                ki.union.keyboardInput.dwFlags |= Win32API.KEYEVENTF_EXTENDEDKEY;
            }
            ki.union.keyboardInput.time = 0;
            ki.union.keyboardInput.dwExtraInfo = IntPtr.Zero;

            Win32API.SendInput(1, ref ki, Marshal.SizeOf(ki));
        }

        private static bool IsExtendedKey( int key )
        {
            // From the SDK:
            // The extended-key flag indicates whether the keystroke message originated from one of
            // the additional keys on the enhanced keyboard. The extended keys consist of the ALT and
            // CTRL keys on the right-hand side of the keyboard; the INS, DEL, HOME, END, PAGE UP,
            // PAGE DOWN, and arrow keys in the clusters to the left of the numeric keypad; the NUM LOCK
            // key; the BREAK (CTRL+PAUSE) key; the PRINT SCRN key; and the divide (/) and ENTER keys in
            // the numeric keypad. The extended-key flag is set if the key is an extended key. 
            //
            // - docs appear to be incorrect. Use of Spy++ indicates that break is not an extended key.
            // Also, menu key and windows keys also appear to be extended.
            return key == Win32API.VK_RMENU
                || key == Win32API.VK_RCONTROL
                || key == Win32API.VK_NUMLOCK
                || key == Win32API.VK_INSERT
                || key == Win32API.VK_DELETE
                || key == Win32API.VK_HOME
                || key == Win32API.VK_END
                || key == Win32API.VK_PRIOR
                || key == Win32API.VK_NEXT
                || key == Win32API.VK_UP
                || key == Win32API.VK_DOWN
                || key == Win32API.VK_LEFT
                || key == Win32API.VK_RIGHT
                || key == Win32API.VK_APPS
                || key == Win32API.VK_RWIN
                || key == Win32API.VK_LWIN;

            // Note that there are no distinct values for the following keys:
            // numpad divide
            // numpad enter
        }

        /// -------------------------------------------------------------------
        /// <summary>
        /// Helper skipping errors in name role and state verifications
        /// </summary>
        public static bool ShouldIgnoreElement(MsaaElement element, Type routine)
        {
            // Client is the deafult role and is almost alway is not important
            AccRole role = element.GetRole();
            if (role == AccRole.Client)
            {
                return true;
            }

            // The user does usually interact with a window but with its children
            // the only eception is top level windows like the notepad window.
            if (role == AccRole.Window)
            {
                if (Win32API.GetParent(element.GetHwnd()) != Win32API.GetDesktopWindow())
                {
                    return true;
                }
            }

            if (routine == typeof(CheckName))
            {
                if (role == AccRole.Grip || role == AccRole.PageTabList)
                {
                    return true;
                }
            }

            return false;
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

