// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Text;
using System.Drawing;
using UIA=UIAutomationClient;
using System.Runtime.InteropServices;

namespace UIAVerifications
{
    internal static class Win32API
    {
        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            public int left;
            public int top;
            public int right;
            public int bottom;

            public RECT(int left, int top, int right, int bottom)
            {
                this.left = left;
                this.top = top;
                this.right = right;
                this.bottom = bottom;
            }

            public RECT(RECT rcSrc)
            {
                this.left = rcSrc.left;
                this.top = rcSrc.top;
                this.right = rcSrc.right;
                this.bottom = rcSrc.bottom;
            }

            // Handy method for converting to a System.Drawing.Rectangle
            public Rectangle ToRectangle()
            { 
                return Rectangle.FromLTRB(left, top, right, bottom);
            }

            public static RECT FromRectangle(Rectangle rectangle)
            {
                return new RECT(rectangle.Left, rectangle.Top, rectangle.Right, rectangle.Bottom);
            }

            public static RECT FromtagRect(UIA.tagRECT rectangle)
            {
                return new RECT(rectangle.left, rectangle.top, rectangle.right, rectangle.bottom);
            }
            
            public UIA.tagRECT TotagRect()
            { 
                UIA.tagRECT rect = new UIA.tagRECT();
                rect.left = left;
                rect.top = top;
                rect.right = right;
                rect.bottom = bottom;
                return rect;
            }

            public static implicit operator Rectangle(RECT rect)
            {
                return rect.ToRectangle();
            }

            public static implicit operator RECT(Rectangle rect)
            {
                return FromRectangle(rect);
            }
            
            public static implicit operator UIA.tagRECT(RECT rect)
            {
                return rect.TotagRect();
            }

            public static implicit operator RECT(UIA.tagRECT rect)
            {
                return FromtagRect(rect);
            }
            
            public bool IsEmpty
            {
                get
                {
                    return left >= right || top >= bottom;
                }
            }
            static public RECT Empty
            {
                get
                {
                    return new RECT(0, 0, 0, 0);
                }
            }
        }
        
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool GetWindowRect(IntPtr hwnd, ref RECT rc);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool IsWindow(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string className, string wndName);


        internal const int E_FAIL = unchecked((int)0x80004005);                   // 2147500037
        internal const int UIA_E_ELEMENTNOTAVAILABLE = unchecked((int)0x80040201);


        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct GUITHREADINFO
        {
            public int cbSize;
            public int dwFlags;
            public IntPtr hwndActive;
            public IntPtr hwndFocus;
            public IntPtr hwndCapture;
            public IntPtr hwndMenuOwner;
            public IntPtr hwndMoveSize;
            public IntPtr hwndCaret;
            public RECT rc;
        }

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool GetGUIThreadInfo(int idThread, ref GUITHREADINFO guiThreadInfo);

        internal const int GA_PARENT = 1;
        internal const int GA_ROOT = 2;
        internal const int GA_ROOTOWNER = 3;

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr GetAncestor(IntPtr hwnd, int gaFlags);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint dwProcessId);

        [DllImport("kernel32.dll")]
        internal static extern int GetCurrentProcessId();


        #region Input

        //
        // INPUT consts and structs 
        //
        internal const int VK_RSHIFT = 0xA1;
        internal const int VK_TAB = 0x09;

        // extended Keys        
        internal const int VK_RMENU = 0xA5;
        internal const int VK_RCONTROL = 0xA3;
        internal const int VK_NUMLOCK = 0x90;
        internal const int VK_INSERT = 0x2D;
        internal const int VK_DELETE = 0x2E;
        internal const int VK_HOME = 0x24;
        internal const int VK_PRIOR = 0x21;
        internal const int VK_NEXT = 0x22;
        internal const int VK_END = 0x23;
        internal const int VK_LEFT = 0x25;
        internal const int VK_UP = 0x26;
        internal const int VK_RIGHT = 0x27;
        internal const int VK_DOWN = 0x28;
        internal const int VK_LWIN = 0x5B;
        internal const int VK_RWIN = 0x5C;
        internal const int VK_APPS = 0x5D;


        internal const int KEYEVENTF_EXTENDEDKEY = 0x0001;
        internal const int KEYEVENTF_KEYUP = 0x0002;
        internal const int KEYEVENTF_SCANCODE = 0x0008;
        internal const int MOUSEEVENTF_VIRTUALDESK = 0x4000;
        internal const int MOUSEEVENTF_MOVE = 0x0001;
        internal const int MOUSEEVENTF_ABSOLUTE = 0x08000;

        internal const int INPUT_MOUSE = 0;
        internal const int INPUT_KEYBOARD = 1;

        [StructLayout(LayoutKind.Sequential)]
        internal struct INPUT
        {
            public int type;
            public INPUTUNION union;
        };

        [StructLayout(LayoutKind.Explicit)]
        internal struct INPUTUNION
        {
            [FieldOffset(0)]
            public MOUSEINPUT mouseInput;
            [FieldOffset(0)]
            public KEYBDINPUT keyboardInput;
        };

        [StructLayout(LayoutKind.Sequential)]
        internal struct MOUSEINPUT
        {
            public int dx;
            public int dy;
            public int mouseData;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        };

        [StructLayout(LayoutKind.Sequential)]
        internal struct KEYBDINPUT
        {
            public short wVk;
            public short wScan;
            public int dwFlags;
            public int time;
            public IntPtr dwExtraInfo;
        };

        // This declaration is good for only one INPUT passed in.
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern uint SendInput(uint nInputs, ref INPUT pInputs, int cbSize);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern int MapVirtualKey(int nVirtKey, int nMapType);
        #endregion

        internal static void SendKeyboardInput(int key, bool press)
        {
            Win32API.INPUT ki = new Win32API.INPUT();
            ki.type = Win32API.INPUT_KEYBOARD;
            ki.union.keyboardInput.wVk = (short)key;
            ki.union.keyboardInput.wScan = (short)Win32API.MapVirtualKey(ki.union.keyboardInput.wVk, 0);
            int dwFlags = 0;
            if (ki.union.keyboardInput.wScan > 0)
                dwFlags |= Win32API.KEYEVENTF_SCANCODE;
            if (!press)
                dwFlags |= Win32API.KEYEVENTF_KEYUP;
            ki.union.keyboardInput.dwFlags = dwFlags;
            if (IsExtendedKey(key))
            {
                ki.union.keyboardInput.dwFlags |= Win32API.KEYEVENTF_EXTENDEDKEY;
            }
            ki.union.keyboardInput.time = 0;
            ki.union.keyboardInput.dwExtraInfo = IntPtr.Zero;

            Win32API.SendInput(1, ref ki, Marshal.SizeOf(ki));
        }

        private static bool IsExtendedKey(int key)
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

    }
}
