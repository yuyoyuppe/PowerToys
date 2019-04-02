// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Text;
using System.Runtime.InteropServices;
using Accessibility;

namespace AccCheckUI
{
    internal static class Win32API
    {
        //
        // Primitives
        //
        [DllImport("user32.dll", ExactSpelling = true)]
        [return: MarshalAs(UnmanagedType.Bool)] internal static extern bool SetProcessDPIAware();

        internal static bool IsBitSet(int flags, int bit)
        {
            return (flags & bit) == bit;
        }

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

        [StructLayout(LayoutKind.Sequential)]
        internal struct POINT
        {
            public int x;
            public int y;

            public POINT(int x, int y)
            {
                this.x = x;
                this.y = y;
            }
        }

        internal const int GW_HWNDNEXT = 2;
        internal const int GW_CHILD = 5;
        internal const int WM_CLOSE = 0x0010;

        public const uint WS_VISIBLE                = 0x10000000;
        
        // GetWindowLong()
        internal const int GWL_EXSTYLE = (-20);
        internal const int GWL_STYLE = (-16);

        internal const int WS_EX_TOOLWINDOW     = 0x00000080;

        internal static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        
        internal const int SWP_NOACTIVATE = 0x0010;
        
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);
        
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError=true)]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool SetWindowPos(IntPtr hWnd, IntPtr hwndAfter, int x, int y, int width, int height, int flags);
        
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool GetCursorPos([In, Out] ref POINT pt);
        
        [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll")]
        internal static extern bool UpdateWindow(IntPtr hwnd);

        public const uint SPI_GETMESSAGEDURATION          = 0x2016;
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        internal static extern bool SystemParametersInfo(uint uiAction, uint uiParam, out int pvParam, UInt32 fWinIni);
        
        internal const int SW_HIDE = 0;
        internal const int SW_SHOWNORMAL = 1;
        internal const int SW_NORMAL = 1;
        internal const int SW_SHOWMINIMIZED = 2;
        internal const int SW_SHOWMAXIMIZED = 3;
        internal const int SW_MAXIMIZE = 3;
        internal const int SW_SHOWNOACTIVATE = 4;
        internal const int SW_SHOW = 5;
        internal const int SW_MINIMIZE = 6;
        internal const int SW_SHOWMINNOACTIVE = 7;
        internal const int SW_SHOWNA = 8;
        internal const int SW_RESTORE = 9;
        
        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public POINT ptMinPosition;
            public POINT ptMaxPosition;
            public RECT rcNormalPosition;
        }

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool GetWindowPlacement(IntPtr hwnd, ref WINDOWPLACEMENT place);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool SetWindowPlacement(IntPtr hwnd, ref WINDOWPLACEMENT place);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError=true)]
        internal static extern IntPtr WindowFromPoint(POINT pt);

        internal const int GA_PARENT = 1;
        internal const int GA_ROOT = 2;
        internal const int GA_ROOTOWNER = 3;
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr GetAncestor(IntPtr hwnd, int gaFlags);

        #region Hotkeys
        [DllImport("user32", SetLastError = true)]
        internal static extern int RegisterHotKey(IntPtr hwnd, int id, int fsModifiers, int vk);
        [DllImport("user32", SetLastError = true)]
        internal static extern int UnregisterHotKey(IntPtr hwnd, int id);

        internal const int MOD_ALT = 1;
        internal const int MOD_CONTROL = 2;
        internal const int MOD_SHIFT = 4;
        internal const int MOD_WIN = 8;
        #endregion

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool GetWindowRect(IntPtr hwnd, ref RECT rc);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int nMax);

        internal const int MAX_PATH = 260;
        internal static string GetWindowText(IntPtr hWnd)
        {
            System.Text.StringBuilder windowName = new System.Text.StringBuilder(MAX_PATH + 1);
            GetWindowText(hWnd, windowName, MAX_PATH);
            return windowName.ToString();
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool IsWindow(IntPtr hwnd);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint dwProcessId);
        
        [DllImport("kernel32.dll")]
        internal static extern int GetCurrentProcessId();

        #region Windows
        // Enum windows delegate
        internal delegate bool WinEnumProc(IntPtr hwnd, ref IntPtr lParam);

        // Enum windows delegate
        internal delegate bool EnumDesktopWindowsProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32.dll")]
        internal static extern bool EnumDesktopWindows(IntPtr hWndParent, EnumDesktopWindowsProc enumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        internal static extern bool EnumChildWindows(IntPtr hWndParent, WinEnumProc enumFunc, ref IntPtr lParam);

        #endregion

        #region Accessibility
        internal const int SELFLAG_NONE = 0x00000000;
        internal const int SELFLAG_TAKEFOCUS = 0x00000001;
        internal const int SELFLAG_TAKESELECTION = 0x00000002;
        internal const int SELFLAG_EXTENDSELECTION	 = 0x00000004;
        internal const int SELFLAG_ADDSELECTION = 0x00000008;
        internal const int SELFLAG_REMOVESELECTION	 = 0x00000010;
        internal const int SELFLAG_VALID = 0x0000001F;
        #endregion



    }
}
