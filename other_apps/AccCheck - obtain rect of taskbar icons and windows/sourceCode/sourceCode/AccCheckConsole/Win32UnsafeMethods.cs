/***************************************************************************\
*
* File: CommonWin32API.cs
*
* Description:
* Win32API imports
*
* Copyright (C) 2002 by Microsoft Corporation.  All rights reserved.
*
\***************************************************************************/
using System;
using System.Text;
using System.Runtime.InteropServices;

namespace AccCheckConsole
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
        
        internal const int SRCCOPY = 0x00CC0020;

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern IntPtr GetDC(IntPtr hwnd);

        // Simplified version of ReleaseDC
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern IntPtr ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool GetWindowRect(IntPtr hwnd, ref RECT rc);
        
        [DllImport("gdi32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest, int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, int dwRop);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool IsWindow(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string className, string wndName);
    }
}
