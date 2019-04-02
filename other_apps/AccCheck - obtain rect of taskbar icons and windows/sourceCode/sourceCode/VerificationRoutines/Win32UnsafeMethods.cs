// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Text;
using System.Runtime.InteropServices;
using Accessibility;

namespace VerificationRoutines
{
    internal static class Win32API
    {
        //
        // Primitives
        //
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

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError=true)]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern bool GetCursorPos([In, Out] ref POINT pt);
        
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        internal extern static bool IsChild(IntPtr parent, IntPtr child);
        
        internal const int GA_PARENT = 1;
        internal const int GA_ROOT = 2;
        internal const int GA_ROOTOWNER = 3;
        
        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        internal static extern IntPtr GetAncestor(IntPtr hwnd, int gaFlags);

        [StructLayout(LayoutKind.Sequential, CharSet=CharSet.Unicode)]
        public struct GUITHREADINFO 
        {
            public int      cbSize;
            public int      dwFlags;
            public IntPtr   hwndActive;
            public IntPtr   hwndFocus;
            public IntPtr   hwndCapture;
            public IntPtr   hwndMenuOwner;
            public IntPtr   hwndMoveSize;
            public IntPtr   hwndCaret;
            public RECT     rc;
        }
        
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool SetForegroundWindow(IntPtr hwnd);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool GetGUIThreadInfo(int idThread, ref GUITHREADINFO guiThreadInfo);
        
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool GetWindowRect(IntPtr hwnd, ref RECT rc);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern IntPtr GetParent(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool IsWindow(IntPtr hwnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int RealGetWindowClass(IntPtr hWnd, StringBuilder classname, int nMax);
        
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern uint GetWindowThreadProcessId(IntPtr hwnd, out uint dwProcessId);
        
        [DllImport("kernel32.dll")]
        internal static extern int GetCurrentProcessId();

        #region Messaging
        [StructLayout(LayoutKind.Sequential)]
        internal struct MSG
        {
            public IntPtr hwnd;
            public uint message;
            public IntPtr wParam;
            public IntPtr lParam;
            public ulong time;
            public POINT pt;
        };

        [DllImport("user32.dll", EntryPoint = "GetMessageW", CharSet = CharSet.Unicode)]
        internal static extern int GetMessage(
            ref MSG msg,
            IntPtr hWnd,
            uint wMsgFilterMin,
            uint wMsgFilterMax
            );

        [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
        internal static extern int TranslateMessage(ref MSG msg);

        [DllImport("user32.dll", EntryPoint = "DispatchMessageW", CharSet = CharSet.Unicode)]
        internal static extern IntPtr DispatchMessage(ref MSG msg);

        public delegate void TIMERPROC(IntPtr hwnd, uint nMsg, UIntPtr idEvent, uint dwTime);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern UIntPtr SetTimer(IntPtr hwnd, UIntPtr idTimer, uint uElapse, TIMERPROC proc);
        
        internal static IntPtr MAKELPARAM( int low, int high )
        {
            return (IntPtr)((high << 16) | (low & 0xffff));
        }

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        public static extern IntPtr SendMessageTimeout(IntPtr hWnd, uint uMsg, IntPtr wParam, IntPtr lParam, uint flags, uint timeout, out uint result);
        
        // Hit-testing
        internal const int HTCLIENT            = 1;
        internal const int HTCAPTION           = 2;
        internal const int HTSYSMENU           = 3;
        internal const int HTSIZE              = 4;
        internal const int HTMENU              = 5;
        internal const int HTHSCROLL           = 6;
        internal const int HTVSCROLL           = 7;
        internal const int HTMINBUTTON         = 8;
        internal const int HTMAXBUTTON         = 9;
        internal const int HTLEFT              = 10;
        internal const int HTRIGHT             = 11;
        internal const int HTTOP               = 12;
        internal const int HTTOPLEFT           = 13;
        internal const int HTTOPRIGHT          = 14;
        internal const int HTBOTTOM            = 15;
        internal const int HTBOTTOMLEFT        = 16;
        internal const int HTBOTTOMRIGHT       = 17;
        internal const int HTBORDER            = 18;
        internal const int HTOBJECT            = 19;
        internal const int HTCLOSE             = 20;
        internal const int HTHELP              = 21;

        public const int WM_NCHITTEST = 0x0084;
        #endregion

        #region Input

        //
        // INPUT consts and structs 
        //
        internal const int  VK_RSHIFT         = 0xA1;
        internal const int  VK_TAB            = 0x09;

        // extended Keys        
        internal const int  VK_RMENU          = 0xA5;
        internal const int  VK_RCONTROL       = 0xA3;
        internal const int  VK_NUMLOCK        = 0x90;
        internal const int  VK_INSERT         = 0x2D;
        internal const int  VK_DELETE         = 0x2E;
        internal const int  VK_HOME           = 0x24;
        internal const int  VK_PRIOR          = 0x21;
        internal const int  VK_NEXT           = 0x22;
        internal const int  VK_END            = 0x23;
        internal const int  VK_LEFT           = 0x25;
        internal const int  VK_UP             = 0x26;
        internal const int  VK_RIGHT          = 0x27;
        internal const int  VK_DOWN           = 0x28;
        internal const int  VK_LWIN           = 0x5B;
        internal const int  VK_RWIN           = 0x5C;
        internal const int  VK_APPS           = 0x5D;
        
        
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
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern uint SendInput(uint nInputs, ref INPUT pInputs, int cbSize);
        
        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern int MapVirtualKey(int nVirtKey, int nMapType);
        #endregion

        #region Windows
        // Enum windows delegate
        internal delegate bool WinEnumProc(IntPtr hwnd, ref IntPtr lParam);

        [DllImport("user32.dll")]
        internal static extern bool EnumChildWindows(IntPtr hWndParent, WinEnumProc enumFunc, ref IntPtr lParam);
        #endregion

        #region WinEvents
        // the delegate passed to USER for receiving a WinEvent
        internal delegate void WinEventProcDef(IntPtr winEventHook, int eventId, IntPtr hwnd, int idObject, int idChild, int eventThread, int eventTime);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern IntPtr SetWinEventHook(int eventMin, int eventMax, int hmodWinEventProc, WinEventProcDef WinEventReentrancyFilter, uint idProcess, uint idThread, int dwFlags);

        [DllImport("user32.dll", CharSet=CharSet.Unicode, SetLastError=true)]
        internal static extern bool UnhookWinEvent(IntPtr winEventHook);

        internal const int WINEVENT_OUTOFCONTEXT = 0x0000;
        internal const int WINEVENT_SKIPOWNTHREAD = 0x0001;
        internal const int WINEVENT_SKIPOWNPROCESS = 0x0002;
        internal const int WINEVENT_INCONTEXT = 0x0004;

        internal const int EVENT_MIN = 0x00000001;
        internal const int EVENT_MAX = 0x7FFFFFFF;

        internal const int EVENT_SYSTEM_SOUND = 0x0001;
        internal const int EVENT_SYSTEM_ALERT = 0x0002;
        internal const int EVENT_SYSTEM_FOREGROUND = 0x0003;
        internal const int EVENT_SYSTEM_MENUSTART = 0x0004;
        internal const int EVENT_SYSTEM_MENUEND = 0x0005;
        internal const int EVENT_SYSTEM_MENUPOPUPSTART = 0x0006;
        internal const int EVENT_SYSTEM_MENUPOPUPEND = 0x0007;
        internal const int EVENT_SYSTEM_CAPTURESTART = 0x0008;
        internal const int EVENT_SYSTEM_CAPTUREEND = 0x0009;
        internal const int EVENT_SYSTEM_MOVESIZESTART = 0x000A;
        internal const int EVENT_SYSTEM_MOVESIZEEND = 0x000B;
        internal const int EVENT_SYSTEM_CONTEXTHELPSTART = 0x000C;
        internal const int EVENT_SYSTEM_CONTEXTHELPEND = 0x000D;
        internal const int EVENT_SYSTEM_DRAGDROPSTART = 0x000E;
        internal const int EVENT_SYSTEM_DRAGDROPEND = 0x000F;
        internal const int EVENT_SYSTEM_DIALOGSTART = 0x0010;
        internal const int EVENT_SYSTEM_DIALOGEND = 0x0011;
        internal const int EVENT_SYSTEM_SCROLLINGSTART = 0x0012;
        internal const int EVENT_SYSTEM_SCROLLINGEND = 0x0013;
        internal const int EVENT_SYSTEM_SWITCHSTART = 0x0014;
        internal const int EVENT_SYSTEM_SWITCHEND = 0x0015;
        internal const int EVENT_SYSTEM_MINIMIZESTART = 0x0016;
        internal const int EVENT_SYSTEM_MINIMIZEEND = 0x0017;
        internal const int EVENT_SYSTEM_DESKTOPSWITCH = 0x0020;
        internal const int EVENT_CONSOLE_CARET = 0x4001;
        internal const int EVENT_CONSOLE_UPDATE_REGION = 0x4002;
        internal const int EVENT_CONSOLE_UPDATE_SIMPLE = 0x4003;
        internal const int EVENT_CONSOLE_UPDATE_SCROLL = 0x4004;
        internal const int EVENT_CONSOLE_LAYOUT = 0x4005;
        internal const int EVENT_CONSOLE_START_APPLICATION = 0x4006;
        internal const int EVENT_CONSOLE_END_APPLICATION = 0x4007;
        internal const int EVENT_OBJECT_CREATE = 0x8000;
        internal const int EVENT_OBJECT_DESTROY = 0x8001;
        internal const int EVENT_OBJECT_SHOW = 0x8002;
        internal const int EVENT_OBJECT_HIDE = 0x8003;
        internal const int EVENT_OBJECT_REORDER = 0x8004;
        internal const int EVENT_OBJECT_FOCUS = 0x8005;
        internal const int EVENT_OBJECT_SELECTION = 0x8006;
        internal const int EVENT_OBJECT_SELECTIONADD = 0x8007;
        internal const int EVENT_OBJECT_SELECTIONREMOVE = 0x8008;
        internal const int EVENT_OBJECT_SELECTIONWITHIN = 0x8009;
        internal const int EVENT_OBJECT_STATECHANGE = 0x800A;
        internal const int EVENT_OBJECT_LOCATIONCHANGE = 0x800B;
        internal const int EVENT_OBJECT_NAMECHANGE = 0x800C;
        internal const int EVENT_OBJECT_DESCRIPTIONCHANGE = 0x800D;
        internal const int EVENT_OBJECT_VALUECHANGE = 0x800E;
        internal const int EVENT_OBJECT_PARENTCHANGE = 0x800F;
        internal const int EVENT_OBJECT_HELPCHANGE = 0x8010;
        internal const int EVENT_OBJECT_DEFACTIONCHANGE = 0x8011;
        internal const int EVENT_OBJECT_ACCELERATORCHANGE = 0x8012;
        internal const int EVENT_OBJECT_INVOKED = 0x8013;
        internal const int EVENT_OBJECT_TEXTSELECTIONCHANGED = 0x8014;
        internal const int EVENT_OBJECT_CONTENTSCROLLED = 0x8015;
        #endregion

        #region Accessibility
        internal const int SELFLAG_NONE = 0x00000000;
        internal const int SELFLAG_TAKEFOCUS = 0x00000001;
        internal const int SELFLAG_TAKESELECTION = 0x00000002;
        internal const int SELFLAG_EXTENDSELECTION	 = 0x00000004;
        internal const int SELFLAG_ADDSELECTION = 0x00000008;
        internal const int SELFLAG_REMOVESELECTION	 = 0x00000010;
        internal const int SELFLAG_VALID = 0x0000001F;

        internal const int ROLE_SYSTEM_TITLEBAR = 0x1;
        internal const int ROLE_SYSTEM_MENUBAR = 0x2;
        internal const int ROLE_SYSTEM_SCROLLBAR = 0x3;
        internal const int ROLE_SYSTEM_GRIP = 0x4;
        internal const int ROLE_SYSTEM_SOUND = 0x5;
        internal const int ROLE_SYSTEM_CURSOR = 0x6;
        internal const int ROLE_SYSTEM_CARET = 0x7;
        internal const int ROLE_SYSTEM_ALERT = 0x8;
        internal const int ROLE_SYSTEM_WINDOW = 0x9;
        internal const int ROLE_SYSTEM_CLIENT = 0xa;
        internal const int ROLE_SYSTEM_MENUPOPUP = 0xb;
        internal const int ROLE_SYSTEM_MENUITEM = 0xc;
        internal const int ROLE_SYSTEM_TOOLTIP = 0xd;
        internal const int ROLE_SYSTEM_APPLICATION = 0xe;
        internal const int ROLE_SYSTEM_DOCUMENT = 0xf;
        internal const int ROLE_SYSTEM_PANE = 0x10;
        internal const int ROLE_SYSTEM_CHART = 0x11;
        internal const int ROLE_SYSTEM_DIALOG = 0x12;
        internal const int ROLE_SYSTEM_BORDER = 0x13;
        internal const int ROLE_SYSTEM_GROUPING = 0x14;
        internal const int ROLE_SYSTEM_SEPARATOR = 0x15;
        internal const int ROLE_SYSTEM_TOOLBAR = 0x16;
        internal const int ROLE_SYSTEM_STATUSBAR = 0x17;
        internal const int ROLE_SYSTEM_TABLE = 0x18;
        internal const int ROLE_SYSTEM_COLUMNHEADER = 0x19;
        internal const int ROLE_SYSTEM_ROWHEADER = 0x1a;
        internal const int ROLE_SYSTEM_COLUMN = 0x1b;
        internal const int ROLE_SYSTEM_ROW = 0x1c;
        internal const int ROLE_SYSTEM_CELL = 0x1d;
        internal const int ROLE_SYSTEM_LINK = 0x1e;
        internal const int ROLE_SYSTEM_HELPBALLOON = 0x1f;
        internal const int ROLE_SYSTEM_CHARACTER = 0x20;
        internal const int ROLE_SYSTEM_LIST = 0x21;
        internal const int ROLE_SYSTEM_LISTITEM = 0x22;
        internal const int ROLE_SYSTEM_OUTLINE = 0x23;
        internal const int ROLE_SYSTEM_OUTLINEITEM = 0x24;
        internal const int ROLE_SYSTEM_PAGETAB = 0x25;
        internal const int ROLE_SYSTEM_PROPERTYPAGE = 0x26;
        internal const int ROLE_SYSTEM_INDICATOR = 0x27;
        internal const int ROLE_SYSTEM_GRAPHIC = 0x28;
        internal const int ROLE_SYSTEM_STATICTEXT = 0x29;
        internal const int ROLE_SYSTEM_TEXT = 0x2a;
        internal const int ROLE_SYSTEM_PUSHBUTTON = 0x2b;
        internal const int ROLE_SYSTEM_CHECKBUTTON = 0x2c;
        internal const int ROLE_SYSTEM_RADIOBUTTON = 0x2d;
        internal const int ROLE_SYSTEM_COMBOBOX = 0x2e;
        internal const int ROLE_SYSTEM_DROPLIST = 0x2f;
        internal const int ROLE_SYSTEM_PROGRESSBAR = 0x30;
        internal const int ROLE_SYSTEM_DIAL = 0x31;
        internal const int ROLE_SYSTEM_HOTKEYFIELD = 0x32;
        internal const int ROLE_SYSTEM_SLIDER = 0x33;
        internal const int ROLE_SYSTEM_SPINBUTTON = 0x34;
        internal const int ROLE_SYSTEM_DIAGRAM = 0x35;
        internal const int ROLE_SYSTEM_ANIMATION = 0x36;
        internal const int ROLE_SYSTEM_EQUATION = 0x37;
        internal const int ROLE_SYSTEM_BUTTONDROPDOWN = 0x38;
        internal const int ROLE_SYSTEM_BUTTONMENU = 0x39;
        internal const int ROLE_SYSTEM_BUTTONDROPDOWNGRID = 0x3a;
        internal const int ROLE_SYSTEM_WHITESPACE = 0x3b;
        internal const int ROLE_SYSTEM_PAGETABLIST = 0x3c;
        internal const int ROLE_SYSTEM_CLOCK = 0x3d;
        internal const int ROLE_SYSTEM_SPLITBUTTON = 0x3e;
        internal const int ROLE_SYSTEM_IPADDRESS = 0x3f;
        internal const int ROLE_SYSTEM_OUTLINEBUTTON = 0x40;

        internal const int STATE_SYSTEM_UNAVAILABLE = 0x00000001;  // Disabled
        internal const int STATE_SYSTEM_SELECTED = 0x00000002;
        internal const int STATE_SYSTEM_FOCUSED = 0x00000004;
        internal const int STATE_SYSTEM_PRESSED = 0x00000008;
        internal const int STATE_SYSTEM_CHECKED = 0x00000010;
        internal const int STATE_SYSTEM_MIXED = 0x00000020;  // 3-state checkbox or toolbar button
        internal const int STATE_SYSTEM_INDETERMINATE = STATE_SYSTEM_MIXED;
        internal const int STATE_SYSTEM_READONLY = 0x00000040;
        internal const int STATE_SYSTEM_HOTTRACKED = 0x00000080;
        internal const int STATE_SYSTEM_DEFAULT = 0x00000100;
        internal const int STATE_SYSTEM_EXPANDED = 0x00000200;
        internal const int STATE_SYSTEM_COLLAPSED = 0x00000400;
        internal const int STATE_SYSTEM_BUSY = 0x00000800;
        internal const int STATE_SYSTEM_FLOATING = 0x00001000;  // Children "owned" not "contained" by parent
        internal const int STATE_SYSTEM_MARQUEED = 0x00002000;
        internal const int STATE_SYSTEM_ANIMATED = 0x00004000;
        internal const int STATE_SYSTEM_INVISIBLE = 0x00008000;
        internal const int STATE_SYSTEM_OFFSCREEN = 0x00010000;
        internal const int STATE_SYSTEM_SIZEABLE = 0x00020000;
        internal const int STATE_SYSTEM_MOVEABLE = 0x00040000;
        internal const int STATE_SYSTEM_SELFVOICING = 0x00080000;
        internal const int STATE_SYSTEM_FOCUSABLE = 0x00100000;
        internal const int STATE_SYSTEM_SELECTABLE = 0x00200000;
        internal const int STATE_SYSTEM_LINKED = 0x00400000;
        internal const int STATE_SYSTEM_TRAVERSED = 0x00800000;
        internal const int STATE_SYSTEM_MULTISELECTABLE = 0x01000000;  // Supports multiple selection
        internal const int STATE_SYSTEM_EXTSELECTABLE = 0x02000000;  // Supports extended selection
        internal const int STATE_SYSTEM_ALERT_LOW = 0x04000000;  // This information is of low priority
        internal const int STATE_SYSTEM_ALERT_MEDIUM = 0x08000000;  // This information is of medium priority
        internal const int STATE_SYSTEM_ALERT_HIGH = 0x10000000;  // This information is of high priority
        internal const int STATE_SYSTEM_PROTECTED = 0x20000000;  // access to this is restricted
        internal const int STATE_SYSTEM_HAS_POPUP = 0x40000000;
        internal const int STATE_SYSTEM_VALID = 0x3FFFFFFF;
        #endregion

        #region COM
        internal static Guid IID_IUnknown = new Guid(0x00000000, 0x0000, 0x0000, 0xc0, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x46);

        internal const int DISP_E_MEMBERNOTFOUND = unchecked((int)0x80020003);    // 2147614723
        internal const int E_NOTIMPL = unchecked((int)0x80004001);                // 2147500033
        internal const int E_FAIL = unchecked((int)0x80004005);                   // 2147500037

        internal const int S_OK = 0;
        #endregion



    }
}
