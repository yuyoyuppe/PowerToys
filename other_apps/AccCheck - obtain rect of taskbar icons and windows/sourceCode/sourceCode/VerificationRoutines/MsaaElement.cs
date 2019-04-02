// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Text;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Threading;
using Accessibility;
using AccCheck.Logging;
using AccCheck.Verification;
using AccCheck;
using System.Windows.Forms;
using System.Reflection;
using VerificationRoutines;

namespace VerificationRoutines
{
    // indicates the state of an accessible element
    public enum AccState
    {
        Checked, Collapsed, Expanded, Focusable, Focused, HasPopup, Invisible,
        Linked, Mixed, Moveable, MultiSelectable, OffScreen, Protected, ReadOnly, Selectable,
        Selected, Sizeable, Unavailable
    };

    // indicates the role of an accessible element
    public enum AccRole
    {
        Titlebar, Menubar, Scrollbar, Grip, Sound, Cursor, Caret, Alert, Window,
        Client, MenuPopup, MenuItem, Tooltip, Application, Document, Pane, Chart, Dialog, Border,
        Grouping, Separator, Toolbar, Statusbar, Table, ColumnHeader, RowHeader, Column, Row, Cell,
        Link, HelpBalloon, Character, List, ListItem, Tree, TreeItem, PageTab, PropertyPage, Indicator,
        Graphic, StaticText, Text, PushButton, CheckButton, RadioButton, Combobox, DropList, ProgressBar,
        Dial, Hotkeyfield, Slider, SpinButton, Diagram, Animation, Equation, ButtonDropdown, ButtonMenu,
        ButtonDropdownGrid, Whitespace, PageTabList, Clock, SplitButton, IPAddress, TreeButton, Invalid
    }

    public class MsaaElement : ICache
    {
        #region Win32 Conversion
        private static Hashtable StateMap = new Hashtable();
        private static Hashtable RoleMap = new Hashtable();
        private static Hashtable EventMap = new Hashtable();
        private static void InitializeStateMap()
        {
            if (StateMap.Count == 0)
            {
                StateMap.Add(Win32API.STATE_SYSTEM_CHECKED, AccState.Checked);
                StateMap.Add(Win32API.STATE_SYSTEM_COLLAPSED, AccState.Collapsed);
                StateMap.Add(Win32API.STATE_SYSTEM_EXPANDED, AccState.Expanded);
                StateMap.Add(Win32API.STATE_SYSTEM_FOCUSABLE, AccState.Focusable);
                StateMap.Add(Win32API.STATE_SYSTEM_FOCUSED, AccState.Focused);
                StateMap.Add(Win32API.STATE_SYSTEM_HAS_POPUP, AccState.HasPopup);
                StateMap.Add(Win32API.STATE_SYSTEM_INVISIBLE, AccState.Invisible);
                StateMap.Add(Win32API.STATE_SYSTEM_MIXED, AccState.Mixed);
                StateMap.Add(Win32API.STATE_SYSTEM_MOVEABLE, AccState.Moveable);
                StateMap.Add(Win32API.STATE_SYSTEM_PROTECTED, AccState.Protected);
                StateMap.Add(Win32API.STATE_SYSTEM_READONLY, AccState.ReadOnly);
                StateMap.Add(Win32API.STATE_SYSTEM_SELECTABLE, AccState.Selectable);
                StateMap.Add(Win32API.STATE_SYSTEM_SIZEABLE, AccState.Sizeable);
                StateMap.Add(Win32API.STATE_SYSTEM_OFFSCREEN, AccState.OffScreen);
                StateMap.Add(Win32API.STATE_SYSTEM_SELECTED, AccState.Selected);
                StateMap.Add(Win32API.STATE_SYSTEM_MULTISELECTABLE, AccState.MultiSelectable);
                StateMap.Add(Win32API.STATE_SYSTEM_UNAVAILABLE, AccState.Unavailable);
                StateMap.Add(Win32API.STATE_SYSTEM_LINKED, AccState.Linked);
            }
        }
        private static void InitializeRoleMap()
        {
            if (RoleMap.Count == 0)
            {
                RoleMap.Add(Win32API.ROLE_SYSTEM_TITLEBAR, AccRole.Titlebar);
                RoleMap.Add(Win32API.ROLE_SYSTEM_MENUBAR, AccRole.Menubar);
                RoleMap.Add(Win32API.ROLE_SYSTEM_SCROLLBAR, AccRole.Scrollbar);
                RoleMap.Add(Win32API.ROLE_SYSTEM_GRIP, AccRole.Grip);
                RoleMap.Add(Win32API.ROLE_SYSTEM_SOUND, AccRole.Sound);
                RoleMap.Add(Win32API.ROLE_SYSTEM_CURSOR, AccRole.Cursor);
                RoleMap.Add(Win32API.ROLE_SYSTEM_CARET, AccRole.Caret);
                RoleMap.Add(Win32API.ROLE_SYSTEM_ALERT, AccRole.Alert);
                RoleMap.Add(Win32API.ROLE_SYSTEM_WINDOW, AccRole.Window);
                RoleMap.Add(Win32API.ROLE_SYSTEM_CLIENT, AccRole.Client);
                RoleMap.Add(Win32API.ROLE_SYSTEM_MENUPOPUP, AccRole.MenuPopup);
                RoleMap.Add(Win32API.ROLE_SYSTEM_MENUITEM, AccRole.MenuItem);
                RoleMap.Add(Win32API.ROLE_SYSTEM_TOOLTIP, AccRole.Tooltip);
                RoleMap.Add(Win32API.ROLE_SYSTEM_APPLICATION, AccRole.Application);
                RoleMap.Add(Win32API.ROLE_SYSTEM_DOCUMENT, AccRole.Document);
                RoleMap.Add(Win32API.ROLE_SYSTEM_PANE, AccRole.Pane);
                RoleMap.Add(Win32API.ROLE_SYSTEM_CHART, AccRole.Chart);
                RoleMap.Add(Win32API.ROLE_SYSTEM_DIALOG, AccRole.Dialog);
                RoleMap.Add(Win32API.ROLE_SYSTEM_BORDER, AccRole.Border);
                RoleMap.Add(Win32API.ROLE_SYSTEM_GROUPING, AccRole.Grouping);
                RoleMap.Add(Win32API.ROLE_SYSTEM_SEPARATOR, AccRole.Separator);
                RoleMap.Add(Win32API.ROLE_SYSTEM_TOOLBAR, AccRole.Toolbar);
                RoleMap.Add(Win32API.ROLE_SYSTEM_STATUSBAR, AccRole.Statusbar);
                RoleMap.Add(Win32API.ROLE_SYSTEM_TABLE, AccRole.Table);
                RoleMap.Add(Win32API.ROLE_SYSTEM_COLUMNHEADER, AccRole.ColumnHeader);
                RoleMap.Add(Win32API.ROLE_SYSTEM_ROWHEADER, AccRole.RowHeader);
                RoleMap.Add(Win32API.ROLE_SYSTEM_COLUMN, AccRole.Column);
                RoleMap.Add(Win32API.ROLE_SYSTEM_ROW, AccRole.Row);
                RoleMap.Add(Win32API.ROLE_SYSTEM_CELL, AccRole.Cell);
                RoleMap.Add(Win32API.ROLE_SYSTEM_LINK, AccRole.Link);
                RoleMap.Add(Win32API.ROLE_SYSTEM_HELPBALLOON, AccRole.HelpBalloon);
                RoleMap.Add(Win32API.ROLE_SYSTEM_CHARACTER, AccRole.Character);
                RoleMap.Add(Win32API.ROLE_SYSTEM_LIST, AccRole.List);
                RoleMap.Add(Win32API.ROLE_SYSTEM_LISTITEM, AccRole.ListItem);
                RoleMap.Add(Win32API.ROLE_SYSTEM_OUTLINE, AccRole.Tree);
                RoleMap.Add(Win32API.ROLE_SYSTEM_OUTLINEITEM, AccRole.TreeItem);
                RoleMap.Add(Win32API.ROLE_SYSTEM_PAGETAB, AccRole.PageTab);
                RoleMap.Add(Win32API.ROLE_SYSTEM_PROPERTYPAGE, AccRole.PropertyPage);
                RoleMap.Add(Win32API.ROLE_SYSTEM_INDICATOR, AccRole.Indicator);
                RoleMap.Add(Win32API.ROLE_SYSTEM_GRAPHIC, AccRole.Graphic);
                RoleMap.Add(Win32API.ROLE_SYSTEM_STATICTEXT, AccRole.StaticText);
                RoleMap.Add(Win32API.ROLE_SYSTEM_TEXT, AccRole.Text);
                RoleMap.Add(Win32API.ROLE_SYSTEM_PUSHBUTTON, AccRole.PushButton);
                RoleMap.Add(Win32API.ROLE_SYSTEM_CHECKBUTTON, AccRole.CheckButton);
                RoleMap.Add(Win32API.ROLE_SYSTEM_RADIOBUTTON, AccRole.RadioButton);
                RoleMap.Add(Win32API.ROLE_SYSTEM_COMBOBOX, AccRole.Combobox);
                RoleMap.Add(Win32API.ROLE_SYSTEM_DROPLIST, AccRole.DropList);
                RoleMap.Add(Win32API.ROLE_SYSTEM_PROGRESSBAR, AccRole.ProgressBar);
                RoleMap.Add(Win32API.ROLE_SYSTEM_DIAL, AccRole.Dial);
                RoleMap.Add(Win32API.ROLE_SYSTEM_HOTKEYFIELD, AccRole.Hotkeyfield);
                RoleMap.Add(Win32API.ROLE_SYSTEM_SLIDER, AccRole.Slider);
                RoleMap.Add(Win32API.ROLE_SYSTEM_SPINBUTTON, AccRole.SpinButton);
                RoleMap.Add(Win32API.ROLE_SYSTEM_DIAGRAM, AccRole.Diagram);
                RoleMap.Add(Win32API.ROLE_SYSTEM_ANIMATION, AccRole.Animation);
                RoleMap.Add(Win32API.ROLE_SYSTEM_EQUATION, AccRole.Equation);
                RoleMap.Add(Win32API.ROLE_SYSTEM_BUTTONDROPDOWN, AccRole.ButtonDropdown);
                RoleMap.Add(Win32API.ROLE_SYSTEM_BUTTONMENU, AccRole.ButtonMenu);
                RoleMap.Add(Win32API.ROLE_SYSTEM_BUTTONDROPDOWNGRID, AccRole.ButtonDropdownGrid);
                RoleMap.Add(Win32API.ROLE_SYSTEM_WHITESPACE, AccRole.Whitespace);
                RoleMap.Add(Win32API.ROLE_SYSTEM_PAGETABLIST, AccRole.PageTabList);
                RoleMap.Add(Win32API.ROLE_SYSTEM_CLOCK, AccRole.Clock);
                RoleMap.Add(Win32API.ROLE_SYSTEM_SPLITBUTTON, AccRole.SplitButton);
                RoleMap.Add(Win32API.ROLE_SYSTEM_IPADDRESS, AccRole.IPAddress);
                RoleMap.Add(Win32API.ROLE_SYSTEM_OUTLINEBUTTON, AccRole.TreeButton);

            }
        }
        private static void InitializeEventMap()
        {
            if (EventMap.Count == 0)
            {
                EventMap.Add(Win32API.EVENT_SYSTEM_SOUND, AccEventIdentifier.SystemSound);
                EventMap.Add(Win32API.EVENT_SYSTEM_ALERT, AccEventIdentifier.SystemAlert);
                EventMap.Add(Win32API.EVENT_SYSTEM_FOREGROUND, AccEventIdentifier.SystemForeground);
                EventMap.Add(Win32API.EVENT_SYSTEM_MENUSTART, AccEventIdentifier.SystemMenuStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_MENUEND, AccEventIdentifier.SystemMenuEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_MENUPOPUPSTART, AccEventIdentifier.SystemMenuPopupStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_MENUPOPUPEND, AccEventIdentifier.SystemMenuPopupEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_CAPTURESTART, AccEventIdentifier.SystemCaptureStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_CAPTUREEND, AccEventIdentifier.SystemCaptureEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_MOVESIZESTART, AccEventIdentifier.SystemMovesizeStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_MOVESIZEEND, AccEventIdentifier.SystemMovesizeEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_CONTEXTHELPSTART, AccEventIdentifier.SystemContexthelpStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_CONTEXTHELPEND, AccEventIdentifier.SystemContexthelpEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_DRAGDROPSTART, AccEventIdentifier.SystemDragdropStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_DRAGDROPEND, AccEventIdentifier.SystemDragdropEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_DIALOGSTART, AccEventIdentifier.SystemDialogStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_DIALOGEND, AccEventIdentifier.SystemDialogEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_SCROLLINGSTART, AccEventIdentifier.SystemScrollingStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_SCROLLINGEND, AccEventIdentifier.SystemScrollingEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_SWITCHSTART, AccEventIdentifier.SystemSwitchStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_SWITCHEND, AccEventIdentifier.SystemSwitchEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_MINIMIZESTART, AccEventIdentifier.SystemMinimizeStart);
                EventMap.Add(Win32API.EVENT_SYSTEM_MINIMIZEEND, AccEventIdentifier.SystemMinimizeEnd);
                EventMap.Add(Win32API.EVENT_SYSTEM_DESKTOPSWITCH, AccEventIdentifier.SystemDesktopSwitch);
                EventMap.Add(Win32API.EVENT_CONSOLE_CARET, AccEventIdentifier.ConsoleCaret);
                EventMap.Add(Win32API.EVENT_CONSOLE_UPDATE_REGION, AccEventIdentifier.ConsoleUpdateRegion);
                EventMap.Add(Win32API.EVENT_CONSOLE_UPDATE_SIMPLE, AccEventIdentifier.ConsoleUpdateSimple);
                EventMap.Add(Win32API.EVENT_CONSOLE_UPDATE_SCROLL, AccEventIdentifier.ConsoleUpdateScroll);
                EventMap.Add(Win32API.EVENT_CONSOLE_LAYOUT, AccEventIdentifier.ConsoleLayout);
                EventMap.Add(Win32API.EVENT_CONSOLE_START_APPLICATION, AccEventIdentifier.ConsoleStartApplication);
                EventMap.Add(Win32API.EVENT_CONSOLE_END_APPLICATION, AccEventIdentifier.ConsoleEndApplication);
                EventMap.Add(Win32API.EVENT_OBJECT_CREATE, AccEventIdentifier.ObjectCreate);
                EventMap.Add(Win32API.EVENT_OBJECT_DESTROY, AccEventIdentifier.ObjectDestroy);
                EventMap.Add(Win32API.EVENT_OBJECT_SHOW, AccEventIdentifier.ObjectShow);
                EventMap.Add(Win32API.EVENT_OBJECT_HIDE, AccEventIdentifier.ObjectHide);
                EventMap.Add(Win32API.EVENT_OBJECT_REORDER, AccEventIdentifier.ObjectReorder);
                EventMap.Add(Win32API.EVENT_OBJECT_FOCUS, AccEventIdentifier.ObjectFocus);
                EventMap.Add(Win32API.EVENT_OBJECT_SELECTION, AccEventIdentifier.ObjectSelection);
                EventMap.Add(Win32API.EVENT_OBJECT_SELECTIONADD, AccEventIdentifier.ObjectSelectionAdd);
                EventMap.Add(Win32API.EVENT_OBJECT_SELECTIONREMOVE, AccEventIdentifier.ObjectSelectionRemove);
                EventMap.Add(Win32API.EVENT_OBJECT_SELECTIONWITHIN, AccEventIdentifier.ObjectSelectionWithin);
                EventMap.Add(Win32API.EVENT_OBJECT_STATECHANGE, AccEventIdentifier.ObjectStateChange);
                EventMap.Add(Win32API.EVENT_OBJECT_LOCATIONCHANGE, AccEventIdentifier.ObjectLocationChange);
                EventMap.Add(Win32API.EVENT_OBJECT_NAMECHANGE, AccEventIdentifier.ObjectNameChange);
                EventMap.Add(Win32API.EVENT_OBJECT_DESCRIPTIONCHANGE, AccEventIdentifier.ObjectDescriptionChange);
                EventMap.Add(Win32API.EVENT_OBJECT_VALUECHANGE, AccEventIdentifier.ObjectValueChange);
                EventMap.Add(Win32API.EVENT_OBJECT_PARENTCHANGE, AccEventIdentifier.ObjectParentChange);
                EventMap.Add(Win32API.EVENT_OBJECT_HELPCHANGE, AccEventIdentifier.ObjectHelpChange);
                EventMap.Add(Win32API.EVENT_OBJECT_DEFACTIONCHANGE, AccEventIdentifier.ObjectDefactionChange);
                EventMap.Add(Win32API.EVENT_OBJECT_ACCELERATORCHANGE, AccEventIdentifier.ObjectAcceleratorChange);
                EventMap.Add(Win32API.EVENT_OBJECT_INVOKED, AccEventIdentifier.ObjectInvoked);
                EventMap.Add(Win32API.EVENT_OBJECT_TEXTSELECTIONCHANGED, AccEventIdentifier.ObjectTextselectionChanged);
                EventMap.Add(Win32API.EVENT_OBJECT_CONTENTSCROLLED, AccEventIdentifier.ObjectContentScrolled);
            }
        }
        #endregion

        #region Private member variables

        // keys are ixElements, values are Elements
        private static MsaaElement _rootElement = null;

        private static IntPtr _rootHwnd = IntPtr.Zero;
        private static Rectangle _rootRectangle;
        private ObjectIdentity _id;
        private Accessible _accessible;

        private MsaaElement _parent;
        private int _indexInParent;

        private object _name;     // Name should be a string can contain an exception object
        private object _role;     // role constant should be a int can contain an exception object
        private object _state;    // state constant should be a int can contain an exception object
        private object _value;    // value should be a string can contain an exception object
        private object _location;    // value should be a string can contain an exception object

        private Exception _childrenException;

        private IntPtr _hwnd;
        private string _windowClass;
        private List<MsaaElement> _children = new List<MsaaElement>();

        private ILogger _logger;

        private const int S_OK = 0;
        private const string PARENT_CHAIN_SEPERATOR = ".";

        /// <summary>Call this to force a refresh of the cached elements when CacheMsaaTree is called/// </summary>
        public void Clear()
        {
            // CacheMsaaTree will not rebuild the cache if the hwnd's are the same.  If
            // they are not the same, or zero, then they cache is rebuilt. 
            MsaaElement._rootHwnd = IntPtr.Zero;
            MsaaElement._rootElement = null;
        }

        public MsaaElement() { }

        static MsaaElement()
        {
            InitializeStateMap();
            InitializeRoleMap();
            InitializeEventMap();
        }

        // this is private use CacheMsaaTree to create the and get the root
        private MsaaElement(ObjectIdentity id, Accessible accessible)
        {
            _id = id;
            _accessible = accessible;
            _parent = null;
            _hwnd = accessible.Hwnd;
        }

        //
        // This method is used when initially indexing the accessibility tree
        // for an application.
        // The element in this case is the objects that get returned from AccessibleChildren
        // these can be either an IAccessible or an child id.
        //
        private static MsaaElement CreateSubTree(Accessible element, MsaaElement parent, int childIndex)
        {
            MsaaElement msaaElement = null;
            
            ObjectIdentity ixMe;

            if (parent == null)
            {
                ixMe = new ObjectIdentity();
            }
            else
            {
                ixMe = new ObjectIdentity(parent._id, childIndex);
            }

            if (ixMe.GetFriendlyPath().Length > 500)
            {
                throw new CircularReferenceInMsaaTree();
            }

            msaaElement = new MsaaElement(ixMe, element);

            msaaElement._parent = parent;
            msaaElement._children = new List<MsaaElement>();
            msaaElement._windowClass = msaaElement.GetWindowClass();
            msaaElement._indexInParent = childIndex;

            UpdateMsaaCachedDynamicProperties(msaaElement);

            try
            {
                Accessible[] accessibleChildren = null;
                if (element.Children(out accessibleChildren) == Win32API.S_OK)
                {
                    int i = 0;
                    foreach (Accessible child in accessibleChildren)
                    {
                       MsaaElement msaaElementChild = MsaaElement.CreateSubTree(child, msaaElement, i++);
                        if (msaaElementChild != null)
                        {
                            msaaElement._children.Add(msaaElementChild);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                if (VerificationHelper.IsCriticalException(exception))
                {
                    throw;
                }

                if (exception is ChildCountInvalidException)
                {
                    // save this exception so that we can report it in a verification.
                    msaaElement._childrenException = exception;
                }
                else
                {
                    TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, typeof(MsaaElement), exception, null, element, "CreateSubTree");
                }
            }

            return msaaElement;
        }

        /// <summary>
        /// Cache/Update the various properties of the MSAAElement that can change.  This
        /// method is called from the tests when the properties could have changed and 
        /// need updated.
        /// </summary>
        /// <param name="element"></param>
        public static void UpdateMsaaCachedDynamicProperties(MsaaElement msaaElement)
        {
            msaaElement._name = msaaElement.GetAccName();
            msaaElement._role = msaaElement.GetAccRole();
            msaaElement._state = msaaElement.GetAccState();
            msaaElement._value = msaaElement.GetAccValue();
            msaaElement._location = msaaElement.GetAccLocation();
        }

        //
        // Compares an MsaaElement based on name, value, child count and location.
        //
        public bool Compare(MsaaElement element)
        {
            // the goal is to abort as quickly as possible if we can, since
            // this method will be called many times when searching the tree,
            // and we want to minimize the amount of time used on non-matches
            if (element == null)
            {
                return false;
            }

            if (this == element)
            {
                return true;
            }

            if (element.GetName() != GetName())
            {
                return false;
            }

            if (element.GetValue() != GetValue())
            {
                return false;
            }

            if (element.GetRole() != GetRole())
            {
                return false;
            }

            if (!GetBoundingBox().Equals(element.GetBoundingBox()))
            {
                return false;
            }

            return true;
        }

        //
        // Compares an MsaaElement with an Accessible based on name, value, child count and location.
        //

        public bool Compare(Accessible acc)
        {
            // the goal is to abort as quickly as possible if we can, since
            // this method will be called many times when searching the tree,
            // and we want to minimize the amount of time used on non-matches
            if (acc == null)
            {
                return false;
            }

            Rectangle loc;
            try
            {
                loc = acc.Location;
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
            if (!GetBoundingBox().Equals(loc))
            {
                return false;
            }


            string name = "";
            try
            {
                name = acc.Name;
                if (name == null)
                {
                    name = "";
                }
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
            if (GetName() != name)
            {
                return false;
            }

            int role = 0;
            try
            {
                role = acc.Role;
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
            if (GetRole() != LookupRole(role))
            {
                return false;
            }

            string value = "";
            try
            {
                value = acc.Value;
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
            if (GetValue() != value)
            {
                return false;
            }

            return true;
        }

        public bool IsStatePresent(AccState state)
        {
            return Array.IndexOf(GetStates(), state) != -1;
        }

        // this is the method that is called to get the name.  One of the problems with 
        // caching the name is that errors encountered could be lost.  Therefore any exceptions
        // are saved in the name object so verification routine that are interested could get this 
        // information.  
        private object GetAccName()
        {
            object name = null;
            try
            {
                name = _accessible.Name;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    name = e;
                }
                else
                {
                    throw;
                }
            }

            return name;
        }

        // this is the method that is called to cache the role.  Role can contain exception objects.
        // Also if the type of the variant passed back is wrong this will not be lost.
        private object GetAccRole()
        {
            object role;
            try
            {
                role = _accessible.Role;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    role = e;
                }
                else
                {
                    throw;
                }
            }

            return role;
        }

        // this is the method that is called to cache the state.  State can contain exceptionobject.
        // Also if the type of the variant passed back is wrong this will not be lost.
        private object GetAccState()
        {
            object state;
            try
            {
                state = _accessible.State;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    state = e;
                }
                else
                {
                    throw;
                }
            }

            return state;
        }

        // this is the method that is called to cache the value.  Value can contain exception object.
        private object GetAccValue()
        {
            object value;
            try
            {
                value = _accessible.Value;
            }
            catch (Exception e)
            {
                // We have seen out of memory return by Value save it here and don't treat it
                // like the system is out of memory.
                if (Accessible.IsExpectedException(e) || e is OutOfMemoryException)
                {
                    value = e;
                }
                else
                {
                    throw;
                }
            }

            return value;
        }

        // this is the method that is called to cache the value.  Value can contain exception object.
        private object GetAccLocation()
        {
            object location;
            try
            {
                location = _accessible.Location;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    location = e;
                }
                else
                {
                    throw;
                }
            }

            return location;
        }

        public String GetWindowClass()
        {
            StringBuilder windowsClass = new StringBuilder(1024);
            try
            {
                Win32API.RealGetWindowClass(_hwnd, windowsClass, windowsClass.Capacity);
            }
            catch (Exception e)
            {
                if (VerificationHelper.IsCriticalException(e))
                {
                    throw;
                }

                // don't need to report errors with this call
                // There are some windows in office that make this overwrite the StringBuild object.
            }

            return windowsClass.ToString();
        }

        #endregion

        public ILogger Logger
        {
            get
            {
                return _logger;
            }
            set
            {
                _logger = value;
            }
        }

        //
        // Sets up everything that is needed for verification of a given hwnd.
        //
        public static MsaaElement CacheMsaaTree(IntPtr hwnd)
        {
            // If this logic changes for how the tree is cached or not cached, update ClearCache accordingly.
            if (hwnd != _rootHwnd)
            {
                Trace.WriteLine("> Updating the cache tree");

                _rootHwnd = hwnd;
                Win32API.RECT win32RootRect = Win32API.RECT.Empty;
                Win32API.GetWindowRect(_rootHwnd, ref win32RootRect);
                _rootRectangle = new Rectangle(win32RootRect.left, win32RootRect.top, win32RootRect.right - win32RootRect.left, win32RootRect.bottom - win32RootRect.top);

                Accessible rootAccessible = null;
                try
                {
                    int ret = Accessible.FromWindow(hwnd, out rootAccessible);

                    _rootElement = MsaaElement.CreateSubTree(rootAccessible, null, 0);
                }
                catch (Exception e)
                {
                    if (VerificationHelper.IsCriticalException(e))
                    {
                        throw;
                    }

                    TestLogger.LogException(EventLevel.Error, LogEvent.MethodExceptionOccured, typeof(MsaaElement), e, null, rootAccessible);
                }
            }

            return RootElement;
        }

        public Accessible Accessible
        {
            get
            {
                return _accessible;
            }
        }

        public int[] TreePath
        {
            get
            {
                return _id.GetFriendlyPath();
            }
        }

        public static MsaaElement RootElement
        {
            get
            {
                return _rootElement;
            }
        }

        public List<MsaaElement> Children
        {
            get
            {
                return _children;
            }
        }

        public bool IsVisible
        {
            get
            {
                return (-1 == Array.IndexOf(GetStates(), AccState.Invisible));
            }
        }

        public MsaaElement Parent
        {
            get
            {
                return _parent;
            }
        }

        public MsaaElement GetNextSibling()
        {

            if (_parent != null)
            {
                int nextSibling = _indexInParent + 1;
                if (nextSibling < _parent._children.Count)
                {
                    return _parent._children[nextSibling];
                }
            }

            return null;
        }

        public MsaaElement GetPrevSibling()
        {
            if (_parent != null)
            {
                int prevSibling = _indexInParent - 1;
                if (prevSibling >= 0)
                {
                    return _parent._children[prevSibling];
                }
            }

            return null;
        }

        // When the tree is created the location is cached.  This returns a sanitized version of location
        public Rectangle GetBoundingBox()
        {
            if (_location is Rectangle)
            {
                return (Rectangle)_location;
            }
            return new Rectangle();
        }

        // When the tree is created the location is cached.  
        // If there was an exception when the location was retreived return it.
        public object GetBoundingBox(out Exception exception)
        {
            exception = _location as Exception;
            return GetBoundingBox();
        }

        // When the tree is created the value is cached.  This returns a sanitized version of value
        public string GetValue()
        {
            if (_value is string)
            {
                return (string)_value;
            }

            return "";
        }

        // When the tree is created the value is cached.  
        // If there was an exception when the value was retreived return it.
        public string GetValue(out Exception exception)
        {
            exception = _value as Exception;
            return GetValue();
        }

        // When the tree is created the name is cached.  This returns a sanitized version of name
        public string GetName()
        {
            if (_name is string)
            {
                return (string)_name;
            }

            return "";
        }

        // When the tree is created the name is cached.  
        // If there was an exception when the name was retreived return it.
        public object GetName(out Exception exception)
        {
            exception = _name as Exception;
            return _name;
        }

        // When the tree is created the role is cached.  This returns a sanitized version of role
        public AccRole GetRole()
        {
            AccRole role = AccRole.Invalid;

            if (_role is int)
            {
                if (RoleMap.Contains((int)_role))
                {
                    role = (AccRole)RoleMap[_role];
                }
            }

            return role;
        }

        // This is used to get enum version the the get_accRole value
        public static AccRole LookupRole(int accRole)
        {
            AccRole role = AccRole.Invalid;


            if (RoleMap.Contains(accRole))
            {
                role = (AccRole)RoleMap[accRole];
            }

            return role;
        }

        // When the tree is created the role is cached.  
        // If there was an exception when the role was retreived return it.
        public AccRole GetRole(out Exception exception)
        {
            exception = _role as Exception;
            return GetRole();
        }

        // When the tree is created the state is cached.  This returns a sanitized version of name
        public AccState[] GetStates()
        {
            List<AccState> states = new List<AccState>();

            if (_state is int)
            {
                foreach (int state in StateMap.Keys)
                {
                    if (Win32API.IsBitSet((int)_state, state))
                    {
                        states.Add((AccState)StateMap[state]);
                    }
                }
            }

            return states.ToArray();
        }

        // When the tree is created the role is cached.  This returns a unsanitized version of state
        public AccState[] GetStates(out Exception exception)
        {
            exception = _state as Exception;
            return GetStates();
        }

        public string GetStatesString()
        {
            AccState[] states = GetStates();
            StringBuilder sb = new StringBuilder();
            foreach (AccState state in states)
            {
                sb.Append(Enum.GetName(typeof(AccState), state));
                sb.Append(" ");
            }
            return sb.ToString();
        }

        // This will convert a IAccessible state to a string
        public static string GetStatesString(int stateFlags)
        {
            StringBuilder sb = new StringBuilder();
            foreach (int state in StateMap.Keys)
            {
                if (Win32API.IsBitSet(stateFlags, state))
                {
                    sb.Append(Enum.GetName(typeof(AccState), state));
                    sb.Append(" ");
                }
            }

            return sb.ToString();
        }


        // When the tree is created the the hwnd for the element is cached.  This simply returns it.
        public IntPtr GetHwnd()
        {
            return _hwnd;
        }

        // This is used for logging to create an ID of sorts that will be fairly reliable in 
        // suppressing errors.  It is only called when an event log is created.
        public string CreateParentChain()
        {
            StringBuilder parentChain = new StringBuilder();
            MsaaElement parent = _parent;

            while (parent != null)
            {
                String name = parent.GetName();
                parentChain.Append(name);
                if (parent._parent != null && name != "")
                {
                    parentChain.Append(PARENT_CHAIN_SEPERATOR);
                }

                parent = parent._parent;
            }

            return parentChain.ToString();
        }


        public static MsaaElement FindMsaaElement(Accessible accessibleObject)
        {
            MsaaElement element = null;

            if (accessibleObject == null)
            {
                return element;
            }

            // first, let's create a chain of this object's parents
            Accessible curAccessible = accessibleObject;

            List<Accessible> accList = new List<Accessible>();
            while (curAccessible != null && !_rootElement.Compare(curAccessible))
            {
                accList.Add(curAccessible);
                try
                {
                    curAccessible = curAccessible.Parent;
                }
                catch (Exception exception)
                {
                    // We need to log to the root since something truly unexception happened and recursion could happen in the logging
                    TestLogger.LogException(EventLevel.Warning, LogEvent.MethodExceptionOccured, typeof(MsaaElement), exception, null, MsaaElement.RootElement, "FindMsaaElement, so cannot identify the correct element, using the root element for visualization purpopes.  Analyze the error after this may clear this issue.", exception.Message);

                    // if we get an exception getting the parents the likelihood of finding the element 
                    // on the way back down is not very good so just return false
                    return element;
                }

            }

            // parentChain now contains accessibleObject and its parents up to, but not including, _accessible
            // we want to start at the root
            accList.Reverse();

            MsaaElement currentElement = MsaaElement.RootElement;
            for (int listIndex = 0; listIndex < accList.Count; listIndex++)
            {
                try
                {
                    foreach (MsaaElement child in currentElement.Children)
                    {
                        if (child.Compare(accList[listIndex]))
                        {
                            element = child;
                            currentElement = child;
                            break;
                        }
                    }
                }
                catch (Exception exception)
                {
                    // We need to log to the root since something truly unexception happened and recursion could happen in the logging
                    TestLogger.LogException(EventLevel.Warning, LogEvent.MethodExceptionOccured, typeof(MsaaElement), exception, null, MsaaElement.RootElement, "FindMsaaElement");
                }
            }

            return element;
        }

        
        public class CircularReferenceInMsaaTree : Exception
        {
            public CircularReferenceInMsaaTree()
            {
            }
        
        }
    }
}
