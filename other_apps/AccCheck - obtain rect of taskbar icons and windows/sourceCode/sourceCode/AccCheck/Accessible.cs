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
using System.Reflection;

namespace AccCheck
{
    // this class is a wrapper for the IAccessible interface and some MSAA APIs
    // It insulates the rest of the code from having to deal with the child Id vs a 
    // real IAccessible usage.  This is easy to get wrong.  An IAccessible is really 
    // the IAccessible interface plus the child id.
    public class Accessible
    {        
        private IAccessible _accessible;
        private int _childId;
        private const int S_OK = 0;

        public Accessible(IAccessible accessible, int childId) 
        {
            System.Diagnostics.Debug.Assert(accessible != null);

            _accessible = accessible;
            _childId = childId;
        }

        public static int FromPoint(Point pt, out Accessible accessible)
        {
            accessible = null;
            IAccessible acc = null;
            object childId = null;
            int hr = Win32API.AccessibleObjectFromPoint(new Win32API.POINT(pt.X, pt.Y), ref acc, ref childId);
            
            if (hr == Win32API.S_OK && acc != null)
            {
                if (!(childId is int))
                {
                    throw new VariantNotIntException(childId);
                }
                accessible = new Accessible(acc, (int)childId);
            }

            return hr;
        }

        public static int FromWindow(IntPtr hwnd, out Accessible accessible)
        {
            accessible = null;
            Guid guid = Win32API.IID_IAccessible;

            object obj = null;
            int hr = Win32API.AccessibleObjectFromWindow(hwnd, Win32API.ObjIdWindow, ref guid, ref obj);

            IAccessible acc = obj as IAccessible;

            if (acc != null)
            {
                accessible  = new Accessible(acc, 0);
            }
            
            return hr;
        }
        
        public static int FromEvent(IntPtr hwnd, int idObject, int idChild, out Accessible accessible)
        {
            accessible = null;
            object childId = null;
            IAccessible acc = null;
            int hr = Win32API.AccessibleObjectFromEvent(hwnd, idObject, idChild, ref acc, ref childId);

            if (acc != null)
            {
                accessible = new Accessible(acc, (int)childId);
            }

            return hr;
        }

        public IAccessible IAccessible
        {
            get
            {
                return _accessible;
            }
        }

        public int ChildId
        {
            get
            {
                return _childId;
            }
        }
        
        public Accessible Parent
        {
            get
            {
                object parent = null;

                if (_childId == Win32API.CHILD_SELF)
                {
                    IAccessible parentIAccessible = _accessible.accParent as IAccessible;
                    if (parentIAccessible == null)
                    {
                        parent = null;
                    }
                    else
                    {
                        parent = new Accessible(parentIAccessible, Win32API.CHILD_SELF);
                    }
                }
                else
                {
                    parent = new Accessible(_accessible, Win32API.CHILD_SELF);
                }

                return parent as Accessible;
            }
        }

        public int Children(out Accessible [] accessible)
        {
            accessible = null;
            if (_childId != Win32API.CHILD_SELF)
            {
                accessible = new Accessible [] {};
                return S_OK;
            }
            
            int childrenCount = _accessible.accChildCount;
            if (childrenCount < 0)
            {
                throw new ChildCountInvalidException(childrenCount);
            }
            
            object [] children = new object[childrenCount];

            int hr = Win32API.AccessibleChildren(_accessible, 0, childrenCount, children, out childrenCount);

            if (hr == Win32API.S_OK)
            {
                accessible = new Accessible[childrenCount];
                int i = 0;
                foreach (object child in children)
                {
                    if (child != null)
                    {
                        if (child is IAccessible)
                        {
                            accessible[i++] = new Accessible((IAccessible)child, Win32API.CHILD_SELF);
                        }
                        else if (child is int)
                        {
                            accessible[i++] = new Accessible(_accessible, (int)child);
                        }
                    }
                }

                // null children don't occur very often but if they do it stops us from going on
                // So keep track of them so we can reallocate the array if necessary
                if (childrenCount != i)
                {
                    // if we had some null chilren create a smaller array to send the 
                    // children back in.
                    Accessible [] accessibleNew = new Accessible[i];
                    Array.Copy(accessible, accessibleNew, i);
                    accessible = accessibleNew;
                }
            }
            
            return hr;

        }
        
        public string Name
        {
            get
            {
                string name = _accessible.get_accName(_childId);

                //This could be returning a valid empty string or null
                return name;
            }
        }

        public int Role
        {
            get
            {
                object role = _accessible.get_accRole(_childId);

                if (!(role is int))
                {
                    throw new VariantNotIntException(role);
                }

                return (int)role;
            }
        }

        public int State
        {
            get
            {
                object state = _accessible.get_accState(_childId);

                if (!(state is int))
                {
                    throw new VariantNotIntException(state);
                }

                return (int)state;
            }
        }

        public string Value
        {
            get
            {
                string value = _accessible.get_accValue(_childId);
                
                // Both null and empty string are ok just force it to be
                // an empty string to make comparing these easier
                if (value == null)
                {
                    value = "";
                }
                
                return value;
            }
        }

        public Rectangle Location
        {
            get
            {
                int left = 0;
                int top = 0;
                int width = 0;
                int height = 0;
                _accessible.accLocation(out left, out top, out width, out height, _childId);

                Rectangle location = new Rectangle(left, top, width, height);

                return location;
            }
        }
        
        public string KeyboardShortcut
        {
            get
            {
                return _accessible.get_accKeyboardShortcut(_childId);
            }
        }

        public void Select(int flag)
        {
            _accessible.accSelect(flag, _childId);
        }
        
        public IntPtr Hwnd
        {
            get
            {
                IntPtr hwnd = IntPtr.Zero;
                Win32API.WindowFromAccessibleObject(_accessible, ref hwnd);

                return hwnd;
            }
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
            Rectangle thisLoc;
            try 
            {
                thisLoc = Location;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    thisLoc = Rectangle.Empty;
                }
                else
                {
                    throw;
                }
            }
            if (!thisLoc.Equals(loc))
            {
                return false;
            }
            

            string name = "";
            try 
            {
                name = acc.Name;
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
            string thisName = "";
            try 
            {
                thisName = Name;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    thisName = "";
                }
                else
                {
                    throw;
                }
            }
            if (thisName != name)
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
            int thisRole = 0;
            try 
            {
                thisRole = Role;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    thisRole = 0;
                }
                else
                {
                    throw;
                }
            }
            if (thisRole != role)
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
            string thisValue = "";
            try 
            {
                thisValue = Value;
            }
            catch (Exception e)
            {
                if (Accessible.IsExpectedException(e))
                {
                    thisValue = "";
                }
                else
                {
                    throw;
                }
            }
            if (thisValue != value)
            {
                return false;
            }
        
            return true;
        }
        
        // Create a sting that has the list of the name of the parents to the root
        public string CreateParentChain()
        {
            List<string> chain = new List<string>();
            Accessible acc = Parent;
            const int sanityCheckLimit = 100;
            int loopCount = 0;
            
            while (acc != null)
            {
                loopCount++;
                string name = "";
                try 
                {
                    name = Name;
                }
                catch (Exception e)
                {
                    if (IsExpectedException(e))
                    {
                        name = "";
                    }
                    else
                    {
                        throw;
                    }
                }

                if (!string.IsNullOrEmpty(name))
                {
                    chain.Add(name);
                }

                try 
                {
                    acc = Parent;
                }
                catch (Exception e)
                {
                    if (IsExpectedException(e))
                    {
                        acc = null;
                    }
                    else
                    {
                        throw;
                    }
                }

                if (loopCount > sanityCheckLimit)
                {
                    break;
                }
            }
            
            chain.Reverse();
            
            return string.Join(".", chain.ToArray());
        }

        override public string ToString()
        {
            string name = "";
            try 
            {
                name = Name;
                if (name == null)
                    name = "";
            }
            catch (Exception e)
            {
                if (IsExpectedException(e))
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
                role = Role;
            }
            catch (Exception e)
            {
                if (IsExpectedException(e))
                {
                    role = 0;
                }
                else
                {
                    throw;
                }
            }

            StringBuilder roleText = new StringBuilder(255);
            Win32API.GetRoleText(role, roleText, (uint)roleText.Capacity);

            return string.Format("[IAccessible.ChildId({0}).Role({1}).Name({2})]", _childId, roleText, name);
        }

        // these are the exceptions that we expect the caller will determine if is ok or not.
        public static bool IsExpectedException(Exception e)
        {
            if ( e is NotImplementedException || 
                e is COMException || 
                e is VariantNotIntException || 
                e is ArgumentException || 
                e is InvalidCastException ||
                e is UnauthorizedAccessException)
            {
                return true;
            }

            return false;
        }
    }

        
    public class VariantNotIntException : Exception
    {
        string _variantType;

        public VariantNotIntException(object variant)
        {
            _variantType = variant == null ? "[NULL]" : variant.GetType().ToString();
        }

        public string VariantType
        {
            get
            {
                return _variantType;
            }
        }
        
    }
    
    public class ChildCountInvalidException : Exception
    {
        int _childCount;

        public ChildCountInvalidException(int childCount)
        {
            _childCount= childCount;
        }

        public int ChildCount
        {
            get
            {
                return _childCount;
            }
        }
        
    }

}
