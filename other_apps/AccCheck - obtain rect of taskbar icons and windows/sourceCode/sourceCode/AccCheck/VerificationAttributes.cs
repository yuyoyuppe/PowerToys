// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;

namespace AccCheck.Verification
{
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false, Inherited = false)]
    public class VerificationAttribute : Attribute
    {
        private String _title;
        private String _description;
        private String _group;
        private bool _needsUI;
        private bool _canVisualize;
        private int _priority;
        private bool _isIntrusive;

        public string UniqueIdentifier
        {
            get { return new StringBuilder(_group + "." + _title).Replace(" ", "").ToString(); }
        }

        public VerificationAttribute(String title, String description, bool needsUI, String group)
        {
            _title = title;
            _description = description;
            _needsUI = needsUI;
            _group = group;
            _canVisualize = false;
            _priority = 0;
            _isIntrusive = false;
        }

        public VerificationAttribute(String title, String description)
        {
            _title = title;
            _description = description;
            _needsUI = false;
            _canVisualize = false;
            _priority = 0;
            _isIntrusive = false;
        }

        public virtual String Title
        {
            get
            {
                return _title;
            }
        }

        public virtual String Description
        {
            get
            {
                return _description;
            }
        }

        public virtual bool NeedsUI
        {
            get
            {
                return _needsUI;
            }
            set
            {
                _needsUI = value;
            }
        }

        public virtual String Group
        {
            get
            {
                return _group;
            }
            set
            {
                _group = value;
            }
        }
        public virtual bool CanVisualize
        {
            get
            {
                return _canVisualize;
            }
            set
            {
                _canVisualize = value;
            }
        }
        public virtual int Priority
        {
            get
            {
                return _priority;
            }
            set
            {
                _priority = value;
            }
        }
        public virtual bool IsIntrusive
        {
            get
            {
                return _isIntrusive;
            }
            set
            {
                _isIntrusive = value;
            }
        }
    }
}
