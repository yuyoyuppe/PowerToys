// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using AccCheck;
using AccCheck.Verification;


namespace AccCheck.Logging
{
    public struct LogMessage
    {
        public LogMessage(String description, String id, String htmlHelpPage)
        {
            Description = description;
            Id = id;
            HelpId = id;
            HtmlHelpPage = htmlHelpPage;
            Priority = 0;
        }

        public LogMessage(String description, String id, String htmlHelpPage, int priority)
        {
            Description = description;
            Id = id;
            HelpId = id;
            HtmlHelpPage = htmlHelpPage;
            Priority = priority;
        }

        public LogMessage(String description, String id)
        {
            Description = description;
            Id = id;
            HelpId = id;
            HtmlHelpPage = string.Empty;
            Priority = 0;
        }

        public LogMessage(String description, String id, int priority)
        {
            Description = description;
            Id = id;
            HelpId = id;
            HtmlHelpPage = string.Empty;
            
            if (priority < 1)
            {
                priority = 1;
            }
            else if (priority > 3)
            {
                priority = 3;
            }                
            Priority = priority;
        }

        public LogMessage(String description)
        {
            Description = description;
            Id = string.Empty;
            HelpId = string.Empty;
            HtmlHelpPage = string.Empty;
            Priority = 0;
        }

        // Id of the log event.
        public String Id;

        // Id within the help file.  In our case currently this is the bookmark.
        public String HelpId;

        // The priority of this error message.
        public int Priority;

        // Format string to display to the userof the nature of the error.
        public String Description;

        // Name of html help file where log message is explained more clearly.
        public String HtmlHelpPage;
    }

    //
    // This class represents a single event in the log.
    // This is serialized into an xml document used in suppression.
    // One of the restrictions for xml serialization is the properties must be read write
    // Therefore there are properties that have get and set on them even though the 
    // set may never be used.
    //
    [XmlRoot(ElementName = "Event")]
    public class LogEvent : IEquatable<LogEvent>, ILogEvent
    {
        /// <summary>Exception occured in ({0})</summary>
        public static LogMessage MethodExceptionOccured = new LogMessage( "Exception occured \"{0}\"", "MethodExceptionOccured");

        /// <summary>A child returned from AccessibleChildren had an invalid type of {0}</summary>
        public static LogMessage InvalidAccessibleChild = new LogMessage("A child returned from AccessibleChildren had an invalid type of {0}", "InvalidAccessibleChild");

        /// <summary>The variant returned from {0} should be an {1} but is a {2}</summary>
        public static LogMessage VariantNotInt = new LogMessage("The variant returned from {0} should be an {1} but is a {2}.", "VariantNotInt");

        #region Private members

        private EventLevel _level;
        private String _eventID;
        private String _extra;
        private String _text;
        private String _helpIndex;
        private String _helpPage;
        private Rectangle _rectangle;
        private String _routine;
        private String _accname;
        private String _accvalue;
        private string _accrole;
        private string _accstate;
        private string _frameworkId = string.Empty;
        private string _providerDescription = string.Empty;
        private String _classname;
        private String _parentChain;
        private string _parentQueryString;
        private string _queryString;
        private string _globalizedQueryString;
        private DateTime _timestamp;
        private int _sequenceNumber;
        private bool _suppressed = false;
        private int _priority = 0;
        private Accessible _accessible;
        private object _automationElement;
        #endregion

        // LogEvent constructor
        public LogEvent(
            EventLevel level,
            String eventID,
            String text,
            String extra,
            Rectangle rect,
            Type routine)
        {
            _timestamp = DateTime.Now;

            _level = level;
            _eventID = eventID;
            _extra = extra;
            _text = text;
            _rectangle = rect;
            _parentChain = "";
            _routine = (routine != null ? routine.FullName : "");

        }

        public LogEvent(
            EventLevel level,
            string routine)
        {
            _timestamp = DateTime.Now;

            _level = level;
            _eventID = string.Empty;
            _extra = string.Empty;
            _text = string.Empty;
            _rectangle = Rectangle.Empty;
            _parentChain = "";
            _routine = routine;
        }

        public LogEvent(
            EventLevel level,
            LogMessage logMessage,
            String text,
            String extra,
            Rectangle rect,
            Type routine)
        {
            _timestamp = DateTime.Now;

            _level = level;
            _eventID = logMessage.Id;
            _extra = extra;
            _text = text;
            _priority = logMessage.Priority;
            _helpIndex = logMessage.HelpId;
            _helpPage = logMessage.HtmlHelpPage;
            _rectangle = rect;
            _parentChain = "";
            _routine = (routine != null ? routine.FullName : "");

        }

        public LogEvent(
            EventLevel level,
            String eventID,
            String text,
            String extra)
        {
            _timestamp = DateTime.Now;

            _level = level;
            _eventID = eventID;
            _extra = extra;
            _text = text;
            _parentChain = "";

        }



        public LogEvent()
        {
            _timestamp = DateTime.Now;
        }

        

        public void SetIAccessibleProps(Accessible acc, string name, string value, string role, string state)
        {
            _accessible = acc;

            // Null causes problems for the xml parsing
            _accname = name.Replace('\0', ' ');
            _accvalue = value.Replace('\0', ' ');
            _accrole = role;

            _accstate = state;
        }

        public void SetUIAProps(object el, string name, string value, string controlType, string state, string frameworkId, string providerDescription)
        {
            _automationElement = el;

            // Null causes problems for the xml parsing
            _accname = name.Replace('\0', ' ');
            _accvalue = value.Replace('\0', ' ');

            _accrole = controlType;
            _accstate = state;
            _frameworkId = frameworkId;
            _providerDescription = providerDescription;
        }

        public void SetFromWindow(IntPtr hwnd, IntPtr rootHwnd)
        {
            StringBuilder sb = new StringBuilder(1024);

            try
            {
                Win32API.GetClassName(hwnd, sb, sb.Capacity);
            }
            catch
            {
                // don't need to report errors with this call
                // There are some windows in office that make this overwrite the StringBuild object.
            }

            _classname = sb.ToString();

            // if we can't find a rectangle using IAccessible, try using GetWindowRect
            if (_rectangle == Rectangle.Empty)
            {
                Win32API.RECT rc = new Win32API.RECT();
                Win32API.GetWindowRect(hwnd, ref rc);
                Win32API.POINT pt = new Win32API.POINT(rc.left, rc.top);
                IntPtr hwndParent = Win32API.GetAncestor(hwnd, Win32API.GA_ROOT);
                Win32API.RECT rcParent = new Win32API.RECT();
                Win32API.GetWindowRect(hwndParent, ref rcParent);
                // calculate rectangle according to ancestral top-level window
                _rectangle = new Rectangle(
                    rc.left - rcParent.left,
                    rc.top - rcParent.top,
                    rc.right - rc.left,
                    rc.bottom - rc.top);
            }
        }

        public override String ToString()
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(String.Format("[{0}] {1}", Enum.GetName(typeof(EventLevel), _level), _eventID));
            if (_routine != null)
            {
                sb.AppendLine(String.Format("\tRoutine: {0}", _routine));
            }
            
            if (_text != String.Empty)
            {
                sb.AppendLine(String.Format("\tText: {0}", _text));
            }

            if (!(String.IsNullOrEmpty(_accname) && String.IsNullOrEmpty(_accrole) && String.IsNullOrEmpty(_accstate)))
            {
                sb.AppendLine(String.Format("\tWhat: Name({0}), Role({1}), State({2})", _accname,  _accrole, _accstate));
            }
            
            if (!_rectangle.Equals(Rectangle.Empty))
            {
                sb.AppendLine(String.Format("\tWhere: {0}", _rectangle));
            }
            
            if (_extra != String.Empty)
            {
                sb.Append(String.Format("\tStack: {0}", _extra));
            }
            
            return sb.ToString();
        }

        #region Properties
        [XmlIgnore]
        public Accessible Accessible
        {
            get
            {
                return _accessible;
            }
        }

        [XmlIgnore]
        public object AutomationElement
        {
            get
            {
                return _automationElement;
            }
        }

        public EventLevel Level
        {
            get
            {
                return _level;
            }
            set
            {
                _level = value;
            }
        }
        public String EventID
        {
            get
            {
                return _eventID;
            }
            set
            {
                _eventID = value;
            }
        }
        public String ExtraInformation
        {
            get
            {
                return _extra;
            }
            set
            {
                _extra = value;
            }
        }
        public String Text
        {
            get
            {
                return _text;
            }
            set
            {
                _text = value;
            }
        }
        public String HelpIndex
        {
            get
            {
                return _helpIndex;
            }
            set
            {
                _helpIndex = value;
            }
        }
        public String HelpPage
        {
            get
            {
                return _helpPage;
            }
            set
            {
                _helpPage = value;
            }
        }
        public Rectangle Rectangle
        {
            get
            {
                return _rectangle;
            }
            set
            {
                _rectangle = value;
            }
        }
        public String ParentChain
        {
            get
            {
                return _parentChain;
            }
            set
            {
                _parentChain = value;
            }
        }
        public String QueryString
        {
            get
            {
                return _queryString;
            }
            set
            {
                _queryString = value;
            }
        }
        public String ParentQueryString
        {
            get
            {
                return _parentQueryString;
            }
            set
            {
                _parentQueryString = value;
            }
        }
        public String GlobalizedQueryString
        {
            get
            {
                return _globalizedQueryString;
            }
            set
            {
                _globalizedQueryString = value;
            }
        }
        public String VerificationRoutine
        {
            get
            {
                return _routine;
            }
            set
            {
                _routine = value;
            }
        }
        public String Classname
        {
            get
            {
                return _classname;
            }
            set
            {
                _classname = value;
            }
        }
        public String AccName
        {
            get
            {
                return _accname;
            }
            set
            {
                // Null causes problems for the xml parsing
                _accname = value.Replace('\0', ' ');
            }
        }
        public String AccValue
        {
            get
            {
                return _accvalue;
            }
            set
            {
                // Null causes problems for the xml parsing
                _accvalue = value.Replace('\0', ' ');
            }
        }
        public string AccRole
        {
            get
            {
                return _accrole;
            }
            set
            {
                _accrole = value;
            }
        }
        public string AccState
        {
            get
            {
                return _accstate;
            }
            set
            {
                _accstate = value;
            }
        }

        public string FrameworkId
        {
            get 
            { 
                return _frameworkId; 
            }
            set 
            { 
                _frameworkId = value; 
            }
        }
        public string ProviderDescription
        {
            get 
            { 
                return _providerDescription; 
            }
            set 
            { 
                _providerDescription = value; 
            }
        }
        public DateTime Timestamp
        {
            get
            {
                return _timestamp;
            }
            set
            {
                _timestamp = value;
            }
        }
        [XmlIgnore]
        public int SequenceNumber
        {
            // Only used by the Accumulating logger and the UI
            get
            {
                return _sequenceNumber;
            }
            set
            {
                _sequenceNumber = value;
            }
        }
        public bool Suppressed
        {
            get
            {
                return _suppressed;
            }
            set
            {
                _suppressed = value;
            }
        }
        public int Priority
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

        #endregion

        #region IEquatable<LogEvent> Members

        public bool Equals(LogEvent other)
        {
            if (other == null)
            {
                return false;
            }
            
            bool isEqual = false;


            // The "this" object is the LogEvent from the suppression file and the "other" object
            // is the LogEvent from the current verification.
            if ((this._eventID == other._eventID) && (this._routine == other._routine))
            {
                IQueryString queryString = VerificationManager.GetQueryStringObject();
                if (!String.IsNullOrEmpty(this.GlobalizedQueryString) && queryString != null)
                {
                    isEqual = queryString.Compare(this.GlobalizedQueryString, other.QueryString);
                }
                // if queryString is null that means the add-in is not being used so other.QueryString would be null
                // So we want to fall back to the parentChain compare.
                else if (!String.IsNullOrEmpty(this.QueryString)&& queryString != null)
                {
                    isEqual = this.QueryString == other.QueryString;
                }
                else if (!String.IsNullOrEmpty(this._parentChain))
                {
                    isEqual = (this._accname == other._accname) && 
                              (this._classname == other._classname) && 
                              (this._parentChain == other._parentChain);
                }
                else
                {
                    // If there is no identifier then treat this as a global suppression
                    isEqual = (this._classname == other._classname) && 
                              (this._accrole == other._accrole);
                }
            }
                
             

            return isEqual;
        }

        #endregion
    }
}
