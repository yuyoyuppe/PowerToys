// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Reflection;
using System.Windows.Forms;
using System.Text;
using System.Drawing;
using System.Diagnostics;
using System.ComponentModel;
using System.IO;
using System.Xml;
using System.Xml.Serialization;
using AccCheck;
using AccCheck.Verification;


namespace AccCheck.Logging
{
    // This class is used to persist the log data that has been gathered for the run of a set of verification on one hwnd.  
    // This can be serialized and de-serialized so that a user can open a log file and view all the errors and navigate 
    // the trees created by the verifications.
    [XmlRoot(ElementName = "LogFile")]
    public class LogFile 
    {
        private const int _versionMajorRelease = 2;
        private const int _versionMinorRelease = 1;
        private int _versionMajor = _versionMajorRelease;
        private int _versionMinor = _versionMinorRelease;
        private byte[] _serializedScreenshot;
        private List<LogEvent> _events;
        private List<SerializedTree> _trees;
        private List<string> _loadedModules;
        private List<string> _ambiguousResources;
        private int _processId = 0;
        private string _windowClass;
        private string _windowName;
        private DateTime _startTime = DateTime.MinValue;
        private DateTime _stopTime = DateTime.MinValue;
        private Point _rootOffset;
        private ProcessFileVersionInfo _fileVersionInfo = new ProcessFileVersionInfo();
        private LogSystemInformation _systemInformation = new LogSystemInformation();

        // used to deserialize
        public LogFile()
        {
        }

        public LogFile(IntPtr hwnd)
        {
            if (Win32API.IsWindow(hwnd))
            {
                Win32API.GetWindowThreadProcessId(hwnd, out _processId);
                if (_processId != 0)
                {
                    _loadedModules = new List<string>();
                    try
                    {
                        Process process = Process.GetProcessById(_processId);
                        if (process != null)
                        {
                            _fileVersionInfo = new ProcessFileVersionInfo(process.MainModule.FileVersionInfo);
                            ProcessModuleCollection modules = process.Modules;
                            foreach (ProcessModule pm in modules)
                            {
                                _loadedModules.Add(pm.FileName);
                            }
                        }
                    }
                    catch (Win32Exception e)
                    {
                        if (e.NativeErrorCode == 5)
                        {
                            MessageBox.Show("Access denied, try running AccChecker elevated");
                        }
                    }
                   
                    StringBuilder sb = new StringBuilder(1024);
                    Win32API.GetClassName(hwnd, sb, sb.Capacity);
                    _windowClass = sb.ToString();
                    _windowName = Win32API.GetWindowText(hwnd);
                    Win32API.RECT rootRect = new Win32API.RECT();
                    Win32API.GetWindowRect(hwnd, ref rootRect);
                    _rootOffset = new Point(-rootRect.left, -rootRect.top);
                }
            }
        }

        void AddQueryString(List<string> queryStrings, string newString)
        {
            if (!queryStrings.Contains(newString))
            {
                queryStrings.Add(newString);
            }
        }

        void AddRangeQueryString(List<string> queryStringList, List<string> newQueryStringList)
        {
            foreach (string queryString in newQueryStringList)
            {
                AddQueryString(queryStringList, queryString);
            }
        }

        public void PrepareToSerialize(List<LogEvent> logEvents)
        {
            _events = new List<LogEvent>();

            _events.AddRange(logEvents);
            
            IQueryString queryString = VerificationManager.GetQueryStringObject();
            if (queryString != null)
            {
                List<string> queryStrings = new List<string>();

                //Add the query strings we have found for the events hit
                if (logEvents != null && logEvents.Count > 0)
                {
                    foreach (LogEvent eventLogged in logEvents)
                    {
                        AddQueryString(queryStrings, eventLogged.QueryString);
                    }
                }

                // Collecting all results to include passes?  THis may go away once we remove the 'include query strings' from the interfaces.
                if (VerificationManager.IncludePassResultsInLog)
                {
                    // Will be enumerating through all the elements for every verification, only do verifications we care about(active).
                    List<VerificationRoutineWrapper> routines = new List<VerificationRoutineWrapper>();
                    foreach (VerificationRoutineWrapper routine in VerificationManager.RoutinesRegistered)
                    {
                        if (routine.Active)
                        {
                            routines.Add(routine);
                        }
                    }

                    // Add elements/passes that are not normally recorded when running verificaitons.  Optimize 
                    // and only add the list after we have gone through all the trees.  Don't want to walk over just added matches.
                    List<LogEvent> passList = new List<LogEvent>();
                    foreach (SerializedTree tree in _trees)
                    {
                        passList.AddRange(GetPassResultsList(logEvents, tree, routines));
                        AddRangeQueryString(queryStrings, tree.ElementsVerifiedQueryString);
                    }

                    _events.AddRange(passList);
                }

                if (_loadedModules.Count > 0 && queryStrings.Count > 0)
                {
                    List<string> ambiguousResources;
                    string [] globalizedQueryStrings = queryString.Globalize(queryStrings, _loadedModules, out ambiguousResources);

                    for (int i = 0; i < _events.Count; i++)
                    {
                        _events[i].GlobalizedQueryString = globalizedQueryStrings[queryStrings.IndexOf(_events[i].QueryString)];
                    }

                    if (ambiguousResources.Count > 0)
                    {
                        _ambiguousResources = ambiguousResources;
                    }
                    else
                    {
                        _ambiguousResources = null;
                    }
                }
            }
        }

        // <summary>Call this to add the results of passes to you log file</summary>
        public List<LogEvent> GetPassResultsList(List<LogEvent> logEvents, SerializedTree tree, List<VerificationRoutineWrapper> routines)
        {
            List<LogEvent> newLogEvents = new List<LogEvent>();

            foreach (string queryString in tree.ElementsVerifiedQueryString)
            {
                foreach (VerificationRoutineWrapper routine in routines)
                {
                    // Look to see if there are any results already added to the log events.  Elements that
                    // passed specific verifications will not log anything in the log events. If there 
                    // are no entries, then this element passed for this verification.
                    if (!HasVerificationBeenRecordedForElement(queryString, routine, logEvents))
                    {
                        // We don't know anything about the properties of the element anymore, so just log a pass.
                        LogEvent log = new LogEvent(EventLevel.Pass, routine.Type.FullName);
                        log.QueryString = queryString;
                        log.Priority = routine.Priority;
                        newLogEvents.Add(log);
                    }
                }
            }

            return newLogEvents;
        }

        private bool HasVerificationBeenRecordedForElement(string queryString, VerificationRoutineWrapper routine, List<LogEvent> logEvents)
        {
            if (logEvents != null)
            {
                foreach (LogEvent le in logEvents)
                {
                    if ((le.QueryString == queryString) && (le.VerificationRoutine == routine.Type.FullName))
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public bool IsVersionOk()
        {
            bool compatibleVersion = false;
            if (_versionMajor == _versionMajorRelease)
            {
                compatibleVersion = true;
            }
            
            return compatibleVersion;
        }

        public int MajorVersion
        {
            get
            {
                return _versionMajor;
            }
            set
            {
                _versionMajor = value;
            }
        }
        public int MinorVersion
        {
            get
            {
                return _versionMinor;
            }
            set
            {
                _versionMinor = value;
            }
        }
        public LogSystemInformation SystemInformation
        {
            get
            {
                return _systemInformation;
            }
            set
            {
                _systemInformation = value;
            }
        }
        public ProcessFileVersionInfo FileVersionInfo 
        {
            get
            {
                return _fileVersionInfo;
            }
            set
            {
                _fileVersionInfo = value;
            }
        }
        [XmlIgnore]
        public Bitmap ScreenShot
        {
            get
            {
                if (_serializedScreenshot == null)
                {
                    return null;
                }
                else
                {
                    return new Bitmap(new MemoryStream(_serializedScreenshot));
                }
            }
            set
            {
                if (value != null)
                {
                    MemoryStream ms = new MemoryStream();
                    value.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    _serializedScreenshot = ms.GetBuffer();
                }
            }
        }

        public List<VerificationRoutineWrapper> Verifications
        {
            get
            {
                return new List<VerificationRoutineWrapper>(VerificationManager.RoutinesRegistered);
            }
            set
            {
                // We need a 'set' to be able to properly use XML serialization.  We only need it on the
                // save so we can see what verification was registered to run, don't care about the load.
            }
        }

        public Point RootOffset
        {
            get
            {
                return _rootOffset;
            }
            set
            {
                _rootOffset = value;
            }
        }
        [XmlIgnore]
        public List<string> LoadedModules
        {
            get
            {
                return _loadedModules;
            }
            set
            {
                _loadedModules = value;
            }
        }
        public string WindowClass
        {
            get
            {
                return _windowClass;
            }
            set
            {
                _windowClass = value;
            }
        }
        public string WindowName
        {
            get
            {
                return _windowName;
            }
            set
            {
                _windowName = value;
            }
        }
        public DateTime StartTime
        {
            get
            {
                return _startTime;
            }
            set
            {
                _startTime = value;
            }
        }
        public DateTime StopTime
        {
            get
            {
                return _stopTime;
            }
            set
            {
                _stopTime = value;
            }
        }
        public byte[] SerializedScreenShot
        {
            get
            {
                return _serializedScreenshot;
            }
            set
            {
                _serializedScreenshot = value;
            }
        }
        public List<SerializedTree> Trees
        {
            get
            {
                return _trees;
            }
            set
            {
                _trees = value;
                if (_trees != null)
                {
                    // go from the back so the index does not change if we remove an item
                    for (int i = _trees.Count - 1; i >= 0; i--)
                    {
                        _trees[i].LoadXmlFromTree();
                        if (String.IsNullOrEmpty(_trees[i].SerializedTreeXml))
                        {
                            _trees.RemoveAt(i);
                        }
                    }
                }
            }
        }
        public List<LogEvent> LogEvents
        {
            get
            {
                return _events;
            }
            set
            {
                _events = value;
            }
        }
        public List<string> AmbiguousResources
        {
            get
            {
                return _ambiguousResources;
            }
            set
            {
                _ambiguousResources = value;
            }
        }
    }

    // This class represents a tree that was created by a verification routine in a verification dll the 
    // verification dll owns the formation of the tree and it may change from dll to dll.
    // For example the MSAA tree will contain different information from the UIA tree and 
    // only the respective verification routine understands the format.
    public class SerializedTree
    {
        private string _verificationDll;
        private string _routine;
        private string _serializedTreeXml;
        private ISerializeTree _serializedTreeObject;
        private List<string> _elementsVerifiedQueryString = new List<string>();

        public SerializedTree(string verificationDll, string routine, ISerializeTree serializedTreeObject)
        {
            _verificationDll = verificationDll;
            _routine = routine;
            _serializedTreeXml = "";
            _serializedTreeObject = serializedTreeObject;            
        }
        
        public SerializedTree()
        {
        }

        public string VerificationDll
        {
            get { return _verificationDll; }
            set { _verificationDll = value; }
        }
        public string Routine
        {
            get { return _routine; }
            set { _routine = value; }
        }
        
        public string SerializedTreeXml
        {
            get { return _serializedTreeXml; }
            set { _serializedTreeXml = value; }
        }
        
        public void LoadXmlFromTree()
        {
            if (_serializedTreeObject != null && String.IsNullOrEmpty(_serializedTreeXml))
            {
                _serializedTreeXml = _serializedTreeObject.GetSerializedTree();
            }
        }

        public void SetXmlForTree(GraphicsHelper graphicsHelper)
        {
            if (_serializedTreeObject == null)
            {
                string verificationDll = "";
                try
                {
                    string currentDirectory = System.IO.Directory.GetCurrentDirectory() + "\\";
                    verificationDll = Path.Combine(currentDirectory, Path.GetFileName(_verificationDll));
                    
                    Assembly assembly = Assembly.LoadFile(verificationDll);
                    foreach (Type type in assembly.GetTypes())
                    {
                        if (Attribute.IsDefined(type, typeof(VerificationAttribute)) &&
                            type.ToString() == _routine &&
                            type.GetInterface("ISerializeTree") != null)
                        {
                            object vrObject = Activator.CreateInstance(type);
                            ISerializeTree verificationRoutine = vrObject as ISerializeTree;
                            if (verificationRoutine != null)
                            {
                                _serializedTreeObject = verificationRoutine;
                            }
                        }
                    }
                }
                catch (FileNotFoundException)
                {
                    MessageBox.Show("Verification dll " + verificationDll + " was referenced in this log but not found on the system");
                }
            }
            
            if (_serializedTreeObject != null)
            {
                _serializedTreeObject.SetSerializedTree(_serializedTreeXml, graphicsHelper);
            }
        }


        [XmlIgnore]
        public List<string> ElementsVerifiedQueryString
        {
            get 
            {
                return _serializedTreeObject == null ? new List<string>() : _serializedTreeObject.ElementsVerifiedQueryString; 
            }
            set 
            { 
                _serializedTreeObject.ElementsVerifiedQueryString = value; 
            }
        }

        public SerializedTree Clone()
        {
            SerializedTree tree = new SerializedTree((string)_verificationDll.Clone(), (string)_routine.Clone(), _serializedTreeObject);
            tree.LoadXmlFromTree();

            return tree;
        }


    }

    /// <summary>
    /// Wrapper class for FileVersionInfo since FileVersionInfo does not have a default constructor required for serialization
    /// </summary>
    public class ProcessFileVersionInfo
    {
        int _fileBuildPart;
        private string _fileDescription;
        int _fileMajorPart;
        int _fileMinorPart;
        string _fileName;
        int _filePrivatePart;
        string _fileVersion;
        string _internalName;
        bool _isDebug;
        bool _isPreRelease;
        string _language;
        string _legalCopyright;
        string _legalTrademarks;
        string _privateBuild;
        int _productBuildPart;
        int _productMajorPart;
        int _productMinorPart;
        string _productName;
        int _productPrivatePart;
        string _productVersion;
        string _specialBuild;

        //
        // Summary:
        //     Gets the build number of the file.
        //
        // Returns:
        //     A value representing the build number of the file or null if the file did
        //     not contain version information.
        public int FileBuildPart 
        {
            get 
            { 
                return _fileBuildPart; 
            }
            set 
            { 
                _fileBuildPart = value; 
            } 
        }
        //
        // Summary:
        //     Gets the description of the file.
        //
        // Returns:
        //     The description of the file or null if the file did not contain version information.
        public string FileDescription 
        {
            get 
            {
                return _fileDescription; 
            }
            set 
            { 
                _fileDescription = value; 
            } 
        }
        //
        // Summary:
        //     Gets the major part of the version number.
        //
        // Returns:
        //     A value representing the major part of the version number or null if the
        //     file did not contain version information.
        public int FileMajorPart 
        {
            get 
            {
                return _fileMajorPart; 
            }
            set 
            { 
                _fileMajorPart = value; 
            } 
        }
        //
        // Summary:
        //     Gets the minor part of the version number of the file.
        //
        // Returns:
        //     A value representing the minor part of the version number of the file or
        //     null if the file did not contain version information.
        public int FileMinorPart 
        {
            get 
            { 
                return _fileMinorPart; 
            }
            set 
            { 
                _fileMinorPart = value; 
            } 
        }
        //
        // Summary:
        //     Gets the name of the file that this instance of System.Diagnostics.FileVersionInfo
        //     describes.
        //
        // Returns:
        //     The name of the file described by this instance of System.Diagnostics.FileVersionInfo.
        public string FileName 
        {
            get 
            {
                return string.IsNullOrEmpty(_fileName) ? "(undefined)" : Path.GetFileName(_fileName);
            }
            set 
            {
                _fileName = value; 
            } 
        }
        //
        // Summary:
        //     Gets the file private part number.
        //
        // Returns:
        //     A value representing the file private part number or null if the file did
        //     not contain version information.
        public int FilePrivatePart 
        {
            get 
            { 
                return _filePrivatePart; 
            }
            set 
            { 
                _filePrivatePart = value; 
            } 
        }
        //
        // Summary:
        //     Gets the file version number.
        //
        // Returns:
        //     The version number of the file or null if the file did not contain version
        //     information.
        public string FileVersion 
        {
            get 
            { 
                return _fileVersion; 
            }
            set 
            { 
                _fileVersion = value; 
            } 
        }
        //
        // Summary:
        //     Gets the internal name of the file, if one exists.
        //
        // Returns:
        //     The internal name of the file. If none exists, this property will contain
        //     the original name of the file without the extension.
        public string InternalName 
        {
            get 
            { 
                return _internalName; 
            }
            set 
            { 
                _internalName = value; 
            } 
        }
        //
        // Summary:
        //     Gets a value that specifies whether the file contains debugging information
        //     or is compiled with debugging features enabled.
        //
        // Returns:
        //     true if the file contains debugging information or is compiled with debugging
        //     features enabled; otherwise, false.
        public bool IsDebug 
        {
            get 
            {
                return _isDebug; 
            }
            set 
            { 
                _isDebug = value; 
            } 
        }
        //
        // Summary:
        //     Gets a value that specifies whether the file is a development version, rather
        //     than a commercially released product.
        //
        // Returns:
        //     true if the file is prerelease; otherwise, false.
        public bool IsPreRelease 
        {
            get 
            {
                return _isPreRelease; 
            }
            set 
            {
                _isPreRelease = value; 
            } 
        }
        //
        // Summary:
        //     Gets the default language string for the version info block.
        //
        // Returns:
        //     The description string for the Microsoft Language Identifier in the version
        //     resource or null if the file did not contain version information.
        public string Language 
        {
            get 
            { 
                return _language; 
            }
            set 
            { 
                _language = value; 
            } 
        }
        //
        // Summary:
        //     Gets all copyright notices that apply to the specified file.
        //
        // Returns:
        //     The copyright notices that apply to the specified file.
        public string LegalCopyright 
        {
            get 
            { 
                return _legalCopyright; 
            }
            set 
            {
                _legalCopyright = value; 
            } 
        }
        //
        // Summary:
        //     Gets the trademarks and registered trademarks that apply to the file.
        //
        // Returns:
        //     The trademarks and registered trademarks that apply to the file or null if
        //     the file did not contain version information.
        public string LegalTrademarks 
        {
            get 
            { 
                return _legalTrademarks; 
            }
            set 
            {
                _legalTrademarks = value; 
            } 
        }
        //
        // Summary:
        //     Gets information about a private version of the file.
        //
        // Returns:
        //     Information about a private version of the file or null if the file did not
        //     contain version information.
        public string PrivateBuild 
        {
            get 
            {
                return _privateBuild; 
            }
            set 
            { 
                _privateBuild = value; 
            } 
        }
        //
        // Summary:
        //     Gets the build number of the product this file is associated with.
        //
        // Returns:
        //     A value representing the build number of the product this file is associated
        //     with or null if the file did not contain version information.
        public int ProductBuildPart 
        {
            get 
            { 
                return _productBuildPart; 
            }
            set 
            { 
                _productBuildPart = value; 
            } 
        }
        //
        // Summary:
        //     Gets the major part of the version number for the product this file is associated
        //     with.
        //
        // Returns:
        //     A value representing the major part of the product version number or null
        //     if the file did not contain version information.
        public int ProductMajorPart 
        {
            get 
            { 
                return _productMajorPart; 
            }
            set 
            {
                _productMajorPart = value; 
            } 
        }
        //
        // Summary:
        //     Gets the minor part of the version number for the product the file is associated
        //     with.
        //
        // Returns:
        //     A value representing the minor part of the product version number or null
        //     if the file did not contain version information.
        public int ProductMinorPart 
        {
            get 
            { 
                return _productMinorPart; 
            }
            set 
            { 
                _productMinorPart = value; 
            } 
        }
        //
        // Summary:
        //     Gets the name of the product this file is distributed with.
        //
        // Returns:
        //     The name of the product this file is distributed with or null if the file
        //     did not contain version information.
        public string ProductName 
        {
            get 
            { 
                return _productName; 
            }
            set 
            { 
                _productName = value;
            } 
        }
        //
        // Summary:
        //     Gets the private part number of the product this file is associated with.
        //
        // Returns:
        //     A value representing the private part number of the product this file is
        //     associated with or null if the file did not contain version information.
        public int ProductPrivatePart 
        {
            get 
            { 
                return _productPrivatePart; 
            }
            set 
            { 
                _productPrivatePart = value; 
            } 
        }
        //
        // Summary:
        //     Gets the version of the product this file is distributed with.
        //
        // Returns:
        //     The version of the product this file is distributed with or null if the file
        //     did not contain version information.
        public string ProductVersion 
        {
            get 
            { 
                return _productVersion; 
            }
            set 
            { 
                _productVersion = value; 
            } 
        }
        //
        // Summary:
        //     Gets the special build information for the file.
        //
        // Returns:
        //     The special build information for the file or null if the file did not contain
        //     version information.
        public string SpecialBuild 
        {
            get 
            { 
                return _specialBuild; 
            } 
            set 
            { 
                _specialBuild = value; 
            } 
        }
       
        // Needed for serialization
        public ProcessFileVersionInfo() { }

        public ProcessFileVersionInfo(FileVersionInfo fvi) 
        { 
            _fileBuildPart = fvi.FileBuildPart;
            _fileDescription = fvi.FileDescription;
            _fileMajorPart = fvi.FileMajorPart;
            _fileMinorPart = fvi.FileMinorPart;
            _fileName = fvi.FileName;
            _filePrivatePart = fvi.FilePrivatePart;
            _fileVersion = fvi.FileVersion;
            _internalName = fvi.InternalName;
            _isDebug = fvi.IsDebug;
            _isPreRelease = fvi.IsPreRelease;
            _language = fvi.Language;
            _legalCopyright = fvi.LegalCopyright;
            _legalTrademarks = fvi.LegalTrademarks;
            _privateBuild = fvi.PrivateBuild;
            _productBuildPart = fvi.ProductBuildPart;
            _productMajorPart = fvi.ProductMajorPart;
            _productMinorPart = fvi.ProductMinorPart;
            _productName = fvi.ProductName;
            _productPrivatePart = fvi.ProductPrivatePart;
            _productVersion = fvi.ProductVersion;
            _specialBuild = fvi.SpecialBuild;
        }
    }
}
