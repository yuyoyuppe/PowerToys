// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using AccCheck.Logging;
using AccCheck.Verification;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Xml.Serialization;

namespace AccCheck.Verification
{
    /*
     * VerificationManager is the glue that holds everything together.
     * It takes care of loading verification routines from DLLs at runtime
     * (It automatically loads the Verification*.dlls from the current dir).
     * It needs a LogList at construction, so you can set up the loggers you want.
     */
    [ComVisible(true), GuidAttribute("2A1B381C-9FAE-41e5-9FA6-F9A86BE02A52")]
    [ProgId("Microsoft.AccCheck.VerificationManager")]
    [ClassInterface(ClassInterfaceType.None)]
    public class VerificationManager : IVerificationManager
    {
        private static readonly String[] s_defaultVerifications = { "VerificationRoutines.dll", "UIAVerifications.dll" };

        // this is a wrapper for the loglist
        private SuppressionFilterLog _logger;

        private LogList _logList;
        private bool _allowUI = false;
        static List<VerificationRoutineWrapper> _routines = new List<VerificationRoutineWrapper>();
        private List<string> _verificationDlls = new List<string>();
        private List<ICache> _cacheMethodList = new List<ICache>();
        private List<SerializedTree> _serializableTreeList = new List<SerializedTree>();
        static IQueryString _queryString;
        private IntPtr _hwnd = IntPtr.Zero;
        private System.Drawing.Bitmap _screenShot = null;
        static private bool _enableQueryStringAddin = false;
        static private int _priority = 3;
        static private bool _includePassResultsInLog = false;
        private bool _processingMonitoredHwnd;
        private IntPtr _monitoredHwnd = IntPtr.Zero;
        private IntPtr _lastMonitoredHwndprocessed;
        private Timer _monitoringTimer = new Timer();
        private MonitorCallback _monitorCallback;
        public delegate void MonitorCallback();

        public VerificationManager()
        {
            _logList = new LogList();
            _logger = new SuppressionFilterLog(_logList);
            string[] fileNames = s_defaultVerifications;
            try
            {
                FindRoutines(fileNames);
            }
            catch
            {
                // ignore errors at creation.  This lets AccChecker start up and verification dll can be loaded manually.
            }
        }

        public VerificationManager(bool allDlls)
        {
            _logList = new LogList();
            _logger = new SuppressionFilterLog(_logList);
            string[] fileNames = s_defaultVerifications;
            try
            {
                if (allDlls)
                    AddAllDllsInCurrentDirectory();
                else
                    FindRoutines(fileNames);
            }
            catch
            {
                // ignore errors at creation.  This lets AccChecker start up and verification dll can be loaded manually.
            }
        }

        public VerificationManager(string[] filenames)
        {
            _logList = new LogList();
            _logger = new SuppressionFilterLog(_logList);
            FindRoutines(filenames);
        }

        /// <summary>Calls the methods that handle clearing the caches</summary>
        public void ClearCache()
        {
            foreach (ICache cache in _cacheMethodList)
            {
                cache.Clear();
            }

            foreach (SerializedTree serializedTree in _serializableTreeList)
            {            
                serializedTree.SerializedTreeXml = "";
            }
        }

        public List<SerializedTree> SerializableTreeList
        {
            get
            {
                return _serializableTreeList;
            }
        }
        
        //Adds all *.dlls from the current dir to 
        public void AddAllDllsInCurrentDirectory()
        {
            string[] dllFiles = this.FindAllDlls();

            foreach (string filename in dllFiles)
            {
                try
                {
                    this.AddVerificationDLL(filename);
                }
                catch (ArgumentException)
                {
                    //ignore error
                }
            }
        }

        //Returns all *.dlls from the current dir
        private string[] FindAllDlls()
        {
            String currentDirectory = System.IO.Directory.GetCurrentDirectory();
            return System.IO.Directory.GetFiles(currentDirectory, "*.dll", SearchOption.TopDirectoryOnly);
        }

        // Find all the routines in the List of DLL names 
        private void FindRoutines(string[] filenames)
        {
            foreach (string filename in filenames)
            {
                String currentDirectory = "";
                if (!Path.IsPathRooted(filename))
                {
                    currentDirectory = System.IO.Directory.GetCurrentDirectory() + "\\";
                }

                string verificationDll = Path.Combine(currentDirectory, filename);
                if (_verificationDlls.IndexOf(verificationDll) >= 0)
                {
                    // we have already added this dll
                    continue;
                }

                try
                {
                    Assembly assembly = Assembly.LoadFile(verificationDll);
                    List<VerificationRoutineWrapper> visualizers = new List<VerificationRoutineWrapper>();
                    foreach (Type type in assembly.GetTypes())
                    {
                        // Load any ICache in the list of caches the ones that are verification routines as well will be added below.
                        // This is currently used to clear the cache when needed.
                        if (type.GetInterface(typeof(ICache).ToString()) != null && type.GetInterface("IVerificationRoutine") == null)
                        {
                            Object cache = Activator.CreateInstance(type);
                            _cacheMethodList.Add((ICache)cache);
                        }

                        // Only load classes that have a VerificationAttribute and IVerificationRoutine
                        // associated with it.  Some classes implement the IVerificationRoutine, but are
                        // only the base class for the tests.  It's the derived class in this situation
                        // that we want to load.
                        if (Attribute.IsDefined(type, typeof(VerificationAttribute)))
                        {
                            if (type.GetInterface("IVerificationRoutine") != null)
                            {
                                object[] attributes = type.GetCustomAttributes(typeof(VerificationAttribute), false);
                                string id = string.Empty;
                                if (attributes.Length == 1)
                                {
                                    id = ((VerificationAttribute)attributes[0]).UniqueIdentifier;
                                }

                                object vrObject = Activator.CreateInstance(type);
                                VerificationRoutineWrapper rvr = new VerificationRoutineWrapper(vrObject, type, id);

                                if (rvr.IsIntrusive || Path.GetFileName(verificationDll) == "VerificationRoutines.dll")
                                {
                                    rvr.IsSlow = true;
                                }

                                if (rvr.CanVisualize)
                                {
                                    // collect the visualizers so we can put them at the end.
                                    visualizers.Add(rvr);
                                }
                                else
                                {
                                    _routines.Add(rvr);
                                }

                                ICache cache = vrObject as ICache;
                                if (cache != null)
                                {
                                    _cacheMethodList.Add(cache);
                                }
                                
                                ISerializeTree serializedTree = vrObject as ISerializeTree;
                                if (serializedTree != null)
                                {
                                    _serializableTreeList.Add(new SerializedTree(verificationDll, type.ToString(), serializedTree));
                                }
                            }
                        }
                    }

                    foreach (VerificationRoutineWrapper rvr in visualizers)
                    {
                        _routines.Add(rvr);
                    }
                }
                catch (ReflectionTypeLoadException error)
                {
                    throw error;
                }
                catch (Exception error)
                {
                    if (verificationDll.Contains("UIAVerifications"))
                    {
                        string reason = "This can happen if you are running 32 bit AccChecker on a 64 bit system.  Solution: Use UIAutomationClient.dll from a 64 bit AccChecker build";
                        throw new ArgumentException(string.Format("Could not load the verification library \"{0}\". Reason : {1}", filename, reason), error.InnerException);
                    }
                    else
                    {
                        throw new ArgumentException(string.Format("Could not load the verification library \"{0}\". Reason : {1}", filename, error.Message), error.InnerException);
                    }
                }

                _verificationDlls.Add(verificationDll);
            }
        }

        //Clears all routines and dlls from lists for a fresh start
        public void ClearAllRoutines()
        {
            _routines.Clear();
            _verificationDlls.Clear();
        }

        // Add a verification dll to AccChecker so that the routines in 
        // that dll can be executed.
        public void AddVerificationDLL(string filename)
        {
            string[] filenames = { filename };
            FindRoutines(filenames);
        }

        // Enables all the verifications routine in all the verification  
        // dlls that are currently known to AccCheck based on the 
        // filter criteria.
        public void EnableVerifications(VerificationFilter filter)
        {
            foreach (VerificationRoutineWrapper rvr in _routines)
            {
                switch (filter)
                {
                    case VerificationFilter.All:
                        {
                            rvr.Active = true;
                            break;
                        }
                    case VerificationFilter.WithoutUI:
                        {
                            rvr.Active = !rvr.NeedsUI;
                            break;
                        }
                    case VerificationFilter.NonIntrusive:
                        {
                            rvr.Active = !rvr.IsIntrusive;
                            break;
                        }
                    case VerificationFilter.NonIntrusive | VerificationFilter.WithoutUI:
                        {
                            rvr.Active = !rvr.IsIntrusive && !rvr.NeedsUI;
                            break;
                        }
                }
            }
        }

        // Disables all the verifications routine in all the verification 
        // dlls that are currently known to AccCheck based on 
        // the filter criteria.
        public void DisableVerifications(VerificationFilter filter)
        {
            foreach (VerificationRoutineWrapper rvr in _routines)
            {
                switch (filter)
                {
                    case VerificationFilter.All:
                        {
                            rvr.Active = false;
                            break;
                        }
                    case VerificationFilter.WithoutUI:
                        {
                            rvr.Active = rvr.NeedsUI;
                            break;
                        }
                    case VerificationFilter.NonIntrusive:
                        {
                            rvr.Active = rvr.IsIntrusive;
                            break;
                        }
                    case VerificationFilter.NonIntrusive | VerificationFilter.WithoutUI:
                        {
                            rvr.Active = rvr.IsIntrusive || rvr.NeedsUI;
                            break;
                        }
                }
            }
        }

        // Enables a specific verification routine by name.  This is the 
        // name that is listed in verifications tab under routines or from 
        // the command line using the -list option.
        public uint EnableRoutine(string verificationRoutine)
        {
            uint found = 0;
            foreach (VerificationRoutineWrapper rvr in _routines)
            {
                if (rvr.Title == verificationRoutine)
                {
                    
                    rvr.Active = true;
                    found++;
                }
            }
            return found;
        }

        // Disables a specific verification routine by name.  This is the 
        // name that is listed in verifications tab under routines or from 
        // the command line using the -list option.
        public uint DisableRoutine(string verificationRoutine)
        {
            uint found = 0;
            foreach (VerificationRoutineWrapper rvr in _routines)
            {
                if (rvr.Title == verificationRoutine)
                {
                    rvr.Active = false;
                    found++;
                }
            }

            return found;
        }

        // Execute the routines that have been enabled as a result 
        // of EnableVerifications, DisableVerifications, 
        // EnableRoutine and DisableRoutine.
        public void ExecuteEnabled(IntPtr hwnd)
        {
            if (!Win32API.IsWindow(hwnd))
            {
                throw new ArgumentException("Not a valid hwnd");
            }

            //Set the Global Hwnd value to be used later for SaveLogToFile
            _hwnd = hwnd; 
            _screenShot = GraphicsHelper.GrabScreenShot(_hwnd);

            UsageReporter.Instance.BeginWindowScan(this, hwnd);

            List<VerificationRoutineWrapper> intrusiveRoutines = new List<VerificationRoutineWrapper>();

            foreach (VerificationRoutineWrapper rvr in _routines)
            {
                if (rvr.Active && rvr.Priority <= _priority)
                {
                    if (rvr.IsIntrusive)
                    {
                        intrusiveRoutines.Add(rvr);
                    }
                    else
                    {
                        rvr.Execute(hwnd, _logger, _allowUI, null);
                    }
                }
            }

            // Run the intrusive ones last so they do not interfere with the rest.
            foreach (VerificationRoutineWrapper rvr in intrusiveRoutines)
            {
                rvr.Execute(hwnd, _logger, _allowUI, null);
            }
            
            UsageReporter.Instance.EndWindowScan(this, hwnd);
        }

        private void MonitorTimerCallback(object sender, EventArgs args)
        {
            if (!_processingMonitoredHwnd)
            {
                _processingMonitoredHwnd = true;

                try
                {
                    Win32API.GUITHREADINFO guiThreadInfo = new Win32API.GUITHREADINFO();
                    guiThreadInfo.cbSize = Marshal.SizeOf(guiThreadInfo.GetType());
                    Win32API.GetGUIThreadInfo(0, ref guiThreadInfo);

                    _monitoredHwnd = Win32API.GetAncestor(guiThreadInfo.hwndFocus, Win32API.GA_PARENT);
                    if (_monitoredHwnd == Win32API.GetDesktopWindow())
                    {
                        _monitoredHwnd = guiThreadInfo.hwndFocus;
                    }

                    if (_lastMonitoredHwndprocessed != _monitoredHwnd && _monitoredHwnd != IntPtr.Zero)
                    {
                        int processId;
                        Win32API.GetWindowThreadProcessId(guiThreadInfo.hwndFocus, out processId);
                        if (processId != Win32API.GetCurrentProcessId())
                        {
                            if (_monitorCallback != null)
                            {
                                _monitorCallback();
                                _lastMonitoredHwndprocessed = _monitoredHwnd;
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    // need to figure out what to do here
                }

                _processingMonitoredHwnd = false;
            }
        }

        public void MonitorEnabled(bool start, MonitorCallback monitorCallback)
        {
            if (start)
            {
                _monitoringTimer.Interval = 500;
                _monitorCallback = monitorCallback;
                _monitoringTimer.Tick += new EventHandler(MonitorTimerCallback); 
                _monitoringTimer.Start();
            }
            else
            {
                _monitoringTimer.Stop();
            }
        }

        //Take filename as input and dump the information in AccumulatingLogger in XML format
        //to the file. 
        //TODO: We need to add screenshot to _logFile.ScreenShot 
        public void SaveLogToFile(string filename)
        {
            AccumulatingLogger acclogger = new AccumulatingLogger();
            XMLSerializingLogger xsl = new XMLSerializingLogger();
            LogFile _logFile;
            
            if (!Win32API.IsWindow(_hwnd))
            {
                throw new ArgumentException("No valid hwnd has been set");
            }
            
            _logFile = new LogFile(_hwnd);

            acclogger = _logList.GetLogger(acclogger.GetType()) as AccumulatingLogger;

            if (acclogger == null)
            {
                throw new ArgumentException("No Accumulating Logger added");
            }

            if (acclogger.Events.Count > 0)
            {
                if (_logFile.StartTime == DateTime.MinValue)
                {
                    _logFile.StartTime = acclogger.Events.First().Timestamp;
                }
                if (_logFile.StopTime == DateTime.MinValue)
                {
                    _logFile.StopTime = acclogger.Events.Last().Timestamp;
                }
            }
            
            acclogger.DumpToLogger(xsl);
            _logFile.ScreenShot = _screenShot;
            _logFile.Trees = this.SerializableTreeList;
            xsl.SerializeToFile(filename, _logFile);
        }

        // Required for COM interface
        public void SetIncludePassResultsInLog(bool includePassResultsInLog)
        {
            VerificationManager._includePassResultsInLog = includePassResultsInLog;
        }

        static public bool IncludePassResultsInLog
        {
            get
            {
                return _includePassResultsInLog;
            }
            set
            {
                _includePassResultsInLog = value;
            }
        }

        static public List<VerificationRoutineWrapper> RoutinesRegistered
        {
            get
            {
                return _routines;
            }
        }

        // Returns an array of routines that are currently known
        // to AccCheck these are of type VerificationRoutineWrapper.
        // [VerificationRoutineWrapper is not COM Visible]
        [ComVisible(false)]
        public VerificationRoutineWrapper[] Routines
        {
            get
            {
                return _routines.ToArray();
            }
        }

        // Returns an array of routine data [exposed for COM]
        IVerificationRoutineData[] IVerificationManager.Routines
        {
            get
            {
                return _routines.ToArray();
            }
        }

        // This returns an array of groups. The group is a property of the
        // verification routine and is shown in the UI on the verifications tab.
        public String[] Groups
        {
            get
            {
                List<String> groups = new List<String>();
                foreach (VerificationRoutineWrapper rvr in _routines)
                {
                    if (!groups.Contains(rvr.Group))
                    {
                        groups.Add(rvr.Group);
                    }
                }
                return groups.ToArray();
            }
        }

        // This value is pased to each verification routine to indicate if it 
        // should show UI.  This is to let verification show UI during 
        // interactive use but not show it in a BVT environment.
        public bool AllowUI
        {
            get
            {
                return _allowUI;
            }
            set
            {
                _allowUI = value;
                UsageReporter.Instance.ShowingUI = value;
            }
        }

        public ILogger Logger
        {
            get
            {
                return _logger;
            }
        }

        static public bool IsQueryStringAddinEnabled
        {
            get
            {
                return _enableQueryStringAddin;
            }
            set
            {
                _enableQueryStringAddin = value;
            }
        }
        
        public ILoggerCallback LoggerCallback
        { 
            get 
            { 
                return _logList.LoggerCallback;
            }
            
            set 
            { 
                _logList.LoggerCallback = value;
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
                if (_priority < 1)
                {
                    _priority = 1;
                }
                else if (_priority > 3)
                {
                    _priority = 3;
                }               
                _logList.Priority = _priority;
            } 
        }

        public IntPtr MonitoredHwnd 
        { 
            get 
            {
                return _monitoredHwnd;
            }
        }
        
        // this is the logger that will be pass to each verification routine 
        // as it is executed.
        public void AddLogger(params ILogger[] logs)
        {
            foreach (ILogger log in logs)
            {
                _logList.AddLogger(log);
                BaseLogger baseLog = log as BaseLogger;
                if (baseLog != null)
                {
                    baseLog.Priority = _priority;
                }
            }
        }

        // These are the suppressions that will be used by each verification routine 
        public void AddSuppressionFiles(params string[] suppressions)
        {
            foreach (string suppression in suppressions)
            {
                _logger.AddSuppressionFile(suppression);
            }
        }

        // These are the supressions that will be used by each verification routine 
        public void ClearSuppressionFiles()
        {
            _logger.Clear();
        }

        // This tries to load a dll that might exist to provide a string that can 
        // uniquely identify an element.
        public static IQueryString GetQueryStringObject()
        {
            if (!_enableQueryStringAddin && !_includePassResultsInLog)
            {
                return null;
            }
            
            if (_queryString == null)
            {
                String currentDirectory = Directory.GetCurrentDirectory();
                string [] dlls = Directory.GetFiles(currentDirectory, "*.dll", SearchOption.TopDirectoryOnly);
                foreach (string filename in dlls)
                {
                    try
                    {
                        Assembly assembly = Assembly.LoadFile(filename);
                        if (assembly != null)
                        {
                            foreach (Type type in assembly.GetTypes())
                            {
                                if (type.GetInterface("IQueryString") != null)
                                {
                                    _queryString = (IQueryString)Activator.CreateInstance(type);

                                    // this call is here to make the assembly pull in all its dependencies so if there are missing
                                    // dlls this will fallback gracefully
                                    _queryString.Generate(Win32API.GetDesktopWindow());
                                    return _queryString;
                                }
                            }
                        }
                    }
                    catch
                    {
                        // its ok for any of this to fail this is an optional dll 
                    }
                }
            }

            return _queryString;
        }
    }
}
