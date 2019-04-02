using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Reflection;
using AccCheck.Verification;
using AccCheck.Logging;

namespace AccCheck.Verification
{
    /*
     * This is a wrapper for verification routines that are fetched
     * via reflection from a DLL at runtime. It can be treated as if
     * the verification routine was in your own project.
     */
    public class VerificationRoutineWrapper : IVerificationRoutine, IVerificationRoutineData
    {
        private object _routine;
        private Type _type;
        private bool _active = false;
        private VerificationAttribute _metadata;
        private string _id = string.Empty;
        private bool _isSlow = false;
        
        // Required for XML Persiting
        public VerificationRoutineWrapper()
        {
        }

        public string Id
        {
            get { return _id; }
            set { _id = value; }
        }

        public bool Active
        {
            get { return _active;  }
            set { _active = value; }
        }

        private void LogException(Exception e, ILogger logger)
        {
            LogException(EventLevel.Warning, e, logger);
        }

        private void LogException(EventLevel eventLevel, Exception e, ILogger logger)
        {
            if (logger != null)
            {
                LogEvent le = new LogEvent(
                    eventLevel,
                    e.Message,
                    e.Message,
                    e.ToString(),
                    System.Drawing.Rectangle.Empty,
                    typeof(VerificationRoutineWrapper));
                logger.Log(le);
            }
        }

        public VerificationRoutineWrapper(object vr, Type t, string id)
        {
            _routine = vr;
            _type = t;
            _id = id;
        }

        public override String ToString()
        {
            return Title;
        }

        #region IVerificationRoutine
        public void Execute(IntPtr hwnd, ILogger logger, bool AllowUI, GraphicsHelper graphics)
        {
            try
            {
                int initialErrorCount = 0;
                int InitialWarningCount = 0;
                ILoggerStatistics logStats = logger as ILoggerStatistics;
                if (logStats != null)
                {
                    initialErrorCount = logStats.ErrorCount;
                    InitialWarningCount = logStats.WarningCount;
                }

                MethodInfo mi = _type.GetMethod("Execute");
                mi.Invoke(_routine, new object[] { hwnd, logger, AllowUI, graphics });

                if (logStats != null)
                {
                    // if there were no errors or warning added during the execute than put out a successful message
                    // Don't bother with this for Visualizers because they always succeed and are just there to provide visual context for errors.
                    int changedErrorCount = logStats.ErrorCount - initialErrorCount;
                    int changedWarningCount = logStats.WarningCount - InitialWarningCount;
                    if (changedErrorCount == 0 && changedWarningCount == 0 && !CanVisualize)
                    {
                        LogEvent le = new LogEvent(EventLevel.Information, "VerificationSuccessful", 
                            "Verification for " + Title + " successful", null,
                            System.Drawing.Rectangle.Empty,
                            typeof(VerificationRoutineWrapper));

                        IQueryString queryString = VerificationManager.GetQueryStringObject();
                        if (queryString != null)
                        {
                            le.QueryString = queryString.Generate(hwnd);
                        }
                        
                        logger.Log(le);
                    }
                }
            }
            catch (TargetInvocationException targetException)
            {
                LogException(EventLevel.Error, new Exception("UNEXPECTED TEST ERROR", targetException.InnerException), logger);
            }
        }
        #endregion
        
        #region Metadata accessors
        private void GetMetadata()
        {
            if (_metadata == null)
            {
                _metadata = (VerificationAttribute)Attribute.GetCustomAttribute(
                     _type, typeof(VerificationAttribute));
            }
        }
        
        // This functions as the name of the verification and is the name
        // that show up in the verification tab of the UI.
        public String Title
        {
            get
            {
                GetMetadata();
                return _metadata.Title;
            }
        }

        // A description of the routine it currently does not show up in the UI
        public String Description
        {
            get
            {
                GetMetadata();
                return _metadata.Description;
            }
        }

        // This is an indication that this verification needs to show UI in
        // order to be useful. It would not make sense to run this in a 
        // bvt environment.
        public bool NeedsUI
        {
            get
            {
                GetMetadata();
                return _metadata.NeedsUI;
            }
        }

        // The group is a property of the verification routine
        // and is shown in the UI on the verifications tab.
        public String Group
        {
            get
            {
                GetMetadata();
                return _metadata.Group;
            }
        }
        
        // This verification routine does visualizations
        public bool CanVisualize
        {
            get
            {
                GetMetadata();
                return _metadata.CanVisualize;
            }
        }
        
        // This is the priority of the verification
        public int Priority
        {
            get
            {
                GetMetadata();
                return _metadata.Priority;
            }
        }

        // This is the priority of the verification
        public bool IsIntrusive 
        {
            get
            {
                GetMetadata();
                return _metadata.IsIntrusive;
            }
        }
        #endregion

        public Type Type
        {
            get
            {
                return _type;
            }
        }

        public bool IsSlow
        {
            get { return _isSlow;  }
            set { _isSlow = value; }
        }

    }
}
