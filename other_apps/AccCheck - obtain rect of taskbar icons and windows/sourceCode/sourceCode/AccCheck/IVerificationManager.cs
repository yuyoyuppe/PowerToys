using System;
using System.IO;
using System.Collections.Generic;
using System.Text;
using System.Reflection;
using AccCheck.Logging;
using AccCheck.Verification;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace AccCheck.Verification
{
    [ComVisible(true)]
    public enum VerificationFilter { All = 0x01, WithoutUI = 0x02, NonIntrusive = 0x04 };

    [ComVisible(true), GuidAttribute("137AD71F-4657-4362-B9E4-D6D734F1F530")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IVerificationManager
    {
        void AddVerificationDLL(string filename);
        void EnableVerifications(VerificationFilter filter);
        void DisableVerifications(VerificationFilter filter);
        uint EnableRoutine(string verificationRoutine);
        uint DisableRoutine(string verificationRoutine);
        void ExecuteEnabled(IntPtr hwnd);
        void AddLogger(params ILogger[] logs);
        void AddSuppressionFiles(params string[] suppressions);
        void ClearSuppressionFiles();
        void SaveLogToFile(string filename);
        void SetIncludePassResultsInLog(bool isLabRun);

        IVerificationRoutineData[] Routines { get; }
        String[] Groups { get; }
        bool AllowUI { get; set; }
        ILogger Logger { get; }
        ILoggerCallback LoggerCallback { get; set; }
    }
}
