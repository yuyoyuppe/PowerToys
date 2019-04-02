// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Diagnostics;
using System.Reflection;
using System.Threading;
using System.Globalization;
using System.Drawing;
using AccCheck;
using AccCheck.Logging;
using AccCheck.Verification;

namespace AccCheckConsole
{
    static class Program
    {
        private static IntPtr s_hwnd;
        private static bool s_quiet;
        private static EventLevel s_loglevel = EventLevel.Information;
        private static List<String> s_enabledRoutines = new List<String>();
        private static List<String> s_disabledRoutines = new List<String>();
        private static bool s_list = false;
        private static bool s_showHelp = false;
        private static VerificationManager s_vm;
        // stores a list of filenames specified on the command line
        private static List<string> s_routineLibs = new List<string>();
        private static List<String> s_xmlLoggerFileNames = new List<String>();
        private static XMLSerializingLogger s_xmlLogger;
        static BaseLogger s_consoleLogger = null;
        static FailureCode s_failure = FailureCode.NoErrorsNoWarnings;
        static bool s_helpShown = false;

        /// <summary>Return error codes for application</summary>
        enum FailureCode
        {
            NoErrorsNoWarnings = 0,
            ShowUsage,
            ErrorsButNoWarnings,
            ErrorsAndWarnings,
            NoErrorsButWarnings,
            CommandLineArgument,
        }

        static void ParseArguments(string[] args)
        {
            s_xmlLogger = new XMLSerializingLogger();
            s_vm.AddLogger(s_xmlLogger);

            int i = 0;
            try
            {
                while (i < args.Length)
                {
                    switch (args[i].ToLower())
                    {
                        case "-break":
                            {
                                Debug.Assert(false);
                                break;
                            }
                        case "-?":
                        case "-help":
                            {
                                s_showHelp = true;
                                break;
                            }
                        case "-list":
                            {
                                s_list = true;
                                break;
                            }
                        case "-log":
                            if (i + 1 < args.Length)
                            {
                                // look at the next argument (is it what we expected?)
                                switch (args[i + 1].ToLower())
                                {
                                    case "trace":
                                        s_loglevel = EventLevel.Trace;
                                        break;
                                    case "info":
                                    case "information":
                                        s_loglevel = EventLevel.Information;
                                        break;
                                    case "warn":
                                    case "warning":
                                        s_loglevel = EventLevel.Warning;
                                        break;
                                    case "err":
                                    case "error":
                                        s_loglevel = EventLevel.Error;
                                        break;
                                    default:
                                        throw new ArgumentException("-log loglevel not valid. Should be either 'info', 'warn' or 'err'.");
                                }
                                i++;
                            }
                            else
                            {
                                throw new ArgumentException("-log not followed by log level.");
                            }
                            break;
                        case "-logfile":
                            if (i + 1 < args.Length)
                            {
                                String filename = args[i + 1];

                                if (Path.GetExtension(filename).ToLower() == ".xml")
                                {
                                    // since the XML logger needs to have it's SerializeToFile() method called
                                    // to output data, we need to store it away, so we can call that method
                                    // when we are done
                                    s_xmlLoggerFileNames.Add(Path.GetFullPath(filename));
                                }
                                else
                                {
                                    // text file loggers output their data right away, so we only have to
                                    // add these to the loglist
                                    s_vm.AddLogger(new TextFileLogger(Path.GetFullPath(filename)));
                                }
                                i++;
                            }
                            else
                            {
                                throw new ArgumentException("-logfile needs a filename.");
                            }
                            break;
                        case "-process":
                            if (i + 1 < args.Length)
                            {
                                string exeName = Path.GetFileNameWithoutExtension(args[i + 1]); 

                                Process[] rgProcesses = Process.GetProcessesByName(exeName);
                                if (rgProcesses.Length == 0)
                                {
                                    throw new ArgumentException(string.Format("No such process to attach to named \"{0}\".  Make sure the application is started before you run AccCheckConsole.exe", exeName));
                                }
                                else if (rgProcesses.Length > 1)
                                {
                                    throw new ArgumentException(string.Format("Multiple processes called \"{0}\".  Please close all but one of them.", exeName));
                                }
                                else
                                {
                                    s_hwnd = rgProcesses[0].MainWindowHandle;
                                    i++;
                                }
                            }
                            else
                            {
                                throw new ArgumentException("Command line argument \"-process\" needs a process name.");
                            }
                            break;
                        case "-window":
                            if (i + 1 < args.Length)
                            {
                                s_hwnd = Win32API.FindWindowEx(IntPtr.Zero, IntPtr.Zero, null, args[i + 1]);
                                bool isValid = Win32API.IsWindow(s_hwnd);
                                if (!isValid)
                                {
                                    throw new ArgumentException(string.Format("There is no window with the caption \"{0}\" that was specified by the command line argument \"-window\"", args[i + 1]));
                                    // don't have to set hwnd = IntPtr.Zero, since the exception causes it not to be used
                                }
                                i++;
                            }
                            else
                            {
                                throw new ArgumentException("Command line argument \"-window\" needs a window title.");
                            }
                            break;
                        case "-hwnd":
                            if (i + 1 < args.Length)
                            {
                                int h;
                                // remove "0x" from string, since int.TryParse doesn't allow it
                                bool res;
                                if (args[i + 1].StartsWith("0x"))
                                {
                                    res = int.TryParse(args[i + 1].Replace("0x", ""), NumberStyles.AllowHexSpecifier, CultureInfo.CurrentCulture.NumberFormat, out h);
                                }
                                else
                                {
                                    res = int.TryParse(args[i + 1], out h);
                                }

                                if (res)
                                {
                                    s_hwnd = new IntPtr(h);
                                    // validate hwnd
                                    bool isValid = Win32API.IsWindow(s_hwnd);
                                    if (!isValid)
                                    {
                                        throw new ArgumentException(string.Format("Command line argument \"-hwnd {0}\" did not point to an actual window", args[i + 1]));
                                    }
                                    // don't have to set hwnd = IntPtr.Zero, since the exception causes it not to be used
                                    i++;
                                }
                                else
                                {
                                    throw new ArgumentException(string.Format("Command line argument \"-hwnd {0}\" was not a decimal or hexadecimal (0x00..).", args[i + 1]));
                                }
                            }
                            else
                            {
                                throw new ArgumentException("Command line argument \"-hwnd\" needs an hwnd.");
                            }
                            break;
                        case "-enable":
                            if (i + 1 < args.Length)
                            {
                                s_enabledRoutines.Add(args[i + 1]);
                                i++;
                            }
                            break;
                        case "-disable":
                            if (i + 1 < args.Length)
                            {
                                s_disabledRoutines.Add(args[i + 1]);
                                i++;
                            }
                            break;
                        case "-suppress":
                            if (i + 1 < args.Length)
                            {
                                String xmlFile = args[i + 1];
                                xmlFile = Path.GetFullPath(xmlFile);
                                if (File.Exists(xmlFile))
                                {
                                    s_vm.AddSuppressionFiles(xmlFile);
                                }
                                i++;
                            }
                            else
                            {
                                throw new ArgumentException("-suppress needs an XML file.");
                            }
                            break;
                        case "-quiet":
                            s_quiet = true;
                            break;
                        default:
                            // interpret as filename
                            if (File.Exists(args[i]))
                            {
                                s_routineLibs.Add(args[i]);
                            }
                            else
                            {
                                // or it might just be an error we didn't catch
                                throw new ArgumentException("Unknown parameter '" + args[i] + "'! It's not a library.");
                            }
                            break;
                    }
                    i++;
                }

                LogToConsole(EventLevel.Information, "Command line argument", string.Join(" ", args));
            }
            catch (ArgumentException exception)
            {
                s_failure = FailureCode.CommandLineArgument;
                throw exception;
            }

        }

        static int Main(string[] args)
        {

            try
            {
                if (args.Length == 0)
                {
                    LogHelpText(EventLevel.Information);
                }
                else
                {
                    // Handle loading the libraries to add to the verification mananager

                    s_vm = new VerificationManager();
                    
                    ParseArguments(args);

                    //If there was a dll specified, clear current routines
                    if(s_routineLibs.Count > 0)
                        s_vm.ClearAllRoutines();

                    //Adds specified dlls
                    foreach (string filename in s_routineLibs)
                        s_vm.AddVerificationDLL(filename);

                    // Display the list of all routines in the loaded libraries
                    if (s_list)
                    {
                        LogListOfRoutines(s_vm, EventLevel.Information);
                    }

                    // Display the help contents.  If there is no hwnd, then display as an error
                    if ((s_showHelp || s_hwnd.Equals(IntPtr.Zero)) && !s_list)
                    {
                        LogHelpText(s_hwnd == IntPtr.Zero ? EventLevel.Error : EventLevel.Information);
                    }

                    if (s_hwnd != IntPtr.Zero)
                    {
                        s_vm.Logger.LogLevel = s_loglevel;

                        if (s_enabledRoutines.Count > 0)
                        {
                            foreach (string routine in s_enabledRoutines)
                            {
                                // If the routine is not found, then print list and exit out.
                                if (s_vm.EnableRoutine(routine) != 1)
                                {
                                    throw new ArgumentException(string.Format("Invalid routine \"{0}\", run AccCheckerConsole with command line \"-list\" to display a list of valid routines", routine));
                                }
                            }
                        }
                        else
                        {
                            s_vm.EnableVerifications(VerificationFilter.WithoutUI);
                            
                            if (s_disabledRoutines.Count > 0)
                            {
                                foreach (string routine in s_disabledRoutines)
                                {
                                    // If the routine is not found, then print list and exit out.
                                    if (s_vm.DisableRoutine(routine) != 1)
                                    {
                                        throw new ArgumentException(string.Format("Invalid routine \"{0}\", run AccCheckerConsole with command line \"-list\" to display a list of valid routines", routine));
                                    }
                                }
                            }
                        }

                        // If you want to produce a log file for uploading to database, uncomment the following line.
                        //s_vm.SetLabRun(true);

                        // Execute the tests
                        s_vm.ExecuteEnabled(s_hwnd);

                        // output XML serialized data
                        foreach (String filename in s_xmlLoggerFileNames)
                        {
                            LogFile logFile = new LogFile();
                            logFile.ScreenShot = GraphicsHelper.GrabScreenShot(s_hwnd);
                            s_xmlLogger.SerializeToFile(filename, logFile);
                        }
                    }
                }
            }
            catch (Exception exception)
            {
                LogToConsole(EventLevel.Error, "ERROR", exception.Message);
            }

            LogResults();

            return (int)GetResultsErrorCode();
        }

        /// <summary>Displays the results of the run to the console logger</summary>
        private static void LogResults()
        {
            if (!s_helpShown && !s_list)
            {
                s_vm.Logger.LogLevel = EventLevel.Information;

                StringBuilder sb = new StringBuilder();

                const int WIDTH = 40;
                sb.Append("\n".PadLeft(WIDTH - "Text: ".Length, '-'));
                sb.Append(string.Format("\tError count         : {0}\n", s_consoleLogger.ErrorCount));
                sb.Append(string.Format("\tWarning count       : {0}\n", s_consoleLogger.WarningCount));
                sb.Append(string.Format("\tInformational count : {0}\n", s_consoleLogger.InformationalCount));
                sb.Append(string.Format("\tTrace count         : {0}\n", s_consoleLogger.TraceCount));
                sb.Append("\t".PadRight(WIDTH, '-'));

                LogToConsole(EventLevel.Information, "TEST RESULTS", sb.ToString());
            }
        }

        /// <summary>Logs the information to the console logger if -quiet is not specific on command line</summary>
        private static void LogToConsole(EventLevel eventLevel, string title, string message)
        {

            // May have dropped out of the parsing command and never started up the logger
            if (s_consoleLogger == null)
            {
                if (s_quiet)
                {
                    // Just keep all the log entries in the logger
                    s_consoleLogger = new AccumulatingLogger();
                }
                else
                {
                    s_consoleLogger = new ConsoleLogger();
                }
                s_vm.AddLogger(s_consoleLogger);
            }

            LogEvent le = new LogEvent(eventLevel, title, message, "");
            s_consoleLogger.Log(le);
        }

        /// <summary>Determine the error code to return based on the run</summary>
        static FailureCode GetResultsErrorCode()
        {
            if (s_failure == FailureCode.NoErrorsNoWarnings)
            {
                if (s_consoleLogger.ErrorCount == 0)
                {
                    if (s_consoleLogger.WarningCount == 0)
                    {
                        s_failure = FailureCode.NoErrorsNoWarnings;
                    }
                    else
                    {
                        s_failure = FailureCode.NoErrorsButWarnings;
                    }
                }
                else if (s_consoleLogger.ErrorCount > 0)
                {
                    if (s_consoleLogger.WarningCount == 0)
                    {
                        s_failure = FailureCode.ErrorsButNoWarnings;
                    }
                    else
                    {
                        s_failure = FailureCode.ErrorsAndWarnings;
                    }
                }
            }
            return s_failure;
        }

        /// <summary>Log the help text to logger</summary>
        static void LogHelpText(EventLevel eventLevel)
        {
            StringBuilder sm = new StringBuilder();

            sm.Append("The syntax of this command is:\n");
            sm.Append("\n\tAccCheckConsole [options] (-hwnd <hwnd> | -process <name>) [<dlls>]\n");
            sm.Append("\n\tOptions:\n");
            sm.Append("\t\t-hwnd <hwnd>          Validates the given hwnd. Can be hex or dec.\n");
            sm.Append("\t\t-window <title>       Validates the window with the title given.\n");
            sm.Append("\t\t-process <name>       Validates the main window of the process with that name.\n");
            sm.Append("\t\t-list                 Lists all the verification routines available.\n");
            sm.Append("\t\t-enable <name>        Runs the given routine. Can be specified more than once\n");
            sm.Append("\t\t-disable <name>       Runs all but the given routine.  Can be specified more than once\n");
            sm.Append("\t\t-log (info|warn|err)  The lowest event rating that will be logged.\n");
            sm.Append("\t\t-logfile <file>       Outputs the log to file. Can be used multiple times.\n");
            sm.Append("\t\t-suppress <file>      Uses the XML file <file> to suppress errors.\n");
            sm.Append("\t\t-quiet                No logging to stdout.\n");
            sm.Append("\t\t-help                 Quick Help.\n");
            sm.Append("\n\tError codes returned from AccCheckConsole when using \"echo %errorlevel%\"\n");
            sm.Append("\t\t0 - No errors and no warnings.\n");
            sm.Append("\t\t1 - Usages statement was requested.\n");
            sm.Append("\t\t2 - Errors and no warnings.\n");
            sm.Append("\t\t3 - Errors and warnings.\n");
            sm.Append("\t\t4 - No errors but Warnings.\n");
            sm.Append("\t\t5 - Invalid command line.\n");
            sm.Append("\n");
            sm.Append("Examples:\n");
            sm.Append("\n");
            sm.Append("1) Run all verifications on a window with a specified name.\n");
            sm.Append("\tAccCheckConsole -window \"Untitled - Notepad\"\n");
            sm.Append("\n");
            sm.Append("2) Run a subset of the verifications against an HWND, specifying a suppression file.\n");
            sm.Append("\tAccCheckConsole -hwnd 0x00382f00 -enable CheckTabbing -enable CheckName -suppress suppress.xml\n");
            sm.Append("\n");
            sm.Append("3) Run all verifications from a new verification DLL.\n");
            sm.Append("\tAccCheckConsole -window \"Untitled - Notepad\" VerificationRoutine1.dll\n");
            
            Console.WriteLine(sm.ToString());
            
            s_helpShown = true;
            s_failure = FailureCode.ShowUsage;

        }

        /// <summary>Print the list of routines to the logger</summary>
        static void LogListOfRoutines(VerificationManager vm, EventLevel eventLevel)
        {
            StringBuilder sb = new StringBuilder("Available verification routines\n");
            foreach (VerificationRoutineWrapper rvr in vm.Routines)
            {
                sb.Append(string.Format("\t\t\"{0}\"\n", rvr.Title));
            }
            LogToConsole(eventLevel, "Verification Routines", sb.ToString());
        }

    }
}
