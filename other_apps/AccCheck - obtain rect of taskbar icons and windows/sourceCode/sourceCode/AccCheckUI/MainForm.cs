using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Reflection;
using System.Diagnostics;
using System.Collections;
using System.Threading;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Runtime.Serialization;
using AccCheck.Verification;
using AccCheck.Logging;
using AccCheck;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace AccCheckUI
{
    public partial class MainForm : Form
    {
        private const int _hotkeySelectId = 0xACC1;
        private const int _hotkeyRunId = 0xACC2;
        private const int _hotkeyParentId = 0xACC3;

        private Hashtable _routines = new Hashtable();
        private AccumulatingLogger _acclogger;
        private VerificationManager _vm;
        private IntPtr _capturedHwnd;
        private IntPtr _typedHwnd;
        private GraphicsHelper _graphics;
        private String _suppressionFile;
        private HighlightRectangle _highlightRectangle = new HighlightRectangle();
        private System.Windows.Forms.Timer _HighlightTimer;
        private int _AccCheckerProcessId = Process.GetCurrentProcess().Id;
        private Win32API.POINT _lastCursorPosition;
        private List<IntPtr> _topLevelHwnds = new List<IntPtr>();
        private static IntPtr _hwndCaptureSearch = IntPtr.Zero;
        private static string _captionCaptureSearch = string.Empty;
        private static string HELPFILE_NAME = "AccCheckerHelp.htm";
        private static object _internetExplorer = null;
        private static Type _internetExplorerType;
        private static bool _runningTests = false;
        private static bool _runningTestsUsingHotKey = false;
        private int _priority = 3;
        private DateTime _startTime; 
        private bool _monitormodeOn = false;

        public MainForm()
        {
            InitializeComponent();

            // Create Verification manager
            _vm = new VerificationManager(Properties.Settings.Default.AutoLoadSelected);
            _vm.AllowUI = true;
        }

        // catches global hotkey keypresses
        // FXCop indicated if a virtual method has a LinkDemand, so should any override of it.
        [System.Security.Permissions.SecurityPermission(System.Security.Permissions.SecurityAction.LinkDemand, Flags = System.Security.Permissions.SecurityPermissionFlag.UnmanagedCode)]
        protected override void WndProc(ref System.Windows.Forms.Message message)
        {
            // let the base class process the message
            base.WndProc(ref message);

            // if this is a WM_HOTKEY message, perform the action
            const int WM_HOTKEY = 0x312;
            if (message.Msg == WM_HOTKEY)
            {
                if ((int)message.WParam == _hotkeyParentId)
                {
                    // click the capture parent button
                    btnCaptureParent_Click(null, null);
                }
                else if ((int)message.WParam == _hotkeyRunId)
                {
                    if (rbMonitoringMode.Checked)
                    {
                        ChangeMonitoringMode(!_monitormodeOn);
                    }
                    else
                    {
                        _runningTestsUsingHotKey = true;
                        RunVerifications();
                        _runningTestsUsingHotKey = false;
                    }
                }
                else if ((int)message.WParam == _hotkeySelectId)
                {
                    Win32API.POINT pt = new Win32API.POINT(MousePosition.X, MousePosition.Y);
                    IntPtr capturedHwnd = Win32API.WindowFromPoint(pt);

                    uint processId;
                    Win32API.GetWindowThreadProcessId(capturedHwnd, out processId);
                    if (_AccCheckerProcessId != processId)
                    {
                        _capturedHwnd = capturedHwnd;
                        UpdateCapturedHwndText();
                        rbCapturedWindow.Checked = true;
                        tabControl.SelectedTab = tabPageVerifications;
                        SetEnableRunVerifications(true);
                    }
                }

            }
        }

        // register global hotkey for window capture
        private void RegisterGlobalHotkey(Keys hotkey, int modifiers, int hotkeyId)
        {
            try
            {
                // register the hotkey, throw if any error
                if (Win32API.RegisterHotKey(this.Handle, hotkeyId, modifiers, (int)hotkey) == 0)
                {
                    throw new Exception("Unable to register hotkey. Error code: " + Marshal.GetLastWin32Error()
                       .ToString());
                }
            }
            catch
            {
                // clean up if hotkey registration failed
                UnregisterGlobalHotKey(hotkeyId);
            }
        }

        // unregister the global hotkey
        private void UnregisterGlobalHotKey(int hotkeyId)
        {
            if (hotkeyId != 0)
            {
                Win32API.UnregisterHotKey(this.Handle, hotkeyId);
            }
        }

        // This method populates our listview of verification routines
        // from the VerificationManager.
        private void PopulateFromVM()
        {
            lvRoutines.Items.Clear();
            lvRoutines.Groups.Clear();
            // find all the groups they belong to
            String[] groups = _vm.Groups;
            Hashtable groupsHash = new Hashtable();
            foreach (String group in groups)
            {
                ListViewGroup lvg = new ListViewGroup(group);
                lvRoutines.Groups.Add(lvg);
                groupsHash.Add(group, lvg);
            }

            foreach (VerificationRoutineWrapper rvr in _vm.Routines)
            {
                // Visualizers are always run so there is no need to put them in the list
                if (rvr.CanVisualize)
                {
                    continue;
                }

                ListViewItem lvi = new ListViewItem(new string [] {rvr.ToString(), rvr.Priority == 0 ? " " : rvr.Priority.ToString()});
                lvi.Tag = rvr;
                ((ListViewGroup)groupsHash[rvr.Group]).Items.Add(lvi);
                lvi.Group = (ListViewGroup)groupsHash[rvr.Group];
                lvi.Checked = rvr.Active;
                lvRoutines.Items.Add(lvi);
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            // create loggers
            _acclogger = new AccumulatingLogger();
            _acclogger.LogLevel = EventLevel.Information;

            _vm.AddLogger(_acclogger);

            // find routines, and add them to the UI
            PopulateFromVM();

            // register hotkey for capturing a window
            RegisterGlobalHotkey(Keys.OemOpenBrackets, Win32API.MOD_CONTROL | Win32API.MOD_SHIFT, _hotkeySelectId);

            // register hotkey for running selected verifications
            RegisterGlobalHotkey(Keys.Oemplus, Win32API.MOD_CONTROL | Win32API.MOD_SHIFT, _hotkeyRunId);

            // register hotkey for navigating to the parent of the current window
            RegisterGlobalHotkey(Keys.OemCloseBrackets, Win32API.MOD_CONTROL | Win32API.MOD_SHIFT, _hotkeyParentId);

            // tell GraphicsHelper how to activate a tab
            GraphicsHelper.ActivateTabDelegate = new ActivateTabPaneCallback(ActivateTab);
            GraphicsHelper.AddTabDelegate = new AddTabPaneCallback(AddTab);

            SetMouseHoverHighlight(rbCapturedWindow.Checked);

            SetEnableRadioButtonsControls();
            cbPriority.SelectedIndex = 0;

            FillTopLevelWindowList();
            LoadSettings();
        }

        /// -------------------------------------------------------------------
        /// <summary>Compares each child window to see if 
        /// _hwndCaptureSearch and text are correct.</summary>
        /// <returns>true - continue enum, false enum should be stopped</returns>
        /// -------------------------------------------------------------------
        private static bool FindWindow(IntPtr hwnd, ref IntPtr lParam)
        {
            if (hwnd == _hwndCaptureSearch)
            {
                lParam = _hwndCaptureSearch;
            }

            // Continue until the two are the same
            return (!(hwnd == lParam));
        }

        private void LoadSettings()
        {
            // Choose window from list combo
            string item = topLevelWindows.Tag.ToString();
            int iLoc = topLevelWindows.Items.IndexOf(item);
            if (iLoc > -1)
            {
                topLevelWindows.SelectedIndex = iLoc;
            }

            // SuppressionFile text box
            if (!File.Exists(tbSuppresionFile.Text))
            {
                tbSuppresionFile.Text = string.Empty;
            }
            else
            {
                SetSuppressionInformation(tbSuppresionFile.Text);
            }

            // Selection radio buttons
            switch (Properties.Settings.Default.SelectionMethod)
            {
                case 0:
                    {
                        rbMainWindow.Checked = true;
                        topLevelWindows_Enter(null, null);
                        break;
                    }
                case 1:
                    {
                        rbCapturedWindow.Checked = true;
                        string searchHwnd = Properties.Settings.Default.CapturedWindow;
                        if (searchHwnd != string.Empty)
                        {
                            try
                            {
                                _hwndCaptureSearch = new IntPtr(Convert.ToInt64(searchHwnd));

                                IntPtr lParam = IntPtr.Zero;
                                Win32API.EnumChildWindows(Win32API.GetDesktopWindow(), new Win32API.WinEnumProc(FindWindow), ref lParam);
                                if (lParam != IntPtr.Zero)
                                {
                                    _capturedHwnd = _hwndCaptureSearch;
                                    SetEnableRunVerifications(true);
                                }
                                else
                                {
                                    SetEnableRunVerifications(false);
                                }
                            }
                            catch (FormatException)
                            {
                                // It's not a number...
                            }
                        }
                        UpdateCapturedHwndText();
                        break;
                    }
                case 2:
                    {
                        rbUseTypedHwnd.Checked = true;
                        tbTypedHwnd_TextChanged(null, null);
                        break;
                    }
                case 3:
                    {
                        rbMonitoringMode.Checked = true;
                        this.gbMonitoringOptions.Enabled = true;
                        if (Properties.Settings.Default.MonitoringModeOptimized)
                        {
                            rbOptimized.Checked = true;
                        }
                        else
                        {
                            rbSelected.Checked = true;
                        }
                        break;
                    }
            }

            cbPriority.SelectedIndex = Properties.Settings.Default.Priority;

            // get the state of the selection of verifications
            StringBuilder sb = new StringBuilder();
            string selectionKey = Properties.Settings.Default.VerificationsSelected as string;
            if (!string.IsNullOrEmpty(selectionKey))
            {
                foreach (ListViewItem lvi in lvRoutines.Items)
                {
                    VerificationRoutineWrapper rvr = (VerificationRoutineWrapper)lvi.Tag;
                    lvi.Checked = (-1 < selectionKey.IndexOf("[" + rvr.Id.ToString() + "]"));
                }
            }

            isIncludePassResultsInLogToolStripMenuItem.Checked = Properties.Settings.Default.PassResults;
            VerificationManager.IncludePassResultsInLog = Properties.Settings.Default.PassResults;
            enableQueryStringAddinStripMenuItem.Checked = Properties.Settings.Default.EnableSuppressions;
            VerificationManager.IsQueryStringAddinEnabled = Properties.Settings.Default.EnableSuppressions;


            //get selection of auto load
            autoLoadToolStripMenuItem.Checked = Properties.Settings.Default.AutoLoadSelected;
        }

        private void SaveSettings()
        {
            // Persist Selection methods
            if (rbMainWindow.Checked)
            {
                Properties.Settings.Default.SelectionMethod = 0;
                if (null != topLevelWindows.SelectedItem)
                {
                    Properties.Settings.Default.SelectedFromList = topLevelWindows.SelectedItem.ToString();
                }
                Properties.Settings.Default.SelectedByHwnd = string.Empty;
            }
            else if (rbCapturedWindow.Checked)
            {
                Properties.Settings.Default.SelectionMethod = 1;
                Properties.Settings.Default.SelectedFromList = string.Empty;
                Properties.Settings.Default.SelectedByHwnd = string.Empty;
                Properties.Settings.Default.CapturedWindow = HwndToVerify().ToString();
            }
            else if (rbUseTypedHwnd.Checked)
            {
                Properties.Settings.Default.SelectionMethod = 2;
                Properties.Settings.Default.SelectedFromList = string.Empty;
            }
            else if (rbMonitoringMode.Checked)
            {
                Properties.Settings.Default.SelectionMethod = 3;
                Properties.Settings.Default.MonitoringModeOptimized = rbOptimized.Checked;
                Properties.Settings.Default.SelectedFromList = string.Empty;
            }

            Properties.Settings.Default.Priority = cbPriority.SelectedIndex;

            // Persist the selection state of the verifications
            StringBuilder sb = new StringBuilder();
            foreach (ListViewItem lvi in lvRoutines.Items)
            {
                if (lvi.Checked)
                {
                    VerificationRoutineWrapper rvr = (VerificationRoutineWrapper)lvi.Tag;
                    sb.Append(string.Format("[{0}]", rvr.Id.ToString()));
                }
            }
            Properties.Settings.Default.VerificationsSelected = sb.ToString();

            Properties.Settings.Default.PassResults = isIncludePassResultsInLogToolStripMenuItem.Checked;
            Properties.Settings.Default.EnableSuppressions = enableQueryStringAddinStripMenuItem.Checked;

            //Persist automatically load verification dlls option
            Properties.Settings.Default.AutoLoadSelected = autoLoadToolStripMenuItem.Checked;

            // Save everything to the config file
            Properties.Settings.Default.Save();
        }
        
        private void enableQueryStringAddinStripMenuItem_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem enableQueryStringAddin = (ToolStripMenuItem)sender;
            enableQueryStringAddin.Checked = !enableQueryStringAddin.Checked;
            VerificationManager.IsQueryStringAddinEnabled = enableQueryStringAddin.Checked;
        }


        private void btnIncludePassResultsInLogRun_Click(object sender, EventArgs e)
        {
            ToolStripMenuItem includePassResultsMenuItem = (ToolStripMenuItem)sender;
            includePassResultsMenuItem.Checked = !includePassResultsMenuItem.Checked;

            VerificationManager.IncludePassResultsInLog = includePassResultsMenuItem.Checked;
            enableQueryStringAddinStripMenuItem.Enabled = !includePassResultsMenuItem.Checked;
        }

        private void btnSelectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lvRoutines.Items.Count; i++)
            {
                VerificationRoutineWrapper rvr = (VerificationRoutineWrapper)lvRoutines.Items[i].Tag;
                lvRoutines.Items[i].Checked = true;
            }

            _vm.EnableVerifications(VerificationFilter.All);
            SetEnableRunVerifications(true);
        }

        private void btnDeselectAll_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < lvRoutines.Items.Count; i++)
            {
                lvRoutines.Items[i].Checked = false;
            }

            _vm.DisableVerifications(VerificationFilter.All);
            SetEnableRunVerifications(false);
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            UnregisterGlobalHotKey(_hotkeyParentId);
            UnregisterGlobalHotKey(_hotkeyRunId);
            UnregisterGlobalHotKey(_hotkeySelectId);
            SetMouseHoverHighlight(false);

            SaveSettings();
        }

        private void UpdateCapturedHwndText()
        {
            String s = Win32API.GetWindowText(_capturedHwnd);
            textBoxCapturedWindow.Text = String.IsNullOrEmpty(s) ? "(no title)" : s;
        }

        private void btnCaptureParent_Click(object sender, EventArgs e)
        {
            IntPtr hwnd = Win32API.GetAncestor(_capturedHwnd, Win32API.GA_PARENT);
            if (!hwnd.Equals(IntPtr.Zero) && !hwnd.Equals(Win32API.GetDesktopWindow()))
            {
                _capturedHwnd = hwnd;
                UpdateCapturedHwndText();
                Win32API.RECT win32Rect = Win32API.RECT.Empty;

                uint processIdParent;
                Win32API.GetWindowThreadProcessId(_capturedHwnd, out processIdParent);
                if (_AccCheckerProcessId != processIdParent && Win32API.GetWindowRect(hwnd, ref win32Rect))
                {
                    Rectangle highlightRect = new Rectangle(win32Rect.left, win32Rect.top, win32Rect.right - win32Rect.left, win32Rect.bottom - win32Rect.top);
                    _highlightRectangle.Location = highlightRect;
                    _highlightRectangle.Visible = true;
                }
            }
        }

        private void tbTypedHwnd_TextChanged(object sender, EventArgs e)
        {
            if (!rbUseTypedHwnd.Checked)
            {
                return;
            }
            
            // try to check if the given handle is valid
            String s = tbTypedHwnd.Text;
            bool res;
            int h;
            if (s.StartsWith("0x"))
            {
                res = int.TryParse(s.Replace("0x", ""), NumberStyles.AllowHexSpecifier, CultureInfo.CurrentCulture.NumberFormat, out h);
            }
            else
            {
                res = int.TryParse(s, NumberStyles.Integer, CultureInfo.CurrentCulture.NumberFormat, out h);
            }

            if (res)
            {
                IntPtr hwnd = new IntPtr(h);
                // validate hwnd
                bool isValid = Win32API.IsWindow(hwnd);

                if (isValid)
                {
                    uint processId;
                    Win32API.GetWindowThreadProcessId(hwnd, out processId);
                    if (processId == Win32API.GetCurrentProcessId())
                    {
                        isValid = false;
                    }
                    else
                    {
                        _typedHwnd = hwnd;
                    }
                }
                tbTypedHwnd.BackColor = isValid ? Color.LightGreen : Color.IndianRed;
                SetEnableRunVerifications(isValid);
            }
            else
            {
                tbTypedHwnd.BackColor = Color.IndianRed;
                SetEnableRunVerifications(false);
            }
        }

        private void openToolStripButton_Click(object sender, EventArgs e)
        {
            // Path can be reset, open the directory where the binary resides when
            // loading verifications.
            openDLLDialog.InitialDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().CodeBase);

            if (openDLLDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    _vm.AddVerificationDLL(openDLLDialog.FileName);
                }
                catch (ArgumentException exception)
                {
                    MessageBox.Show(exception.Message);
                }

                PopulateFromVM();
            }
        }

        private void openLogFileStripButton_Click(object sender, EventArgs e)
        {
            if (openSuppresionFileDialog.ShowDialog() == DialogResult.OK)
            {
                _acclogger.Clear();
                _acclogger.LogLevel = EventLevel.Information;
                
                // clear the tab pages from the previous run. 
                ClearVerificationAddedTabPages();
                

                foreach (string fileName in openSuppresionFileDialog.FileNames)
                {
                    try
                    {
                        using (FileStream fs = new FileStream(fileName, FileMode.Open))
                        {

                            XmlSerializer xml = new XmlSerializer(typeof(LogFile));

                            LogFile logFile = (LogFile)xml.Deserialize(fs);
                            if (!logFile.IsVersionOk())
                            {
                                MessageBox.Show("Suppression file format is too old");
                                return;
                            }

                            fs.Flush();
                            fs.Close();

                            saveSuppressionToolStripMenuItem.Enabled = true;
                            saveToolStripMenuItem.Enabled = true;

                            _acclogger.Clear();
                            int sequenceNumber = 0;
                            foreach(LogEvent logEntry in logFile.LogEvents)
                            {
                                logEntry.SequenceNumber = sequenceNumber++;
                                _acclogger.Log(logEntry);
                            }

                            _graphics = new GraphicsHelper(logFile.ScreenShot, logFile.RootOffset);

                            foreach (SerializedTree tree in logFile.Trees)
                            {
                                tree.SetXmlForTree(_graphics);
                            }

                            tabControl.SelectedTab = tabPageResults;
                            guiLogList.AddGuiLogger(_acclogger, logFile, _graphics);
                        }
                    }
                    catch (InvalidOperationException)
                    {
                        MessageBox.Show("Log file " + fileName + " is invalid or from a previous version of AccChecker");
                    }
                    catch (ArgumentException exception)
                    {
                        MessageBox.Show(exception.Message);
                    }
                    catch (UnauthorizedAccessException exception)
                    {
                        MessageBox.Show(exception.Message + "- Try copying the file locally");
                    }
                }
            }
        }

        private void SetEnableRunVerifications(bool fEnabled)
        {
            // Only allow the button to be enabled when there is a valid hwnd and verification selected.
            bool state = fEnabled;

            if (!Win32API.IsWindow(HwndToVerify()) || lvRoutines.CheckedItems.Count < 1)
            {
                state = false;
            }

            buttonRunVerifications.Enabled = state;
            runNowToolStripMenuItem.Enabled = state;
        }

        private void SetEnableRadioButtonsControls()
        {
            topLevelWindows.Enabled = false;
            textBoxCapturedWindow.Enabled = false;
            btnCaptureParent.Enabled = false;
            tbTypedHwnd.Enabled = false;

            if (rbMainWindow.Checked)
            {
                topLevelWindows.Enabled = true;
            }
            else if (rbCapturedWindow.Checked)
            {
                textBoxCapturedWindow.Enabled = true;
                btnCaptureParent.Enabled = true;
            }
            else if (rbUseTypedHwnd.Checked)
            {
                tbTypedHwnd.Enabled = true;
            }
        }

        private void rbCapturedWindow_Click(object sender, EventArgs e)
        {
            SetEnableRunVerifications(_capturedHwnd != IntPtr.Zero);
        }

        private void rbMainWindow_CheckedChanged(object sender, EventArgs e)
        {
            // only enable the control related to this radio button
            SetEnableRadioButtonsControls();
            SetEnableRunVerifications(topLevelWindows.SelectedIndex != -1);
        }

        private void rbCapturedWindow_CheckedChanged(Object sender, EventArgs e)
        {
            // only enable the control related to this radio button
            SetEnableRadioButtonsControls();
            SetMouseHoverHighlight(rbCapturedWindow.Checked);
            UpdateCapturedHwndText();
        }


        private void rbUseTypedHwnd_CheckedChanged(object sender, EventArgs e)
        {
            // only enable the control related to this radio button
            SetEnableRadioButtonsControls();
            if (rbUseTypedHwnd.Checked)
            {
                tbTypedHwnd_TextChanged(null, null);
            }
            else
            {
                tbTypedHwnd.BackColor = Color.White;

            }
        }

        private void HighlightCallback(Object obj, EventArgs eventArgs)
        {
            // check to make sure we still need to do this in case we killed the timer at the same moment it fired.
            // Only do highlighting on the first page
            if (_HighlightTimer != null && tabControl.SelectedIndex == 0)
            {
                Win32API.POINT point = new Win32API.POINT();
                if (Win32API.GetCursorPos(ref point) && !_lastCursorPosition.Equals(point))
                {
                    IntPtr windowUnderMouse = Win32API.WindowFromPoint(point);

                    Win32API.RECT win32Rect = Win32API.RECT.Empty;

                    uint processIdUnderMouse;
                    Win32API.GetWindowThreadProcessId(windowUnderMouse, out processIdUnderMouse);
                    if (_AccCheckerProcessId != processIdUnderMouse && Win32API.GetWindowRect(windowUnderMouse, ref win32Rect))
                    {
                        Rectangle highlightRect = new Rectangle(win32Rect.left, win32Rect.top, win32Rect.right - win32Rect.left, win32Rect.bottom - win32Rect.top);
                        _highlightRectangle.Location = highlightRect;
                        _highlightRectangle.Visible = true;
                    }
                    _lastCursorPosition = point;
                }
            }
        }

        private void tabControlTabChanged(Object sender, EventArgs e)
        {
            if (tabControl.SelectedIndex == 0 && rbCapturedWindow.Checked)
            {
                SetMouseHoverHighlight(true);
            }
            else
            {
                SetMouseHoverHighlight(false);
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(saveFileDialog.FileName).ToLower() == ".xml")
                {
                    guiLogList.SaveLogs(saveFileDialog.FileName, false);
                }
                else
                {
                    TextFileLogger tfl = new TextFileLogger(saveFileDialog.FileName);
                    _acclogger.DumpToLogger(tfl);
                }
            }

        }
        protected Bitmap GetImage(IntPtr hwnd)
        {
            // make sure the window is on top (restore and then bring to top)
            Win32API.WINDOWPLACEMENT wp = new Win32API.WINDOWPLACEMENT();
            wp.length = Marshal.SizeOf(typeof(Win32API.WINDOWPLACEMENT));
            Win32API.GetWindowPlacement(hwnd, ref wp);
            if (wp.showCmd != Win32API.SW_SHOWMAXIMIZED)
            {
                wp.showCmd = Win32API.SW_RESTORE;
            }
            Win32API.SetWindowPlacement(hwnd, ref wp);

            // Don't set focus if they are using a hot key because it dismisses menus
            // In monitoring mode this does not make sense either.
            if (!_runningTestsUsingHotKey && 
                !_monitormodeOn)
            {
                if (!Win32API.SetForegroundWindow(hwnd))
                {
                    Accessible accessible = null;
                    Accessible.FromWindow(hwnd, out accessible);
                    try
                    {
                        accessible.Select(Win32API.SELFLAG_TAKEFOCUS);
                    }
                    catch (Exception exception)
                    {
                        if (!Accessible.IsExpectedException(exception))
                        {
                            throw;
                        }
                    }
                }
                Thread.Sleep(500);
            }


            return GraphicsHelper.GrabScreenShot(hwnd);
        }

        #region Threading for verifications
        private struct ProgressUpdate
        {
            public String status;
            public Type verificationRoutine;
        }

        private void backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {

            _cursor = Cursor.Current;

            Cursor.Current = Cursors.WaitCursor;

            // Turn the controls back off so users cannot manipulate the controls.
            this.BeginInvoke(new SetControlsEnableStateDelegate(SetControlsEnableState), false);

            IntPtr hwnd = (IntPtr)e.Argument;

            UsageReporter.Instance.BeginWindowScan(_vm, hwnd);

            List<VerificationRoutineWrapper> intrusiveRoutines = new List<VerificationRoutineWrapper>();
            List<VerificationRoutineWrapper> routines = new List<VerificationRoutineWrapper>();
            List<VerificationRoutineWrapper> visualizers = new List<VerificationRoutineWrapper>();
            foreach (VerificationRoutineWrapper rvr in _vm.Routines)
            {
                if ((rvr.Active || rvr.CanVisualize) && rvr.Priority <= _priority)
                {
                    if (_monitormodeOn && rbOptimized.Checked && rvr.IsSlow)
                    {
                        continue;
                    }
                    
                    if (rvr.IsIntrusive)
                    {
                        intrusiveRoutines.Add(rvr);
                    }
                    else if (rvr.CanVisualize)
                    {
                        visualizers.Add(rvr);
                    }
                    else
                    {
                       routines.Add(rvr);
                    }
                }
            }

            
            // The visualizers go after the intrusive routines
            routines.AddRange(intrusiveRoutines);
            routines.AddRange(visualizers);

            if (_monitormodeOn)
            {
                Win32API.RECT win32Rect = Win32API.RECT.Empty;
                Win32API.GetWindowRect(hwnd, ref win32Rect);
                Rectangle highlightRect = new Rectangle(win32Rect.left, win32Rect.top, win32Rect.right - win32Rect.left, win32Rect.bottom - win32Rect.top);
                _highlightRectangle.Location = highlightRect;
                _highlightRectangle.Visible = true;
            }
            
            _startTime = DateTime.Now;
            
            int progress = 0;
            ProgressUpdate p = new ProgressUpdate();
            p.status = "Looking at accessibility tree";
            backgroundWorker.ReportProgress(0, p);
            foreach (VerificationRoutineWrapper rvr in routines)
            {
                p = new ProgressUpdate();
                p.status = rvr.Title;
                p.verificationRoutine = null;
                
                backgroundWorker.ReportProgress(progress * 100 / routines.Count, p);
                

                rvr.Execute(hwnd, _vm.Logger, true, _graphics);
                
                progress++;
                p = new ProgressUpdate();
                p.status = rvr.Title;
                p.verificationRoutine = rvr.Type;
                backgroundWorker.ReportProgress(progress * 100 / routines.Count, p);
                if (backgroundWorker.CancellationPending)
                {
                    break;
                }
            }
            
            if (_monitormodeOn)
            {
                _highlightRectangle.Visible = false;
            }

            UsageReporter.Instance.EndWindowScan(_vm, hwnd);
        }

        Cursor _cursor;
        private void backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            LogFile logFile = new LogFile(HwndToVerify());
            logFile.StartTime = _startTime;
            logFile.StopTime = DateTime.Now;
            logFile.ScreenShot = _graphics.Screenshot;
            logFile.Trees = new List<SerializedTree>(_vm.SerializableTreeList);
            
            guiLogList.AddGuiLogger(_acclogger, logFile, _graphics);
            if (rbMonitoringMode.Checked)
            {
                buttonRunVerifications.Text = "Stop Background Verifications";
            }
            else
            {
                tabControl.SelectedTab = tabPageResults;
                this.Activate();
                buttonRunVerifications.Text = "Run Verifications";
            }
            progressBarVerifications.Value = 0;
            verificationStatus.Text = "Verifications complete";

            Cursor.Current = _cursor;

            // Turn the controls back on for user
            this.BeginInvoke(new SetControlsEnableStateDelegate(SetControlsEnableState), true);
            _runningTests = false;
        }

        private void backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarVerifications.Value = e.ProgressPercentage;
            ProgressUpdate pu = (ProgressUpdate)e.UserState;
            verificationStatus.Text = "Currently running - " + pu.status + "\r\nWarning: Focus change during execution may invalidate results.";
        }
        #endregion

        delegate void SetControlsEnableStateDelegate(bool enableState);
        private void SetControlsEnableState(bool enableState)
        {

            Control[] containers = new Control[] { 
                specifyWindow,
                suppressionFile,
                selectVerificationRoutines,
                mainMenu,
                topLevelWindows,
                btnCaptureParent,
                textBoxCapturedWindow
            };

            if (enableState == false)
            {
                foreach (Control control in containers)
                {
                    control.Tag = control.Enabled;
                    control.Enabled = false;
                }
            }
            else
            {
                foreach (Control control in containers)
                {

                    if (control.Tag is bool && (bool)control.Tag == true)
                    {
                        control.Enabled = true;
                        control.Tag = null;
                    }
                }
            }
        }

        private void lvRoutines_ItemCheck(object sender, ItemCheckedEventArgs e)
        {
            if (e.Item.Checked)
            {
                _vm.EnableRoutine(e.Item.Text);
            }
            else
            {
                _vm.DisableRoutine(e.Item.Text);
            }
            SetEnableRunVerifications(lvRoutines.CheckedItems.Count > 0);
        }

        private void btnLoadSuppresionFile_Click(object sender, EventArgs e)
        {
            if (openSuppresionFileDialog.ShowDialog() == DialogResult.OK)
            {
                SetSuppressionInformation(openSuppresionFileDialog.FileName);
            }
        }

        private void SetSuppressionInformation(string filename)
        {
            tbSuppresionFile.Text = filename;
            _suppressionFile = filename;
        }

        private void _btnClear_Click(object sender, EventArgs e)
        {
            tbSuppresionFile.Text = string.Empty;
            _suppressionFile = string.Empty;
            _vm.ClearSuppressionFiles();
        }

        private void RunVerifications()
        {
            lock (this)
            {
                if (buttonRunVerifications.Enabled == false)
                {
                    return;
                }
                if (_runningTests)
                {
                    return;
                }
                _runningTests = true;
            }

            try
            {
                IntPtr hwnd = HwndToVerify();
                if (Win32API.IsWindow(hwnd))
                {
                    if (!String.IsNullOrEmpty(_suppressionFile))
                    {
                        _vm.ClearSuppressionFiles();

                        try
                        {
                            _vm.AddSuppressionFiles(_suppressionFile);
                        }
                        catch (FileNotFoundException exception)
                        {
                            MessageBox.Show(exception.Message);
                        }
                        catch (InvalidOperationException)
                        {
                            MessageBox.Show("Invalid suppression file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    // clear the tab pages from the previous run. 
                    ClearVerificationAddedTabPages();

                    // clear the logs for a new run
                    _acclogger.Clear();

                    // fetch screenshot
                    Bitmap screenshot = GetImage(hwnd);
                    Win32API.RECT rootRect = new Win32API.RECT();
                    Win32API.GetWindowRect(hwnd, ref rootRect);
                    Point rootOffset = new Point(-rootRect.left, -rootRect.top);
                    _graphics = new GraphicsHelper(screenshot, rootOffset);

                    buttonRunVerifications.Text = "Cancel Verifications";
                    SetMouseHoverHighlight(false);

                    _vm.ClearCache();
                    backgroundWorker.RunWorkerAsync(hwnd);
                    
                    saveSuppressionToolStripMenuItem.Enabled = true;
                    saveToolStripMenuItem.Enabled = true;
                }
                else
                {
                    MessageBox.Show("Hwnd " + hwnd.ToString("X") + " is not a valid hwnd or no longer exists");
                }
            }
            catch (Exception error)
            {
                MessageBox.Show("Error: " + error);
                _runningTests = false;
            }
        }

        private IntPtr HwndToVerify()
        {
            // decide which hwnd to use
            IntPtr hwnd = IntPtr.Zero;
            if (rbCapturedWindow.Checked)
            {
                hwnd = _capturedHwnd;
            }
            else if (rbMainWindow.Checked)
            {
                if (_topLevelHwnds.Count > topLevelWindows.SelectedIndex && topLevelWindows.SelectedIndex >= 0)
                {
                    hwnd = _topLevelHwnds[topLevelWindows.SelectedIndex];
                }
            }
            else if (rbUseTypedHwnd.Checked)
            {
                hwnd = _typedHwnd;
            }
            else if (rbMonitoringMode.Checked)
            {
                hwnd = _vm.MonitoredHwnd;
                if (!_monitormodeOn && hwnd == IntPtr.Zero)
                {
                    // Set this to a valid hwnd so the run button
                    // stays enabled so the user can start
                    // monitoring mode.
                    hwnd = Win32API.GetDesktopWindow();
                }
            }

            return hwnd;
        }

        public void ActivateTab(Type routine)
        {
            foreach (TabPage tabPage in tabControl.TabPages)
            {
                if (tabPage.Tag == routine)
                {
                    tabControl.SelectedTab = tabPage;
                }
            }
        }

        public Panel AddTab(String title, Type routine)
        {
            // If we are not on the right thread call ourseves again on the right one
            if (InvokeRequired)
            {
                IAsyncResult asyncResult = tabControl.BeginInvoke(GraphicsHelper.AddTabDelegate, title, routine);

                // Wait for the result and return it
                return (Panel)tabControl.EndInvoke(asyncResult);
            }
            else
            {
                if (tabControl.TabPages.ContainsKey(title))
                {
                    tabControl.TabPages.RemoveByKey(title);
                }
                tabControl.TabPages.Add(title, title);
                tabControl.TabPages[title].Tag = routine;
                return tabControl.TabPages[title];
            }
        }


        public void AddVerificationDlls(string[] verificationDlls)
        {
            foreach (string dll in verificationDlls)
            {
                try
                {
                    _vm.AddVerificationDLL(dll);
                }
                catch (ArgumentException exception)
                {
                    MessageBox.Show(exception.Message);
                }
            }
        }

        // Go thru all the top level windows and get there title
        private void FillTopLevelWindowList()
        {
            Trace.WriteLine(DateTime.Now.Ticks + " : ############## FillTopLevelWindowList");

            // Save the current selected item to we can reset it
            // once we have rebuilt the list.
            string preSelectedItem = string.Empty;
            if (topLevelWindows.SelectedItem != null)
            {
                preSelectedItem = topLevelWindows.SelectedItem.ToString();
            }

            _topLevelHwnds.Clear();
            topLevelWindows.Items.Clear();
            Win32API.EnumDesktopWindows(IntPtr.Zero, new Win32API.EnumDesktopWindowsProc(TopLevelWindows), IntPtr.Zero);

            // Set the list back to the originally selected item
            if (preSelectedItem != string.Empty)
            {
                int iLoc = topLevelWindows.Items.IndexOf(preSelectedItem);
                if (iLoc > -1)
                {
                    topLevelWindows.SelectedIndex = iLoc;
                }
            }
        }

        // Called by EnumWindows in FillTopLevelWindowList for each toplevel window
        private bool TopLevelWindows(IntPtr hwnd, IntPtr lParam)
        {
            // Is it visible? 
            int style = Win32API.GetWindowLong(hwnd, Win32API.GWL_STYLE);

            if (Win32API.IsBitSet(style, (int)Win32API.WS_VISIBLE))
            {

                string title = Win32API.GetWindowText(hwnd);
                if (string.IsNullOrEmpty(title))
                {
                    uint processId;
                    Win32API.GetWindowThreadProcessId(hwnd, out processId);
                    Process process = Process.GetProcessById((int)processId);
                    title = "Untitled [process " + process.ProcessName + " (" + processId + ")]";
                }

                // exclude AccCheck from this list
                if (title != this.Text)
                {
                    _topLevelHwnds.Add(hwnd);
                    topLevelWindows.Items.Add(title);
                }
            }

            return true;
        }

        private void runNowToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (rbMonitoringMode.Checked)
            {
                ChangeMonitoringMode(!_monitormodeOn);
            }
            else
            {
                RunVerifications();
            }
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void saveSuppressionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string preFilter = saveFileDialog.Filter;
            saveFileDialog.Filter = openSuppresionFileDialog.Filter;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                guiLogList.SaveLogs(saveFileDialog.FileName, true);
            }

            saveFileDialog.Filter = preFilter;
        }

        private void RefreshWindowsList_Click(object sender, EventArgs e)
        {
            FillTopLevelWindowList();
            SetEnableRunVerifications(true);
        }

        private void buttonRunVerificationClick(object sender, EventArgs e)
        {
            if (_runningTests)
            {
                backgroundWorker.CancelAsync();
                return;
            }
            if (rbMonitoringMode.Checked)
            {
                ChangeMonitoringMode(!_monitormodeOn);
            }
            else
            {
                RunVerifications();
            }
        }

        void ChangeMonitoringMode(bool monitoringMode)
        {
            _vm.MonitorEnabled(monitoringMode, new VerificationManager.MonitorCallback(RunVerifications));
            _monitormodeOn = monitoringMode;
            if (monitoringMode)
            {
                runNowToolStripMenuItem.Text = "&Stop Now         Ctrl+Shift+=";
                buttonRunVerifications.Text = "Stop Background Verifications";
            }
            else
            {
                runNowToolStripMenuItem.Text = "&Run Now         Ctrl+Shift+=";
                buttonRunVerifications.Text = "Run Background Verifications";
            }
        }

        private void SetMouseHoverHighlight(bool on)
        {
            if (on)
            {
                if (_HighlightTimer != null)
                {
                    // there is already a timer going so there is nothing to do
                    return;
                }

                int messageDurationValue;
                Win32API.SystemParametersInfo(Win32API.SPI_GETMESSAGEDURATION, 0, out messageDurationValue, 0);

                // base the hover delay on message duration unless it;s too short
                int hoverDelay = (messageDurationValue / 2) * 1000;
                if (hoverDelay < 2)
                {
                    hoverDelay = 2;
                }

                _HighlightTimer = new System.Windows.Forms.Timer();
                _HighlightTimer.Tick += new EventHandler(HighlightCallback);
                _HighlightTimer.Interval = hoverDelay;
                _HighlightTimer.Start();
            }
            else
            {
                if (_HighlightTimer != null)
                {
                    _HighlightTimer.Dispose();
                    _HighlightTimer = null;
                }

                _highlightRectangle.Visible = false;
            }
        }

        private void tbSuppresionFile_TextChanged(object sender, EventArgs e)
        {
            _btnClear.Enabled = (tbSuppresionFile.Text.Length != 0);
        }

        private void menu_About_Click(object sender, EventArgs e)
        {
            AboutBox aboutBox = new AboutBox();
            aboutBox.ShowDialog(this);
        }

        private void menu_Help_Click(object sender, EventArgs e)
        {
            MainForm.OpenHelpFile(string.Empty, string.Empty);
        }

        /// -------------------------------------------------------------------
        /// <summary>Opens the help file if it is not already opened.  If a 
        /// book mark is supplied, it will navigate to the bookmark.</summary>
        /// -------------------------------------------------------------------
        public static void OpenHelpFile(string helpFileName, string bookMark)
        {
            string fullPathAndName = string.Empty;
            Cursor prevCursor = Cursor.Current;
            string fileName;

            if (string.IsNullOrEmpty(helpFileName))
                fileName = HELPFILE_NAME;
            else
                fileName = helpFileName;

            // Use the path of the executing assembly to determine where the help file is located
            fullPathAndName = Assembly.GetExecutingAssembly().Location;
            fullPathAndName = Path.Combine(fullPathAndName.Substring(0, fullPathAndName.LastIndexOf(@"\")), fileName);

            if (!File.Exists(fullPathAndName))
            {
                MessageBox.Show(string.Format("Could not find help file at {0}", fullPathAndName), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                
                try
                {
                    Cursor.Current = Cursors.WaitCursor;

                    if (!IsInternetExplorerAvailable)
                    {
                        StartInternetExplorer();
                    }

                    fullPathAndName += string.Format("#{0}", bookMark);
                    InvokeIEMethod("Visible", BindingFlags.SetProperty, new Object[] { true });
                    InvokeIEMethod("Navigate", BindingFlags.InvokeMethod, new Object[] { fullPathAndName });
                    
                    // Set the title. Unfortunately, if the help file is already displayed, and we just supply the 
                    // fullpath again, then the title changes to the full path + help file name instead of just the 
                    // help file name.  We need to reset the title name.
                    //MessageBox.Show("1");
                    //Exception handler being added in order to avoid unknown error from coming in pop-up, even
                    //though window is already open
                    try
                    {
                        Object mainDocument = InvokeIEMethod("Document", BindingFlags.GetProperty, null);
                        InvokeIEMethod(mainDocument, "title", BindingFlags.SetProperty, new Object[] { fileName });
                    }
                    catch (Exception)
                    {
                    }

                    IntPtr hwndPtr = GetHwnd();
                    Accessible accessible = null;
                    Accessible.FromWindow(hwndPtr, out accessible);
                    accessible.Select(Win32API.SELFLAG_TAKEFOCUS);


                }
                catch (Exception error)
                {
                    //MessageBox.Show("!!error!!");
                    MessageBox.Show(error.Message);
                    // IE may be up and running but invisible. If anything happens to this object,
                    // then we need to kill it off so the IE is not still hanging around which the 
                    // user cannot remove from memory.
                    QuitInternetExplorer();
                }
                finally
                {
                    Cursor.Current = prevCursor;
                }
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Get the hwnd for the IE window</summary>
        /// -------------------------------------------------------------------
        private static IntPtr GetHwnd()
        {
            object hwnd = InvokeIEMethod("HWND", BindingFlags.GetProperty, null);
            IntPtr hwndPtr = IntPtr.Zero;
            if (hwnd is Int64)
            {
                hwndPtr = new IntPtr((long)hwnd);
            }
            else if (hwnd is Int32)
            {
                hwndPtr = new IntPtr((int)hwnd);
            }
            else
            {
                throw new ArgumentException("Value returned for hwnd is not correct data type but is " + hwnd.GetType().ToString());
            }

            return hwndPtr;
        }

        /// -------------------------------------------------------------------
        /// <summary>Close Internet Explorer.  Eat any exceptions that might 
        /// happen.</summary>
        /// -------------------------------------------------------------------
        private static void QuitInternetExplorer()
        {
            try
            {
                InvokeIEMethod("Quit", BindingFlags.InvokeMethod, new Object[] { });
            }
            catch
            {
                // Just eat everything
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Determines whether _internetExplorer is usable</summary>
        /// -------------------------------------------------------------------
        private static bool IsInternetExplorerAvailable
        {
            get
            {
                try
                {
                    if (_internetExplorer == null)
                    {
                        return false;
                    }
                    else
                    {
                        // try to see if we can call something on the object 
                        InvokeIEMethod("Visible", BindingFlags.SetProperty, new Object[] { true });
                        return true;
                    }
                }
                catch (COMException)
                {
                    return false;
                }
            }
        }

        /// -------------------------------------------------------------------
        /// <summary>Opens and positions Internet Explorer.  Invoke 
        /// 'visible' to show the form.</summary>
        /// -------------------------------------------------------------------
        private static void StartInternetExplorer()
        {
            _internetExplorerType = Type.GetTypeFromProgID("InternetExplorer.Application");
            _internetExplorer = Activator.CreateInstance(_internetExplorerType);

            InvokeIEMethod("Width", BindingFlags.SetProperty, 600);
            InvokeIEMethod("Height", BindingFlags.SetProperty, 450);
        }

        /// -------------------------------------------------------------------
        /// <summary>Invoke a method/property in DHTML object model</summary>
        /// <param name="name">The name of API</param>
        /// <param name="invokeType">Type of invokation</param>
        /// <param name="parameters">Parameter array</param>
        /// <returns>The return object from invoked APIs</returns>
        /// -------------------------------------------------------------------
        static Object InvokeIEMethod(String name, BindingFlags invokeType, params Object[] parameters)
        {
            return InvokeIEMethod(_internetExplorer, name, invokeType, parameters);
        }

        /// -------------------------------------------------------------------
        /// <summary>Invoke a method/property in DHTML object model</summary>
        /// <param name="windowObject"></param>
        /// <param name="name">The name of API</param>
        /// <param name="invokeType">Type of invokation</param>
        /// <param name="parameters">Parameter array</param>
        /// <returns>The return object from invoked APIs</returns>
        /// -------------------------------------------------------------------
        static Object InvokeIEMethod(object windowObject, String name, BindingFlags invokeType, params Object[] parameters)
        {
            return _internetExplorerType.InvokeMember(name, BindingFlags.IgnoreCase | invokeType, null, windowObject, parameters);
        }

        private void topLevelWindows_TextChanged(object sender, EventArgs e)
        {
            if (rbMainWindow.Checked == true)
            {
                SetEnableRunVerifications(true);
            }
        }

        private void topLevelWindows_KeyDown(object sender, KeyEventArgs e)
        {
            if (rbMainWindow.Checked == true)
            {
                SetEnableRunVerifications(true);
            }
        }

        private void topLevelWindows_KeyUp(object sender, KeyEventArgs e)
        {
            if (rbMainWindow.Checked == true)
            {
                SetEnableRunVerifications(true);
            }
        }

        private void topLevelWindows_Enter(object sender, EventArgs e)
        {
            if (rbMainWindow.Checked == true)
            {
                SetEnableRunVerifications(true);
            }
        }

        private void MainForm_GotFocus(object sender, EventArgs e)
        {
            // Whenever the form gets the focus, refresh 
            // the list of windows since the user could have 
            // added or deleted a window from the desktop.
            RefreshWindowsList_Click(null, null);
        }

        private void autoLoadToolStripMenuItem_CheckStateChanged(object sender, EventArgs e)
        {
            if (autoLoadToolStripMenuItem.Checked)
            {
                SaveSettings();
                _vm.AddAllDllsInCurrentDirectory();
                PopulateFromVM();
                LoadSettings();
            }
        }

        private void ClearVerificationAddedTabPages()
        {
            for (int i = tabControl.TabPages.Count - 1; i >= 0; i--)
            {
                if (tabControl.TabPages[i].Text != "Verifications" && tabControl.TabPages[i].Text != "Results")
                {
                    tabControl.TabPages.RemoveAt(i);
                }
            }
        }

        private void rbMonitoringMode_CheckedChanged(object sender, EventArgs e)
        {
            if (rbMonitoringMode.Checked)
            {
                buttonRunVerifications.Text = "Run Background Verifications";
                SetEnableRunVerifications(true);

                if (rbOptimized.Checked)
                {
                    rbOptimized_CheckedChanged(sender, e);
                }
            }
            else
            {
                buttonRunVerifications.Text = "Run Verifications";
                if (rbOptimized.Checked)
                {
                    PopulateFromVM();
                }
            }
            gbMonitoringOptions.Enabled = rbMonitoringMode.Checked;
        }

        private void cbPriority_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (cbPriority.SelectedIndex)
            {
                case 0: _priority = 3; break;
                case 1: _priority = 1; break;
                case 2: _priority = 2; break;
            }
            _acclogger.Priority = _priority;
            
            SetEnableRunVerifications(true);
        }

        private void rbOptimized_CheckedChanged(object sender, EventArgs e)
        {
            // go back to front so removal does end up skipping things
            for (int i = lvRoutines.Items.Count - 1; i >= 0; i--)
            {
                VerificationRoutineWrapper rvr = (VerificationRoutineWrapper)lvRoutines.Items[i].Tag;
                if (rvr.IsSlow)
                {
                    lvRoutines.Items[i].Remove();
                }
            }

            SetEnableRunVerifications(true);
        }

        private void rbSelected_CheckedChanged(object sender, EventArgs e)
        {
            PopulateFromVM();
        }

    }
}
