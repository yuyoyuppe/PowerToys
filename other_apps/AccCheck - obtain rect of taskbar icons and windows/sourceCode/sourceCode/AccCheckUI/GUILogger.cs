using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Text;
using System.Windows.Forms;
using AccCheck.Logging;
using AccCheck;

namespace AccCheckUI
{
    public partial class GUILogger : UserControl, ILogger
    {
        private AccumulatingLogger _acclogger;
        private GraphicsHelper _graphicsHelper;
        private List<EventLevel> _eventTypes = new List<EventLevel>();

        // this array contain the count of errors, warning and messeages
        // The index to use is the enum EventLevel
        private int[] _eventCounts = new int[Enum.GetValues(typeof(EventLevel)).Length];
        private int _cSuppressions = 0;
        private EventLevel _logLevel;
        
        private List<ListViewItem> _listitems = new List<ListViewItem>();
        public GUILogger()
        {
            InitializeComponent();
            SetIncludedEventLevels();
        }

        public GraphicsHelper GraphicsHelper
        {
            get
            {
                return _graphicsHelper;
            }
            set
            {
                _graphicsHelper = value;
                if (value != null)
                {
                    pbScreenshot.Image = (Image)_graphicsHelper.Screenshot;
                }
            }
        }

        public AccumulatingLogger AccLogger
        {
            get
            {
                return _acclogger;
            }
            set
            {
                _acclogger = value;
            }
        }

        private void btnErrors_Click(object sender, EventArgs e)
        {
            SetIncludedEventLevels();
            UpdateUI();
        }

        public void UpdateUI()
        {
            Clear();
            if (_acclogger != null)
            {
                _acclogger.DumpToLogger(this);
                lvLog.Hide();
                lvLog.Items.AddRange(_listitems.ToArray());
                if (lvLog.Items.Count != 0)
                {
                    lvLog.Items[0].Selected = true;
                }
                lvLog.Show();
                _listitems.Clear();
                btnErrors.Text = _eventCounts[(int)EventLevel.Error] + " &Errors";
                btnWarnings.Text = _eventCounts[(int)EventLevel.Warning] + " &Warnings";
                btnMessages.Text = _eventCounts[(int)EventLevel.Information] + " &Messages";
                btnSuppressions.Text = _cSuppressions + " &Suppressions";
            }
        }
        
        private void lvLog_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvLog.SelectedItems.Count == 1)
            {
                //int ixSelected = lvLog.SelectedIndices[0];
                LogEvent ev = _acclogger.Get((int)lvLog.SelectedItems[0].Tag);
                if (ev != null)
                {
                    lvLog.AccessibleName = ev.ToString();
                    pbScreenshot.Image = _graphicsHelper.Screenshot;

                    if (pbScreenshot.Image != null)
                    {
                        _graphicsHelper.HighLightRectangle(pbScreenshot.Image, ev.Rectangle);
                    }

                    pbScreenshot.Refresh();

                    lblID.Text = ev.EventID;
                    if (ev.VerificationRoutine != null)
                        lblRoutine.Text = ev.VerificationRoutine;
                    else
                        lblRoutine.Text = "";
                    lblText.Text = ev.Text;

                    lblAccName.Text = ev.AccName;
                    lblValue.Text = ev.AccValue;
                    lblRole.Text = ev.AccRole;
                    lblState.Text = ev.AccState;
                    lblBoundingRectangle.Text = ev.Rectangle.ToString();
                    lblWindowClass.Text = ev.Classname;
                    tbStack.Text = ev.ExtraInformation;
                }
            }
            btnVisualize.Enabled = (lvLog.SelectedItems.Count == 1);
        }

        private void lvLog_ColumnClick(object sender, ColumnClickEventArgs args)
        {
            lvLog.ListViewItemSorter = new ListViewItemSorter(args.Column);
        }

        private void lvLog_KeyDown(object sender, KeyEventArgs args)
        {
            if (args.Control && (args.KeyCode == Keys.A))
            {
                foreach (ListViewItem item in lvLog.Items)
                {
                    item.Selected = true;
                }
            }
        }

        #region Logger Members
        public int InformationalCount
        {
            get { return _eventCounts[(int)EventLevel.Information]; }
        }

        public int TraceCount
        {
            get { return _eventCounts[(int)EventLevel.Trace]; }
        }

        public int ErrorCount
        {
            get { return _eventCounts[(int)EventLevel.Error]; }
        }

        public int WarningCount
        {
            get { return _eventCounts[(int)EventLevel.Warning]; }
        }

        public void Log(LogEvent e)
        {
            if (e.Level < _logLevel)
            {
                return;
            }

            if (e.Suppressed)
            {
                _cSuppressions++;
                if (!btnSuppressions.Checked)
                {
                    return;
                }
            }
            else
            {
                _eventCounts[(int)e.Level]++;
            }

            if (_eventTypes.Contains(e.Level))
            {
                ListViewItem lvi = new ListViewItem();
                SyncListItemImage(lvi, e);
                lvi.SubItems.Add(e.Timestamp.ToLongTimeString());
                lvi.SubItems.Add(e.EventID);
                lvi.SubItems.Add(e.Text);
                lvi.SubItems.Add(e.Priority.ToString());
                lvi.Tag = e.SequenceNumber;
                if (e.VerificationRoutine != null)
                {
                    lvi.SubItems.Add(e.VerificationRoutine);
                }
                _listitems.Add(lvi);
            }

            
        }

        public EventLevel LogLevel
        {
            get { return _logLevel; }
            set { _logLevel = value; }
        }

        #endregion

        public void Clear()
        {
            for (int i = 0; i < _eventCounts.Length; i++)
            {
                _eventCounts[i] = 0;
            }
            _cSuppressions = 0;
            _listitems.Clear();
            lvLog.Items.Clear();

            pbScreenshot.Image = null;

            lblID.Text = "";
            lblRoutine.Text = "";
            lblText.Text = "";
            lblAccName.Text = "";
            lblValue.Text = "";
            lblRole.Text = "";
            lblState.Text = "";
            lblBoundingRectangle.Text = "";
            lblWindowClass.Text = "";
            tbStack.Text = "";
            
        }

        private void GUILogger_Load(object sender, EventArgs e)
        {
            SetIncludedEventLevels();
            UpdateUI();
        }

        private void btnVisualize_Click(object sender, EventArgs e)
        {
            if (lvLog.SelectedItems.Count == 1)
            {
                LogEvent ev = _acclogger.Get((int)lvLog.SelectedItems[0].Tag);
                if (ev != null)
                {
                    _graphicsHelper.Visualize(new VisualizeInfo(ev.AccName, ev.AccRole, ev.AccState, ev.AccValue, ev.Rectangle));
                }
            }
        }

        private void contextMenuStripLogEntries_Opening(object sender, CancelEventArgs e)
        {
            if (lvLog.SelectedItems.Count == 1)
            {
                ListViewItem focusedItem = lvLog.FocusedItem;
                LogEvent logEvent = _acclogger.Get((int)focusedItem.Tag) as LogEvent;
                if (logEvent != null)
                {
                    if (logEvent.Suppressed)
                    {
                        suppressToolStripMenuItem.Text = "Unsuppress";
                    }
                    else
                    {
                        suppressToolStripMenuItem.Text = "Suppress";
                    }

                    if (logEvent.Level == EventLevel.Error || logEvent.Level == EventLevel.Warning)
                    {
                        suppressToolStripMenuItem.Enabled = true;
                    }
                    else
                    {
                        suppressToolStripMenuItem.Enabled = false;
                    }
                }
            }
            else
            {
                suppressToolStripMenuItem.Text = "Toggle suppressions";
            }
            
            return;
        }

        private void suppressToolStripMenuItem_Click(object sender, EventArgs e)
        {

            foreach (ListViewItem item in lvLog.SelectedItems)
            {
                LogEvent logEvent = _acclogger.Get((int)item.Tag) as LogEvent;
                if (logEvent != null && logEvent.Level > EventLevel.Information)
                {
                    logEvent.Suppressed = !logEvent.Suppressed;
                    SyncListItemImage(lvLog.SelectedItems[0], logEvent);

                    if (logEvent.Suppressed)
                    {
                        _cSuppressions++;
                        _eventCounts[(int)logEvent.Level]--;
                    }
                    else
                    {
                        _cSuppressions--;
                        _eventCounts[(int)logEvent.Level]++;
                    }
                }
            }

            UpdateUI();
            
        }
        
        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lvLog.SelectedItems.Count == 1)
            {
                LogEvent logEvent = _acclogger.Get((int)lvLog.SelectedItems[0].Tag) as LogEvent;
                if (logEvent != null)
                {
                    string clipboardData = 
                        "AccChecker: Detailed result information.\r\n" +
                        "EventID: " + logEvent.EventID + "\r\n" +
                        "VerificationRoutine: " + logEvent.VerificationRoutine + "\r\n" +
                        "Text: " + logEvent.Text + "\r\n" +
                        "Rectangle: " + logEvent.Rectangle + "\r\n" +
                        "Classname: " + logEvent.Classname + "\r\n" +
                        "AccName: " + logEvent.AccName + "\r\n" +
                        "AccValue: " + logEvent.AccValue + "\r\n" +
                        "AccRole: " + logEvent.AccRole + "\r\n" +
                        "AccState: " + logEvent.AccState + "\r\n" +
                        "ExtraInformation: " + logEvent.ExtraInformation;
                    
                    Clipboard.SetText(clipboardData);
                }
            }
        }
    
        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lvLog.SelectedItems.Count == 1)
            {
                LogEvent logEvent = _acclogger.Get((int)lvLog.SelectedItems[0].Tag) as LogEvent;
                if (logEvent != null)
                {
                    MainForm.OpenHelpFile(logEvent.HelpPage, logEvent.HelpIndex);
                }
            }
        }

        private void SetIncludedEventLevels()
        {
            _eventTypes.Clear();
            if (btnErrors.Checked)
            {
                _eventTypes.Add(EventLevel.Error);
            }
            if (btnWarnings.Checked)
            {
                _eventTypes.Add(EventLevel.Warning);
            }
            if (btnMessages.Checked)
            {
                _eventTypes.Add(EventLevel.Information);
            }
        }

        private void SyncListItemImage(ListViewItem lvi, LogEvent logEvent)
        {
            if (logEvent.Suppressed)
            {
                lvi.ImageIndex = 3;
            }
            else
            {
                switch (logEvent.Level)
                {
                    case EventLevel.Error:          lvi.ImageIndex = 0; break;
                    case EventLevel.Warning:        lvi.ImageIndex = 1; break;
                    case EventLevel.Information:    lvi.ImageIndex = 2; break;
                }
            }
        }
    }
}
