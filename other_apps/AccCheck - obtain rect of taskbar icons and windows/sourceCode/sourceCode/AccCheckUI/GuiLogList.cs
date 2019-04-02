using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using System.Xml.Serialization;
using System.Drawing;
using System.Data;
using System.Text;
using System.IO;
using System.Windows.Forms;
using AccCheck.Verification;
using AccCheck.Logging;
using AccCheck;

namespace AccCheckUI
{
    public partial class GuiLogList : UserControl
    {
        private class LogData
        {
            public LogData(AccumulatingLogger logger, LogFile logFile, GraphicsHelper graphicsHelper)
            {
                _logger = logger;
                _logFile = logFile;
                _graphicsHelper = graphicsHelper;
            }

            public AccumulatingLogger _logger;
            public LogFile _logFile;
            public GraphicsHelper _graphicsHelper;
        }

        private const int s_threshholdForPagingOutLogs = 50;
        
        private BackgroundWorker _backgroundWorker = null;
        private int _activeListEntries = 0;
        private int _tempLogSequence = 0;
        
        public GuiLogList()
        {
            InitializeComponent();

            string tempFolder = Path.Combine(Path.GetTempPath(), "AccChecker");
            if (Directory.Exists(tempFolder))
            {
                DirectoryInfo tempDirectory = new DirectoryInfo(tempFolder);
                FileInfo [] files = tempDirectory.GetFiles();
                foreach (FileInfo file in files)
                {
                    File.Delete(file.FullName);
                }
            }
        }

        public void AddGuiLogger(AccumulatingLogger logger, LogFile logFile, GraphicsHelper graphicsHelper)
        {
            AccumulatingLogger newLogger = new AccumulatingLogger();
            logger.DumpToLogger(newLogger);
            
            List<SerializedTree> trees = new List<SerializedTree>();
            foreach (SerializedTree tree in logFile.Trees)
            {
                trees.Add(tree.Clone());
            }

            logFile.Trees = trees;
            
            LogData logData = new LogData(newLogger, logFile, graphicsHelper);

            guiLogger.AccLogger = logger;
            guiLogger.GraphicsHelper = graphicsHelper;

            lvLogs.SuspendLayout();
            lvLogs.Hide();
            ListViewItem lvi = new ListViewItem(logFile.StartTime.ToLongTimeString());
            lvi.SubItems.Add(logFile.WindowName);
            lvi.SubItems.Add(logFile.FileVersionInfo.FileName);
            lvi.SubItems.Add(logFile.WindowClass);
            lvi.Tag = logData;             
            
            lvLogs.Items.Add(lvi);
            lvLogs.Show();

            foreach (ListViewItem item in lvLogs.Items)
            {
                item.Selected = false;
            }

            lvLogs.Items[lvi.Index].Selected = true;
            lvLogs.EnsureVisible(lvi.Index);
            _activeListEntries++;
            lvLogs.ResumeLayout();
            if (_activeListEntries > s_threshholdForPagingOutLogs)
            {
                DoBackgroundWork();
            }
            
        }

        public void SaveLogs(string filename, bool isForSuppressions)
        {
            if (lvLogs.Items.Count == 0)
            {
                return;
            }
            
            if (lvLogs.SelectedItems.Count == 0)
            {
                lvLogs.Items[0].Selected = true;
            }

            int i = 1;
            foreach (ListViewItem item in lvLogs.SelectedItems)
            {
                LogData logData = GetLogData(item);

                string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(filename);
                int index = filename.LastIndexOf(fileNameWithoutExtension) + fileNameWithoutExtension.Length;
                if (Path.GetExtension(filename) == ".xml")
                {
                    XMLSerializingLogger xsl = new XMLSerializingLogger();

                    logData._logger.DumpToLogger(xsl);
                    try
                    {
                        Cursor.Current = Cursors.WaitCursor;
                        if (lvLogs.SelectedItems.Count == 1)
                        {
                            xsl.SerializeToFile(filename, logData._logFile);
                        }
                        else
                        {
                            xsl.SerializeToFile(filename.Insert(index, (i++).ToString()), logData._logFile);
                        }
                    }
                    catch (OutOfMemoryException )
                    {
                        MessageBox.Show("AccChecker ran out of memory. Try running with less verifications.");
                    }
                    finally
                    {
                        Cursor.Current = Cursors.Default;
                    }

                    if (isForSuppressions)
                    {
                        xsl.LogLevel = EventLevel.Warning;
                        if (logData._logFile.AmbiguousResources != null && VerificationManager.IsQueryStringAddinEnabled)
                        {
                            MessageBox.Show("There are ambiguous resources that may cause suppressions to fail in other languages.  Running verifications on Pseudo loc build should resolve this", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
                else
                {
                    TextFileLogger tfl = new TextFileLogger(filename.Insert(index, (i++).ToString()));
                    logData._logger.DumpToLogger(tfl);
                }
            }
        }

        private void DoBackgroundWork()
        {
            if (_backgroundWorker == null)
            {
                _backgroundWorker = new BackgroundWorker();
                _backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(SaveTemporaryLog);
            }
            
            if (_backgroundWorker.IsBusy)
            {
                return;
            }
            
            for (int i = 0; i < lvLogs.Items.Count; i++)
            {
                ListViewItem item = lvLogs.Items[i];
                
                if (i < 5)
                {
                    // don't page out the first page because people often look
                    // at the first page.
                    continue;
                }
            
                if (lvLogs.Bounds.IntersectsWith(item.Bounds))
                {
                    // don't page out items that are showing right now.
                    continue;
                }
                
                if (item.Selected == true)
                {
                    // don't page out items that are selected.
                    continue;
                }
            
                LogData logData = item.Tag as LogData;
                // If this is not a LogData then the log has already been saved to a 
                // temporary file.
                if (logData != null)
                {
                    if (!_backgroundWorker.IsBusy)
                    {
                        _backgroundWorker.RunWorkerAsync(item);
                    }
                    break;
                }
            }
        }
        
        private void SaveTemporaryLog(object sender, DoWorkEventArgs e)
        {
            ListViewItem item = (ListViewItem)e.Argument;
            LogData logData = item.Tag as LogData;

            string tempFolder = Path.Combine(Path.GetTempPath(), "AccChecker");
            if (!Directory.Exists(tempFolder))
            {
                Directory.CreateDirectory(tempFolder);
            }
            string fileName = tempFolder + "\\AccCheckerLog" + _tempLogSequence++ + "_" + logData._logFile.FileVersionInfo.FileName + ".xml";
            
            XMLSerializingLogger xsl = new XMLSerializingLogger();
            logData._logger.DumpToLogger(xsl);
            xsl.SerializeToFile(fileName, logData._logFile);

            item.Tag = fileName;
            
            _activeListEntries--;
        }

        private LogData GetLogData(ListViewItem item)
        {
            LogData logData = item.Tag as LogData;
            if (logData == null)
            {
                string fileName = item.Tag as string;
                if (fileName != null)
                {
                    using (FileStream fs = new FileStream(fileName, FileMode.Open))
                    {
                        XmlSerializer xml = new XmlSerializer(typeof(LogFile));
                        
                        LogFile logFile = (LogFile)xml.Deserialize(fs);

                        AccumulatingLogger logger = new AccumulatingLogger();
                        foreach(LogEvent entry in logFile.LogEvents)
                        {
                            logger.Log(entry);
                        }
                        logData = new LogData(logger, logFile, new GraphicsHelper(logFile.ScreenShot, logFile.RootOffset));

                        item.Tag = logData;
                        
                        _activeListEntries++;
                        
                        fs.Flush();
                        fs.Close();
                        
                        File.Delete(fileName);
                    }
                }
            }
            
            return logData;
        }
        
        private void lvLogs_ColumnClick(object sender, ColumnClickEventArgs args)
        {
            lvLogs.ListViewItemSorter = new ListViewItemSorter(args.Column);
        }

        private void lvLogs_KeyDown(object sender, KeyEventArgs args)
        {
            if (args.Control && (args.KeyCode == Keys.A))
            {
                foreach (ListViewItem item in lvLogs.Items)
                {
                    item.Selected = true;
                }
            }
        }
        
        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lvLogs.Items.Count > 0)
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Filter = "XML files|*.xml|Text files|*.txt";
                saveFileDialog.DefaultExt = "xml";
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    SaveLogs(saveFileDialog.FileName, false);
                }
            }
        }

        private void removeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            int index = -1;

            lvLogs.SuspendLayout();

            foreach (ListViewItem item in lvLogs.SelectedItems)
            {
                index = item.Index;
                lvLogs.Items.Remove(item);
            }

            guiLogger.AccLogger = new AccumulatingLogger();
            guiLogger.UpdateUI();

            // Select the another element after removing one if there are any left.
            if (index >= 0 && lvLogs.Items.Count > 0)
            {
                if (index > lvLogs.Items.Count - 1)
                {
                    lvLogs.Items[lvLogs.Items.Count - 1].Selected = true;
                }
                else
                {
                    lvLogs.Items[index].Selected = true;
                }
            }

            lvLogs.ResumeLayout();

        }

        private void lvLogs_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Changing this condition to == 1 to avoid updating UI unnecessarily
            if (lvLogs.SelectedItems.Count == 1)
            {
                // the tab control and remove the tabs so the when we call  
                TabControl tabControl = (TabControl)this.Parent.Parent;
                for (int i = tabControl.TabPages.Count - 1; i >= 0; i--)
                {
                    if (tabControl.TabPages[i].Text != "Verifications" && tabControl.TabPages[i].Text != "Results")
                    {
                        tabControl.TabPages.RemoveAt(i);
                    }
                }
                
                ListViewItem item = lvLogs.SelectedItems[0];
                if (item != null)
                {
                    LogData logData = GetLogData(item);

                    guiLogger.AccLogger = logData._logger;
                    guiLogger.GraphicsHelper = logData._graphicsHelper;
                    guiLogger.UpdateUI();

                    foreach(SerializedTree serializedTree in logData._logFile.Trees)
                    {
                        serializedTree.SetXmlForTree(logData._graphicsHelper);
                    }
                }
                if (_activeListEntries > s_threshholdForPagingOutLogs)
                {
                    DoBackgroundWork();
                }
            }
        }
    }

    internal class ListViewItemSorter : System.Collections.IComparer
    {
        public ListViewItemSorter(int columnToSort)
        {
            _column = columnToSort;
        }

        public int Compare(object x, object y)
        {
            string string1 = "";
            string string2 = "";
            if (x != null)
            {
                string1 = ((ListViewItem)x).SubItems[_column].Text;
            }

            if (y != null)
            {
                string2 = ((ListViewItem)y).SubItems[_column].Text;
            }

            return String.Compare(string1, string2);
        }

        private int _column;
    }
}
