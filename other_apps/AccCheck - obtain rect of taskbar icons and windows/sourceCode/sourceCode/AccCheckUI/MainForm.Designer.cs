// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

namespace AccCheckUI
{
    partial class MainForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.mainMenu = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.verificationsDLLToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.logFileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.autoLoadToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator = new System.Windows.Forms.ToolStripSeparator();
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.saveSuppressionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.verificationsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runNowToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.enableQueryStringAddinStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.enableAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.disableAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.isIncludePassResultsInLogToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openDLLDialog = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            this.backgroundWorker = new System.ComponentModel.BackgroundWorker();
            this.openSuppresionFileDialog = new System.Windows.Forms.OpenFileDialog();
            this.TopToolStripPanel = new System.Windows.Forms.ToolStripPanel();
            this.RightToolStripPanel = new System.Windows.Forms.ToolStripPanel();
            this.LeftToolStripPanel = new System.Windows.Forms.ToolStripPanel();
            this.ContentPanel = new System.Windows.Forms.ToolStripContentPanel();
            this.tabControl = new System.Windows.Forms.TabControl();
            this.tabPageVerifications = new System.Windows.Forms.TabPage();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.cbPriority = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.topLevelWindows = new System.Windows.Forms.ComboBox();
            this.textBoxCapturedWindow = new System.Windows.Forms.TextBox();
            this.btnCaptureParent = new System.Windows.Forms.Button();
            this.tbTypedHwnd = new System.Windows.Forms.TextBox();
            this.specifyWindow = new System.Windows.Forms.GroupBox();
            this.gbMonitoringOptions = new System.Windows.Forms.GroupBox();
            this.rbSelected = new System.Windows.Forms.RadioButton();
            this.rbOptimized = new System.Windows.Forms.RadioButton();
            this.rbMonitoringMode = new System.Windows.Forms.RadioButton();
            this.rbMainWindow = new System.Windows.Forms.RadioButton();
            this.rbCapturedWindow = new System.Windows.Forms.RadioButton();
            this.rbUseTypedHwnd = new System.Windows.Forms.RadioButton();
            this.suppressionFile = new System.Windows.Forms.GroupBox();
            this._btnClear = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.tbSuppresionFile = new System.Windows.Forms.TextBox();
            this.btnLoadSuppresionFile = new System.Windows.Forms.Button();
            this.runSelectedVerification = new System.Windows.Forms.GroupBox();
            this.buttonRunVerifications = new System.Windows.Forms.Button();
            this.progressBarVerifications = new System.Windows.Forms.ProgressBar();
            this.verificationStatus = new System.Windows.Forms.Label();
            this.selectVerificationRoutines = new System.Windows.Forms.GroupBox();
            this.lvRoutines = new System.Windows.Forms.ListView();
            this.chRoutine = new System.Windows.Forms.ColumnHeader();
            this.Priority = new System.Windows.Forms.ColumnHeader();
            this.tabPageResults = new System.Windows.Forms.TabPage();
            this.guiLogList = new AccCheckUI.GuiLogList();
            this.mainMenu.SuspendLayout();
            this.tabControl.SuspendLayout();
            this.tabPageVerifications.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.specifyWindow.SuspendLayout();
            this.gbMonitoringOptions.SuspendLayout();
            this.suppressionFile.SuspendLayout();
            this.runSelectedVerification.SuspendLayout();
            this.selectVerificationRoutines.SuspendLayout();
            this.tabPageResults.SuspendLayout();
            this.SuspendLayout();
            // 
            // mainMenu
            // 
            this.mainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.verificationsToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.mainMenu.Location = new System.Drawing.Point(0, 0);
            this.mainMenu.Name = "mainMenu";
            this.mainMenu.Size = new System.Drawing.Size(680, 24);
            this.mainMenu.TabIndex = 4;
            this.mainMenu.Text = "mainMenuStrip";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.autoLoadToolStripMenuItem,
            this.toolStripSeparator,
            this.saveToolStripMenuItem,
            this.saveSuppressionToolStripMenuItem,
            this.toolStripSeparator2,
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.AccessibleName = "Open";
            this.openToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.verificationsDLLToolStripMenuItem,
            this.logFileToolStripMenuItem});
            this.openToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("openToolStripMenuItem.Image")));
            this.openToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(330, 22);
            this.openToolStripMenuItem.Text = "Open";
            // 
            // verificationsDLLToolStripMenuItem
            // 
            this.verificationsDLLToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.verificationsDLLToolStripMenuItem.Name = "verificationsDLLToolStripMenuItem";
            this.verificationsDLLToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.D)));
            this.verificationsDLLToolStripMenuItem.Size = new System.Drawing.Size(203, 22);
            this.verificationsDLLToolStripMenuItem.Text = "Verifications &DLL";
            this.verificationsDLLToolStripMenuItem.Click += new System.EventHandler(this.openToolStripButton_Click);
            // 
            // logFileToolStripMenuItem
            // 
            this.logFileToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.logFileToolStripMenuItem.Name = "logFileToolStripMenuItem";
            this.logFileToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.O)));
            this.logFileToolStripMenuItem.Size = new System.Drawing.Size(203, 22);
            this.logFileToolStripMenuItem.Text = "L&og file";
            this.logFileToolStripMenuItem.Click += new System.EventHandler(this.openLogFileStripButton_Click);
            // 
            // autoLoadToolStripMenuItem
            // 
            this.autoLoadToolStripMenuItem.AccessibleName = "Automatically load available verifications";
            this.autoLoadToolStripMenuItem.CheckOnClick = true;
            this.autoLoadToolStripMenuItem.Name = "autoLoadToolStripMenuItem";
            this.autoLoadToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.L)));
            this.autoLoadToolStripMenuItem.Size = new System.Drawing.Size(330, 22);
            this.autoLoadToolStripMenuItem.Text = "Automatically &load available verifications";
            this.autoLoadToolStripMenuItem.CheckStateChanged += new System.EventHandler(this.autoLoadToolStripMenuItem_CheckStateChanged);
            // 
            // toolStripSeparator
            // 
            this.toolStripSeparator.Name = "toolStripSeparator";
            this.toolStripSeparator.Size = new System.Drawing.Size(327, 6);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.AccessibleName = "Save Log Ctrl+S";
            this.saveToolStripMenuItem.Enabled = false;
            this.saveToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("saveToolStripMenuItem.Image")));
            this.saveToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.S)));
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(330, 22);
            this.saveToolStripMenuItem.Text = "&Save Log";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // saveSuppressionToolStripMenuItem
            // 
            this.saveSuppressionToolStripMenuItem.AccessibleName = "Save Suppression Ctrl+Shift+S";
            this.saveSuppressionToolStripMenuItem.Enabled = false;
            this.saveSuppressionToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("saveSuppressionToolStripMenuItem.Image")));
            this.saveSuppressionToolStripMenuItem.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.saveSuppressionToolStripMenuItem.Name = "saveSuppressionToolStripMenuItem";
            this.saveSuppressionToolStripMenuItem.ShortcutKeys = ((System.Windows.Forms.Keys)(((System.Windows.Forms.Keys.Control | System.Windows.Forms.Keys.Shift)
                        | System.Windows.Forms.Keys.S)));
            this.saveSuppressionToolStripMenuItem.Size = new System.Drawing.Size(330, 22);
            this.saveSuppressionToolStripMenuItem.Text = "Save Suppression";
            this.saveSuppressionToolStripMenuItem.Click += new System.EventHandler(this.saveSuppressionToolStripMenuItem_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(327, 6);
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.AccessibleName = "Exit Alt+f4";
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(330, 22);
            this.exitToolStripMenuItem.Text = "E&xit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // verificationsToolStripMenuItem
            // 
            this.verificationsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.runNowToolStripMenuItem,
            this.enableQueryStringAddinStripMenuItem,
            this.enableAllToolStripMenuItem,
            this.disableAllToolStripMenuItem,
            this.toolStripSeparator1,
            this.isIncludePassResultsInLogToolStripMenuItem});
            this.verificationsToolStripMenuItem.Name = "verificationsToolStripMenuItem";
            this.verificationsToolStripMenuItem.Size = new System.Drawing.Size(84, 20);
            this.verificationsToolStripMenuItem.Text = "Ver&ifications";
            // 
            // runNowToolStripMenuItem
            // 
            this.runNowToolStripMenuItem.Enabled = false;
            this.runNowToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("runNowToolStripMenuItem.Image")));
            this.runNowToolStripMenuItem.Name = "runNowToolStripMenuItem";
            this.runNowToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            this.runNowToolStripMenuItem.Text = "&Run Now         Ctrl+Shift+=";
            this.runNowToolStripMenuItem.Click += new System.EventHandler(this.runNowToolStripMenuItem_Click);
            // 
            // enableQueryStringAddinStripMenuItem
            // 
            this.enableQueryStringAddinStripMenuItem.Name = "enableQueryStringAddinStripMenuItem";
            this.enableQueryStringAddinStripMenuItem.Size = new System.Drawing.Size(217, 22);
            this.enableQueryStringAddinStripMenuItem.Text = "Enable suppressions across &languages";
            this.enableQueryStringAddinStripMenuItem.Click += new System.EventHandler(this.enableQueryStringAddinStripMenuItem_Click);
            // 
            // enableAllToolStripMenuItem
            // 
            this.enableAllToolStripMenuItem.Name = "enableAllToolStripMenuItem";
            this.enableAllToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            this.enableAllToolStripMenuItem.Text = "&Enable All";
            this.enableAllToolStripMenuItem.Click += new System.EventHandler(this.btnSelectAll_Click);
            // 
            // disableAllToolStripMenuItem
            // 
            this.disableAllToolStripMenuItem.Name = "disableAllToolStripMenuItem";
            this.disableAllToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            this.disableAllToolStripMenuItem.Text = "&Disable All";
            this.disableAllToolStripMenuItem.Click += new System.EventHandler(this.btnDeselectAll_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(214, 6);
            // 
            // isIncludePassResultsInLogToolStripMenuItem
            // 
            this.isIncludePassResultsInLogToolStripMenuItem.Name = "isIncludePassResultsInLogToolStripMenuItem";
            this.isIncludePassResultsInLogToolStripMenuItem.Size = new System.Drawing.Size(217, 22);
            this.isIncludePassResultsInLogToolStripMenuItem.Text = "&Include Pass Results";
            this.isIncludePassResultsInLogToolStripMenuItem.Click += new System.EventHandler(this.btnIncludePassResultsInLogRun_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.viewToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // viewToolStripMenuItem
            // 
            this.viewToolStripMenuItem.AccessibleName = "Help F1";
            this.viewToolStripMenuItem.Name = "viewToolStripMenuItem";
            this.viewToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1;
            this.viewToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.viewToolStripMenuItem.Text = "&Help";
            this.viewToolStripMenuItem.Click += new System.EventHandler(this.menu_Help_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
            this.aboutToolStripMenuItem.Text = "&About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.menu_About_Click);
            // 
            // openDLLDialog
            // 
            this.openDLLDialog.Filter = "Libraries|*.dll";
            // 
            // saveFileDialog
            // 
            this.saveFileDialog.Filter = "XML files|*.xml|Text files|*.txt";
            // 
            // backgroundWorker
            // 
            this.backgroundWorker.WorkerReportsProgress = true;
            this.backgroundWorker.WorkerSupportsCancellation = true;
            this.backgroundWorker.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker_DoWork);
            this.backgroundWorker.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            this.backgroundWorker.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            // 
            // openSuppresionFileDialog
            // 
            this.openSuppresionFileDialog.Filter = "XML files|*.xml";
            this.openSuppresionFileDialog.Multiselect = true;
            // 
            // TopToolStripPanel
            // 
            this.TopToolStripPanel.Location = new System.Drawing.Point(0, 0);
            this.TopToolStripPanel.Name = "TopToolStripPanel";
            this.TopToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal;
            this.TopToolStripPanel.RowMargin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.TopToolStripPanel.Size = new System.Drawing.Size(0, 0);
            // 
            // RightToolStripPanel
            // 
            this.RightToolStripPanel.Location = new System.Drawing.Point(0, 0);
            this.RightToolStripPanel.Name = "RightToolStripPanel";
            this.RightToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal;
            this.RightToolStripPanel.RowMargin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.RightToolStripPanel.Size = new System.Drawing.Size(0, 0);
            // 
            // LeftToolStripPanel
            // 
            this.LeftToolStripPanel.Location = new System.Drawing.Point(0, 0);
            this.LeftToolStripPanel.Name = "LeftToolStripPanel";
            this.LeftToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal;
            this.LeftToolStripPanel.RowMargin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.LeftToolStripPanel.Size = new System.Drawing.Size(0, 0);
            // 
            // ContentPanel
            // 
            this.ContentPanel.Size = new System.Drawing.Size(680, 500);
            // 
            // tabControl
            // 
            this.tabControl.Controls.Add(this.tabPageVerifications);
            this.tabControl.Controls.Add(this.tabPageResults);
            this.tabControl.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tabControl.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tabControl.Location = new System.Drawing.Point(0, 24);
            this.tabControl.Name = "tabControl";
            this.tabControl.SelectedIndex = 0;
            this.tabControl.Size = new System.Drawing.Size(680, 599);
            this.tabControl.TabIndex = 6;
            this.tabControl.SelectedIndexChanged += new System.EventHandler(this.tabControlTabChanged);
            // 
            // tabPageVerifications
            // 
            this.tabPageVerifications.Controls.Add(this.splitContainer1);
            this.tabPageVerifications.Location = new System.Drawing.Point(4, 22);
            this.tabPageVerifications.Name = "tabPageVerifications";
            this.tabPageVerifications.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageVerifications.Size = new System.Drawing.Size(672, 573);
            this.tabPageVerifications.TabIndex = 0;
            this.tabPageVerifications.Text = "Verifications";
            this.tabPageVerifications.UseVisualStyleBackColor = true;
            // 
            // splitContainer1
            // 
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(3, 3);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.groupBox1);
            this.splitContainer1.Panel1.Controls.Add(this.topLevelWindows);
            this.splitContainer1.Panel1.Controls.Add(this.textBoxCapturedWindow);
            this.splitContainer1.Panel1.Controls.Add(this.btnCaptureParent);
            this.splitContainer1.Panel1.Controls.Add(this.tbTypedHwnd);
            this.splitContainer1.Panel1.Controls.Add(this.specifyWindow);
            this.splitContainer1.Panel1.Controls.Add(this.suppressionFile);
            this.splitContainer1.Panel1.Controls.Add(this.runSelectedVerification);
            this.splitContainer1.Panel1.Margin = new System.Windows.Forms.Padding(3);
            this.splitContainer1.Panel1.Padding = new System.Windows.Forms.Padding(3);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.selectVerificationRoutines);
            this.splitContainer1.Size = new System.Drawing.Size(666, 567);
            this.splitContainer1.SplitterDistance = 388;
            this.splitContainer1.TabIndex = 11;
            this.splitContainer1.TabStop = false;
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.groupBox1.Controls.Add(this.cbPriority);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Location = new System.Drawing.Point(2, 382);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(380, 58);
            this.groupBox1.TabIndex = 15;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Show prioritized results";
            // 
            // cbPriority
            // 
            this.cbPriority.FormattingEnabled = true;
            this.cbPriority.Items.AddRange(new object[] {
            "All",
            "P1 only",
            "P1 – P2 only"});
            this.cbPriority.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cbPriority.Location = new System.Drawing.Point(64, 21);
            this.cbPriority.Name = "cbPriority";
            this.cbPriority.Size = new System.Drawing.Size(121, 21);
            this.cbPriority.TabIndex = 2;
            this.cbPriority.SelectedIndexChanged += new System.EventHandler(this.cbPriority_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(43, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Priority";
            // 
            // topLevelWindows
            // 
            this.topLevelWindows.AccessibleName = "Top level window list";
            this.topLevelWindows.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.topLevelWindows.DataBindings.Add(new System.Windows.Forms.Binding("Tag", global::AccCheckUI.Properties.Settings.Default, "SelectedFromList", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.topLevelWindows.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.topLevelWindows.FormattingEnabled = true;
            this.topLevelWindows.ItemHeight = 13;
            this.topLevelWindows.Items.AddRange(new object[] {
            "<Select a top level window>"});
            this.topLevelWindows.Location = new System.Drawing.Point(35, 46);
            this.topLevelWindows.Name = "topLevelWindows";
            this.topLevelWindows.Size = new System.Drawing.Size(306, 21);
            this.topLevelWindows.TabIndex = 3;
            this.topLevelWindows.Tag = global::AccCheckUI.Properties.Settings.Default.SelectedFromList;
            this.topLevelWindows.Enter += new System.EventHandler(this.topLevelWindows_Enter);
            this.topLevelWindows.KeyUp += new System.Windows.Forms.KeyEventHandler(this.topLevelWindows_KeyUp);
            this.topLevelWindows.KeyDown += new System.Windows.Forms.KeyEventHandler(this.topLevelWindows_KeyDown);
            this.topLevelWindows.TextChanged += new System.EventHandler(this.topLevelWindows_TextChanged);
            // 
            // textBoxCapturedWindow
            // 
            this.textBoxCapturedWindow.AccessibleName = "Captured window";
            this.textBoxCapturedWindow.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.textBoxCapturedWindow.Location = new System.Drawing.Point(35, 102);
            this.textBoxCapturedWindow.Name = "textBoxCapturedWindow";
            this.textBoxCapturedWindow.ReadOnly = true;
            this.textBoxCapturedWindow.Size = new System.Drawing.Size(306, 22);
            this.textBoxCapturedWindow.TabIndex = 6;
            this.textBoxCapturedWindow.Text = "(none yet)";
            // 
            // btnCaptureParent
            // 
            this.btnCaptureParent.Location = new System.Drawing.Point(35, 128);
            this.btnCaptureParent.Name = "btnCaptureParent";
            this.btnCaptureParent.Size = new System.Drawing.Size(160, 23);
            this.btnCaptureParent.TabIndex = 7;
            this.btnCaptureParent.Text = "Go to &parent ( Ctrl+Shift+] )";
            this.btnCaptureParent.UseVisualStyleBackColor = true;
            this.btnCaptureParent.Click += new System.EventHandler(this.btnCaptureParent_Click);
            // 
            // tbTypedHwnd
            // 
            this.tbTypedHwnd.AccessibleName = "Use this hwnd";
            this.tbTypedHwnd.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tbTypedHwnd.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::AccCheckUI.Properties.Settings.Default, "SelectedByHwnd", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.tbTypedHwnd.Location = new System.Drawing.Point(35, 190);
            this.tbTypedHwnd.Name = "tbTypedHwnd";
            this.tbTypedHwnd.Size = new System.Drawing.Size(306, 22);
            this.tbTypedHwnd.TabIndex = 9;
            this.tbTypedHwnd.Text = global::AccCheckUI.Properties.Settings.Default.SelectedByHwnd;
            this.tbTypedHwnd.TextChanged += new System.EventHandler(this.tbTypedHwnd_TextChanged);
            // 
            // specifyWindow
            // 
            this.specifyWindow.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.specifyWindow.Controls.Add(this.gbMonitoringOptions);
            this.specifyWindow.Controls.Add(this.rbMonitoringMode);
            this.specifyWindow.Controls.Add(this.rbMainWindow);
            this.specifyWindow.Controls.Add(this.rbCapturedWindow);
            this.specifyWindow.Controls.Add(this.rbUseTypedHwnd);
            this.specifyWindow.Dock = System.Windows.Forms.DockStyle.Top;
            this.specifyWindow.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.specifyWindow.Location = new System.Drawing.Point(3, 3);
            this.specifyWindow.Name = "specifyWindow";
            this.specifyWindow.Size = new System.Drawing.Size(380, 275);
            this.specifyWindow.TabIndex = 1;
            this.specifyWindow.TabStop = false;
            this.specifyWindow.Text = "Select window to verify";
            // 
            // gbMonitoringOptions
            // 
            this.gbMonitoringOptions.Controls.Add(this.rbSelected);
            this.gbMonitoringOptions.Controls.Add(this.rbOptimized);
            this.gbMonitoringOptions.Location = new System.Drawing.Point(138, 215);
            this.gbMonitoringOptions.Name = "gbMonitoringOptions";
            this.gbMonitoringOptions.Size = new System.Drawing.Size(200, 54);
            this.gbMonitoringOptions.TabIndex = 10;
            this.gbMonitoringOptions.TabStop = false;
            this.gbMonitoringOptions.Enabled = false;
            // 
            // rbSelected
            // 
            this.rbSelected.AutoSize = true;
            this.rbSelected.Location = new System.Drawing.Point(7, 29);
            this.rbSelected.Name = "rbSelected";
            this.rbSelected.Size = new System.Drawing.Size(68, 17);
            this.rbSelected.TabIndex = 1;
            this.rbSelected.TabStop = true;
            this.rbSelected.Text = "Selected";
            this.rbSelected.UseVisualStyleBackColor = true;
            this.rbSelected.Checked = true;
            this.rbSelected.CheckedChanged += new System.EventHandler(this.rbSelected_CheckedChanged);
            // 
            // rbOptimized
            // 
            this.rbOptimized.AutoSize = true;
            this.rbOptimized.Location = new System.Drawing.Point(7, 11);
            this.rbOptimized.Name = "rbOptimized";
            this.rbOptimized.Size = new System.Drawing.Size(78, 17);
            this.rbOptimized.TabIndex = 0;
            this.rbOptimized.TabStop = true;
            this.rbOptimized.Text = "Optimized";
            this.rbOptimized.UseVisualStyleBackColor = true;
            this.rbOptimized.CheckedChanged += new System.EventHandler(this.rbOptimized_CheckedChanged);
            // 
            // rbMonitoringMode
            // 
            this.rbMonitoringMode.AutoSize = true;
            this.rbMonitoringMode.Location = new System.Drawing.Point(10, 225);
            this.rbMonitoringMode.Name = "rbMonitoringMode";
            this.rbMonitoringMode.Size = new System.Drawing.Size(125, 17);
            this.rbMonitoringMode.TabIndex = 9;
            this.rbMonitoringMode.TabStop = true;
            this.rbMonitoringMode.Text = "Run in background";
            this.rbMonitoringMode.UseVisualStyleBackColor = true;
            this.rbMonitoringMode.CheckedChanged += new System.EventHandler(this.rbMonitoringMode_CheckedChanged);
            // 
            // rbMainWindow
            // 
            this.rbMainWindow.AutoSize = true;
            this.rbMainWindow.Checked = true;
            this.rbMainWindow.Cursor = System.Windows.Forms.Cursors.Default;
            this.rbMainWindow.Location = new System.Drawing.Point(10, 19);
            this.rbMainWindow.Name = "rbMainWindow";
            this.rbMainWindow.Size = new System.Drawing.Size(146, 21);
            this.rbMainWindow.TabIndex = 2;
            this.rbMainWindow.TabStop = true;
            this.rbMainWindow.Text = "Choose window from &list";
            this.rbMainWindow.UseCompatibleTextRendering = true;
            this.rbMainWindow.UseVisualStyleBackColor = true;
            this.rbMainWindow.CheckedChanged += new System.EventHandler(this.rbMainWindow_CheckedChanged);
            // 
            // rbCapturedWindow
            // 
            this.rbCapturedWindow.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.rbCapturedWindow.AutoSize = true;
            this.rbCapturedWindow.Location = new System.Drawing.Point(10, 79);
            this.rbCapturedWindow.Name = "rbCapturedWindow";
            this.rbCapturedWindow.Size = new System.Drawing.Size(348, 17);
            this.rbCapturedWindow.TabIndex = 5;
            this.rbCapturedWindow.Text = "Choose window: hover with the &mouse then press Ctrl+Shift+[";
            this.rbCapturedWindow.UseVisualStyleBackColor = true;
            this.rbCapturedWindow.Click += new System.EventHandler(this.rbCapturedWindow_Click);
            this.rbCapturedWindow.CheckedChanged += new System.EventHandler(this.rbCapturedWindow_CheckedChanged);
            // 
            // rbUseTypedHwnd
            // 
            this.rbUseTypedHwnd.AutoSize = true;
            this.rbUseTypedHwnd.Location = new System.Drawing.Point(10, 167);
            this.rbUseTypedHwnd.Name = "rbUseTypedHwnd";
            this.rbUseTypedHwnd.Size = new System.Drawing.Size(107, 17);
            this.rbUseTypedHwnd.TabIndex = 8;
            this.rbUseTypedHwnd.Text = "Use this H&WND:";
            this.rbUseTypedHwnd.UseVisualStyleBackColor = true;
            this.rbUseTypedHwnd.CheckedChanged += new System.EventHandler(this.rbUseTypedHwnd_CheckedChanged);
            // 
            // suppressionFile
            // 
            this.suppressionFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.suppressionFile.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.suppressionFile.Controls.Add(this._btnClear);
            this.suppressionFile.Controls.Add(this.label1);
            this.suppressionFile.Controls.Add(this.tbSuppresionFile);
            this.suppressionFile.Controls.Add(this.btnLoadSuppresionFile);
            this.suppressionFile.Location = new System.Drawing.Point(2, 278);
            this.suppressionFile.Name = "suppressionFile";
            this.suppressionFile.Size = new System.Drawing.Size(380, 101);
            this.suppressionFile.TabIndex = 10;
            this.suppressionFile.TabStop = false;
            this.suppressionFile.Text = "Specify a suppression file (optional)";
            // 
            // _btnClear
            // 
            this._btnClear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this._btnClear.Location = new System.Drawing.Point(182, 69);
            this._btnClear.Name = "_btnClear";
            this._btnClear.Size = new System.Drawing.Size(75, 23);
            this._btnClear.TabIndex = 12;
            this._btnClear.Text = "&Clear";
            this._btnClear.UseVisualStyleBackColor = true;
            this._btnClear.Click += new System.EventHandler(this._btnClear_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(310, 13);
            this.label1.TabIndex = 21;
            this.label1.Text = "Load suppression file to ignore specific errors or warnings.";
            // 
            // tbSuppresionFile
            // 
            this.tbSuppresionFile.AccessibleName = "Suppression Filename";
            this.tbSuppresionFile.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tbSuppresionFile.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::AccCheckUI.Properties.Settings.Default, "SuppressionFileSelected", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.tbSuppresionFile.Location = new System.Drawing.Point(19, 43);
            this.tbSuppresionFile.Name = "tbSuppresionFile";
            this.tbSuppresionFile.ReadOnly = true;
            this.tbSuppresionFile.Size = new System.Drawing.Size(322, 22);
            this.tbSuppresionFile.TabIndex = 11;
            this.tbSuppresionFile.Text = global::AccCheckUI.Properties.Settings.Default.SuppressionFileSelected;
            this.tbSuppresionFile.TextChanged += new System.EventHandler(this.tbSuppresionFile_TextChanged);
            // 
            // btnLoadSuppresionFile
            // 
            this.btnLoadSuppresionFile.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnLoadSuppresionFile.Location = new System.Drawing.Point(263, 69);
            this.btnLoadSuppresionFile.Name = "btnLoadSuppresionFile";
            this.btnLoadSuppresionFile.Size = new System.Drawing.Size(78, 23);
            this.btnLoadSuppresionFile.TabIndex = 13;
            this.btnLoadSuppresionFile.Text = "&Browse...";
            this.btnLoadSuppresionFile.UseVisualStyleBackColor = true;
            this.btnLoadSuppresionFile.Click += new System.EventHandler(this.btnLoadSuppresionFile_Click);
            // 
            // runSelectedVerification
            // 
            this.runSelectedVerification.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.runSelectedVerification.Controls.Add(this.buttonRunVerifications);
            this.runSelectedVerification.Controls.Add(this.progressBarVerifications);
            this.runSelectedVerification.Controls.Add(this.verificationStatus);
            this.runSelectedVerification.Location = new System.Drawing.Point(3, 443);
            this.runSelectedVerification.Name = "runSelectedVerification";
            this.runSelectedVerification.Size = new System.Drawing.Size(380, 120);
            this.runSelectedVerification.TabIndex = 14;
            this.runSelectedVerification.TabStop = false;
            this.runSelectedVerification.Text = "Run selected verifications";
            // 
            // buttonRunVerifications
            // 
            this.buttonRunVerifications.Enabled = false;
            this.buttonRunVerifications.Location = new System.Drawing.Point(7, 18);
            this.buttonRunVerifications.Name = "buttonRunVerifications";
            this.buttonRunVerifications.Size = new System.Drawing.Size(200, 28);
            this.buttonRunVerifications.TabIndex = 15;
            this.buttonRunVerifications.Text = "&Run Verifications";
            this.buttonRunVerifications.UseVisualStyleBackColor = true;
            this.buttonRunVerifications.Click += new System.EventHandler(this.buttonRunVerificationClick);
            // 
            // progressBarVerifications
            // 
            this.progressBarVerifications.AccessibleName = "Progress of verifications";
            this.progressBarVerifications.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.progressBarVerifications.Location = new System.Drawing.Point(7, 56);
            this.progressBarVerifications.Name = "progressBarVerifications";
            this.progressBarVerifications.Size = new System.Drawing.Size(334, 23);
            this.progressBarVerifications.TabIndex = 16;
            // 
            // verificationStatus
            // 
            this.verificationStatus.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.verificationStatus.AutoSize = true;
            this.verificationStatus.Location = new System.Drawing.Point(16, 86);
            this.verificationStatus.Name = "verificationStatus";
            this.verificationStatus.Size = new System.Drawing.Size(10, 13);
            this.verificationStatus.TabIndex = 28;
            this.verificationStatus.Text = " ";
            // 
            // selectVerificationRoutines
            // 
            this.selectVerificationRoutines.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.selectVerificationRoutines.Controls.Add(this.lvRoutines);
            this.selectVerificationRoutines.Location = new System.Drawing.Point(3, 3);
            this.selectVerificationRoutines.Name = "selectVerificationRoutines";
            this.selectVerificationRoutines.Size = new System.Drawing.Size(265, 563);
            this.selectVerificationRoutines.TabIndex = 17;
            this.selectVerificationRoutines.TabStop = false;
            this.selectVerificationRoutines.Text = "Select verification routines";
            // 
            // lvRoutines
            // 
            this.lvRoutines.AccessibleName = "Verification Routines";
            this.lvRoutines.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.lvRoutines.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lvRoutines.CheckBoxes = true;
            this.lvRoutines.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chRoutine,
            this.Priority});
            this.lvRoutines.DataBindings.Add(new System.Windows.Forms.Binding("Tag", global::AccCheckUI.Properties.Settings.Default, "VerificationsSelected", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.lvRoutines.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable;
            this.lvRoutines.HideSelection = false;
            this.lvRoutines.Location = new System.Drawing.Point(3, 16);
            this.lvRoutines.Name = "lvRoutines";
            this.lvRoutines.Size = new System.Drawing.Size(257, 541);
            this.lvRoutines.TabIndex = 18;
            this.lvRoutines.Tag = global::AccCheckUI.Properties.Settings.Default.VerificationsSelected;
            this.lvRoutines.UseCompatibleStateImageBehavior = false;
            this.lvRoutines.View = System.Windows.Forms.View.Details;
            this.lvRoutines.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.lvRoutines_ItemCheck);
            // 
            // chRoutine
            // 
            this.chRoutine.Text = "Verification Routines";
            this.chRoutine.Width = 180;
            // 
            // Priority
            // 
            this.Priority.Text = "Priority";
            // 
            // tabPageResults
            // 
            this.tabPageResults.Controls.Add(this.guiLogList);
            this.tabPageResults.Location = new System.Drawing.Point(4, 22);
            this.tabPageResults.Name = "tabPageResults";
            this.tabPageResults.Padding = new System.Windows.Forms.Padding(3);
            this.tabPageResults.Size = new System.Drawing.Size(672, 573);
            this.tabPageResults.TabIndex = 1;
            this.tabPageResults.Text = "Results";
            this.tabPageResults.UseVisualStyleBackColor = true;
            // 
            // guiLogList
            // 
            this.guiLogList.Dock = System.Windows.Forms.DockStyle.Fill;
            this.guiLogList.Location = new System.Drawing.Point(3, 3);
            this.guiLogList.Name = "guiLogList";
            this.guiLogList.Size = new System.Drawing.Size(666, 567);
            this.guiLogList.TabIndex = 0;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.ClientSize = new System.Drawing.Size(680, 623);
            this.Controls.Add(this.tabControl);
            this.Controls.Add(this.mainMenu);
            this.DataBindings.Add(new System.Windows.Forms.Binding("Location", global::AccCheckUI.Properties.Settings.Default, "FormLocation", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Location = global::AccCheckUI.Properties.Settings.Default.FormLocation;
            this.MainMenuStrip = this.mainMenu;
            this.Name = "MainForm";
            this.Text = "UI Accessibility Checker";
            this.Load += new System.EventHandler(this.MainForm_Load);
            this.Activated += new System.EventHandler(this.MainForm_GotFocus);
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.mainMenu.ResumeLayout(false);
            this.mainMenu.PerformLayout();
            this.tabControl.ResumeLayout(false);
            this.tabPageVerifications.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.PerformLayout();
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.specifyWindow.ResumeLayout(false);
            this.specifyWindow.PerformLayout();
            this.gbMonitoringOptions.ResumeLayout(false);
            this.gbMonitoringOptions.PerformLayout();
            this.suppressionFile.ResumeLayout(false);
            this.suppressionFile.PerformLayout();
            this.runSelectedVerification.ResumeLayout(false);
            this.runSelectedVerification.PerformLayout();
            this.selectVerificationRoutines.ResumeLayout(false);
            this.tabPageResults.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.MenuStrip mainMenu;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem verificationsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runNowToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem enableQueryStringAddinStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem enableAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem isIncludePassResultsInLogToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem disableAllToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openDLLDialog;
        private System.Windows.Forms.SaveFileDialog saveFileDialog;
        private System.ComponentModel.BackgroundWorker backgroundWorker;
        private System.Windows.Forms.OpenFileDialog openSuppresionFileDialog;
        private System.Windows.Forms.ToolStripMenuItem saveSuppressionToolStripMenuItem;
        private System.Windows.Forms.TabControl tabControl;
        private System.Windows.Forms.TabPage tabPageVerifications;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TextBox tbSuppresionFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox suppressionFile;
        private System.Windows.Forms.Button btnLoadSuppresionFile;
        private System.Windows.Forms.GroupBox specifyWindow;
        private System.Windows.Forms.RadioButton rbMainWindow;
        private System.Windows.Forms.Button btnCaptureParent;
        private System.Windows.Forms.TextBox textBoxCapturedWindow;
        private System.Windows.Forms.RadioButton rbUseTypedHwnd;
        private System.Windows.Forms.TextBox tbTypedHwnd;
        private System.Windows.Forms.RadioButton rbCapturedWindow;
        private System.Windows.Forms.ListView lvRoutines;
        private System.Windows.Forms.ColumnHeader chRoutine;
        private System.Windows.Forms.TabPage tabPageResults;
        private System.Windows.Forms.ToolStripPanel TopToolStripPanel;
        private System.Windows.Forms.ToolStripPanel RightToolStripPanel;
        private System.Windows.Forms.ToolStripPanel LeftToolStripPanel;
        private System.Windows.Forms.ToolStripContentPanel ContentPanel;
        private System.Windows.Forms.ComboBox topLevelWindows;
        private System.Windows.Forms.GroupBox selectVerificationRoutines;
        private System.Windows.Forms.GroupBox runSelectedVerification;
        private System.Windows.Forms.Button buttonRunVerifications;
        private System.Windows.Forms.ProgressBar progressBarVerifications;
        private System.Windows.Forms.Label verificationStatus;
        private System.Windows.Forms.Button _btnClear;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem autoLoadToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem verificationsDLLToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem logFileToolStripMenuItem;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ColumnHeader Priority;
        private GuiLogList guiLogList;
        private System.Windows.Forms.ComboBox cbPriority;
        private System.Windows.Forms.GroupBox gbMonitoringOptions;
        private System.Windows.Forms.RadioButton rbSelected;
        private System.Windows.Forms.RadioButton rbOptimized;
        private System.Windows.Forms.RadioButton rbMonitoringMode;
    }
}
