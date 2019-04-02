// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

namespace AccCheckUI
{
    partial class GUILogger
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(GUILogger));
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.lvLog = new System.Windows.Forms.ListView();
            this.chTypeImage = new System.Windows.Forms.ColumnHeader();
            this.chTimestamp = new System.Windows.Forms.ColumnHeader();
            this.chID = new System.Windows.Forms.ColumnHeader();
            this.chText = new System.Windows.Forms.ColumnHeader();
            this.chTest = new System.Windows.Forms.ColumnHeader();
            this.contextMenuStripLogEntries = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.suppressToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.copyToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.visualizeStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.imageListLog = new System.Windows.Forms.ImageList(this.components);
            this.splitContainer3 = new System.Windows.Forms.SplitContainer();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label4 = new System.Windows.Forms.Label();
            this.lblID = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lblRoutine = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.lblText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.lblAccName = new System.Windows.Forms.TextBox();
            this.label9 = new System.Windows.Forms.Label();
            this.lblValue = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.lblRole = new System.Windows.Forms.TextBox();
            this.label10 = new System.Windows.Forms.Label();
            this.lblState = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lblWindowClass = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.lblBoundingRectangle = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.tbStack = new System.Windows.Forms.RichTextBox();
            this.btnVisualize = new System.Windows.Forms.Button();
            this.pbScreenshot = new System.Windows.Forms.PictureBox();
            this.panel2 = new System.Windows.Forms.Panel();
            this.btnSuppressions = new System.Windows.Forms.CheckBox();
            this.btnMessages = new System.Windows.Forms.CheckBox();
            this.btnWarnings = new System.Windows.Forms.CheckBox();
            this.btnErrors = new System.Windows.Forms.CheckBox();
            this.BottomToolStripPanel = new System.Windows.Forms.ToolStripPanel();
            this.TopToolStripPanel = new System.Windows.Forms.ToolStripPanel();
            this.RightToolStripPanel = new System.Windows.Forms.ToolStripPanel();
            this.LeftToolStripPanel = new System.Windows.Forms.ToolStripPanel();
            this.ContentPanel = new System.Windows.Forms.ToolStripContentPanel();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.chPriority = new System.Windows.Forms.ColumnHeader();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.contextMenuStripLogEntries.SuspendLayout();
            this.splitContainer3.Panel1.SuspendLayout();
            this.splitContainer3.Panel2.SuspendLayout();
            this.splitContainer3.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbScreenshot)).BeginInit();
            this.panel2.SuspendLayout();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer2
            // 
            this.splitContainer2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            this.splitContainer2.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.lvLog);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.splitContainer3);
            this.splitContainer2.Size = new System.Drawing.Size(685, 511);
            this.splitContainer2.SplitterDistance = 189;
            this.splitContainer2.TabIndex = 2;
            this.splitContainer2.TabStop = false;
            // 
            // lvLog
            // 
            this.lvLog.AccessibleName = "Log entries";
            this.lvLog.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.lvLog.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chTypeImage,
            this.chTimestamp,
            this.chID,
            this.chText,
            this.chPriority,
            this.chTest});
            this.lvLog.ContextMenuStrip = this.contextMenuStripLogEntries;
            this.lvLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvLog.FullRowSelect = true;
            this.lvLog.HideSelection = false;
            this.lvLog.LargeImageList = this.imageListLog;
            this.lvLog.Location = new System.Drawing.Point(0, 0);
            this.lvLog.Name = "lvLog";
            this.lvLog.Size = new System.Drawing.Size(683, 187);
            this.lvLog.SmallImageList = this.imageListLog;
            this.lvLog.TabIndex = 4;
            this.lvLog.UseCompatibleStateImageBehavior = false;
            this.lvLog.View = System.Windows.Forms.View.Details;
            this.lvLog.SelectedIndexChanged += new System.EventHandler(this.lvLog_SelectedIndexChanged);
            this.lvLog.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lvLog_ColumnClick);
            this.lvLog.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lvLog_KeyDown);
            // 
            // chTypeImage
            // 
            this.chTypeImage.Text = "";
            this.chTypeImage.Width = 24;
            // 
            // chTimestamp
            // 
            this.chTimestamp.Text = "Time";
            // 
            // chID
            // 
            this.chID.Text = "ID";
            this.chID.Width = 90;
            // 
            // chText
            // 
            this.chText.Text = "Text";
            this.chText.Width = 350;
            // 
            // chTest
            // 
            this.chTest.Text = "Verification";
            this.chTest.Width = 90;
            // 
            // contextMenuStripLogEntries
            // 
            this.contextMenuStripLogEntries.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.suppressToolStripMenuItem,
            this.copyToolStripMenuItem,
            this.visualizeStripMenuItem,
            this.helpToolStripMenuItem});
            this.contextMenuStripLogEntries.Name = "contextMenuStripLogEntries";
            this.contextMenuStripLogEntries.Size = new System.Drawing.Size(170, 70);
            this.contextMenuStripLogEntries.Opening += new System.ComponentModel.CancelEventHandler(this.contextMenuStripLogEntries_Opening);
            // 
            // suppressToolStripMenuItem
            // 
            this.suppressToolStripMenuItem.Name = "suppressToolStripMenuItem";
            this.suppressToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.suppressToolStripMenuItem.Text = "Suppress";
            this.suppressToolStripMenuItem.Click += new System.EventHandler(this.suppressToolStripMenuItem_Click);
            // 
            // copyToolStripMenuItem
            // 
            this.copyToolStripMenuItem.Name = "copyToolStripMenuItem";
            this.copyToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.copyToolStripMenuItem.Text = "Copy to clipboard";
            this.copyToolStripMenuItem.Click += new System.EventHandler(this.copyToolStripMenuItem_Click);
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.helpToolStripMenuItem.Text = "Help";
            this.helpToolStripMenuItem.Click += new System.EventHandler(this.helpToolStripMenuItem_Click);
            // 
            // visualizeStripMenuItem
            // 
            this.visualizeStripMenuItem.Name = "visualizeStripMenuItem";
            this.visualizeStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.visualizeStripMenuItem.Text = "Visualize";
            this.visualizeStripMenuItem.Click += new System.EventHandler(this.btnVisualize_Click);
            // 
            // imageListLog
            // 
            this.imageListLog.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageListLog.ImageStream")));
            this.imageListLog.TransparentColor = System.Drawing.Color.Transparent;
            this.imageListLog.Images.SetKeyName(0, "Error.png");
            this.imageListLog.Images.SetKeyName(1, "Warning.png");
            this.imageListLog.Images.SetKeyName(2, "Information.png");
            this.imageListLog.Images.SetKeyName(3, "delete.png");
            // 
            // splitContainer3
            // 
            this.splitContainer3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitContainer3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer3.Location = new System.Drawing.Point(0, 0);
            this.splitContainer3.Name = "splitContainer3";
            // 
            // splitContainer3.Panel1
            // 
            this.splitContainer3.Panel1.Controls.Add(this.tableLayoutPanel1);
            // 
            // splitContainer3.Panel2
            // 
            this.splitContainer3.Panel2.Controls.Add(this.pbScreenshot);
            this.splitContainer3.Size = new System.Drawing.Size(685, 318);
            this.splitContainer3.SplitterDistance = 369;
            this.splitContainer3.TabIndex = 0;
            this.splitContainer3.TabStop = false;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoScroll = true;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 26.75676F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 73.24324F));
            this.tableLayoutPanel1.Controls.Add(this.label4, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.lblID, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label5, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.lblRoutine, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label6, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.lblText, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.lblAccName, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.label9, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.lblValue, 1, 4);
            this.tableLayoutPanel1.Controls.Add(this.label7, 0, 5);
            this.tableLayoutPanel1.Controls.Add(this.lblRole, 1, 5);
            this.tableLayoutPanel1.Controls.Add(this.label10, 0, 6);
            this.tableLayoutPanel1.Controls.Add(this.lblState, 1, 6);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 8);
            this.tableLayoutPanel1.Controls.Add(this.lblWindowClass, 1, 8);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 7);
            this.tableLayoutPanel1.Controls.Add(this.lblBoundingRectangle, 1, 7);
            this.tableLayoutPanel1.Controls.Add(this.label8, 0, 9);
            this.tableLayoutPanel1.Controls.Add(this.tbStack, 1, 9);
            this.tableLayoutPanel1.Controls.Add(this.btnVisualize, 1, 10);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 11;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(367, 316);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(18, 13);
            this.label4.TabIndex = 0;
            this.label4.Text = "ID";
            this.label4.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lblID
            // 
            this.lblID.AccessibleName = "ID";
            this.lblID.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblID.Location = new System.Drawing.Point(101, 0);
            this.lblID.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblID.Name = "lblID";
            this.lblID.ReadOnly = true;
            this.lblID.Size = new System.Drawing.Size(263, 22);
            this.lblID.TabIndex = 5;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(3, 25);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(66, 13);
            this.label5.TabIndex = 2;
            this.label5.Text = "Verification";
            // 
            // lblRoutine
            // 
            this.lblRoutine.AccessibleName = "Verification";
            this.lblRoutine.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblRoutine.Location = new System.Drawing.Point(101, 25);
            this.lblRoutine.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblRoutine.Name = "lblRoutine";
            this.lblRoutine.ReadOnly = true;
            this.lblRoutine.Size = new System.Drawing.Size(263, 22);
            this.lblRoutine.TabIndex = 6;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(3, 50);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(27, 13);
            this.label6.TabIndex = 11;
            this.label6.Text = "Text";
            // 
            // lblText
            // 
            this.lblText.AccessibleName = "Text";
            this.lblText.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblText.Location = new System.Drawing.Point(101, 50);
            this.lblText.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblText.Multiline = true;
            this.lblText.Name = "lblText";
            this.lblText.ReadOnly = true;
            this.lblText.Size = new System.Drawing.Size(263, 38);
            this.lblText.TabIndex = 7;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 91);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(36, 13);
            this.label1.TabIndex = 13;
            this.label1.Text = "Name";
            // 
            // lblAccName
            // 
            this.lblAccName.AccessibleName = "Name";
            this.lblAccName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblAccName.Location = new System.Drawing.Point(101, 91);
            this.lblAccName.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblAccName.Name = "lblAccName";
            this.lblAccName.ReadOnly = true;
            this.lblAccName.Size = new System.Drawing.Size(263, 22);
            this.lblAccName.TabIndex = 8;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(3, 116);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(36, 13);
            this.label9.TabIndex = 22;
            this.label9.Text = "Value";
            // 
            // lblValue
            // 
            this.lblValue.AccessibleName = "Value";
            this.lblValue.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblValue.Location = new System.Drawing.Point(101, 119);
            this.lblValue.Name = "lblValue";
            this.lblValue.ReadOnly = true;
            this.lblValue.Size = new System.Drawing.Size(263, 22);
            this.lblValue.TabIndex = 9;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(3, 144);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(30, 13);
            this.label7.TabIndex = 21;
            this.label7.Text = "Role";
            // 
            // lblRole
            // 
            this.lblRole.AccessibleName = "Role";
            this.lblRole.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblRole.Location = new System.Drawing.Point(101, 147);
            this.lblRole.Name = "lblRole";
            this.lblRole.ReadOnly = true;
            this.lblRole.Size = new System.Drawing.Size(263, 22);
            this.lblRole.TabIndex = 10;
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(3, 172);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(33, 13);
            this.label10.TabIndex = 23;
            this.label10.Text = "State";
            // 
            // lblState
            // 
            this.lblState.AccessibleName = "State";
            this.lblState.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblState.Location = new System.Drawing.Point(101, 175);
            this.lblState.Name = "lblState";
            this.lblState.ReadOnly = true;
            this.lblState.Size = new System.Drawing.Size(263, 22);
            this.lblState.TabIndex = 11;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 225);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(78, 13);
            this.label3.TabIndex = 18;
            this.label3.Text = "Window class";
            // 
            // lblWindowClass
            // 
            this.lblWindowClass.AccessibleName = "Window class";
            this.lblWindowClass.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblWindowClass.Location = new System.Drawing.Point(101, 225);
            this.lblWindowClass.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblWindowClass.Name = "lblWindowClass";
            this.lblWindowClass.ReadOnly = true;
            this.lblWindowClass.Size = new System.Drawing.Size(263, 22);
            this.lblWindowClass.TabIndex = 13;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 200);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(58, 13);
            this.label2.TabIndex = 15;
            this.label2.Text = "Rectangle";
            // 
            // lblBoundingRectangle
            // 
            this.lblBoundingRectangle.AccessibleName = "Rectangle";
            this.lblBoundingRectangle.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblBoundingRectangle.Location = new System.Drawing.Point(101, 200);
            this.lblBoundingRectangle.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblBoundingRectangle.Name = "lblBoundingRectangle";
            this.lblBoundingRectangle.ReadOnly = true;
            this.lblBoundingRectangle.Size = new System.Drawing.Size(263, 22);
            this.lblBoundingRectangle.TabIndex = 12;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(3, 250);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(59, 13);
            this.label8.TabIndex = 8;
            this.label8.Text = "Extra info.";
            // 
            // tbStack
            // 
            this.tbStack.AccessibleName = "Extra info.";
            this.tbStack.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tbStack.Location = new System.Drawing.Point(101, 253);
            this.tbStack.Name = "tbStack";
            this.tbStack.ReadOnly = true;
            this.tbStack.Size = new System.Drawing.Size(263, 112);
            this.tbStack.TabIndex = 14;
            this.tbStack.Text = "";
            // 
            // btnVisualize
            // 
            this.btnVisualize.Enabled = false;
            this.btnVisualize.Location = new System.Drawing.Point(101, 371);
            this.btnVisualize.Name = "btnVisualize";
            this.btnVisualize.Size = new System.Drawing.Size(75, 23);
            this.btnVisualize.TabIndex = 17;
            this.btnVisualize.Text = "&Visualize";
            this.btnVisualize.UseVisualStyleBackColor = true;
            this.btnVisualize.Click += new System.EventHandler(this.btnVisualize_Click);
            // 
            // pbScreenshot
            // 
            this.pbScreenshot.AccessibleName = "Screen Shot";
            this.pbScreenshot.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pbScreenshot.Location = new System.Drawing.Point(0, 0);
            this.pbScreenshot.Name = "pbScreenshot";
            this.pbScreenshot.Size = new System.Drawing.Size(310, 316);
            this.pbScreenshot.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbScreenshot.TabIndex = 0;
            this.pbScreenshot.TabStop = false;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.btnSuppressions);
            this.panel2.Controls.Add(this.btnMessages);
            this.panel2.Controls.Add(this.btnWarnings);
            this.panel2.Controls.Add(this.btnErrors);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(685, 35);
            this.panel2.TabIndex = 3;
            // 
            // btnSuppressions
            // 
            this.btnSuppressions.AutoSize = true;
            this.btnSuppressions.Location = new System.Drawing.Point(413, 15);
            this.btnSuppressions.Name = "btnSuppressions";
            this.btnSuppressions.Size = new System.Drawing.Size(104, 17);
            this.btnSuppressions.TabIndex = 3;
            this.btnSuppressions.Text = "0 Suppressions";
            this.btnSuppressions.UseVisualStyleBackColor = true;
            this.btnSuppressions.Click += new System.EventHandler(this.btnErrors_Click);
            // 
            // btnMessages
            // 
            this.btnMessages.AutoSize = true;
            this.btnMessages.Location = new System.Drawing.Point(278, 15);
            this.btnMessages.Name = "btnMessages";
            this.btnMessages.Size = new System.Drawing.Size(85, 17);
            this.btnMessages.TabIndex = 2;
            this.btnMessages.Text = "0 Messages";
            this.btnMessages.UseVisualStyleBackColor = true;
            this.btnMessages.Click += new System.EventHandler(this.btnErrors_Click);
            // 
            // btnWarnings
            // 
            this.btnWarnings.AutoSize = true;
            this.btnWarnings.Checked = true;
            this.btnWarnings.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btnWarnings.Location = new System.Drawing.Point(143, 15);
            this.btnWarnings.Name = "btnWarnings";
            this.btnWarnings.Size = new System.Drawing.Size(85, 17);
            this.btnWarnings.TabIndex = 1;
            this.btnWarnings.Text = "0 Warnings";
            this.btnWarnings.UseVisualStyleBackColor = true;
            this.btnWarnings.Click += new System.EventHandler(this.btnErrors_Click);
            // 
            // btnErrors
            // 
            this.btnErrors.AutoSize = true;
            this.btnErrors.Checked = true;
            this.btnErrors.CheckState = System.Windows.Forms.CheckState.Checked;
            this.btnErrors.Location = new System.Drawing.Point(6, 15);
            this.btnErrors.Name = "btnErrors";
            this.btnErrors.Size = new System.Drawing.Size(65, 17);
            this.btnErrors.TabIndex = 0;
            this.btnErrors.Text = "0 Errors";
            this.btnErrors.UseVisualStyleBackColor = true;
            this.btnErrors.Click += new System.EventHandler(this.btnErrors_Click);
            // 
            // BottomToolStripPanel
            // 
            this.BottomToolStripPanel.Location = new System.Drawing.Point(0, 0);
            this.BottomToolStripPanel.Name = "BottomToolStripPanel";
            this.BottomToolStripPanel.Orientation = System.Windows.Forms.Orientation.Horizontal;
            this.BottomToolStripPanel.RowMargin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.BottomToolStripPanel.Size = new System.Drawing.Size(0, 0);
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
            this.ContentPanel.Size = new System.Drawing.Size(685, 550);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.IsSplitterFixed = true;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.panel2);
            this.splitContainer1.Panel1MinSize = 34;
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(685, 550);
            this.splitContainer1.SplitterDistance = 35;
            this.splitContainer1.TabIndex = 3;
            this.splitContainer1.TabStop = false;
            // 
            // chPriority
            // 
            this.chPriority.Text = "Priority";
            // 
            // GUILogger
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.Controls.Add(this.splitContainer1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "GUILogger";
            this.Size = new System.Drawing.Size(685, 550);
            this.Load += new System.EventHandler(this.GUILogger_Load);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.ResumeLayout(false);
            this.contextMenuStripLogEntries.ResumeLayout(false);
            this.splitContainer3.Panel1.ResumeLayout(false);
            this.splitContainer3.Panel2.ResumeLayout(false);
            this.splitContainer3.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbScreenshot)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.ListView lvLog;
        private System.Windows.Forms.ColumnHeader chTypeImage;
        private System.Windows.Forms.ColumnHeader chID;
        private System.Windows.Forms.ColumnHeader chText;
        private System.Windows.Forms.ColumnHeader chTest;
        private System.Windows.Forms.SplitContainer splitContainer3;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.TextBox lblID;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox lblRoutine;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.RichTextBox tbStack;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox lblText;
        private System.Windows.Forms.PictureBox pbScreenshot;
        private System.Windows.Forms.ImageList imageListLog;
        private System.Windows.Forms.ColumnHeader chTimestamp;
        private System.Windows.Forms.TextBox lblAccName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox lblBoundingRectangle;
        private System.Windows.Forms.Button btnVisualize;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox lblWindowClass;
        private System.Windows.Forms.ContextMenuStrip contextMenuStripLogEntries;
        private System.Windows.Forms.ToolStripMenuItem suppressToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripPanel BottomToolStripPanel;
        private System.Windows.Forms.ToolStripPanel TopToolStripPanel;
        private System.Windows.Forms.ToolStripPanel RightToolStripPanel;
        private System.Windows.Forms.ToolStripPanel LeftToolStripPanel;
        private System.Windows.Forms.ToolStripContentPanel ContentPanel;
        private System.Windows.Forms.CheckBox btnErrors;
        private System.Windows.Forms.CheckBox btnMessages;
        private System.Windows.Forms.CheckBox btnWarnings;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.TextBox lblValue;
        private System.Windows.Forms.TextBox lblRole;
        private System.Windows.Forms.TextBox lblState;
        private System.Windows.Forms.CheckBox btnSuppressions;
        private System.Windows.Forms.ToolStripMenuItem copyToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem visualizeStripMenuItem;
        private System.Windows.Forms.ColumnHeader chPriority;
    }
}
