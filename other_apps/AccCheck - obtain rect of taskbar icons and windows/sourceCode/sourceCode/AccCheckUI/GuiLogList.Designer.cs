// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

namespace AccCheckUI
{
    partial class GuiLogList
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
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.lvLogs = new System.Windows.Forms.ListView();
            this.chName = new System.Windows.Forms.ColumnHeader();
            this.chProcessName = new System.Windows.Forms.ColumnHeader();
            this.chWindowClass = new System.Windows.Forms.ColumnHeader();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.saveToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.removeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.guiLogger = new AccCheckUI.GUILogger();
            this.chStartTime = new System.Windows.Forms.ColumnHeader();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.contextMenuStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.lvLogs);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.guiLogger);
            this.splitContainer1.Size = new System.Drawing.Size(700, 600);
            this.splitContainer1.SplitterDistance = 120;
            this.splitContainer1.TabIndex = 0;
            // 
            // lvLogs
            // 
            this.lvLogs.AccessibleName = "List of active logs";
            this.lvLogs.AllowDrop = true;
            this.lvLogs.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chStartTime,
            this.chName,
            this.chProcessName,
            this.chWindowClass});
            this.lvLogs.ContextMenuStrip = this.contextMenuStrip1;
            this.lvLogs.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvLogs.FullRowSelect = true;
            this.lvLogs.HideSelection = false;
            this.lvLogs.Location = new System.Drawing.Point(0, 0);
            this.lvLogs.Name = "lvLogs";
            this.lvLogs.Size = new System.Drawing.Size(700, 120);
            this.lvLogs.TabIndex = 0;
            this.lvLogs.UseCompatibleStateImageBehavior = false;
            this.lvLogs.View = System.Windows.Forms.View.Details;
            this.lvLogs.SelectedIndexChanged += new System.EventHandler(this.lvLogs_SelectedIndexChanged);
            this.lvLogs.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lvLogs_ColumnClick);
            this.lvLogs.KeyDown += new System.Windows.Forms.KeyEventHandler(this.lvLogs_KeyDown);
            // 
            // chName
            // 
            this.chName.Text = "Name";
            this.chName.Width = 350;
            // 
            // chProcessName
            // 
            this.chProcessName.Text = "Process Name";
            this.chProcessName.Width = 100;
            // 
            // chWindowClass
            // 
            this.chWindowClass.Text = "Window Class";
            this.chWindowClass.Width = 100;
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.saveToolStripMenuItem,
            this.removeToolStripMenuItem});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(118, 48);
            // 
            // saveToolStripMenuItem
            // 
            this.saveToolStripMenuItem.Name = "saveToolStripMenuItem";
            this.saveToolStripMenuItem.Size = new System.Drawing.Size(117, 22);
            this.saveToolStripMenuItem.Text = "Save";
            this.saveToolStripMenuItem.Click += new System.EventHandler(this.saveToolStripMenuItem_Click);
            // 
            // removeToolStripMenuItem
            // 
            this.removeToolStripMenuItem.Name = "removeToolStripMenuItem";
            this.removeToolStripMenuItem.Size = new System.Drawing.Size(117, 22);
            this.removeToolStripMenuItem.Text = "Remove";
            this.removeToolStripMenuItem.Click += new System.EventHandler(this.removeToolStripMenuItem_Click);
            // 
            // guiLogger
            // 
            this.guiLogger.AccLogger = null;
            this.guiLogger.AutoSize = true;
            this.guiLogger.Dock = System.Windows.Forms.DockStyle.Fill;
            this.guiLogger.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.guiLogger.GraphicsHelper = null;
            this.guiLogger.Location = new System.Drawing.Point(0, 0);
            this.guiLogger.LogLevel = AccCheck.Logging.EventLevel.Trace;
            this.guiLogger.Name = "guiLogger";
            this.guiLogger.Size = new System.Drawing.Size(700, 476);
            this.guiLogger.TabIndex = 0;
            // 
            // chStartTime
            // 
            this.chStartTime.Text = "Start Time";
            this.chStartTime.Width = 70;
            // 
            // GuiLogList
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.splitContainer1);
            this.Name = "GuiLogList";
            this.Size = new System.Drawing.Size(700, 600);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.Panel2.PerformLayout();
            this.splitContainer1.ResumeLayout(false);
            this.contextMenuStrip1.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ListView lvLogs;
        private System.Windows.Forms.ColumnHeader chName;
        private System.Windows.Forms.ColumnHeader chProcessName;
        private System.Windows.Forms.ColumnHeader chWindowClass;
        private GUILogger guiLogger;
        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem saveToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem removeToolStripMenuItem;
        private System.Windows.Forms.ColumnHeader chStartTime;
    }
}
