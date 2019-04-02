// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

namespace VerificationRoutines
{
    partial class EventViewPanel
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
                if (_events != null)
                {
                    _events.Dispose();
                }
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
            this.timer = new System.Windows.Forms.Timer(this.components);
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.lvEvents = new System.Windows.Forms.ListView();
            this.chTimestamp = new System.Windows.Forms.ColumnHeader();
            this.chEvent = new System.Windows.Forms.ColumnHeader();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.btnVisualize = new System.Windows.Forms.Button();
            this.pbScreenshot = new System.Windows.Forms.PictureBox();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbScreenshot)).BeginInit();
            this.SuspendLayout();
            // 
            // timer
            // 
            this.timer.Enabled = true;
            this.timer.Interval = 250;
            this.timer.Tick += new System.EventHandler(this.timer_Tick);
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
            this.splitContainer1.Panel1.Controls.Add(this.lvEvents);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(582, 395);
            this.splitContainer1.SplitterDistance = 197;
            this.splitContainer1.TabIndex = 4;
            // 
            // lvEvents
            // 
            this.lvEvents.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.chTimestamp,
            this.chEvent});
            this.lvEvents.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lvEvents.FullRowSelect = true;
            this.lvEvents.Location = new System.Drawing.Point(0, 0);
            this.lvEvents.MultiSelect = false;
            this.lvEvents.Name = "lvEvents";
            this.lvEvents.Size = new System.Drawing.Size(582, 197);
            this.lvEvents.TabIndex = 4;
            this.lvEvents.UseCompatibleStateImageBehavior = false;
            this.lvEvents.View = System.Windows.Forms.View.Details;
            this.lvEvents.SelectedIndexChanged += new System.EventHandler(this.lvEvents_SelectedIndexChanged);
            // 
            // chTimestamp
            // 
            this.chTimestamp.Text = "Timestamp";
            this.chTimestamp.Width = 100;
            // 
            // chEvent
            // 
            this.chEvent.Text = "Event";
            this.chEvent.Width = 476;
            // 
            // splitContainer2
            // 
            this.splitContainer2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer2.Location = new System.Drawing.Point(0, 0);
            this.splitContainer2.Name = "splitContainer2";
            // 
            // splitContainer2.Panel1
            // 
            this.splitContainer2.Panel1.Controls.Add(this.btnVisualize);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.pbScreenshot);
            this.splitContainer2.Size = new System.Drawing.Size(582, 194);
            this.splitContainer2.SplitterDistance = 297;
            this.splitContainer2.TabIndex = 0;
            // 
            // btnVisualize
            // 
            this.btnVisualize.Location = new System.Drawing.Point(3, 3);
            this.btnVisualize.Name = "btnVisualize";
            this.btnVisualize.Size = new System.Drawing.Size(75, 23);
            this.btnVisualize.TabIndex = 0;
            this.btnVisualize.Text = "Visualize";
            this.btnVisualize.UseVisualStyleBackColor = true;
            this.btnVisualize.Click += new System.EventHandler(this.btnVisualize_Click);
            // 
            // pbScreenshot
            // 
            this.pbScreenshot.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pbScreenshot.Location = new System.Drawing.Point(0, 0);
            this.pbScreenshot.Name = "pbScreenshot";
            this.pbScreenshot.Size = new System.Drawing.Size(281, 194);
            this.pbScreenshot.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbScreenshot.TabIndex = 0;
            this.pbScreenshot.TabStop = false;
            // 
            // EventViewPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.Controls.Add(this.splitContainer1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "EventViewPanel";
            this.Size = new System.Drawing.Size(582, 395);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.pbScreenshot)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Timer timer;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.ListView lvEvents;
        private System.Windows.Forms.ColumnHeader chTimestamp;
        private System.Windows.Forms.ColumnHeader chEvent;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.PictureBox pbScreenshot;
        private System.Windows.Forms.Button btnVisualize;
    }
}
