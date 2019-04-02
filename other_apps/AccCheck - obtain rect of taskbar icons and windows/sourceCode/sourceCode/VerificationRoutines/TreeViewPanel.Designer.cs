// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

namespace VerificationRoutines
{
    partial class TreeViewPanel
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.tree = new System.Windows.Forms.TreeView();
            this.splitContainer2 = new System.Windows.Forms.SplitContainer();
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.label2 = new System.Windows.Forms.Label();
            this.lblName = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.lblValue = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.lblRole = new System.Windows.Forms.TextBox();
            this.label6 = new System.Windows.Forms.Label();
            this.lblState = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.lblLocation = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.pbScreenshot = new System.Windows.Forms.PictureBox();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.splitContainer2.Panel1.SuspendLayout();
            this.splitContainer2.Panel2.SuspendLayout();
            this.splitContainer2.SuspendLayout();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbScreenshot)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Location = new System.Drawing.Point(285, 6);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(263, 251);
            this.panel1.TabIndex = 10;
            // 
            // splitContainer1
            // 
            this.splitContainer1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.tree);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.splitContainer2);
            this.splitContainer1.Size = new System.Drawing.Size(654, 341);
            this.splitContainer1.SplitterDistance = 218;
            this.splitContainer1.TabIndex = 0;
            this.splitContainer1.TabStop = false;
            // 
            // tree
            // 
            this.tree.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.tree.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tree.FullRowSelect = true;
            this.tree.HideSelection = false;
            this.tree.Location = new System.Drawing.Point(0, 0);
            this.tree.Name = "tree";
            this.tree.Size = new System.Drawing.Size(216, 339);
            this.tree.TabIndex = 10;
            this.tree.AfterSelect += new System.Windows.Forms.TreeViewEventHandler(this.tree_AfterSelect);
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
            this.splitContainer2.Panel1.Controls.Add(this.tableLayoutPanel1);
            this.splitContainer2.Panel1.Controls.Add(this.label1);
            // 
            // splitContainer2.Panel2
            // 
            this.splitContainer2.Panel2.Controls.Add(this.pbScreenshot);
            this.splitContainer2.Size = new System.Drawing.Size(432, 341);
            this.splitContainer2.SplitterDistance = 170;
            this.splitContainer2.TabIndex = 23;
            this.splitContainer2.TabStop = false;
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.tableLayoutPanel1.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle());
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.lblName, 1, 0);
            this.tableLayoutPanel1.Controls.Add(this.label3, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.lblValue, 1, 1);
            this.tableLayoutPanel1.Controls.Add(this.label5, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.lblRole, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.label6, 0, 3);
            this.tableLayoutPanel1.Controls.Add(this.lblState, 1, 3);
            this.tableLayoutPanel1.Controls.Add(this.label4, 0, 4);
            this.tableLayoutPanel1.Controls.Add(this.lblLocation, 1, 4);
            this.tableLayoutPanel1.Location = new System.Drawing.Point(9, 21);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 5;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(418, 144);
            this.tableLayoutPanel1.TabIndex = 24;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.label2.Size = new System.Drawing.Size(36, 16);
            this.label2.TabIndex = 13;
            this.label2.Text = "Name";
            // 
            // lblName
            // 
            this.lblName.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblName.Location = new System.Drawing.Point(60, 0);
            this.lblName.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblName.Name = "lblName";
            this.lblName.ReadOnly = true;
            this.lblName.Size = new System.Drawing.Size(355, 22);
            this.lblName.TabIndex = 16;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(3, 25);
            this.label3.Name = "label3";
            this.label3.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.label3.Size = new System.Drawing.Size(36, 16);
            this.label3.TabIndex = 17;
            this.label3.Text = "Value";
            // 
            // lblValue
            // 
            this.lblValue.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblValue.Location = new System.Drawing.Point(60, 25);
            this.lblValue.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblValue.Name = "lblValue";
            this.lblValue.ReadOnly = true;
            this.lblValue.Size = new System.Drawing.Size(355, 22);
            this.lblValue.TabIndex = 18;
            // 
            // label5
            // 
            this.label5.Location = new System.Drawing.Point(3, 50);
            this.label5.Name = "label5";
            this.label5.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.label5.Size = new System.Drawing.Size(30, 22);
            this.label5.TabIndex = 19;
            this.label5.Text = "Role";
            // 
            // lblRole
            // 
            this.lblRole.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblRole.Location = new System.Drawing.Point(60, 50);
            this.lblRole.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblRole.Name = "lblRole";
            this.lblRole.ReadOnly = true;
            this.lblRole.Size = new System.Drawing.Size(355, 22);
            this.lblRole.TabIndex = 20;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(3, 75);
            this.label6.Name = "label6";
            this.label6.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.label6.Size = new System.Drawing.Size(33, 16);
            this.label6.TabIndex = 21;
            this.label6.Text = "State";
            // 
            // lblState
            // 
            this.lblState.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblState.Location = new System.Drawing.Point(60, 75);
            this.lblState.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblState.Name = "lblState";
            this.lblState.ReadOnly = true;
            this.lblState.Size = new System.Drawing.Size(355, 22);
            this.lblState.TabIndex = 22;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(3, 100);
            this.label4.Name = "label4";
            this.label4.Padding = new System.Windows.Forms.Padding(0, 3, 0, 0);
            this.label4.Size = new System.Drawing.Size(51, 16);
            this.label4.TabIndex = 25;
            this.label4.Text = "Location";
            // 
            // lblLocation
            // 
            this.lblLocation.AccessibleName = "Location";
            this.lblLocation.Dock = System.Windows.Forms.DockStyle.Fill;
            this.lblLocation.Location = new System.Drawing.Point(60, 100);
            this.lblLocation.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.lblLocation.Name = "lblLocation";
            this.lblLocation.ReadOnly = true;
            this.lblLocation.Size = new System.Drawing.Size(355, 22);
            this.lblLocation.TabIndex = 24;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 5);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(70, 13);
            this.label1.TabIndex = 23;
            this.label1.Text = "Information";
            // 
            // pbScreenshot
            // 
            this.pbScreenshot.Dock = System.Windows.Forms.DockStyle.Fill;
            this.pbScreenshot.Location = new System.Drawing.Point(0, 0);
            this.pbScreenshot.Name = "pbScreenshot";
            this.pbScreenshot.Size = new System.Drawing.Size(430, 165);
            this.pbScreenshot.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pbScreenshot.TabIndex = 1;
            this.pbScreenshot.TabStop = false;
            // 
            // TreeViewPanel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.Controls.Add(this.splitContainer1);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "TreeViewPanel";
            this.Size = new System.Drawing.Size(654, 341);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            this.splitContainer1.ResumeLayout(false);
            this.splitContainer2.Panel1.ResumeLayout(false);
            this.splitContainer2.Panel1.PerformLayout();
            this.splitContainer2.Panel2.ResumeLayout(false);
            this.splitContainer2.ResumeLayout(false);
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pbScreenshot)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.TreeView tree;
        private System.Windows.Forms.SplitContainer splitContainer2;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.PictureBox pbScreenshot;
        private System.Windows.Forms.TextBox lblName;
        private System.Windows.Forms.TextBox lblValue;
        private System.Windows.Forms.TextBox lblRole;
        private System.Windows.Forms.TextBox lblState;
        private System.Windows.Forms.TextBox lblLocation;
        private System.Windows.Forms.Label label4;
    }
}
