// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

namespace VerificationRoutines
{
    partial class SimpleScreenReader
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
            this._screenReaderOutput = new System.Windows.Forms.ListBox();
            this._visualize = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // _screenReaderOutput
            // 
            this._screenReaderOutput.AccessibleName = "Elements as a screen reader may read the UI to a user.";
            this._screenReaderOutput.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
                        | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this._screenReaderOutput.FormattingEnabled = true;
            this._screenReaderOutput.Location = new System.Drawing.Point(0, 69);
            this._screenReaderOutput.Name = "_screenReaderOutput";
            this._screenReaderOutput.Size = new System.Drawing.Size(552, 368);
            this._screenReaderOutput.TabIndex = 0;
            // 
            // _visualize
            // 
            this._visualize.Location = new System.Drawing.Point(3, 14);
            this._visualize.MinimumSize = new System.Drawing.Size(0, 23);
            this._visualize.Name = "_visualize";
            this._visualize.Size = new System.Drawing.Size(75, 23);
            this._visualize.TabIndex = 1;
            this._visualize.Text = "&Visualize";
            this._visualize.UseVisualStyleBackColor = true;
            this._visualize.Click += new System.EventHandler(this.visualize_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(3, 53);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(346, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "List displays elements as a screen reader may read the UI to a user.";
            // 
            // SimpleScreenReader
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(96F, 96F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Dpi;
            this.AutoSize = true;
            this.Controls.Add(this.label1);
            this.Controls.Add(this._visualize);
            this.Controls.Add(this._screenReaderOutput);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Name = "SimpleScreenReader";
            this.Size = new System.Drawing.Size(561, 467);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox _screenReaderOutput;
        private System.Windows.Forms.Button _visualize;
        private System.Windows.Forms.Label label1;

    }
}
