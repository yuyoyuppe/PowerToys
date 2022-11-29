namespace MouseWithoutBorders
{
    partial class FrmScreen
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmScreen));
            this.MainMenu = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.mnuExit = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.mnuAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuHelp = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuGenDumpFile = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.mnuGetScreenCapture = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuGetFromAll = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuSendScreenCapture = new System.Windows.Forms.ToolStripMenuItem();
            this.allToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuSend2Myself = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.mnuWindowsPhone = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuWindowsPhoneEnable = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuWindowsPhoneDownload = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuWindowsPhoneInformation = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuReinstallKeyboardAndMouseHook = new System.Windows.Forms.ToolStripMenuItem();
            this.mnuMachineMatrix = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.MnuALLPC = new System.Windows.Forms.ToolStripMenuItem();
            this.dUCTDOToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tRUONG2DToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.NotifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
            this.imgListIcon = new System.Windows.Forms.ImageList(this.components);
            this.picLogonLogo = new System.Windows.Forms.PictureBox();
            this.MainMenu.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.picLogonLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // MainMenu
            // 
            this.MainMenu.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuExit,
            this.toolStripSeparator2,
            this.mnuAbout,
            this.mnuHelp,
            this.mnuGenDumpFile,
            this.toolStripSeparator1,
            this.mnuGetScreenCapture,
            this.mnuSendScreenCapture,
            this.toolStripMenuItem1,
            this.mnuWindowsPhone,
            this.mnuReinstallKeyboardAndMouseHook,
            this.mnuMachineMatrix,
            this.toolStripMenuItem2,
            this.MnuALLPC,
            this.dUCTDOToolStripMenuItem,
            this.tRUONG2DToolStripMenuItem});
            this.MainMenu.Name = "MainMenu";
            this.MainMenu.Size = new System.Drawing.Size(197, 292);
            this.MainMenu.Opening += new System.ComponentModel.CancelEventHandler(this.MainMenu_Opening);
            this.MainMenu.MouseLeave += new System.EventHandler(this.MainMenu_MouseLeave);
            // 
            // mnuExit
            // 
            this.mnuExit.ForeColor = System.Drawing.Color.Black;
            this.mnuExit.Name = "mnuExit";
            this.mnuExit.Size = new System.Drawing.Size(196, 22);
            this.mnuExit.Text = "&Exit";
            this.mnuExit.Click += new System.EventHandler(this.MnuExit_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(193, 6);
            // 
            // mnuAbout
            // 
            this.mnuAbout.ForeColor = System.Drawing.Color.Black;
            this.mnuAbout.Name = "mnuAbout";
            this.mnuAbout.Size = new System.Drawing.Size(196, 22);
            this.mnuAbout.Text = "A&bout";
            this.mnuAbout.Click += new System.EventHandler(this.MnuAbout_Click);
            // 
            // mnuHelp
            // 
            this.mnuHelp.Name = "mnuHelp";
            this.mnuHelp.Size = new System.Drawing.Size(196, 22);
            this.mnuHelp.Text = "&Help && Questions";
            this.mnuHelp.Click += new System.EventHandler(this.MnuHelp_Click);
            // 
            // mnuGenDumpFile
            // 
            this.mnuGenDumpFile.Name = "mnuGenDumpFile";
            this.mnuGenDumpFile.Size = new System.Drawing.Size(196, 22);
            this.mnuGenDumpFile.Text = "&Generate log";
            this.mnuGenDumpFile.ToolTipText = "Create logfile for triage, logfile will be generated under program directory.";
            this.mnuGenDumpFile.Visible = false;
            this.mnuGenDumpFile.Click += new System.EventHandler(this.MnuGenDumpFile_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(193, 6);
            // 
            // mnuGetScreenCapture
            // 
            this.mnuGetScreenCapture.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.mnuGetFromAll});
            this.mnuGetScreenCapture.ForeColor = System.Drawing.Color.Black;
            this.mnuGetScreenCapture.Name = "mnuGetScreenCapture";
            this.mnuGetScreenCapture.Size = new System.Drawing.Size(196, 22);
            this.mnuGetScreenCapture.Text = "&Get Screen Capture from";
            // 
            // mnuGetFromAll
            // 
            this.mnuGetFromAll.Enabled = false;
            this.mnuGetFromAll.Name = "mnuGetFromAll";
            this.mnuGetFromAll.Size = new System.Drawing.Size(85, 22);
            this.mnuGetFromAll.Text = "All";
            this.mnuGetFromAll.Click += new System.EventHandler(this.MnuGetScreenCaptureClick);
            // 
            // mnuSendScreenCapture
            // 
            this.mnuSendScreenCapture.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.allToolStripMenuItem,
            this.mnuSend2Myself});
            this.mnuSendScreenCapture.ForeColor = System.Drawing.Color.Black;
            this.mnuSendScreenCapture.Name = "mnuSendScreenCapture";
            this.mnuSendScreenCapture.Size = new System.Drawing.Size(196, 22);
            this.mnuSendScreenCapture.Text = "Send Screen &Capture to";
            // 
            // allToolStripMenuItem
            // 
            this.allToolStripMenuItem.Name = "allToolStripMenuItem";
            this.allToolStripMenuItem.Size = new System.Drawing.Size(105, 22);
            this.allToolStripMenuItem.Text = "All";
            this.allToolStripMenuItem.Click += new System.EventHandler(this.MnuSendScreenCaptureClick);
            // 
            // mnuSend2Myself
            // 
            this.mnuSend2Myself.Name = "mnuSend2Myself";
            this.mnuSend2Myself.Size = new System.Drawing.Size(105, 22);
            this.mnuSend2Myself.Text = "Myself";
            this.mnuSend2Myself.Click += new System.EventHandler(this.MnuSendScreenCaptureClick);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(193, 6);
            // 
            // mnuReinstallKeyboardAndMouseHook
            // 
            this.mnuReinstallKeyboardAndMouseHook.Name = "mnuReinstallKeyboardAndMouseHook";
            this.mnuReinstallKeyboardAndMouseHook.Size = new System.Drawing.Size(196, 22);
            this.mnuReinstallKeyboardAndMouseHook.Text = "&Reinstall Input hook";
            this.mnuReinstallKeyboardAndMouseHook.ToolTipText = "This might help when keyboard/Mouse redirection stops working";
            this.mnuReinstallKeyboardAndMouseHook.Visible = false;
            this.mnuReinstallKeyboardAndMouseHook.Click += new System.EventHandler(this.MnuReinstallKeyboardAndMouseHook_Click);
            // 
            // mnuMachineMatrix
            // 
            this.mnuMachineMatrix.Font = new System.Drawing.Font(System.Windows.Forms.Control.DefaultFont.Name, System.Windows.Forms.Control.DefaultFont.Size, System.Drawing.FontStyle.Bold);
            this.mnuMachineMatrix.ForeColor = System.Drawing.Color.Black;
            this.mnuMachineMatrix.Name = "mnuMachineMatrix";
            this.mnuMachineMatrix.Size = new System.Drawing.Size(196, 22);
            this.mnuMachineMatrix.Text = "&Settings";
            this.mnuMachineMatrix.Click += new System.EventHandler(this.MnuMachineMatrix_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(193, 6);
            // 
            // mnuALLPC
            // 
            this.MnuALLPC.CheckOnClick = true;
            this.MnuALLPC.Name = "mnuALLPC";
            this.MnuALLPC.Size = new System.Drawing.Size(196, 22);
            this.MnuALLPC.Text = "&ALL COMPUTERS";
            this.MnuALLPC.ToolTipText = "Repeat Mouse/keyboard in all machines.";
            this.MnuALLPC.Click += new System.EventHandler(this.MnuAllPC_Click);
            // 
            // dUCTDOToolStripMenuItem
            // 
            this.dUCTDOToolStripMenuItem.Name = "dUCTDOToolStripMenuItem";
            this.dUCTDOToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.dUCTDOToolStripMenuItem.Tag = "MACHINE: TEST1";
            this.dUCTDOToolStripMenuItem.Text = "DUCTDO";
            // 
            // tRUONG2DToolStripMenuItem
            // 
            this.tRUONG2DToolStripMenuItem.Name = "tRUONG2DToolStripMenuItem";
            this.tRUONG2DToolStripMenuItem.Size = new System.Drawing.Size(196, 22);
            this.tRUONG2DToolStripMenuItem.Tag = "MACHINE: TEST2";
            this.tRUONG2DToolStripMenuItem.Text = "TRUONG2D";
            // 
            // notifyIcon
            // 
            this.NotifyIcon.BalloonTipIcon = System.Windows.Forms.ToolTipIcon.Info;
            this.NotifyIcon.BalloonTipText = "Microsoft® Visual Studio® 2010";
            this.NotifyIcon.BalloonTipTitle = "Microsoft® Visual Studio® 2010";
            this.NotifyIcon.ContextMenuStrip = this.MainMenu;
            this.NotifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));
            this.NotifyIcon.Text = "Microsoft® Visual Studio® 2010";
            this.NotifyIcon.MouseDown += new System.Windows.Forms.MouseEventHandler(this.NotifyIcon_MouseDown);
            // 
            // imgListIcon
            // 
            this.imgListIcon.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imgListIcon.ImageStream")));
            this.imgListIcon.TransparentColor = System.Drawing.Color.Transparent;
            this.imgListIcon.Images.SetKeyName(0, "Logo.ico");
            // 
            // picLogonLogo
            // 
            this.picLogonLogo.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.picLogonLogo.Image = global::MouseWithoutBorders.Properties.Resources.MouseWithoutBorders;
            this.picLogonLogo.Location = new System.Drawing.Point(99, 62);
            this.picLogonLogo.Name = "picLogonLogo";
            this.picLogonLogo.Size = new System.Drawing.Size(95, 17);
            this.picLogonLogo.TabIndex = 1;
            this.picLogonLogo.TabStop = false;
            this.picLogonLogo.Visible = false;
            // 
            // frmScreen
            // 
            this.AllowDrop = true;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Black;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.None;
            this.ClientSize = new System.Drawing.Size(10, 10);
            this.ControlBox = false;
            this.Controls.Add(this.picLogonLogo);
            this.Cursor = System.Windows.Forms.Cursors.Default;
            this.Enabled = false;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmScreen";
            this.Opacity = 0.5D;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.Manual;
            this.Text = "Microsoft® Visual Studio® 2010 Application";
            this.TopMost = true;
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FrmScreen_FormClosing);
            this.Load += new System.EventHandler(this.FrmScreen_Load);
            this.Shown += new System.EventHandler(this.FrmScreen_Shown);
            this.MouseMove += new System.Windows.Forms.MouseEventHandler(this.FrmScreen_MouseMove);
            this.MainMenu.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.picLogonLogo)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ToolStripMenuItem mnuExit;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem mnuMachineMatrix;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem dUCTDOToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tRUONG2DToolStripMenuItem;
        private System.Windows.Forms.ImageList imgListIcon;
        private System.Windows.Forms.ToolStripMenuItem mnuSendScreenCapture;
        private System.Windows.Forms.ToolStripMenuItem mnuSend2Myself;
        private System.Windows.Forms.ToolStripMenuItem mnuGetScreenCapture;
        private System.Windows.Forms.ToolStripMenuItem mnuGetFromAll;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripMenuItem allToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripMenuItem mnuAbout;
        private System.Windows.Forms.PictureBox picLogonLogo;
        private System.Windows.Forms.ToolStripMenuItem mnuHelp;
        private System.Windows.Forms.ToolStripMenuItem mnuReinstallKeyboardAndMouseHook;
        private System.Windows.Forms.ToolStripMenuItem mnuGenDumpFile;
        private System.Windows.Forms.ToolStripMenuItem mnuWindowsPhone;
        private System.Windows.Forms.ToolStripMenuItem mnuWindowsPhoneEnable;
        private System.Windows.Forms.ToolStripMenuItem mnuWindowsPhoneInformation;
        private System.Windows.Forms.ToolStripMenuItem mnuWindowsPhoneDownload;
    }
}

