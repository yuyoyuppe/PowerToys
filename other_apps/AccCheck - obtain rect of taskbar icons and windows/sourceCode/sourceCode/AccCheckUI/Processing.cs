using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AccCheckUI
{
    public partial class Processing : Form
    {
        public Processing()
        {
            InitializeComponent();
        }

        public String Status
        {
            get
            {
                return lblVerification.Text;
            }
            set
            {
                lblVerification.Text = value;
            }
        }

        public int Value
        {
            get
            {
                return progressBar.Value;
            }
            set
            {
                progressBar.Value = value;
            }
        }
    }
}