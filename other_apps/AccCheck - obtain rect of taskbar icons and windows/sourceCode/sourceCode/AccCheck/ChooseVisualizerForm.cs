// (c) Copyright Microsoft Corporation.
// This source is subject to the Microsoft Permissive License.
// See http://www.microsoft.com/opensource/licenses.mspx#Ms-PL.
// All other rights reserved.

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace AccCheck
{
    public partial class ChooseVisualizerForm : Form
    {
        public ChooseVisualizerForm()
        {
            InitializeComponent();
        }

        public void AddItem(Type t)
        {
            lbVisualizers.Items.Add(t);
        }

        public void Clear()
        {
            lbVisualizers.Items.Clear();
        }

        public int GetSelectedIndex()
        {
            return lbVisualizers.SelectedIndex;
        }

        private void lbVisualizers_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnOK.Enabled = (lbVisualizers.SelectedIndex != -1);
        }
    }
}
