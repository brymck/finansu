/* 
 * FinAnSu
 * http://code.google.com/p/finansu/
 * 
 * Copyright 2011, Bryan McKelvey
 * Licensed under the MIT license
 * http://www.opensource.org/licenses/mit-license.php
 */

using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace FinAnSu
{
    [ComVisible(true)]
    public partial class AboutForm : Form
    {
        public AboutForm()
        {
            InitializeComponent();
            this.VersionLabel.Text = "Version " + Main.InstalledVersion();
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void UrlLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel label = (LinkLabel)sender;
            Process.Start(label.Text);
        }

        private void MailLinkLabel_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            LinkLabel label = (LinkLabel)sender;
            Process.Start("mailto:" + label.Text);
        }
    }
}