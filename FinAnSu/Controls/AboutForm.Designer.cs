namespace FinAnSu
{
    partial class AboutForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AboutForm));
            this.ProductNameLabel = new System.Windows.Forms.Label();
            this.VersionLabel = new System.Windows.Forms.Label();
            this.CloseButton = new System.Windows.Forms.Button();
            this.CopyrightLabel = new System.Windows.Forms.Label();
            this.HomepageLinkLabel = new System.Windows.Forms.LinkLabel();
            this.MailLinkLabel = new System.Windows.Forms.LinkLabel();
            this.LicenseLabel = new System.Windows.Forms.Label();
            this.LicenseLinkLabel = new System.Windows.Forms.LinkLabel();
            this.ContactLabel = new System.Windows.Forms.Label();
            this.ExcelDnaLabel = new System.Windows.Forms.Label();
            this.ExcelDnaCopyrightLabel = new System.Windows.Forms.Label();
            this.DownloadLinkLabel = new System.Windows.Forms.LinkLabel();
            this.ExcelDnaLicenseLabel = new System.Windows.Forms.Label();
            this.ExcelDnaLicenseLinkLabel = new System.Windows.Forms.LinkLabel();
            this.ExcelDnaDownloadLinkLabel = new System.Windows.Forms.LinkLabel();
            this.LogoImage = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.LogoImage)).BeginInit();
            this.SuspendLayout();
            // 
            // ProductNameLabel
            // 
            this.ProductNameLabel.AutoSize = true;
            this.ProductNameLabel.Font = new System.Drawing.Font("Segoe UI", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ProductNameLabel.Location = new System.Drawing.Point(130, 12);
            this.ProductNameLabel.Margin = new System.Windows.Forms.Padding(3, 0, 0, 0);
            this.ProductNameLabel.Name = "ProductNameLabel";
            this.ProductNameLabel.Size = new System.Drawing.Size(109, 32);
            this.ProductNameLabel.TabIndex = 1;
            this.ProductNameLabel.Text = "FinAnSu";
            // 
            // VersionLabel
            // 
            this.VersionLabel.AutoSize = true;
            this.VersionLabel.Location = new System.Drawing.Point(239, 27);
            this.VersionLabel.Margin = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.VersionLabel.Name = "VersionLabel";
            this.VersionLabel.Size = new System.Drawing.Size(46, 13);
            this.VersionLabel.TabIndex = 2;
            this.VersionLabel.Text = "Version";
            // 
            // CloseButton
            // 
            this.CloseButton.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.CloseButton.Location = new System.Drawing.Point(114, 276);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(75, 23);
            this.CloseButton.TabIndex = 15;
            this.CloseButton.Text = "Close";
            this.CloseButton.UseVisualStyleBackColor = true;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // CopyrightLabel
            // 
            this.CopyrightLabel.AutoSize = true;
            this.CopyrightLabel.Location = new System.Drawing.Point(12, 76);
            this.CopyrightLabel.Margin = new System.Windows.Forms.Padding(3, 6, 3, 3);
            this.CopyrightLabel.Name = "CopyrightLabel";
            this.CopyrightLabel.Size = new System.Drawing.Size(179, 13);
            this.CopyrightLabel.TabIndex = 4;
            this.CopyrightLabel.Text = "Copyright © 2011 Bryan McKelvey";
            // 
            // HomepageLinkLabel
            // 
            this.HomepageLinkLabel.ActiveLinkColor = System.Drawing.Color.Blue;
            this.HomepageLinkLabel.AutoSize = true;
            this.HomepageLinkLabel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.HomepageLinkLabel.Location = new System.Drawing.Point(12, 238);
            this.HomepageLinkLabel.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.HomepageLinkLabel.Name = "HomepageLinkLabel";
            this.HomepageLinkLabel.Size = new System.Drawing.Size(130, 13);
            this.HomepageLinkLabel.TabIndex = 13;
            this.HomepageLinkLabel.TabStop = true;
            this.HomepageLinkLabel.Text = "http://www.brymck.com";
            this.HomepageLinkLabel.VisitedLinkColor = System.Drawing.Color.Blue;
            this.HomepageLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.UrlLinkLabel_LinkClicked);
            // 
            // MailLinkLabel
            // 
            this.MailLinkLabel.ActiveLinkColor = System.Drawing.Color.Blue;
            this.MailLinkLabel.AutoSize = true;
            this.MailLinkLabel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.MailLinkLabel.Location = new System.Drawing.Point(12, 254);
            this.MailLinkLabel.Margin = new System.Windows.Forms.Padding(3, 0, 3, 6);
            this.MailLinkLabel.Name = "MailLinkLabel";
            this.MailLinkLabel.Size = new System.Drawing.Size(147, 13);
            this.MailLinkLabel.TabIndex = 14;
            this.MailLinkLabel.TabStop = true;
            this.MailLinkLabel.Text = "bryan.mckelvey@gmail.com";
            this.MailLinkLabel.VisitedLinkColor = System.Drawing.Color.Blue;
            this.MailLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.MailLinkLabel_LinkClicked);
            // 
            // LicenseLabel
            // 
            this.LicenseLabel.AutoSize = true;
            this.LicenseLabel.Location = new System.Drawing.Point(12, 92);
            this.LicenseLabel.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.LicenseLabel.Name = "LicenseLabel";
            this.LicenseLabel.Size = new System.Drawing.Size(167, 13);
            this.LicenseLabel.TabIndex = 5;
            this.LicenseLabel.Text = "Licensed under the MIT license:";
            // 
            // LicenseLinkLabel
            // 
            this.LicenseLinkLabel.ActiveLinkColor = System.Drawing.Color.Blue;
            this.LicenseLinkLabel.AutoSize = true;
            this.LicenseLinkLabel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.LicenseLinkLabel.Location = new System.Drawing.Point(12, 108);
            this.LicenseLinkLabel.Margin = new System.Windows.Forms.Padding(3, 0, 3, 6);
            this.LicenseLinkLabel.Name = "LicenseLinkLabel";
            this.LicenseLinkLabel.Size = new System.Drawing.Size(279, 13);
            this.LicenseLinkLabel.TabIndex = 6;
            this.LicenseLinkLabel.TabStop = true;
            this.LicenseLinkLabel.Text = "http://www.opensource.org/licenses/mit-license.php";
            this.LicenseLinkLabel.VisitedLinkColor = System.Drawing.Color.Blue;
            this.LicenseLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.UrlLinkLabel_LinkClicked);
            // 
            // ContactLabel
            // 
            this.ContactLabel.AutoSize = true;
            this.ContactLabel.Location = new System.Drawing.Point(12, 222);
            this.ContactLabel.Margin = new System.Windows.Forms.Padding(3, 6, 3, 3);
            this.ContactLabel.Name = "ContactLabel";
            this.ContactLabel.Size = new System.Drawing.Size(50, 13);
            this.ContactLabel.TabIndex = 12;
            this.ContactLabel.Text = "Contact:";
            // 
            // ExcelDnaLabel
            // 
            this.ExcelDnaLabel.AutoSize = true;
            this.ExcelDnaLabel.Location = new System.Drawing.Point(12, 133);
            this.ExcelDnaLabel.Margin = new System.Windows.Forms.Padding(3, 6, 3, 3);
            this.ExcelDnaLabel.Name = "ExcelDnaLabel";
            this.ExcelDnaLabel.Size = new System.Drawing.Size(97, 13);
            this.ExcelDnaLabel.TabIndex = 7;
            this.ExcelDnaLabel.Text = "Utilizes Excel-Dna";
            // 
            // ExcelDnaCopyrightLabel
            // 
            this.ExcelDnaCopyrightLabel.AutoSize = true;
            this.ExcelDnaCopyrightLabel.Location = new System.Drawing.Point(12, 165);
            this.ExcelDnaCopyrightLabel.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.ExcelDnaCopyrightLabel.Name = "ExcelDnaCopyrightLabel";
            this.ExcelDnaCopyrightLabel.Size = new System.Drawing.Size(244, 13);
            this.ExcelDnaCopyrightLabel.TabIndex = 9;
            this.ExcelDnaCopyrightLabel.Text = "Copyright © 2005‒2011 Govert van Drimmelen";
            // 
            // DownloadLinkLabel
            // 
            this.DownloadLinkLabel.ActiveLinkColor = System.Drawing.Color.Blue;
            this.DownloadLinkLabel.AutoSize = true;
            this.DownloadLinkLabel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.DownloadLinkLabel.Location = new System.Drawing.Point(133, 46);
            this.DownloadLinkLabel.Margin = new System.Windows.Forms.Padding(3, 6, 3, 0);
            this.DownloadLinkLabel.Name = "DownloadLinkLabel";
            this.DownloadLinkLabel.Size = new System.Drawing.Size(187, 13);
            this.DownloadLinkLabel.TabIndex = 3;
            this.DownloadLinkLabel.TabStop = true;
            this.DownloadLinkLabel.Text = "https://github.com/brymck/finansu";
            this.DownloadLinkLabel.VisitedLinkColor = System.Drawing.Color.Blue;
            this.DownloadLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.UrlLinkLabel_LinkClicked);
            // 
            // ExcelDnaLicenseLabel
            // 
            this.ExcelDnaLicenseLabel.AutoSize = true;
            this.ExcelDnaLicenseLabel.Location = new System.Drawing.Point(12, 181);
            this.ExcelDnaLicenseLabel.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.ExcelDnaLicenseLabel.Name = "ExcelDnaLicenseLabel";
            this.ExcelDnaLicenseLabel.Size = new System.Drawing.Size(177, 13);
            this.ExcelDnaLicenseLabel.TabIndex = 10;
            this.ExcelDnaLicenseLabel.Text = "Released under a custom license:";
            // 
            // ExcelDnaLicenseLinkLabel
            // 
            this.ExcelDnaLicenseLinkLabel.ActiveLinkColor = System.Drawing.Color.Blue;
            this.ExcelDnaLicenseLinkLabel.AutoSize = true;
            this.ExcelDnaLicenseLinkLabel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.ExcelDnaLicenseLinkLabel.Location = new System.Drawing.Point(12, 197);
            this.ExcelDnaLicenseLinkLabel.Margin = new System.Windows.Forms.Padding(3, 0, 3, 6);
            this.ExcelDnaLicenseLinkLabel.Name = "ExcelDnaLicenseLinkLabel";
            this.ExcelDnaLicenseLinkLabel.Size = new System.Drawing.Size(197, 13);
            this.ExcelDnaLicenseLinkLabel.TabIndex = 11;
            this.ExcelDnaLicenseLinkLabel.TabStop = true;
            this.ExcelDnaLicenseLinkLabel.Text = "http://exceldna.codeplex.com/license";
            this.ExcelDnaLicenseLinkLabel.VisitedLinkColor = System.Drawing.Color.Blue;
            this.ExcelDnaLicenseLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.UrlLinkLabel_LinkClicked);
            // 
            // ExcelDnaDownloadLinkLabel
            // 
            this.ExcelDnaDownloadLinkLabel.ActiveLinkColor = System.Drawing.Color.Blue;
            this.ExcelDnaDownloadLinkLabel.AutoSize = true;
            this.ExcelDnaDownloadLinkLabel.LinkBehavior = System.Windows.Forms.LinkBehavior.HoverUnderline;
            this.ExcelDnaDownloadLinkLabel.Location = new System.Drawing.Point(12, 149);
            this.ExcelDnaDownloadLinkLabel.Margin = new System.Windows.Forms.Padding(3, 0, 3, 3);
            this.ExcelDnaDownloadLinkLabel.Name = "ExcelDnaDownloadLinkLabel";
            this.ExcelDnaDownloadLinkLabel.Size = new System.Drawing.Size(162, 13);
            this.ExcelDnaDownloadLinkLabel.TabIndex = 8;
            this.ExcelDnaDownloadLinkLabel.TabStop = true;
            this.ExcelDnaDownloadLinkLabel.Text = "http://exceldna.codeplex.com/";
            this.ExcelDnaDownloadLinkLabel.VisitedLinkColor = System.Drawing.Color.Blue;
            this.ExcelDnaDownloadLinkLabel.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.UrlLinkLabel_LinkClicked);
            // 
            // LogoImage
            // 
            this.LogoImage.Image = global::FinAnSu.Properties.Resources.logo;
            this.LogoImage.Location = new System.Drawing.Point(12, 12);
            this.LogoImage.Name = "LogoImage";
            this.LogoImage.Size = new System.Drawing.Size(112, 55);
            this.LogoImage.TabIndex = 0;
            this.LogoImage.TabStop = false;
            // 
            // AboutForm
            // 
            this.AcceptButton = this.CloseButton;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.CloseButton;
            this.ClientSize = new System.Drawing.Size(332, 302);
            this.Controls.Add(this.ExcelDnaLicenseLabel);
            this.Controls.Add(this.DownloadLinkLabel);
            this.Controls.Add(this.ExcelDnaCopyrightLabel);
            this.Controls.Add(this.ExcelDnaLabel);
            this.Controls.Add(this.ContactLabel);
            this.Controls.Add(this.ExcelDnaLicenseLinkLabel);
            this.Controls.Add(this.ExcelDnaDownloadLinkLabel);
            this.Controls.Add(this.LicenseLinkLabel);
            this.Controls.Add(this.LicenseLabel);
            this.Controls.Add(this.MailLinkLabel);
            this.Controls.Add(this.HomepageLinkLabel);
            this.Controls.Add(this.CopyrightLabel);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.VersionLabel);
            this.Controls.Add(this.ProductNameLabel);
            this.Controls.Add(this.LogoImage);
            this.Font = new System.Drawing.Font("Segoe UI", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "AboutForm";
            this.ShowInTaskbar = false;
            this.Text = "About FinAnSu";
            ((System.ComponentModel.ISupportInitialize)(this.LogoImage)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox LogoImage;
        private System.Windows.Forms.Label ProductNameLabel;
        private System.Windows.Forms.Label VersionLabel;
        private System.Windows.Forms.Button CloseButton;
        private System.Windows.Forms.Label CopyrightLabel;
        private System.Windows.Forms.LinkLabel HomepageLinkLabel;
        private System.Windows.Forms.LinkLabel MailLinkLabel;
        private System.Windows.Forms.Label LicenseLabel;
        private System.Windows.Forms.LinkLabel LicenseLinkLabel;
        private System.Windows.Forms.Label ContactLabel;
        private System.Windows.Forms.Label ExcelDnaLabel;
        private System.Windows.Forms.Label ExcelDnaCopyrightLabel;
        private System.Windows.Forms.LinkLabel DownloadLinkLabel;
        private System.Windows.Forms.Label ExcelDnaLicenseLabel;
        private System.Windows.Forms.LinkLabel ExcelDnaLicenseLinkLabel;
        private System.Windows.Forms.LinkLabel ExcelDnaDownloadLinkLabel;
    }
}