namespace WordToDWiki
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
            this.rtxtAbout = new System.Windows.Forms.RichTextBox();
            this.pcbLogo = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.pcbLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // rtxtAbout
            // 
            this.rtxtAbout.BorderStyle = System.Windows.Forms.BorderStyle.None;
            this.rtxtAbout.Location = new System.Drawing.Point(146, 12);
            this.rtxtAbout.Name = "rtxtAbout";
            this.rtxtAbout.ReadOnly = true;
            this.rtxtAbout.Size = new System.Drawing.Size(301, 128);
            this.rtxtAbout.TabIndex = 1;
            this.rtxtAbout.Text = "";
            // 
            // pcbLogo
            // 
            this.pcbLogo.Image = global::WordToDWiki.Properties.Resources.logoWDWiki;
            this.pcbLogo.Location = new System.Drawing.Point(0, 0);
            this.pcbLogo.Name = "pcbLogo";
            this.pcbLogo.Size = new System.Drawing.Size(140, 140);
            this.pcbLogo.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.pcbLogo.TabIndex = 2;
            this.pcbLogo.TabStop = false;
            // 
            // AboutForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(459, 151);
            this.Controls.Add(this.rtxtAbout);
            this.Controls.Add(this.pcbLogo);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MaximumSize = new System.Drawing.Size(475, 190);
            this.MinimizeBox = false;
            this.MinimumSize = new System.Drawing.Size(475, 190);
            this.Name = "AboutForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "About";
            this.Load += new System.EventHandler(this.About_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pcbLogo)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.PictureBox pcbLogo;
        private System.Windows.Forms.RichTextBox rtxtAbout;
    }
}