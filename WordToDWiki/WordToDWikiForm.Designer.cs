namespace WordToDWiki
{
    partial class WordToDWikiForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(WordToDWikiForm));
            this.lstListOfFiles = new System.Windows.Forms.ListBox();
            this.cms_FileList = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.cmsi_RSItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cmsi_RAItems = new System.Windows.Forms.ToolStripMenuItem();
            this.btnAddFile = new System.Windows.Forms.Button();
            this.btnProcessStart = new System.Windows.Forms.Button();
            this.rtxtLogProcess = new System.Windows.Forms.RichTextBox();
            this.mnsSoftware = new System.Windows.Forms.MenuStrip();
            this.tsmiSoftware = new System.Windows.Forms.ToolStripMenuItem();
            this.tsmiAbout = new System.Windows.Forms.ToolStripMenuItem();
            this.lblFiles = new System.Windows.Forms.Label();
            this.lblPLog = new System.Windows.Forms.Label();
            this.lblSLog = new System.Windows.Forms.Label();
            this.rtxtLogControl = new System.Windows.Forms.RichTextBox();
            this.cms_FileList.SuspendLayout();
            this.mnsSoftware.SuspendLayout();
            this.SuspendLayout();
            // 
            // lstListOfFiles
            // 
            this.lstListOfFiles.ContextMenuStrip = this.cms_FileList;
            this.lstListOfFiles.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lstListOfFiles.FormattingEnabled = true;
            this.lstListOfFiles.ItemHeight = 15;
            this.lstListOfFiles.Location = new System.Drawing.Point(12, 42);
            this.lstListOfFiles.Name = "lstListOfFiles";
            this.lstListOfFiles.Size = new System.Drawing.Size(195, 64);
            this.lstListOfFiles.TabIndex = 0;
            // 
            // cms_FileList
            // 
            this.cms_FileList.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cmsi_RSItem,
            this.cmsi_RAItems});
            this.cms_FileList.Name = "cms_FileList";
            this.cms_FileList.Size = new System.Drawing.Size(191, 48);
            this.cms_FileList.Opening += new System.ComponentModel.CancelEventHandler(this.cms_FileList_Opening);
            // 
            // cmsi_RSItem
            // 
            this.cmsi_RSItem.Name = "cmsi_RSItem";
            this.cmsi_RSItem.Size = new System.Drawing.Size(190, 22);
            this.cmsi_RSItem.Text = "Remove selected item";
            this.cmsi_RSItem.Click += new System.EventHandler(this.cmsi_RSItem_Click);
            // 
            // cmsi_RAItems
            // 
            this.cmsi_RAItems.Name = "cmsi_RAItems";
            this.cmsi_RAItems.Size = new System.Drawing.Size(190, 22);
            this.cmsi_RAItems.Text = "Remove all items";
            this.cmsi_RAItems.Click += new System.EventHandler(this.cmsi_RAItems_Click);
            // 
            // btnAddFile
            // 
            this.btnAddFile.BackColor = System.Drawing.SystemColors.Control;
            this.btnAddFile.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnAddFile.Location = new System.Drawing.Point(223, 41);
            this.btnAddFile.Name = "btnAddFile";
            this.btnAddFile.Size = new System.Drawing.Size(90, 25);
            this.btnAddFile.TabIndex = 1;
            this.btnAddFile.Text = "Add File(s)";
            this.btnAddFile.UseVisualStyleBackColor = true;
            this.btnAddFile.Click += new System.EventHandler(this.btnAddFile_Click);
            // 
            // btnProcessStart
            // 
            this.btnProcessStart.Font = new System.Drawing.Font("Consolas", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.btnProcessStart.Location = new System.Drawing.Point(223, 72);
            this.btnProcessStart.Name = "btnProcessStart";
            this.btnProcessStart.Size = new System.Drawing.Size(90, 25);
            this.btnProcessStart.TabIndex = 3;
            this.btnProcessStart.Text = "Start";
            this.btnProcessStart.UseVisualStyleBackColor = true;
            this.btnProcessStart.Click += new System.EventHandler(this.btnProcess_Click);
            // 
            // rtxtLogProcess
            // 
            this.rtxtLogProcess.BackColor = System.Drawing.Color.Black;
            this.rtxtLogProcess.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtxtLogProcess.ForeColor = System.Drawing.Color.Lime;
            this.rtxtLogProcess.Location = new System.Drawing.Point(12, 268);
            this.rtxtLogProcess.Name = "rtxtLogProcess";
            this.rtxtLogProcess.ReadOnly = true;
            this.rtxtLogProcess.Size = new System.Drawing.Size(340, 76);
            this.rtxtLogProcess.TabIndex = 4;
            this.rtxtLogProcess.Text = "";
            // 
            // mnsSoftware
            // 
            this.mnsSoftware.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiSoftware});
            this.mnsSoftware.Location = new System.Drawing.Point(0, 0);
            this.mnsSoftware.Name = "mnsSoftware";
            this.mnsSoftware.Size = new System.Drawing.Size(364, 24);
            this.mnsSoftware.TabIndex = 5;
            this.mnsSoftware.Text = "Software";
            // 
            // tsmiSoftware
            // 
            this.tsmiSoftware.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tsmiAbout});
            this.tsmiSoftware.Name = "tsmiSoftware";
            this.tsmiSoftware.Size = new System.Drawing.Size(65, 20);
            this.tsmiSoftware.Text = "Software";
            // 
            // tsmiAbout
            // 
            this.tsmiAbout.Image = global::WordToDWiki.Properties.Resources.info;
            this.tsmiAbout.Name = "tsmiAbout";
            this.tsmiAbout.Size = new System.Drawing.Size(152, 22);
            this.tsmiAbout.Text = "About";
            this.tsmiAbout.Click += new System.EventHandler(this.tsmiAbout_Click);
            // 
            // lblFiles
            // 
            this.lblFiles.AutoSize = true;
            this.lblFiles.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblFiles.Location = new System.Drawing.Point(12, 24);
            this.lblFiles.Name = "lblFiles";
            this.lblFiles.Size = new System.Drawing.Size(49, 15);
            this.lblFiles.TabIndex = 6;
            this.lblFiles.Text = "Files:";
            // 
            // lblPLog
            // 
            this.lblPLog.AutoSize = true;
            this.lblPLog.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblPLog.Location = new System.Drawing.Point(9, 250);
            this.lblPLog.Name = "lblPLog";
            this.lblPLog.Size = new System.Drawing.Size(91, 15);
            this.lblPLog.TabIndex = 7;
            this.lblPLog.Text = "Process Log:";
            // 
            // lblSLog
            // 
            this.lblSLog.AutoSize = true;
            this.lblSLog.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblSLog.Location = new System.Drawing.Point(12, 132);
            this.lblSLog.Name = "lblSLog";
            this.lblSLog.Size = new System.Drawing.Size(91, 15);
            this.lblSLog.TabIndex = 9;
            this.lblSLog.Text = "Control Log:";
            // 
            // rtxtLogControl
            // 
            this.rtxtLogControl.BackColor = System.Drawing.Color.Black;
            this.rtxtLogControl.Font = new System.Drawing.Font("Consolas", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rtxtLogControl.ForeColor = System.Drawing.Color.Aqua;
            this.rtxtLogControl.Location = new System.Drawing.Point(15, 150);
            this.rtxtLogControl.Name = "rtxtLogControl";
            this.rtxtLogControl.Size = new System.Drawing.Size(340, 76);
            this.rtxtLogControl.TabIndex = 8;
            this.rtxtLogControl.Text = "";
            // 
            // WordToDWikiForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(364, 471);
            this.Controls.Add(this.lblSLog);
            this.Controls.Add(this.rtxtLogControl);
            this.Controls.Add(this.lblPLog);
            this.Controls.Add(this.lblFiles);
            this.Controls.Add(this.rtxtLogProcess);
            this.Controls.Add(this.btnProcessStart);
            this.Controls.Add(this.btnAddFile);
            this.Controls.Add(this.lstListOfFiles);
            this.Controls.Add(this.mnsSoftware);
            this.ForeColor = System.Drawing.Color.Black;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.mnsSoftware;
            this.MinimumSize = new System.Drawing.Size(380, 510);
            this.Name = "WordToDWikiForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Word To Doku Wiki";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.frmDokuWikiFC_FormClosing);
            this.Load += new System.EventHandler(this.frmDokuWikiFC_Load);
            this.SizeChanged += new System.EventHandler(this.frmDokuWikiFC_SizeChanged);
            this.cms_FileList.ResumeLayout(false);
            this.mnsSoftware.ResumeLayout(false);
            this.mnsSoftware.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox lstListOfFiles;
        private System.Windows.Forms.Button btnAddFile;
        private System.Windows.Forms.Button btnProcessStart;
        private System.Windows.Forms.RichTextBox rtxtLogProcess;
        private System.Windows.Forms.MenuStrip mnsSoftware;
        private System.Windows.Forms.ToolStripMenuItem tsmiSoftware;
        private System.Windows.Forms.ToolStripMenuItem tsmiAbout;
        private System.Windows.Forms.Label lblFiles;
        private System.Windows.Forms.Label lblPLog;
        private System.Windows.Forms.ContextMenuStrip cms_FileList;
        private System.Windows.Forms.ToolStripMenuItem cmsi_RSItem;
        private System.Windows.Forms.ToolStripMenuItem cmsi_RAItems;
        private System.Windows.Forms.Label lblSLog;
        private System.Windows.Forms.RichTextBox rtxtLogControl;
    }
}

