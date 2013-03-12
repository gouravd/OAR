namespace OutlookAttachmentReminder
{
    partial class Options
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Options));
            this.cntxMenuWordList = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.deleteTheWordToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cntxMenuSubjectList = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.deleteTheSubjectToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.txtSize = new System.Windows.Forms.TextBox();
            this.lblSize = new System.Windows.Forms.Label();
            this.chkbxEmptyMessage = new System.Windows.Forms.CheckBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.chkbxDisallowAttachment = new System.Windows.Forms.CheckBox();
            this.lnklblFeedback = new System.Windows.Forms.LinkLabel();
            this.cntxMenuFileTypes = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.deleteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.chkbxRestrictFileTypes = new System.Windows.Forms.CheckBox();
            this.gpBxManage = new System.Windows.Forms.GroupBox();
            this.lblRestrictFileTypes = new System.Windows.Forms.Label();
            this.lstbxFileTypes = new System.Windows.Forms.ListBox();
            this.txtNewWord = new System.Windows.Forms.TextBox();
            this.btnAddFileTypes = new System.Windows.Forms.Button();
            this.lblSubject = new System.Windows.Forms.Label();
            this.llbWordList = new System.Windows.Forms.Label();
            this.btnAddToSubjectList = new System.Windows.Forms.Button();
            this.lstBxSubject = new System.Windows.Forms.ListBox();
            this.lstBxWordList = new System.Windows.Forms.ListBox();
            this.btnAddToWordList = new System.Windows.Forms.Button();
            this.btnReset = new System.Windows.Forms.Button();
            this.rBtnEMM = new System.Windows.Forms.RadioButton();
            this.rBtnPMM = new System.Windows.Forms.RadioButton();
            this.chkBxAutoSaveIncomingAttachments = new System.Windows.Forms.CheckBox();
            this.btnAutoSave = new System.Windows.Forms.Button();
            this.chkBxDeleteAttachments = new System.Windows.Forms.CheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.notifyIcon1 = new System.Windows.Forms.NotifyIcon(this.components);
            this.btnSave = new System.Windows.Forms.Button();
            this.cntxMenuWordList.SuspendLayout();
            this.cntxMenuSubjectList.SuspendLayout();
            this.cntxMenuFileTypes.SuspendLayout();
            this.gpBxManage.SuspendLayout();
            this.SuspendLayout();
            // 
            // cntxMenuWordList
            // 
            this.cntxMenuWordList.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.deleteTheWordToolStripMenuItem});
            this.cntxMenuWordList.Name = "cntxMenuWordList";
            this.cntxMenuWordList.Size = new System.Drawing.Size(158, 26);
            // 
            // deleteTheWordToolStripMenuItem
            // 
            this.deleteTheWordToolStripMenuItem.Name = "deleteTheWordToolStripMenuItem";
            this.deleteTheWordToolStripMenuItem.Size = new System.Drawing.Size(157, 22);
            this.deleteTheWordToolStripMenuItem.Text = "Delete the word";
            this.deleteTheWordToolStripMenuItem.Click += new System.EventHandler(this.deleteTheWordToolStripMenuItem_Click);
            // 
            // cntxMenuSubjectList
            // 
            this.cntxMenuSubjectList.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.deleteTheSubjectToolStripMenuItem});
            this.cntxMenuSubjectList.Name = "cntxMenuWordList";
            this.cntxMenuSubjectList.Size = new System.Drawing.Size(170, 26);
            // 
            // deleteTheSubjectToolStripMenuItem
            // 
            this.deleteTheSubjectToolStripMenuItem.Name = "deleteTheSubjectToolStripMenuItem";
            this.deleteTheSubjectToolStripMenuItem.Size = new System.Drawing.Size(169, 22);
            this.deleteTheSubjectToolStripMenuItem.Text = "Delete the Subject";
            this.deleteTheSubjectToolStripMenuItem.Click += new System.EventHandler(this.deleteTheSubjectToolStripMenuItem_Click);
            // 
            // txtSize
            // 
            this.txtSize.Location = new System.Drawing.Point(15, 326);
            this.txtSize.Name = "txtSize";
            this.txtSize.Size = new System.Drawing.Size(124, 20);
            this.txtSize.TabIndex = 2;
            this.txtSize.Text = "100000";
            // 
            // lblSize
            // 
            this.lblSize.AutoSize = true;
            this.lblSize.Location = new System.Drawing.Point(13, 310);
            this.lblSize.Name = "lblSize";
            this.lblSize.Size = new System.Drawing.Size(124, 13);
            this.lblSize.TabIndex = 5;
            this.lblSize.Text = "Size of Attachment in KB";
            // 
            // chkbxEmptyMessage
            // 
            this.chkbxEmptyMessage.AutoSize = true;
            this.chkbxEmptyMessage.Checked = true;
            this.chkbxEmptyMessage.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkbxEmptyMessage.Location = new System.Drawing.Point(156, 401);
            this.chkbxEmptyMessage.Name = "chkbxEmptyMessage";
            this.chkbxEmptyMessage.Size = new System.Drawing.Size(138, 17);
            this.chkbxEmptyMessage.TabIndex = 6;
            this.chkbxEmptyMessage.Text = "Warn on Empty Subject";
            this.chkbxEmptyMessage.UseVisualStyleBackColor = true;
            // 
            // richTextBox1
            // 
            this.richTextBox1.BackColor = System.Drawing.Color.DarkGray;
            this.richTextBox1.Enabled = false;
            this.richTextBox1.Location = new System.Drawing.Point(5, 424);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(508, 58);
            this.richTextBox1.TabIndex = 13;
            this.richTextBox1.Text = resources.GetString("richTextBox1.Text");
            // 
            // chkbxDisallowAttachment
            // 
            this.chkbxDisallowAttachment.AutoSize = true;
            this.chkbxDisallowAttachment.Location = new System.Drawing.Point(156, 379);
            this.chkbxDisallowAttachment.Name = "chkbxDisallowAttachment";
            this.chkbxDisallowAttachment.Size = new System.Drawing.Size(146, 17);
            this.chkbxDisallowAttachment.TabIndex = 14;
            this.chkbxDisallowAttachment.Text = "Do not allow attachments";
            this.chkbxDisallowAttachment.UseVisualStyleBackColor = true;
            this.chkbxDisallowAttachment.CheckedChanged += new System.EventHandler(this.chkbxDisallowAttachment_CheckedChanged);
            // 
            // lnklblFeedback
            // 
            this.lnklblFeedback.AutoSize = true;
            this.lnklblFeedback.Location = new System.Drawing.Point(458, 374);
            this.lnklblFeedback.Name = "lnklblFeedback";
            this.lnklblFeedback.Size = new System.Drawing.Size(55, 13);
            this.lnklblFeedback.TabIndex = 15;
            this.lnklblFeedback.TabStop = true;
            this.lnklblFeedback.Text = "Feedback";
            this.lnklblFeedback.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnklblFeedback_LinkClicked);
            // 
            // cntxMenuFileTypes
            // 
            this.cntxMenuFileTypes.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.deleteToolStripMenuItem});
            this.cntxMenuFileTypes.Name = "cntxMenuFileTypes";
            this.cntxMenuFileTypes.Size = new System.Drawing.Size(108, 26);
            // 
            // deleteToolStripMenuItem
            // 
            this.deleteToolStripMenuItem.Name = "deleteToolStripMenuItem";
            this.deleteToolStripMenuItem.Size = new System.Drawing.Size(107, 22);
            this.deleteToolStripMenuItem.Text = "Delete";
            this.deleteToolStripMenuItem.Click += new System.EventHandler(this.deleteToolStripMenuItem_Click);
            // 
            // chkbxRestrictFileTypes
            // 
            this.chkbxRestrictFileTypes.AutoSize = true;
            this.chkbxRestrictFileTypes.Location = new System.Drawing.Point(156, 310);
            this.chkbxRestrictFileTypes.Name = "chkbxRestrictFileTypes";
            this.chkbxRestrictFileTypes.Size = new System.Drawing.Size(214, 17);
            this.chkbxRestrictFileTypes.TabIndex = 18;
            this.chkbxRestrictFileTypes.Text = "Restrict above Filetypes as attachments";
            this.chkbxRestrictFileTypes.UseVisualStyleBackColor = true;
            // 
            // gpBxManage
            // 
            this.gpBxManage.BackColor = System.Drawing.Color.Gainsboro;
            this.gpBxManage.Controls.Add(this.lblRestrictFileTypes);
            this.gpBxManage.Controls.Add(this.lstbxFileTypes);
            this.gpBxManage.Controls.Add(this.txtNewWord);
            this.gpBxManage.Controls.Add(this.btnAddFileTypes);
            this.gpBxManage.Controls.Add(this.lblSubject);
            this.gpBxManage.Controls.Add(this.llbWordList);
            this.gpBxManage.Controls.Add(this.btnAddToSubjectList);
            this.gpBxManage.Controls.Add(this.lstBxSubject);
            this.gpBxManage.Controls.Add(this.lstBxWordList);
            this.gpBxManage.Controls.Add(this.btnAddToWordList);
            this.gpBxManage.Location = new System.Drawing.Point(5, 6);
            this.gpBxManage.Name = "gpBxManage";
            this.gpBxManage.Size = new System.Drawing.Size(508, 246);
            this.gpBxManage.TabIndex = 19;
            this.gpBxManage.TabStop = false;
            this.gpBxManage.Text = "Manage";
            // 
            // lblRestrictFileTypes
            // 
            this.lblRestrictFileTypes.AutoSize = true;
            this.lblRestrictFileTypes.Location = new System.Drawing.Point(194, 100);
            this.lblRestrictFileTypes.Name = "lblRestrictFileTypes";
            this.lblRestrictFileTypes.Size = new System.Drawing.Size(98, 13);
            this.lblRestrictFileTypes.TabIndex = 29;
            this.lblRestrictFileTypes.Text = "FileTypes to restrict";
            // 
            // lstbxFileTypes
            // 
            this.lstbxFileTypes.ContextMenuStrip = this.cntxMenuFileTypes;
            this.lstbxFileTypes.FormattingEnabled = true;
            this.lstbxFileTypes.Location = new System.Drawing.Point(197, 119);
            this.lstbxFileTypes.Name = "lstbxFileTypes";
            this.lstbxFileTypes.Size = new System.Drawing.Size(106, 121);
            this.lstbxFileTypes.TabIndex = 28;
            this.lstbxFileTypes.SelectedIndexChanged += new System.EventHandler(this.lstbxFileTypes_SelectedIndexChanged);
            // 
            // txtNewWord
            // 
            this.txtNewWord.Location = new System.Drawing.Point(174, 41);
            this.txtNewWord.Name = "txtNewWord";
            this.txtNewWord.Size = new System.Drawing.Size(153, 20);
            this.txtNewWord.TabIndex = 24;
            this.txtNewWord.Text = "Enter new word/phrase";
            // 
            // btnAddFileTypes
            // 
            this.btnAddFileTypes.Location = new System.Drawing.Point(226, 67);
            this.btnAddFileTypes.Name = "btnAddFileTypes";
            this.btnAddFileTypes.Size = new System.Drawing.Size(50, 23);
            this.btnAddFileTypes.TabIndex = 27;
            this.btnAddFileTypes.Text = "vvvv";
            this.btnAddFileTypes.UseVisualStyleBackColor = true;
            this.btnAddFileTypes.Click += new System.EventHandler(this.btnAddFileTypes_Click);
            // 
            // lblSubject
            // 
            this.lblSubject.AutoSize = true;
            this.lblSubject.Location = new System.Drawing.Point(330, 25);
            this.lblSubject.Name = "lblSubject";
            this.lblSubject.Size = new System.Drawing.Size(62, 13);
            this.lblSubject.TabIndex = 23;
            this.lblSubject.Text = "Subject List";
            // 
            // llbWordList
            // 
            this.llbWordList.AutoSize = true;
            this.llbWordList.Location = new System.Drawing.Point(7, 24);
            this.llbWordList.Name = "llbWordList";
            this.llbWordList.Size = new System.Drawing.Size(52, 13);
            this.llbWordList.TabIndex = 22;
            this.llbWordList.Text = "Word List";
            // 
            // btnAddToSubjectList
            // 
            this.btnAddToSubjectList.Location = new System.Drawing.Point(277, 67);
            this.btnAddToSubjectList.Name = "btnAddToSubjectList";
            this.btnAddToSubjectList.Size = new System.Drawing.Size(50, 23);
            this.btnAddToSubjectList.TabIndex = 26;
            this.btnAddToSubjectList.Text = ">>>>";
            this.btnAddToSubjectList.UseVisualStyleBackColor = true;
            this.btnAddToSubjectList.Click += new System.EventHandler(this.btnAddToSubjectList_Click);
            // 
            // lstBxSubject
            // 
            this.lstBxSubject.ContextMenuStrip = this.cntxMenuSubjectList;
            this.lstBxSubject.FormattingEnabled = true;
            this.lstBxSubject.Location = new System.Drawing.Point(333, 41);
            this.lstBxSubject.Name = "lstBxSubject";
            this.lstBxSubject.Size = new System.Drawing.Size(169, 199);
            this.lstBxSubject.TabIndex = 21;
            this.lstBxSubject.SelectedIndexChanged += new System.EventHandler(this.lstBxSubject_SelectedIndexChanged);
            // 
            // lstBxWordList
            // 
            this.lstBxWordList.ContextMenuStrip = this.cntxMenuWordList;
            this.lstBxWordList.FormattingEnabled = true;
            this.lstBxWordList.Location = new System.Drawing.Point(10, 40);
            this.lstBxWordList.Name = "lstBxWordList";
            this.lstBxWordList.Size = new System.Drawing.Size(158, 199);
            this.lstBxWordList.TabIndex = 20;
            this.lstBxWordList.SelectedIndexChanged += new System.EventHandler(this.lstBxWordList_SelectedIndexChanged);
            // 
            // btnAddToWordList
            // 
            this.btnAddToWordList.Location = new System.Drawing.Point(174, 67);
            this.btnAddToWordList.Name = "btnAddToWordList";
            this.btnAddToWordList.Size = new System.Drawing.Size(50, 23);
            this.btnAddToWordList.TabIndex = 25;
            this.btnAddToWordList.Text = "<<<<";
            this.btnAddToWordList.UseVisualStyleBackColor = true;
            this.btnAddToWordList.Click += new System.EventHandler(this.btnAddToWordList_Click);
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(438, 391);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(75, 23);
            this.btnReset.TabIndex = 34;
            this.btnReset.Text = "Reset";
            this.btnReset.UseVisualStyleBackColor = true;
            // 
            // rBtnEMM
            // 
            this.rBtnEMM.AutoSize = true;
            this.rBtnEMM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rBtnEMM.Location = new System.Drawing.Point(248, 272);
            this.rBtnEMM.Name = "rBtnEMM";
            this.rBtnEMM.Size = new System.Drawing.Size(200, 20);
            this.rBtnEMM.TabIndex = 36;
            this.rBtnEMM.Text = "Exact Match Mode (EMM)";
            this.rBtnEMM.UseVisualStyleBackColor = true;
            // 
            // rBtnPMM
            // 
            this.rBtnPMM.AutoSize = true;
            this.rBtnPMM.Checked = true;
            this.rBtnPMM.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rBtnPMM.Location = new System.Drawing.Point(31, 272);
            this.rBtnPMM.Name = "rBtnPMM";
            this.rBtnPMM.Size = new System.Drawing.Size(211, 20);
            this.rBtnPMM.TabIndex = 35;
            this.rBtnPMM.TabStop = true;
            this.rBtnPMM.Text = "Pattern Match Mode (PMM)";
            this.rBtnPMM.UseVisualStyleBackColor = true;
            // 
            // chkBxAutoSaveIncomingAttachments
            // 
            this.chkBxAutoSaveIncomingAttachments.AutoSize = true;
            this.chkBxAutoSaveIncomingAttachments.Location = new System.Drawing.Point(156, 333);
            this.chkBxAutoSaveIncomingAttachments.Name = "chkBxAutoSaveIncomingAttachments";
            this.chkBxAutoSaveIncomingAttachments.Size = new System.Drawing.Size(184, 17);
            this.chkBxAutoSaveIncomingAttachments.TabIndex = 37;
            this.chkBxAutoSaveIncomingAttachments.Text = "Auto Save Incoming Attachments";
            this.chkBxAutoSaveIncomingAttachments.UseVisualStyleBackColor = true;
            // 
            // btnAutoSave
            // 
            this.btnAutoSave.BackColor = System.Drawing.Color.Tomato;
            this.btnAutoSave.Location = new System.Drawing.Point(338, 329);
            this.btnAutoSave.Name = "btnAutoSave";
            this.btnAutoSave.Size = new System.Drawing.Size(24, 23);
            this.btnAutoSave.TabIndex = 38;
            this.btnAutoSave.Text = "A";
            this.btnAutoSave.UseVisualStyleBackColor = false;
            this.btnAutoSave.Click += new System.EventHandler(this.btnAutoSave_Click);
            // 
            // chkBxDeleteAttachments
            // 
            this.chkBxDeleteAttachments.AutoSize = true;
            this.chkBxDeleteAttachments.Location = new System.Drawing.Point(156, 356);
            this.chkBxDeleteAttachments.Name = "chkBxDeleteAttachments";
            this.chkBxDeleteAttachments.Size = new System.Drawing.Size(177, 17);
            this.chkBxDeleteAttachments.TabIndex = 39;
            this.chkBxDeleteAttachments.Text = "Delete Attachments after saving";
            this.toolTip1.SetToolTip(this.chkBxDeleteAttachments, "Delete attacments from Email after saving on disk");
            this.chkBxDeleteAttachments.UseVisualStyleBackColor = true;
            // 
            // notifyIcon1
            // 
            this.notifyIcon1.Text = "notifyIcon1";
            this.notifyIcon1.Visible = true;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(14, 360);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(125, 54);
            this.btnSave.TabIndex = 40;
            this.btnSave.Text = "Save Settings";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // Options
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.Gainsboro;
            this.ClientSize = new System.Drawing.Size(519, 485);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.rBtnEMM);
            this.Controls.Add(this.rBtnPMM);
            this.Controls.Add(this.btnReset);
            this.Controls.Add(this.gpBxManage);
            this.Controls.Add(this.lnklblFeedback);
            this.Controls.Add(this.lblSize);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.txtSize);
            this.Controls.Add(this.chkBxDeleteAttachments);
            this.Controls.Add(this.btnAutoSave);
            this.Controls.Add(this.chkbxEmptyMessage);
            this.Controls.Add(this.chkBxAutoSaveIncomingAttachments);
            this.Controls.Add(this.chkbxRestrictFileTypes);
            this.Controls.Add(this.chkbxDisallowAttachment);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Options";
            this.Text = "OAR v1.0b Options - http://oar.codeplex.com";
            this.Load += new System.EventHandler(this.Options_Load);
            this.cntxMenuWordList.ResumeLayout(false);
            this.cntxMenuSubjectList.ResumeLayout(false);
            this.cntxMenuFileTypes.ResumeLayout(false);
            this.gpBxManage.ResumeLayout(false);
            this.gpBxManage.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.TextBox txtSize;
        public System.Windows.Forms.Label lblSize;
        public System.Windows.Forms.CheckBox chkbxEmptyMessage;
        public System.Windows.Forms.ContextMenuStrip cntxMenuWordList;
        public System.Windows.Forms.ToolStripMenuItem deleteTheWordToolStripMenuItem;
        public System.Windows.Forms.ContextMenuStrip cntxMenuSubjectList;
        public System.Windows.Forms.ToolStripMenuItem deleteTheSubjectToolStripMenuItem;
        public System.Windows.Forms.RichTextBox richTextBox1;
        public System.Windows.Forms.CheckBox chkbxDisallowAttachment;
        public System.Windows.Forms.LinkLabel lnklblFeedback;
        public System.Windows.Forms.CheckBox chkbxRestrictFileTypes;
        public System.Windows.Forms.ContextMenuStrip cntxMenuFileTypes;
        public System.Windows.Forms.ToolStripMenuItem deleteToolStripMenuItem;
        public System.Windows.Forms.GroupBox gpBxManage;
        public System.Windows.Forms.Label lblRestrictFileTypes;
        public System.Windows.Forms.ListBox lstbxFileTypes;
        public System.Windows.Forms.TextBox txtNewWord;
        public System.Windows.Forms.Button btnAddFileTypes;
        public System.Windows.Forms.Label lblSubject;
        public System.Windows.Forms.Label llbWordList;
        public System.Windows.Forms.Button btnAddToSubjectList;
        public System.Windows.Forms.ListBox lstBxSubject;
        public System.Windows.Forms.ListBox lstBxWordList;
        public System.Windows.Forms.Button btnAddToWordList;
        private System.Windows.Forms.Button btnReset;
        public System.Windows.Forms.RadioButton rBtnEMM;
        public System.Windows.Forms.RadioButton rBtnPMM;
        public System.Windows.Forms.CheckBox chkBxAutoSaveIncomingAttachments;
        private System.Windows.Forms.Button btnAutoSave;
        private System.Windows.Forms.ToolTip toolTip1;
        public System.Windows.Forms.CheckBox chkBxDeleteAttachments;
        private System.Windows.Forms.NotifyIcon notifyIcon1;
        public System.Windows.Forms.Button btnSave;
    }
}