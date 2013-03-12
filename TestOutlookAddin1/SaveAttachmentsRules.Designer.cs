namespace OutlookAttachmentReminder
{
    partial class SaveAttachmentsRules
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
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btnAddRule = new System.Windows.Forms.Button();
            this.btnDeleteRule = new System.Windows.Forms.Button();
            this.txtRuleName = new System.Windows.Forms.TextBox();
            this.txtDest = new System.Windows.Forms.TextBox();
            this.txtSub = new System.Windows.Forms.TextBox();
            this.txtTo = new System.Windows.Forms.TextBox();
            this.txtFrom = new System.Windows.Forms.TextBox();
            this.chkBxRemoveAttFromMail = new System.Windows.Forms.CheckBox();
            this.chkBxOverwriteAttWSameName = new System.Windows.Forms.CheckBox();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.chkBxRuleActive = new System.Windows.Forms.CheckBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.HorizontalScrollbar = true;
            this.listBox1.Location = new System.Drawing.Point(3, 6);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(120, 199);
            this.listBox1.Sorted = true;
            this.listBox1.TabIndex = 0;
            this.listBox1.SelectedIndexChanged += new System.EventHandler(this.listBox1_SelectedIndexChanged);
            // 
            // btnAddRule
            // 
            this.btnAddRule.Location = new System.Drawing.Point(392, 136);
            this.btnAddRule.Name = "btnAddRule";
            this.btnAddRule.Size = new System.Drawing.Size(75, 27);
            this.btnAddRule.TabIndex = 1;
            this.btnAddRule.Text = "Add Rule";
            this.btnAddRule.UseVisualStyleBackColor = true;
            this.btnAddRule.Click += new System.EventHandler(this.btnAddRule_Click);
            // 
            // btnDeleteRule
            // 
            this.btnDeleteRule.Location = new System.Drawing.Point(392, 163);
            this.btnDeleteRule.Name = "btnDeleteRule";
            this.btnDeleteRule.Size = new System.Drawing.Size(75, 19);
            this.btnDeleteRule.TabIndex = 2;
            this.btnDeleteRule.Text = "Delete Rule";
            this.btnDeleteRule.UseVisualStyleBackColor = true;
            this.btnDeleteRule.Click += new System.EventHandler(this.btnDeleteRule_Click);
            // 
            // txtRuleName
            // 
            this.txtRuleName.Location = new System.Drawing.Point(129, 6);
            this.txtRuleName.Name = "txtRuleName";
            this.txtRuleName.Size = new System.Drawing.Size(212, 20);
            this.txtRuleName.TabIndex = 3;
            this.toolTip1.SetToolTip(this.txtRuleName, "Enter Rule Name");
            // 
            // txtDest
            // 
            this.txtDest.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend;
            this.txtDest.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.FileSystemDirectories;
            this.txtDest.Location = new System.Drawing.Point(129, 32);
            this.txtDest.Name = "txtDest";
            this.txtDest.Size = new System.Drawing.Size(338, 20);
            this.txtDest.TabIndex = 4;
            this.toolTip1.SetToolTip(this.txtDest, "Enter destination folder where the attachments should be saved.");
            // 
            // txtSub
            // 
            this.txtSub.Location = new System.Drawing.Point(129, 58);
            this.txtSub.Name = "txtSub";
            this.txtSub.Size = new System.Drawing.Size(338, 20);
            this.txtSub.TabIndex = 5;
            this.toolTip1.SetToolTip(this.txtSub, "Entr subject to scan. Multiple Subjects have to be semi-colon separated.");
            // 
            // txtTo
            // 
            this.txtTo.Location = new System.Drawing.Point(129, 84);
            this.txtTo.Name = "txtTo";
            this.txtTo.Size = new System.Drawing.Size(338, 20);
            this.txtTo.TabIndex = 6;
            this.toolTip1.SetToolTip(this.txtTo, "Enter Recipient. Should be email IDs. Multiple IDs have to be semi-colon separate" +
                    "d.");
            // 
            // txtFrom
            // 
            this.txtFrom.Location = new System.Drawing.Point(129, 110);
            this.txtFrom.Name = "txtFrom";
            this.txtFrom.Size = new System.Drawing.Size(338, 20);
            this.txtFrom.TabIndex = 7;
            this.toolTip1.SetToolTip(this.txtFrom, "Enter Sender\'s name. Multiple names have to be semi-colon separated.");
            // 
            // chkBxRemoveAttFromMail
            // 
            this.chkBxRemoveAttFromMail.AutoSize = true;
            this.chkBxRemoveAttFromMail.Location = new System.Drawing.Point(129, 169);
            this.chkBxRemoveAttFromMail.Name = "chkBxRemoveAttFromMail";
            this.chkBxRemoveAttFromMail.Size = new System.Drawing.Size(167, 17);
            this.chkBxRemoveAttFromMail.TabIndex = 8;
            this.chkBxRemoveAttFromMail.Text = "Remove attachment from Mail";
            this.toolTip1.SetToolTip(this.chkBxRemoveAttFromMail, "Check to remove attachment from original mail.");
            this.chkBxRemoveAttFromMail.UseVisualStyleBackColor = true;
            this.chkBxRemoveAttFromMail.CheckedChanged += new System.EventHandler(this.chkBxRemoveAttFromMail_CheckedChanged);
            // 
            // chkBxOverwriteAttWSameName
            // 
            this.chkBxOverwriteAttWSameName.AutoSize = true;
            this.chkBxOverwriteAttWSameName.Location = new System.Drawing.Point(129, 192);
            this.chkBxOverwriteAttWSameName.Name = "chkBxOverwriteAttWSameName";
            this.chkBxOverwriteAttWSameName.Size = new System.Drawing.Size(212, 17);
            this.chkBxOverwriteAttWSameName.TabIndex = 9;
            this.chkBxOverwriteAttWSameName.Text = "Overwrite Attachments with same name";
            this.toolTip1.SetToolTip(this.chkBxOverwriteAttWSameName, "Check to overwrite attachments with same name.");
            this.chkBxOverwriteAttWSameName.UseVisualStyleBackColor = true;
            this.chkBxOverwriteAttWSameName.CheckedChanged += new System.EventHandler(this.chkBxOverwriteAttWSameName_CheckedChanged);
            // 
            // chkBxRuleActive
            // 
            this.chkBxRuleActive.AutoSize = true;
            this.chkBxRuleActive.Location = new System.Drawing.Point(129, 146);
            this.chkBxRuleActive.Name = "chkBxRuleActive";
            this.chkBxRuleActive.Size = new System.Drawing.Size(92, 17);
            this.chkBxRuleActive.TabIndex = 12;
            this.chkBxRuleActive.Text = "Rule Is Active";
            this.chkBxRuleActive.UseVisualStyleBackColor = true;
            this.chkBxRuleActive.CheckedChanged += new System.EventHandler(this.chkBxRuleActive_CheckedChanged);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(392, 182);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 27);
            this.btnSave.TabIndex = 13;
            this.btnSave.Text = "Save Rule";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // SaveAttachmentsRules
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveBorder;
            this.ClientSize = new System.Drawing.Size(472, 213);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.chkBxRuleActive);
            this.Controls.Add(this.chkBxOverwriteAttWSameName);
            this.Controls.Add(this.chkBxRemoveAttFromMail);
            this.Controls.Add(this.txtFrom);
            this.Controls.Add(this.txtTo);
            this.Controls.Add(this.txtSub);
            this.Controls.Add(this.txtDest);
            this.Controls.Add(this.txtRuleName);
            this.Controls.Add(this.btnDeleteRule);
            this.Controls.Add(this.btnAddRule);
            this.Controls.Add(this.listBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SaveAttachmentsRules";
            this.Text = "SaveAttachmentsRules v1.0b";
            this.Load += new System.EventHandler(this.SaveAttachmentsRules_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btnAddRule;
        private System.Windows.Forms.Button btnDeleteRule;
        private System.Windows.Forms.TextBox txtRuleName;
        private System.Windows.Forms.TextBox txtDest;
        private System.Windows.Forms.TextBox txtSub;
        private System.Windows.Forms.TextBox txtTo;
        private System.Windows.Forms.TextBox txtFrom;
        private System.Windows.Forms.CheckBox chkBxRemoveAttFromMail;
        private System.Windows.Forms.CheckBox chkBxOverwriteAttWSameName;
        private System.Windows.Forms.ToolTip toolTip1;
        private System.Windows.Forms.CheckBox chkBxRuleActive;
        private System.Windows.Forms.Button btnSave;
    }
}