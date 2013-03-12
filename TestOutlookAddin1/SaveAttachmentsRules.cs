using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAttachmentReminder
{
    public partial class SaveAttachmentsRules : Form
    {
        int idx = -1;
        int result = 0x00000000;
        LogMessageToOARDiag.LogMessage cLogMessage = new LogMessageToOARDiag.LogMessage();
        const string OARRuleFile = "OARRules1000b.oar";
        const string OARDIagFIle = "OARDiag1000b.log";
        string SubDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + "OARsFiles";
        private OutlookAttachmentReminder.OARSaveFileRules OARRules = new OARSaveFileRules();
        private string ss;
       
        public SaveAttachmentsRules()
        {
            InitializeComponent();
        }

        private void btnAddRule_Click(object sender, EventArgs e)
        {
            if (txtRuleName.Text == string.Empty)
            {
                MessageBox.Show("Please Enter Rule Name", "Invalid Entry", MessageBoxButtons.OK, MessageBoxIcon.Stop);
                return;
            }
            
            else
            {
                try
                {
                    DirectoryInfo dirInfo = new DirectoryInfo(txtDest.Text + "\\" + txtRuleName.Text);
                    if (!dirInfo.Exists)
                        dirInfo.Create();
                    else
                    {
                        DialogResult mb=MessageBox.Show("Rule with same name exists. Proceeding will replace entries. Proceed?", "Info",MessageBoxButtons.YesNo);

                        switch (mb)
                        {
                            case DialogResult.Yes:
                                foreach (string ss in OARRules.sRuleName)
                                {
                                    if (ss == txtRuleName.Text)
                                    {
                                        System.Predicate<string> pred = new Predicate<string>(findIndx);
                                        int idx1 = OARRules.sRuleName.FindIndex(pred);
                                        OARRules.sDestination[idx1]=txtDest.Text;
                                        OARRules.sFromRule[idx1]=txtFrom.Text;
                                        OARRules.sToRule[idx1]=txtTo.Text;
                                        OARRules.sRuleName[idx1]=txtRuleName.Text;
                                        OARRules.sSubjectRule[idx1]=txtSub.Text;
                                        OARRules.bIsActive[idx1]=(chkBxRuleActive.Checked)? true : false;
                                        OARRules.bOverwriteAttachmentsWSameName[idx1]=(chkBxOverwriteAttWSameName.Checked) ? true : false;
                                        OARRules.bRemoveAttachmentFromMail[idx1]=(chkBxRemoveAttFromMail.Checked) ? true : false;
                                        OARRules.iLowerSize[idx1] = 0;
                                        OARRules.iUpperSize[idx1] = 0;
                                        saveRule();
                                    }
                                }
                                return;
                            case DialogResult.No:
                                MessageBox.Show("Please provide a different name for the rule"); ;
                                break;
                        }

                        return;
                    }
                }
                catch (System.IO.IOException ex)
                {
                 
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);

                    result=cLogMessage.fnLogExceptions(ex, OARDIagFIle);
                    if (result != 0x11111111)
                    {
                        MessageBox.Show("HRESULT = " + result);
                    }
                    result = 0x00000000;
                    
                    return;
                }
            }

            try
            {
                OARRules.sRuleName.Add(txtRuleName.Text);
                OARRules.sDestination.Add(txtDest.Text);
                OARRules.iLowerSize.Add(0);
                OARRules.iUpperSize.Add(0);
                OARRules.sFromRule.Add(txtFrom.Text);
                OARRules.sToRule.Add(txtTo.Text);
                OARRules.sSubjectRule.Add(txtSub.Text);
                OARRules.bOverwriteAttachmentsWSameName.Add(chkBxOverwriteAttWSameName.Checked ? true : false);
                OARRules.bRemoveAttachmentFromMail.Add(chkBxRemoveAttFromMail.Checked ? true : false);
                OARRules.bIsActive.Add(chkBxRuleActive.Checked ? true : false);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);

                result = cLogMessage.fnLogExceptions(ex, OARDIagFIle);
                if (result != 0x11111111)
                {
                    MessageBox.Show("Error logging exception message.HRESULT = " + result + " .Report to developer");
                }
                result = 0x00000000;
                
                return;
            }

            listBox1.Items.Add(OARRules.sRuleName[OARRules.sRuleName.Count-1]);
            saveRule();
        }

        private void saveRule()
        {
            FileStream FS = new FileStream(SubDir + "\\" + OARRuleFile,
            FileMode.OpenOrCreate, FileAccess.Write);

            try
            {
                BinaryFormatter bf = new BinaryFormatter();
                bf.Serialize(FS, OARRules);
                OutlookAttachmentReminder.Globals.OutlookAttachmentReminderAddin.notifyicon.ShowBalloonTip(5000, "Info", "Rule Saved", ToolTipIcon.Info);
                OutlookAttachmentReminder.Globals.OutlookAttachmentReminderAddin.OARRuleCreated = true;
                OutlookAttachmentReminder.Globals.OutlookAttachmentReminderAddin.notifyicon.ShowBalloonTip(5000, "Restart required", "A new rule is saved. This requires restart of outlook to take effect", ToolTipIcon.Info);
 
            }
            catch (System.Exception ex)
            {
                result = cLogMessage.fnLogExceptions(ex, OARDIagFIle);
                if(result != 0x11111111)
                {
                    MessageBox.Show("Error logging exception message.HRESULT = " + result + " .Report to developer");
                }
                result = 0x00000000;

                OutlookAttachmentReminder.Globals.OutlookAttachmentReminderAddin.notifyicon.ShowBalloonTip(5000, "Error", "Error Saving Rule", ToolTipIcon.Error); 
            }
            finally
            {
                FS.Close();
                for (int i = 0; i < OARRules.sRuleName.Count; i++)
                {
                    DirectoryInfo dirInf = new DirectoryInfo(OARRules.sDestination[i] + "\\" + OARRules.sRuleName);
                    if (!dirInf.Exists)
                        dirInf.Create();
                }
            }

        }

        private void SaveAttachmentsRules_Load(object sender, EventArgs e)
        {
            listBox1.Items.Clear();

            //Added this fileInfo to check if rule file exists each time this form starts.
            FileInfo finfo = new FileInfo(SubDir + "\\" + OARRuleFile);
            if (finfo.Exists)
            {
                FileStream flStream = new FileStream(SubDir + "\\" + OARRuleFile,
                    FileMode.Open, FileAccess.Read);
                try
                {
                    BinaryFormatter binFormatter = new BinaryFormatter();
                    OARRules = (OARSaveFileRules)binFormatter.Deserialize(flStream);
                }
                finally
                {
                    flStream.Close();
                }

                foreach (string s in OARRules.sRuleName)
                {
                    listBox1.Items.Add(s);
                }
            }
            else
                return;
        
        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                ss = listBox1.SelectedItem.ToString();
            }
            catch
            {
                return;
            }

            finally
            {
                System.Predicate<string> pred = new Predicate<string>(findIndx);
                idx = OARRules.sRuleName.FindIndex(pred);
                txtDest.Text = OARRules.sDestination[idx];
                txtFrom.Text = OARRules.sFromRule[idx];
                txtTo.Text = OARRules.sToRule[idx];
                txtRuleName.Text = OARRules.sRuleName[idx];
                txtSub.Text = OARRules.sSubjectRule[idx];
                chkBxRuleActive.Checked=(OARRules.bIsActive[idx])?true:false;
                chkBxOverwriteAttWSameName.Checked=(OARRules.bOverwriteAttachmentsWSameName[idx])?true:false;
                chkBxRemoveAttFromMail.Checked=(OARRules.bRemoveAttachmentFromMail[idx])?true:false;
                
                //MessageBox.Show(idx.ToString());
            }
                        
        }

        private bool findIndx(string sss)
        {
            {
                if (sss == ss)
                    return true;
                else
                    return false;
            }
        }

        private void btnDeleteRule_Click(object sender, EventArgs e)
        {
            if (!(listBox1.SelectedIndex == -1))
            {
                OARRules.sRuleName.RemoveAt(idx);
                OARRules.sDestination.RemoveAt(idx);
                OARRules.sToRule.RemoveAt(idx);
                OARRules.sFromRule.RemoveAt(idx);
                OARRules.sSubjectRule.RemoveAt(idx);
                OARRules.iLowerSize.RemoveAt(idx);

                OARRules.iUpperSize.RemoveAt(idx);
                OARRules.bIsActive.RemoveAt(idx);
                OARRules.bOverwriteAttachmentsWSameName.RemoveAt(idx);
                OARRules.bRemoveAttachmentFromMail.RemoveAt(idx);
                listBox1.Items.RemoveAt(idx);

                DirectoryInfo dirInf = new DirectoryInfo(OARRules.sDestination[idx] + "\\" + OARRules.sRuleName[idx]);
                if (dirInf.Exists)
                {
                    dirInf.Delete();
                }
            }
        }

        private void chkBxRuleActive_CheckedChanged(object sender, EventArgs e)
        {
            if(listBox1.SelectedIndex > -1)
                saveRule();
        }

        private void chkBxRemoveAttFromMail_CheckedChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex > -1)
                saveRule();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            saveRule();
        }

        private void chkBxOverwriteAttWSameName_CheckedChanged(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex > -1)
                saveRule();
        }
    }
}
