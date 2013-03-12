using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
namespace OutlookAttachmentReminder
{
  
    public partial class Options : Form
    {
        const string OARDIagFIle = "OARDiag1000RC.log";
        const string OARWordListFile = "OARWordList1000RC.txt";
        const string OARSaveAttachLog = "OARSaveAttachments1000RC.log";
        const string OARRuleFile = "OARRules1000RC.oar";
        const string OARHelpFile = "OARHelpFile1000RC.txt";
        string SubDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + "OARsFiles";

        SaveAttachmentsRules sarWnd = new SaveAttachmentsRules();
        public Options()
        {
            InitializeComponent();

        }


        private void btnAddToWordList_Click(object sender, EventArgs e)
        {
            FileInfo fInfoTmp = new FileInfo(SubDir + "\\" + OARWordListFile+ ".tmp");
            FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);
            TextReader OARSr = new StreamReader(fInfo.FullName);
            string tmpString=string.Empty;
            StreamWriter OARSw = new StreamWriter(fInfoTmp.FullName);

            while ((tmpString = OARSr.ReadLine()) != null)
            {
                OARSw.WriteLine(tmpString);
            }

            OARSw.WriteLine(txtNewWord.Text);
            lstBxWordList.Items.Add(txtNewWord.Text);
            OARSw.Close();
            OARSr.Close();

            fInfo.Delete();
            fInfoTmp.MoveTo(SubDir + "\\" + OARWordListFile);

        }

        private void btnAddToSubjectList_Click(object sender, EventArgs e)
        {
            FileInfo fInfoTmp = new FileInfo(SubDir + "\\" + OARWordListFile + ".tmp");
            FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);
            TextReader OARSr = new StreamReader(fInfo.FullName);
            string tmpString = string.Empty;
            StreamWriter OARSw = new StreamWriter(fInfoTmp.FullName);

            while ((tmpString = OARSr.ReadLine()) != null)
            {
                OARSw.WriteLine(tmpString);
            }

            OARSw.WriteLine("SUB:-:" + txtNewWord.Text);
            lstBxSubject.Items.Add(txtNewWord.Text);
            OARSw.Close();
            OARSr.Close();

            fInfo.Delete();
            fInfoTmp.MoveTo(SubDir + "\\" + OARWordListFile);

        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            FileInfo fInfoTmp = new FileInfo(SubDir + "\\" + OARWordListFile+ ".tmp");
            FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);
            TextReader OARSr = new StreamReader(fInfo.FullName);

            string tmpString = string.Empty;
            StreamWriter OARSw = new StreamWriter(fInfoTmp.FullName);

            bool bIsPresentOPTDISALLOWATTACHMENTS = false;
            bool bIsPresentOPTEMPTYSUBJECT = false;
            bool bIsPresentOPTRESTRICTFILETYPES = false;
            bool bIsPresentOPTAUTOSAVEINCOMING = false;
            bool bIsPresentOPTDELETEATTACHMENTS = false;
            bool bOPTMATCHMODE = false; //indicates PMM

            while ((tmpString = OARSr.ReadLine()) != null)
            {
                if (tmpString.Contains("SIZE:-:"))
                {
                    OARSw.WriteLine("SIZE:-:" + txtSize.Text);
                }
                else if ((tmpString.Contains("OPTDISALLOWATTACHMENTS:-:Yes")))
                {
                    bIsPresentOPTDISALLOWATTACHMENTS = true;
                    if (chkbxDisallowAttachment.Checked)
                        OARSw.WriteLine(tmpString);
                    else
                        OARSw.WriteLine("OPTDISALLOWATTACHMENTS:-:No");
                }
                else if ((tmpString.Contains("OPTDISALLOWATTACHMENTS:-:No")))
                {
                    bIsPresentOPTDISALLOWATTACHMENTS = true;
                    if (chkbxDisallowAttachment.Checked)
                        OARSw.WriteLine("OPTDISALLOWATTACHMENTS:-:Yes");
                    else
                        OARSw.WriteLine(tmpString);
                }

                else if (tmpString.Contains("OPTEMPTYSUBJECT:-:Yes"))
                {
                    bIsPresentOPTEMPTYSUBJECT = true;
                    if (chkbxEmptyMessage.Checked)
                        OARSw.WriteLine(tmpString);
                    else
                        OARSw.WriteLine("OPTEMPTYSUBJECT:-:No");

                }
                else if (tmpString.Contains("OPTEMPTYSUBJECT:-:No"))
                {
                    bIsPresentOPTEMPTYSUBJECT = true;
                    if (chkbxEmptyMessage.Checked)
                        OARSw.WriteLine("OPTEMPTYSUBJECT:-:Yes");
                    else
                        OARSw.WriteLine(tmpString);

                }

                else if (tmpString.Contains("OPTRESTRICTFILETYPES:-:Yes"))
                {
                    bIsPresentOPTRESTRICTFILETYPES = true;
                    if (chkbxRestrictFileTypes.Checked)
                        OARSw.WriteLine(tmpString);
                    else
                        OARSw.WriteLine("OPTRESTRICTFILETYPES:-:No");

                }
                else if (tmpString.Contains("OPTRESTRICTFILETYPES:-:No"))
                {
                    bIsPresentOPTRESTRICTFILETYPES = true;
                    if (chkbxRestrictFileTypes.Checked)
                        OARSw.WriteLine("OPTRESTRICTFILETYPES:-:Yes");
                    else
                        OARSw.WriteLine(tmpString);

                }

                else if (tmpString.Contains("OPTAUTOSAVEINCOMING:-:Yes"))
                {
                    bIsPresentOPTAUTOSAVEINCOMING = true;
                    if (chkBxAutoSaveIncomingAttachments.Checked)
                        OARSw.WriteLine(tmpString);
                    else
                        OARSw.WriteLine("OPTAUTOSAVEINCOMING:-:No");

                }
                else if (tmpString.Contains("OPTAUTOSAVEINCOMING:-:No"))
                {
                    bIsPresentOPTAUTOSAVEINCOMING = true;
                    if (chkBxAutoSaveIncomingAttachments.Checked)
                        OARSw.WriteLine("OPTAUTOSAVEINCOMING:-:Yes");
                    else
                        OARSw.WriteLine(tmpString);

                }

                else if (tmpString.Contains("OPTDELETEATTACHMENTS:-:Yes"))
                {
                    bIsPresentOPTDELETEATTACHMENTS = true;
                    if (chkBxDeleteAttachments.Checked)
                        OARSw.WriteLine(tmpString);
                    else
                        OARSw.WriteLine("OPTDELETEATTACHMENTS:-:No");

                }
                else if (tmpString.Contains("OPTDELETEATTACHMENTS:-:No"))
                {
                    bIsPresentOPTDELETEATTACHMENTS = true;
                    if (chkBxDeleteAttachments.Checked)
                        OARSw.WriteLine("OPTDELETEATTACHMENTS:-:Yes");
                    else
                        OARSw.WriteLine(tmpString);

                }
                else if (tmpString.Contains("OPTMATCHMODE:-:PMM"))
                {
                    bOPTMATCHMODE = true;
                    if (rBtnEMM.Checked)
                        OARSw.WriteLine("OPTMATCHMODE:-:EMM");
                    else
                        OARSw.WriteLine(tmpString);
                }
                else if (tmpString.Contains("OPTMATCHMODE:-:EMM"))
                {
                    bOPTMATCHMODE = true;
                    if (rBtnPMM.Checked)
                        OARSw.WriteLine("OPTMATCHMODE:-:PMM");
                    else
                        OARSw.WriteLine(tmpString);
                }
                else
                    OARSw.WriteLine(tmpString);

            }

            if (!bIsPresentOPTAUTOSAVEINCOMING)
            {
                if (chkBxAutoSaveIncomingAttachments.Checked)
                    OARSw.WriteLine("OPTAUTOSAVEINCOMING:-:Yes");
                else
                    OARSw.WriteLine("OPTAUTOSAVEINCOMING:-:No");
            }
            if (!bIsPresentOPTDELETEATTACHMENTS)
            {
                if (chkBxDeleteAttachments.Checked)
                    OARSw.WriteLine("OPTDELETEATTACHMENTS:-:Yes");
                else
                    OARSw.WriteLine("OPTDELETEATTACHMENTS:-:No");
            }
            if (!bIsPresentOPTDISALLOWATTACHMENTS)
            {
                if (chkbxDisallowAttachment.Checked)
                    OARSw.WriteLine("OPTDISALLOWATTACHMENTS:-:Yes");
                else
                    OARSw.WriteLine("OPTDISALLOWATTACHMENTS:-:No");
            }
            if (!bIsPresentOPTEMPTYSUBJECT)
            {
                if (chkbxEmptyMessage.Checked)
                    OARSw.WriteLine("OPTEMPTYSUBJECT:-:Yes");
                else
                    OARSw.WriteLine("OPTEMPTYSUBJECT:-:No");
            }
            if (!bIsPresentOPTRESTRICTFILETYPES)
            {
                if (chkbxRestrictFileTypes.Checked)
                    OARSw.WriteLine("OPTRESTRICTFILETYPES:-:Yes");
                else
                    OARSw.WriteLine("OPTRESTRICTFILETYPES:-:No");
            }
            if (!bOPTMATCHMODE)
            {
                if (rBtnPMM.Checked)
                    OARSw.WriteLine("OPTMATCHMODE:-:PMM");
                else if(rBtnEMM.Checked)
                    OARSw.WriteLine("OPTMATCHMODE:-:EMM");
            }

            OARSw.Close();
            OARSr.Close();

            fInfo.Delete();
            fInfoTmp.MoveTo(SubDir + "\\" + OARWordListFile);
            OutlookAttachmentReminder.Globals.OutlookAttachmentReminderAddin.notifyicon.ShowBalloonTip(5000, "Settings Saved", "Sucessfully settings are saved", ToolTipIcon.Info);
            
        }

        private void deleteTheWordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lstBxWordList.SelectedIndex > -1)
            {
                FileInfo fInfoTmp = new FileInfo(SubDir + "\\" + OARWordListFile + ".tmp");
                FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);
                TextReader OARSr = new StreamReader(fInfo.FullName);

                int index = lstBxWordList.SelectedIndex;
                
                string tmpString = string.Empty;

                StreamWriter OARSw = new StreamWriter(fInfoTmp.FullName);
                string holderString = lstBxWordList.SelectedItem.ToString();
                while ((tmpString = OARSr.ReadLine()) != null)
                {
                    if (tmpString == holderString)
                    {
                        lstBxWordList.Items.RemoveAt(index);
                    }
                    else
                    {
                        OARSw.WriteLine(tmpString);
                    }
                }

                OARSw.Close();
                OARSr.Close();

                fInfo.Delete();
                fInfoTmp.MoveTo(SubDir + "\\" + OARWordListFile);
            }
        }

        private void deleteTheSubjectToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (lstBxSubject.SelectedIndex > -1)
            {
                FileInfo fInfoTmp = new FileInfo(SubDir + "\\" + OARWordListFile+ ".tmp");
                FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);
                TextReader OARSr = new StreamReader(fInfo.FullName);

                string tmpString = string.Empty;

                StreamWriter OARSw = new StreamWriter(fInfoTmp.FullName);
                int index = lstBxSubject.SelectedIndex;

                string holderString = lstBxSubject.SelectedItem.ToString();
                while ((tmpString = OARSr.ReadLine()) != null)
                {
                    if (tmpString == "SUB:-:" + holderString)
                    {
                        lstBxSubject.Items.RemoveAt(index);
                    }
                    else
                    {
                        OARSw.WriteLine(tmpString);
                    }
                }

                OARSw.Close();
                OARSr.Close();

                fInfo.Delete();
                fInfoTmp.MoveTo(SubDir + "\\" + OARWordListFile);
            }
        }

        private void lnklblFeedback_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start("Iexplore.exe", "http://oar.codeplex.com/Thread/List.aspx");    
        }

        private void chkbxDisallowAttachment_CheckedChanged(object sender, EventArgs e)
        {
            if (chkbxDisallowAttachment.Checked)
            {
                txtSize.Enabled = false;
                chkbxRestrictFileTypes.Enabled = false;
                btnAddFileTypes.Enabled = false;
                lstbxFileTypes.Enabled = false;
            }
            else
            {
                txtSize.Enabled = true;
                chkbxRestrictFileTypes.Enabled = true;
                btnAddFileTypes.Enabled = true;
                lstbxFileTypes.Enabled = true;
            }
        }

        private void btnAddFileTypes_Click(object sender, EventArgs e)
        {
            FileInfo fInfoTmp = new FileInfo(SubDir + "\\" + OARWordListFile + ".tmp");
            FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);
            TextReader OARSr = new StreamReader(fInfo.FullName);
            string tmpString = string.Empty;
            StreamWriter OARSw = new StreamWriter(fInfoTmp.FullName);

            while ((tmpString = OARSr.ReadLine()) != null)
            {
                OARSw.WriteLine(tmpString);
            }

            OARSw.WriteLine("FILET:-:"+txtNewWord.Text);
            lstbxFileTypes.Items.Add(txtNewWord.Text);
            OARSw.Close();
            OARSr.Close();

            fInfo.Delete();
            fInfoTmp.MoveTo(SubDir + "\\" + OARWordListFile);

        }

        private void deleteToolStripMenuItem_Click(object sender, EventArgs e)
        {            
            if ( lstbxFileTypes.SelectedIndex > -1)
            {
                FileInfo fInfoTmp = new FileInfo(SubDir + "\\" + OARWordListFile + ".tmp");
                FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);
                TextReader OARSr = new StreamReader(fInfo.FullName);

                string tmpString = string.Empty;

                StreamWriter OARSw = new StreamWriter(fInfoTmp.FullName);
                int index = lstbxFileTypes.SelectedIndex;

                string holderString = lstbxFileTypes.SelectedItem.ToString();
                while ((tmpString = OARSr.ReadLine()) != null)
                {
                    if (tmpString == "FILET:-:" + holderString)
                    {
                        //Do nothing
                        lstbxFileTypes.Items.RemoveAt(index);
                    }
                    else
                    {
                        OARSw.WriteLine(tmpString);
                    }
                }

                OARSw.Close();
                OARSr.Close();

                fInfo.Delete();
                fInfoTmp.MoveTo(SubDir + "\\" + OARWordListFile);
            }
        }

        private void lstBxWordList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstBxWordList.SelectedIndex > -1)
                txtNewWord.Text = lstBxWordList.SelectedItem.ToString();
        }

        private void lstBxSubject_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstBxSubject.SelectedIndex > -1)
                txtNewWord.Text = lstBxSubject.SelectedItem.ToString();
        }

        private void lstbxFileTypes_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstbxFileTypes.SelectedIndex > -1)
                txtNewWord.Text = lstbxFileTypes.SelectedItem.ToString();
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);

            if (!fInfo.Exists)
            {
                StreamWriter OARSw = fInfo.CreateText();

                //Adding some default words
                OARSw.WriteLine("attach");
                OARSw.WriteLine("attached");
                OARSw.WriteLine("attaching");
                OARSw.WriteLine("attachment");
                OARSw.WriteLine("enclose");
                OARSw.WriteLine("enclosing");
                OARSw.WriteLine("enclosure");

                OARSw.WriteLine("SUB:-:attach");
                OARSw.WriteLine("SUB:-:attached");
                OARSw.WriteLine("SUB:-:attaching");
                OARSw.WriteLine("SUB:-:attachment");
                OARSw.WriteLine("SUB:-:enclose");
                OARSw.WriteLine("SUB:-:enclosing");
                OARSw.WriteLine("SUB:-:enclosure");

                OARSw.WriteLine("SIZE:-:100000");

                //Added in v1.0b
                OARSw.WriteLine("OPTRESTRICTFILETYPES:-:No");
                OARSw.WriteLine("OPTDISALLOWATTACHMENTS:-:No");
                OARSw.WriteLine("OPTEMPTYSUBJECT:-:Yes");
                OARSw.WriteLine("OPTAUTOSAVEINCOMING:-:No");
                OARSw.WriteLine("OPTDELETEATTACHMENTS:-:No");
                OARSw.WriteLine("OPTMATCHMODE:-:PMM");

                OARSw.Close();
            }
            else
            {
                fInfo.Delete();
                StreamWriter OARSw = fInfo.CreateText();

                //Adding some default words
                OARSw.WriteLine("attach");
                OARSw.WriteLine("attached");
                OARSw.WriteLine("attaching");
                OARSw.WriteLine("attachment");
                OARSw.WriteLine("enclose");
                OARSw.WriteLine("enclosing");
                OARSw.WriteLine("enclosure");

                OARSw.WriteLine("SUB:-:attach");
                OARSw.WriteLine("SUB:-:attached");
                OARSw.WriteLine("SUB:-:attaching");
                OARSw.WriteLine("SUB:-:attachment");
                OARSw.WriteLine("SUB:-:enclose");
                OARSw.WriteLine("SUB:-:enclosing");
                OARSw.WriteLine("SUB:-:enclosure");

                OARSw.WriteLine("SIZE:-:100000");

                //Added in v1.0b
                OARSw.WriteLine("OPTRESTRICTFILETYPES:-:No");
                OARSw.WriteLine("OPTDISALLOWATTACHMENTS:-:No");
                OARSw.WriteLine("OPTEMPTYSUBJECT:-:Yes");
                OARSw.WriteLine("OPTAUTOSAVEINCOMING:-:No");
                OARSw.WriteLine("OPTDELETEATTACHMENTS:-:No");
                OARSw.WriteLine("OPTMATCHMODE:-:PMM");

                OARSw.Close();
            }

        }


        private void btnAutoSave_Click(object sender, EventArgs e)
        {
            sarWnd.ShowDialog();
        }

        private void Options_Load(object sender, EventArgs e)
        {
            lstBxSubject.Items.Clear();
            lstBxWordList.Items.Clear();
            lstbxFileTypes.Items.Clear();
            txtSize.Text = "";
            FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);
            TextReader OARSw = new StreamReader(fInfo.FullName);
            string tmpString = string.Empty;
            while ((tmpString = OARSw.ReadLine()) != null)
            {

                if (tmpString == string.Empty || tmpString.Contains("SUB:-:") || tmpString.Contains("SIZE:-:") || tmpString.Contains("FILET:-:")
                    || tmpString.Contains("OPTRESTRICTFILETYPES:-:")
                    || tmpString.Contains("OPTDISALLOWATTACHMENTS:-:")
                    || tmpString.Contains("OPTEMPTYSUBJECT:-:")
                    || tmpString.Contains("OPTAUTOSAVEINCOMING:-:")
                    || tmpString.Contains("OPTDELETEATTACHMENTS:-:")
                    || tmpString.Contains("OPTMATCHMODE:-:")
                    )
                {
                    if (tmpString.Contains("SUB:-:"))
                    {
                        lstBxSubject.Items.Add(tmpString.Substring(6, tmpString.Length - 6));
                    }
                    else if (tmpString.Contains("FILET:-:"))
                    {
                        lstbxFileTypes.Items.Add(tmpString.Substring(8, tmpString.Length - 8));
                    }
                    else if (tmpString.Contains("SIZE:-:"))
                    {
                        txtSize.Text = (tmpString.Substring(7, tmpString.Length - 7));
                    }
                    else if (tmpString.Contains("OPTRESTRICTFILETYPES:-:"))
                    {
                        if (tmpString.Contains("Yes"))
                            chkbxRestrictFileTypes.Checked = true;
                        else
                            chkbxRestrictFileTypes.Checked = false;
                    }
                    else if (tmpString.Contains("OPTDISALLOWATTACHMENTS:-:"))
                    {
                        if (tmpString.Contains("Yes"))
                        {
                            txtSize.Enabled = false;
                            chkbxRestrictFileTypes.Enabled = false;
                            btnAddFileTypes.Enabled = false;
                            lstbxFileTypes.Enabled = false;
                            chkbxDisallowAttachment.Checked = true;

                        }
                        else
                        {
                            txtSize.Enabled = true;
                            chkbxRestrictFileTypes.Enabled = true;
                            btnAddFileTypes.Enabled = true;
                            lstbxFileTypes.Enabled = true;
                            chkbxDisallowAttachment.Checked = false;

                        }
                    
                    }
                    else if (tmpString.Contains("OPTEMPTYSUBJECT:-:"))
                    {
                        if (tmpString.Contains("Yes"))
                            chkbxEmptyMessage.Checked = true;
                        else
                            chkbxEmptyMessage.Checked = false;
                    }
                    else if (tmpString.Contains("OPTAUTOSAVEINCOMING:-:"))
                    {
                        if (tmpString.Contains("Yes"))
                            chkBxAutoSaveIncomingAttachments.Checked = true;
                        else
                            chkBxAutoSaveIncomingAttachments.Checked = false;
                    }
                    else if (tmpString.Contains("OPTDELETEATTACHMENTS:-:"))
                    {
                        if (tmpString.Contains("Yes"))
                            chkBxDeleteAttachments.Checked = true;
                        else
                            chkBxDeleteAttachments.Checked = false;
                    }
                    else if (tmpString.Contains("OPTMATCHMODE:-:"))
                    {
                        if (tmpString.Contains("PMM"))
                            rBtnPMM.Checked = true;
                        else if (tmpString.Contains("EMM"))
                            rBtnEMM.Checked = true;
                    }
                }
                else
                {
                    lstBxWordList.Items.Add(tmpString);
                }
            }

            OARSw.Close();

        }

    }
}
