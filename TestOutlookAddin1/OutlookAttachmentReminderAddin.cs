//Outlook Attachment Reminder 1.0.0.0 Beta in progress
//Last Modified Date : 03/03/2010 PST
//Author: Gourav Das

using System;
using System.Collections.Generic;
using System.Text;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using Microsoft.VisualBasic;
using System.Runtime.Serialization.Formatters.Binary;
using System.Threading;


namespace OutlookAttachmentReminder
{
    //Created this Mail Class for future
    class OARMailItem
    {
        public Outlook.MailItem oMailItem;
        public string OARConversationIndex { get; set; } //introduced in 0.9.7.6

        public OARMailItem()
        {
            oMailItem = null;
            OARConversationIndex = null;
        }
    }

    //added in 0.9.9.8
    [Serializable]
    public class OARSaveFileRules
    {
        public List<string> sRuleName { get; set; }
        public List<string> sDestination { get; set; }
        public List<string> sSubjectRule { get; set; }
        public List<string> sFromRule { get; set; }
        public List<string> sToRule { get; set; }
        public List<bool> bRemoveAttachmentFromMail { get; set;}
        public List<Int64> iLowerSize { get; set; }
        public List<Int64> iUpperSize { get; set; }
        public List<bool> bOverwriteAttachmentsWSameName { get; set; }
        public List<bool> bIsActive { get; set; }

        public OARSaveFileRules()
        {
            sRuleName=new List<string>();
            sDestination = new List<string>();
            sSubjectRule = new List<string>();
            sFromRule = new List<string>();
            sToRule = new List<string>();
            bRemoveAttachmentFromMail = new List<bool>();
            iLowerSize = new List<Int64>();
            iUpperSize = new List<Int64>();
            bOverwriteAttachmentsWSameName = new List<bool>();
            bIsActive = new List<bool>();

        }

    }
    
 
    public partial class OutlookAttachmentReminderAddin
    {
        public NotifyIcon notifyicon = new NotifyIcon();
        public ContextMenuStrip ctxtMenu = new ContextMenuStrip();

        const string OARDIagFIle = "OARDiag1000RC.log";
        const string OARWordListFile = "OARWordList1000RC.txt";
        const string OARSaveAttachLog = "OARSaveAttachments1000RC.log";
        const string OARRuleFile = "OARRules1000RC.oar";
        const string OARHelpFile = "OARHelpFile1000RC.txt";
        string SubDir = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + "OARsFiles";
        
        LogMessageToOARDiag.LogMessage logMessage = new LogMessageToOARDiag.LogMessage();
        Thread bckThreadForSaveByUser;
        Thread bckThreadForSaveByDate;
        
        public bool OARRuleCreated = false;
        Options myForm = new Options();
        OARSaveFileRules OSFR = new OARSaveFileRules();
        int iCounter = 0;
        int iNumOfStringMatchesBody = 0;        
        int iNumOfStringMatchesSub = 0;  // Added this in Build 0.9.7.3 as we are tracking words in subjects separately than in body
        char[] sep = { ' ', '.', ':', ';', '-' };
        Microsoft.Office.Interop.Outlook.MAPIFolder gFolder;

        int iAttachmentThresholdSizeInKB = 100000; //Added in 0.9.7.6

        //removed in 0.9.9.8 as it is redundant. As we already have Try blocks saving the section of code in AddMenu and RemoveMenu
        //bool bActiveWindowPresent = false; //Added in build 0.9.7.3. This will keep track if we have an open window or not (for outlook). If not then we should notify user
        //that the menu would be loaded only when we stop the outlook and start it in normal mode (with Oulook Window)

        private Office.CommandBar activeMenuBar;
        private Office.CommandBarPopup OARMenuBar;
        private Office.CommandBarButton OARNewMailButton;
        private Office.CommandBarButton OARManageWordsButton;
        private Office.CommandBarButton OARReadDiagLogButton;
        private Office.CommandBarButton OARHelpButton;

        private string OARMenuTag = "OARTag";             //This is required to later find the Menu, while removing

        private OARMailItem OARMail = new OARMailItem();
        
        Outlook::MailItem mailItem = null;

        private void OutlookAttachmentReminderAddin_Startup(object sender, System.EventArgs e)
        {
            DirectoryInfo dirInfStartup = new DirectoryInfo(SubDir);
            if (!dirInfStartup.Exists)
                dirInfStartup.Create();

            notifyicon.Text = "OAR Notification area";
            notifyicon.Visible = true;
            notifyicon.Icon = OutlookAttachmentReminder.Properties.Resources.Icon128x128;

            RemoveOARMenuBar();
            AddOARMenuBar();

            this.Application.FolderContextMenuDisplay += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_FolderContextMenuDisplayEventHandler(Application_FolderContextMenuDisplay);
            this.Application.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(Application_ItemSend);
        
            this.Application.ItemLoad += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_ItemLoadEventHandler(Application_ItemLoad);
            this.Application.NewMailEx += new Microsoft.Office.Interop.Outlook.ApplicationEvents_11_NewMailExEventHandler(Application_NewMailEx);
            
            int result = logMessage.fnLogStartUPMessage(OARDIagFIle);
            if (result != 0x11111111)
            {
                MessageBox.Show("HRESULT = " + result);
            }
            result = 0x00000000;
            
            FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);

            if (!fInfo.Exists)
            {
                #region Code to copy data from older version Wordlist file to latest version.

                DirectoryInfo dirInf = new DirectoryInfo(SubDir);
                
                FileInfo[] fInf = dirInf.GetFiles("OARWordList*.txt", SearchOption.AllDirectories);
                FileInfo[] fInfAllFiles = dirInf.GetFiles("OAR*.*", SearchOption.AllDirectories);
                List<FileInfo> fileversion= new List<FileInfo>();
                List<FileInfo> fileversionAllFiles = new List<FileInfo>();
                foreach (FileInfo fl in fInf)
                {
                    //Code to copy details from older file to newer file
                    fileversion.Add(fl);                  
                }

                foreach (FileInfo fl in fInfAllFiles)
                {
                    //Code to copy details from older file to newer file
                    fileversionAllFiles.Add(fl);
                }

                if (fileversion.Count > 0)
                {
                    //fileversion.Sort();
                    FileInfo newestFileName = fileversion[fileversion.Count-1];
                    File.Copy(newestFileName.FullName, fInfo.FullName);
                    //FileInfo fnfo = new FileInfo(newestFileName);
                    notifyicon.ShowBalloonTip(5000, "File Copied", "Data copied from " + newestFileName.Name + " to " + fInfo.Name + ". Old file will be deleted.", ToolTipIcon.Info);
                    try
                    {
                        foreach (FileInfo fin in fileversion)
                            fin.Delete();

                        foreach (FileInfo fin in fileversionAllFiles)
                            fin.Delete();
                    }
                    catch(System.Exception)
                    {
                        notifyicon.ShowBalloonTip(5000, "Error", "Some or any of the old files could not be deleted, please delete is manually.", ToolTipIcon.Warning);
                    
                    }
                    
                }

                #endregion

                else
                {
                    notifyicon.ShowBalloonTip(5000, "Information", "Use Options menu to manage words, Subject and Size of attachment", ToolTipIcon.Info);
                    
                    logMessage.fnCreateWordList(ref fInfo);
                }
            }

            StreamReader OARtr = new StreamReader(SubDir + "\\" + OARWordListFile);
            string S1Size = null;
            while ((S1Size = OARtr.ReadLine()) != null)
            {
                //Inluded the below check in Build 0.9.7.3 for Empty string to avoid treating of 
                //empty lines as a valid string. In the OR condition we have logic where, if the line in the OARWordlist
                //contains SUB:-: then only we check for matches.
                if (S1Size == string.Empty || (!S1Size.Contains("SIZE:-:")))
                {
                    continue;
                }
                else
                {
                    try
                    {
                        iAttachmentThresholdSizeInKB = (Convert.ToInt32(S1Size.Substring(7, S1Size.Length - 7)));
                    }
                    catch
                    {
                        notifyicon.ShowBalloonTip(5000, "Oversized", "Please enter a valid value for Attachment threshold size. Defaulting to 100000 KB", ToolTipIcon.Warning);
                        iAttachmentThresholdSizeInKB = 100000;
                    }
                }
            }
            OARtr.Close();

            try
            {
                FileStream flStream = new FileStream(SubDir + "\\" + OARRuleFile,
                    FileMode.Open, FileAccess.Read);
                try
                {
                    BinaryFormatter binFormatter = new BinaryFormatter();
                    OSFR = (OARSaveFileRules)binFormatter.Deserialize(flStream);
                }
                finally
                {
                    flStream.Close();
                }
                OARRuleCreated = true;

            }
            catch (System.IO.FileNotFoundException fnf)
            {
                notifyicon.ShowBalloonTip(5000, "Tip", "You can create a rule of your own to save incoming attachments from the options window.", ToolTipIcon.Info);
                OARRuleCreated = false;
            }
            catch (System.Exception ex)
            {
                result = logMessage.fnLogExceptions(ex, OARDIagFIle);
                if (result != 0x11111111)
                {
                    MessageBox.Show("HRESULT = " + result);
                }
                result = 0x00000000;
            }

            #region LOad OADWordList into Options Windows
            myForm.Show();//This will call Form_Load for Options Window
            myForm.Hide();
            #endregion

        }


        //Added for autosave of incoming and outgoing attachments
        void Application_NewMailEx(string EntryIDCollection)
        {
            if (myForm.chkBxAutoSaveIncomingAttachments.Checked)
            {
                try
                {
                    //creating a thread for each incoming mail
                    ThreadPool.QueueUserWorkItem(new WaitCallback(fnSaveIncomingAttach), EntryIDCollection);
                    
                }
                catch (System.Exception ex)
                {
                    notifyicon.ShowBalloonTip(5000, "Critical System error", ex.Message, ToolTipIcon.Error);
                }

            }

        }

        //introduced in 0.9.9.0
        void Application_FolderContextMenuDisplay(Microsoft.Office.Core.CommandBar CommandBar, Microsoft.Office.Interop.Outlook.MAPIFolder Folder)
        {
            gFolder = Folder;
            try
            {
                Microsoft.Office.Core.CommandBarControl newCommandBarControl = CommandBar.Controls.Add(Office.MsoControlType.msoControlPopup, missing, missing, 1, true);
                newCommandBarControl.Visible = true;
                newCommandBarControl.Caption = "OAR Attachments";

                Microsoft.Office.Core.CommandBarPopup popup= (Microsoft.Office.Core.CommandBarPopup)newCommandBarControl;
                Microsoft.Office.Core.CommandBar bar = popup.CommandBar;
                Microsoft.Office.Core.CommandBarButton commandbarcontrol = (Microsoft.Office.Core.CommandBarButton )
                    bar.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, missing, missing, 1, true);
                commandbarcontrol.Caption = "Save attachments before a specific date";
                commandbarcontrol.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(commandbarcontrol_Click);

                Microsoft.Office.Core.CommandBarPopup popup1 = (Microsoft.Office.Core.CommandBarPopup)newCommandBarControl;
                Microsoft.Office.Core.CommandBar bar1 = popup1.CommandBar;
                Microsoft.Office.Core.CommandBarButton commandbarcontrol1 = (Microsoft.Office.Core.CommandBarButton )
                    bar1.Controls.Add(Microsoft.Office.Core.MsoControlType.msoControlButton, missing, missing, 1, true);
                commandbarcontrol1.Caption = "Save attachments from certain user";
                commandbarcontrol1.Click += new Microsoft.Office.Core._CommandBarButtonEvents_ClickEventHandler(commandbarcontrol1_Click);

            }
            catch
            {
                MessageBox.Show("Exception while Adding context menus - 0x001");
            }
        }

        //code based on http://msdn.microsoft.com/en-us/library/bb612664.aspx. Modified in 0.9.9.8
        void commandbarcontrol1_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            bckThreadForSaveByUser=new Thread(new ThreadStart(fnSaveAttachmentByUser));
            bckThreadForSaveByUser.IsBackground = true;
            bckThreadForSaveByUser.Start();
        }

        //code based on http://msdn.microsoft.com/en-us/library/bb612664.aspx introduced in 0.9.9.8
        void commandbarcontrol_Click(Microsoft.Office.Core.CommandBarButton Ctrl, ref bool CancelDefault)
        {
            bckThreadForSaveByDate = new Thread(new ThreadStart(fnSaveAttachmentByDate));
            bckThreadForSaveByDate.IsBackground = true;
            bckThreadForSaveByDate.Start();
        }


        void Application_ItemLoad(object item)
        {
            if (item is Outlook::MailItem)
            {
                mailItem = item as Outlook::MailItem;

                ///<summary>
                ///Adding this so that Attachment add is called even when we forward, reply etc..and add attachment before sending (calling send).
                /// Previously since attachment add was only called afte we clicked on send. Now (as it should), it will be called whenever,
                ///we attach file, in a new mail, reply mail etc.
                ///</summary>
                mailItem.AttachmentAdd += new Outlook.ItemEvents_10_AttachmentAddEventHandler(mailItem_AttachmentAdd);
                mailItem.BeforeAttachmentAdd += new Microsoft.Office.Interop.Outlook.ItemEvents_10_BeforeAttachmentAddEventHandler(mailItem_BeforeAttachmentAdd);
                mailItem.AttachmentRemove += new Microsoft.Office.Interop.Outlook.ItemEvents_10_AttachmentRemoveEventHandler(mailItem_AttachmentRemove);
                
            }
        }


        private void AddOARMenuBar()
        {
            ///<Summary> about the immediate next Try catch
            ///After digging into the error reported by http://oar.codeplex.com/WorkItem/View.aspx?WorkItemId=1456, found that if
            ///we have outlook running in Background(without an Outlook window), the Application.ActiveExplorer() returns null. The Outlook has to start with a Window(not as background process) and the Menu would load. If
            ///before the menu is loaded the the Outlook starts as background process, now it prompts user to close the Outlook and start 
            ///normally. I might consider removing this prompt completely in future versions, because the code changes allows things to work normally and hence this prompt woudl be only informational.
            ///</Summary>

            try
            {
                Microsoft.Office.Interop.Outlook.Explorer activeExplorer = this.Application.ActiveExplorer();
                //bActiveWindowPresent = true;
            }
            catch
            {
                //bActiveWindowPresent = false;
                return;
            }

            try
            {
                activeMenuBar = this.Application.ActiveExplorer().CommandBars.ActiveMenuBar;
                OARMenuBar = (Office.CommandBarPopup)activeMenuBar.Controls.Add(Office.MsoControlType.msoControlPopup,
                                                                            missing, missing, missing, false);
                if (OARMenuBar != null)
                {
                    OARMenuBar.Caption = "OAR Addin";
                    OARMenuBar.Tag = OARMenuTag;

                    //Adding Button for Help
                    ////////////////////////////////////////////
                    OARHelpButton = (Office.CommandBarButton)OARMenuBar.Controls.Add
                        (Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    OARHelpButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                    OARHelpButton.Caption = "Help";
                    //OARNewMailButton.ShortcutText = "Ctrl+J";//Not sure how to use this shortcut "easily"
                    OARHelpButton.Tag = "OARHelpBtn";
                    OARHelpButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(OARHelpButton_Click);


                    //Adding Button for Reading Diag Log
                    ////////////////////////////////////
                    OARReadDiagLogButton = (Office.CommandBarButton)OARMenuBar.Controls.Add
                        (Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    OARReadDiagLogButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                    OARReadDiagLogButton.Caption = "Read OAR Logfile";
                    //OARManageWordsButton.ShortcutText = "Ctrl+M+W";//Not sure how to use this shortcut "easily"
                    OARReadDiagLogButton.Tag = "OARReadDiagLogBtnTag";
                    OARReadDiagLogButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(OARReadDiagLogButton_Click);

                    //Adding Button for opening Options window
                    //////////////////////////////////////////
                    OARManageWordsButton = (Office.CommandBarButton)OARMenuBar.Controls.Add
                        (Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    OARManageWordsButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                    OARManageWordsButton.Caption = "Options";
                    //OARManageWordsButton.ShortcutText = "Ctrl+M+W";//Not sure how to use this shortcut "easily"
                    OARManageWordsButton.Tag = "OAROptionsBtnTag";
                    OARManageWordsButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(OAROptionsButton_Click);

                    //Adding Button for New Mail with Attachment
                    ////////////////////////////////////////////
                    OARNewMailButton = (Office.CommandBarButton)OARMenuBar.Controls.Add
                        (Office.MsoControlType.msoControlButton, missing, missing, 1, true);
                    OARNewMailButton.Style = Office.MsoButtonStyle.msoButtonCaption;
                    OARNewMailButton.Caption = "New Mail with Attachment";
                    //OARNewMailButton.ShortcutText = "Ctrl+J";//Not sure how to use this shortcut "easily"
                    OARNewMailButton.Tag = "OARNewMailBtnTag";
                    OARNewMailButton.Click += new Office._CommandBarButtonEvents_ClickEventHandler(OARNewMailButton_Click);


                    OARMenuBar.Visible = true;

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                int result = logMessage.fnLogExceptions(ex, OARDIagFIle);
                if (result != 0x11111111)
                {
                    MessageBox.Show("HRESULT = " + result);
                }
                result = 0x00000000;
                
            }
        }

        private void OARHelpButton_Click(Office.CommandBarButton item, ref bool cancel)
        {

            FileInfo fHelpInfo = new FileInfo(SubDir + "\\" + OARHelpFile);

            if (!fHelpInfo.Exists)
            {
                StreamWriter HelpStrWriter = fHelpInfo.CreateText();

                HelpStrWriter.WriteLine("Outlook Attachment Reminder Suite 1.0.0.0 RC Instructions");
                HelpStrWriter.WriteLine("================================================");
                HelpStrWriter.WriteLine("1) You can manage word list and Subject list using the Options Windows");
                HelpStrWriter.WriteLine("2) If you face any problems, please check in OARDiag log first which is located in My Documents\\OARsFile. Then please report any messages that you see. Ideally you can upload to the Codeplex site");
                HelpStrWriter.WriteLine("3) There is a known issue, that when you forward an email with an attachment already, it will still prompt. I will be working on it slowly. This is not much of a problem as of now.");
                HelpStrWriter.WriteLine("4) With Every Version of OAR Addin, the file name changes too. In earlier versions you were required to delete the files of older version. But starting with 1.0.0.0 RC these are taken care by the application. File names start with OAR. Current version is found in this help menu itself.");
                HelpStrWriter.WriteLine("5) With Size option, you can set the size limit for the attachment. If the attachment size is greater than this threshold, then Addin prompts asking if you want to add it or not. If you click no, the attachment is not added. This size mentioned is in KB. Default is 100000");
                HelpStrWriter.WriteLine("6) You can create rules, based on which Incoming attachments would be saved. If no rule is present, it prompt gently about that everytime Outlook starts.");
                HelpStrWriter.WriteLine("7) You can disallow attachments altogether");
                HelpStrWriter.WriteLine("8) We have option to restrict attachments of specific files only");
                HelpStrWriter.WriteLine("9) Everytime a new rule is created or existing rule is modified, outlook has to be restarted to take the modified rule to take into effect");
                HelpStrWriter.WriteLine("9) In general we can set option to delete attachments after saving using OAR attachments option in the Right-Click context menu");
                HelpStrWriter.WriteLine("9) Any change in other Options does not require restart. But needs to be saved.");
                HelpStrWriter.WriteLine("*************************");
                HelpStrWriter.WriteLine("=========================");

                HelpStrWriter.Close();
            }
            try
            {
                //Nothing Special here. 
                Process.Start("notepad.exe", SubDir + "\\" + OARHelpFile);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);

                int result = logMessage.fnLogExceptions(ex, OARDIagFIle);
                if (result != 0x11111111)
                {
                    MessageBox.Show("HRESULT = " + result);
                }

                result = 0x00000000;
                
            }

        }

        private void OARNewMailButton_Click(Office.CommandBarButton item, ref bool cancel)
        {
            OARMail.oMailItem = (Outlook.MailItem)this.Application.CreateItem(Outlook.OlItemType.olMailItem);

            /// <summary>
            /// Using this below, so that when new mail with attachment is used to create a new mail and we add, 
            /// the counter should be added otherwise attachment_add doesnt get called first time, since the event handler
            /// is created in Item send. So unless send is clicked once and then we add, the attachment doesnt get
            /// called and Icounter is not updated.
            /// </summary>
            OARMail.oMailItem.Subject = null;
            OARMail.oMailItem.Display(true);
            iCounter = 0;
            /// <summary>
            /// Note: All properties of oMailItem is not filled unless the modal comes up. So setting values before true
            /// has no benefit. Setting those after, is not benefecial, since the next the control will come out of 
            /// OARMail.oMailItem.Display only after the window exits. hence set the EntryID and ConvesationIndex
            /// in Application_ItemSend
            /// </summary>
        }

        private void OAROptionsButton_Click(Office.CommandBarButton item, ref bool cancel)
        {

            myForm.ShowDialog();

        }

        private void OARReadDiagLogButton_Click(Office.CommandBarButton item, ref bool cancel)
        {
            try
            {
                //Nothing Special here. 
                Process.Start("notepad.exe", SubDir + "\\" + OARDIagFIle);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);

                int result = logMessage.fnLogExceptions(ex,OARDIagFIle);
                if (result != 0x11111111)
                {
                    MessageBox.Show("HRESULT = " + result);
                }
                result = 0x00000000;

            }
        }

        // If the menu already exists, remove it.
        private void RemoveOARMenuBar()
        {
            ///<Summary> about the immediate next Try catch
            ///After digging into the error reported by http://oar.codeplex.com/WorkItem/View.aspx?WorkItemId=1456, found that if
            ///we have outlook running in Background(without an Outlook window), the Application.ActiveExplorer() returns null. The Outlook has to start with a Window(not as background process) and the Menu would load. If
            ///before the menu is loaded the the Outlook starts as background process, now it prompts user to close the Outlook and start 
            ///normally. I might consider removing this prompt completely in future versions, because the code changes allows things to work normally and hence this prompt woudl be only informational.
            ///</Summary>
            try
            {
                Microsoft.Office.Interop.Outlook.Explorer eActiveExplorer = this.Application.ActiveExplorer();
                //bActiveWindowPresent = true;
            }
            catch
            {
                //bActiveWindowPresent = false;
                MessageBox.Show(@"We have Outlook running but do not have a window for it (may be its running in background). Due to this, the Menu would not be created. You need to stop the Outlook and restart so that we have the Outlook window open. Only after that, you can use the Menu in the Outlook. \n Note this is just an informational message in Beta builds. Final builds might not have this prompt.", "Information", MessageBoxButtons.OK);
                return;
            }

            try
            {
                Office.CommandBarPopup foundMenu = (Office.CommandBarPopup)
                    this.Application.ActiveExplorer().CommandBars.ActiveMenuBar.
                    FindControl(Office.MsoControlType.msoControlPopup,
                    missing, OARMenuTag, true, true);
                if (foundMenu != null)
                {
                    foundMenu.Delete(true);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                int result = logMessage.fnLogExceptions(ex,OARDIagFIle);
                if (result != 0x11111111)
                {
                    MessageBox.Show("HRESULT = " + result);
                }
                result = 0x00000000;
                
            }
        }

        void Application_ItemSend(object Item, ref bool cancel)
        {
            try
            {
                OARMail.OARConversationIndex = OARMail.oMailItem.ConversationIndex;
            }
            catch
            {
                //Ignore
                //This is a very dirty way, since we know for every reply/forward for a created mail with "New mail w attachment", OARMailItem is 
                //non-existent, hence for every new mail using the addin new mail will not hit exception, but
                //all replies to those parent messages will hit Exception.
                OARMail.OARConversationIndex = string.Empty;

            }

            //Added this IF Block in 0.9.9.8 to get rid of HRResult 0x80040108 which we got normally trying to respond to meeting requests and appointments
            if (Item is Outlook::MailItem)
            {
                //Added in 1.0.0.0b. 
                mailItem = Item as Outlook::MailItem;

                try
                {
                    //If the mail we are sending is opened through the Add in "New mail" button then...
                    //We are also checking if this is the main email and not a reply mail. if same email, then the
                    //conversation index for both messages will match.
                    //If we are replying, then we will hit the exception above and the value would be equal Empty set above and below
                    //would never equate to true

                    if (mailItem.ConversationIndex == OARMail.OARConversationIndex)
                    {
                        if (iCounter == 0)
                        {
                            System.Windows.Forms.DialogResult mb;
                            mb = MessageBox.Show("Did you forget to add an attachment?", "OAR - Missing Attachment", MessageBoxButtons.YesNo);
                            switch (mb)
                            {
                                case DialogResult.Yes:
                                    cancel = true;
                                    break;
                                case DialogResult.No:
                                    cancel = false;
                                    break;
                            }

                        }

                    }
                    else
                    {

                        int iNumOfCharofBody = 0;
                        int iNumOfCharofSub = 0;

                        //I introduced the below fix :) in Build 0.9.5.0. The below Try/Catch is added to get rid of the issue
                        //when we vote and Body of Email is Null and we get an exception in the below statement.
                        try
                        {
                            iNumOfCharofBody = mailItem.Body.IndexOf("From: ") - 9; //Number of characters in the email we are typing
                            //there will be a problem with "From:" exists in our current email :)
                            //Subtracting 9, because this function adds 8.
                        }
                        catch
                        {
                            //Another dirty way of circumventing around the problem when we 
                            //just reply a vote. The problem occurs when we only vote (and not reply).
                            //In that case, the body is NULL and the above mailItem.Body.IndexOf("From: ") - 9;
                            //would face an exception. Since the body is null, no point going through all the other processing
                            //and hence we get out of this by return;

                            return;
                        }

                        //This new feature for scanning subject was introduced in 0.9.7.0
                        try
                        {
                            iNumOfCharofSub = mailItem.Subject.Length;
                        }
                        catch
                        {
                            if (myForm.chkbxEmptyMessage.Checked)
                            {
                                DialogResult mb;
                                mb = MessageBox.Show("This has an empty Subject line. Do you want to continue?", "OAR - Empty Subject Warning", MessageBoxButtons.YesNo);
                                switch (mb)
                                {
                                    case DialogResult.Yes:
                                        cancel = false;
                                        break;
                                    case DialogResult.No:
                                        cancel = true;
                                        return;
                                    //break;
                                }
                            }
                            //Inserting a temporary Blank subject to be removed later This has 10 spaces.
                            mailItem.Subject = "          ";
                        }

                        if (iNumOfCharofBody <= 0)
                            iNumOfCharofBody = mailItem.Body.Length;//If we dont do this, for new message (when we are not replying), we will face error
                        //as iNumOfChar would be -ve

                        //All the following File operations, because we want to read list from file, which can be 
                        //maintained by user.
                        FileInfo fInfo = new FileInfo(SubDir + "\\" + OARWordListFile);

                        if (!fInfo.Exists)
                        {
                            MessageBox.Show("Use Options menu to manage words, Subject and Size of attachment", "Information", MessageBoxButtons.OK);
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

                        iNumOfStringMatchesBody = 0; //Ensuring this resets, everytime we enter here
                        iNumOfStringMatchesSub = 0; //Ensuring this resets too.

                        string matchedSubWords = null;
                        string matchedBodyWords = null;

                        TextReader OARtr = new StreamReader(fInfo.FullName);
                        string S1Body = null;
                        string S1Subject = null;

                        //bool bAdvancedMatch = false;

                        while ((S1Body = OARtr.ReadLine()) != null)
                        {
                            //Inluded the below check in Build 0.9.7.3 for Empty string to avoid treating of 
                            //empty lines as a valid string. In the OR condition we have logic where, if the line in the OARWordlist
                            //contains SUB:-: then we do "not" check for matches.
                            if (S1Body == string.Empty || S1Body.Contains("SUB:-:") || S1Body.Contains("SIZE:-:") || S1Body.Contains("FILET:-:"))
                            {
                                continue;
                            }

                            string[] sWordsInBody = (mailItem.Body.Substring(0, iNumOfCharofBody)).ToLower().Split(sep);
                            string[] sWordsInOARList = S1Body.Split(sep);

                            int iBodyWordsCtr = sWordsInBody.Length - 1;
                            int iOARListWordsCtr = sWordsInOARList.Length - 1;

                            //Added the below IF Block for 0.9.9.0
                            if (!myForm.rBtnEMM.Checked)
                            {
                                if (mailItem.Body.Substring(0, iNumOfCharofBody).ToLower().Contains(S1Body.ToLower()))
                                {
                                    iNumOfStringMatchesBody++;
                                    matchedBodyWords = S1Body + ", " + matchedBodyWords;
                                }
                            }
                            else //EMM Mode
                            {
                                if (sWordsInOARList.Length >= 2)
                                {
                                    for (int counter = 0; counter <= iBodyWordsCtr; counter++)
                                    {
                                        int x = 0;
                                        int tmpcounter = counter;
                                        while (x <= iOARListWordsCtr && tmpcounter <= iBodyWordsCtr)
                                        {
                                            if (sWordsInOARList[x] == sWordsInBody[tmpcounter])
                                            {
                                                //bAdvancedMatch = true;
                                                iNumOfStringMatchesBody++;
                                                matchedBodyWords = sWordsInBody[tmpcounter] + ", " + matchedBodyWords;
                                                if (x < iOARListWordsCtr && tmpcounter < iBodyWordsCtr)
                                                {
                                                    x++;
                                                    tmpcounter++;
                                                    continue;
                                                }
                                            }
                                            else
                                            {
                                                iNumOfStringMatchesBody = 0; ;
                                                matchedBodyWords = string.Empty;
                                                //bAdvancedMatch = false;
                                                break;
                                            }
                                            break;
                                        }
                                    }
                                }
                                else
                                {
                                    for (int counter = 0; counter <= iBodyWordsCtr; counter++)
                                    {
                                        for (int counter1 = 0; counter1 <= iOARListWordsCtr; counter1++)
                                        {
                                            if (sWordsInBody[counter] == sWordsInOARList[counter1])
                                            {
                                                iNumOfStringMatchesBody++;
                                                matchedBodyWords = sWordsInBody[counter] + ", " + matchedBodyWords;

                                                //Trying to make the matched words to bold :(
                                                //int index = mailItem.Body.IndexOf(sWordsInBody[counter]);
                                                //mailItem.Body = mailItem.Body.Replace((mailItem.Body.Substring(index, sWordsInBody[counter].Length)), "\b" + (mailItem.Body.Substring(index, sWordsInBody[counter].Length)) + "\b");

                                            }
                                        }
                                    }
                                }
                            }
                        }

                        OARtr.Close(); //Closing the filestream
                        OARtr = new StreamReader(fInfo.FullName);

                        while ((S1Subject = OARtr.ReadLine()) != null)
                        {
                            //Inluded the below check in Build 0.9.7.3 for Empty string to avoid treating of 
                            //empty lines as a valid string. In the OR condition we have logic where, if the line in the OARWordlist
                            //contains SUB:-: then only we check for matches.
                            if (S1Subject == string.Empty || (!S1Subject.Contains("SUB:-:")))
                            {
                                continue;
                            }

                            string[] sWordsInSubject = (mailItem.Subject.Substring(0, iNumOfCharofSub)).ToLower().Split(sep);
                            string[] sSubInOARList = (S1Subject.Substring(6, S1Subject.Length - 6)).ToLower().Split(sep);

                            int iSubjectWordsCtr = sWordsInSubject.Length - 1;
                            int iOARListSubCtr = sSubInOARList.Length - 1;

                            if (!myForm.rBtnEMM.Checked)
                            {

                                if ((mailItem.Subject.ToLower()).Contains(S1Subject.Substring(6, S1Subject.Length - 6)))
                                {

                                    iNumOfStringMatchesSub++;
                                    matchedSubWords = S1Subject.Substring(6, S1Subject.Length - 6) + ", " + matchedSubWords;
                                }
                            }
                            else
                            {

                                if (sSubInOARList.Length >= 2)
                                {

                                    for (int counter = 0; counter <= iSubjectWordsCtr; counter++)
                                    {

                                        int tmpcounter = counter;
                                        int x = 0;
                                        while (x <= iOARListSubCtr && tmpcounter <= iSubjectWordsCtr)
                                        {

                                            if (sSubInOARList[x] == sWordsInSubject[tmpcounter])
                                            {

                                                //bAdvancedMatch = true;
                                                iNumOfStringMatchesSub++;
                                                matchedSubWords = sWordsInSubject[tmpcounter] + ", " + matchedSubWords;

                                                if (x < iOARListSubCtr && tmpcounter < iSubjectWordsCtr)
                                                {

                                                    x++;
                                                    tmpcounter++;
                                                    continue;
                                                }

                                            }
                                            else
                                            {

                                                iNumOfStringMatchesSub = 0;
                                                matchedSubWords = string.Empty;
                                                //bAdvancedMatch = false;
                                                break;
                                            }
                                            break;
                                        }

                                    }
                                }
                                else
                                {

                                    for (int counter = 0; counter <= iSubjectWordsCtr; counter++)
                                    {
                                        for (int counter1 = 0; counter1 <= iOARListSubCtr; counter1++)
                                        {
                                            if (sWordsInSubject[counter] == (sSubInOARList[counter1]))
                                            {
                                                iNumOfStringMatchesSub++;
                                                matchedSubWords = sWordsInSubject[counter] + ", " + matchedSubWords;
                                            }
                                        }
                                    }
                                }

                            }
                        }

                        OARtr.Close(); //Closing the filestream

                        if (iNumOfStringMatchesBody > 0 || iNumOfStringMatchesSub > 0)
                        {
                            if (iCounter == 0)
                            {
                                System.Windows.Forms.DialogResult mb;
                                string message = "Did you forget to add an attachment? " + "\n\n" + iNumOfStringMatchesBody + " nos of match in body." + "\n-->" + matchedBodyWords + "\n\n" + iNumOfStringMatchesSub + " nos of match in subject." + "\n-->" + matchedSubWords;

                                mb = MessageBox.Show(message, "MISSING ATTACHMENT", MessageBoxButtons.YesNo);
                                switch (mb)
                                {
                                    case DialogResult.Yes:
                                        cancel = true;

                                        break;
                                    case DialogResult.No:
                                        cancel = false;
                                        break;
                                }

                            }

                        }

                    }
                    if (mailItem.Subject == "          ")
                        mailItem.Subject = null;
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK);
                    int result = logMessage.fnLogExceptions(ex,OARDIagFIle);
                    if (result != 0x11111111)
                    {
                        MessageBox.Show("HRESULT = " + result);
                    }
                    result = 0x00000000;
                    
                }

                iCounter = 0;//To ensure that the number of attachments is set to 0. Else once you increment counter, forr
                //all future checks icounter will not enter the above loop
                //}
            }
        }

        void mailItem_BeforeAttachmentAdd(Microsoft.Office.Interop.Outlook.Attachment Attachment, ref bool Cancel)
        {
            if (myForm.chkbxDisallowAttachment.Checked)
            {
                MessageBox.Show("You have restricted sending of attachments. You can change it through the options","Information");
                Cancel = true;
            }
            else
            {
                if (!myForm.chkbxRestrictFileTypes.Checked)
                {
                    #region Code to retrieve Attachment size at runtime
                    StreamReader OARtr = new StreamReader(SubDir + "\\" + OARWordListFile);
                    string S1Size = null;
                    while ((S1Size = OARtr.ReadLine()) != null)
                    {
                        //Inluded the below check in Build 0.9.7.3 for Empty string to avoid treating of 
                        //empty lines as a valid string. In the OR condition we have logic where, if the line in the OARWordlist
                        //contains SUB:-: then only we check for matches.
                        if (S1Size == string.Empty || (!S1Size.Contains("SIZE:-:")))
                        {
                            continue;
                        }
                        else
                        {
                            try
                            {
                                iAttachmentThresholdSizeInKB = (Convert.ToInt32(S1Size.Substring(7, S1Size.Length - 7)));
                            }
                            catch
                            {
                                notifyicon.ShowBalloonTip(5000, "Oversized", "Please enter a valid value for Attachment threshold size. Defaulting to 100000 KB", ToolTipIcon.Warning);
                                iAttachmentThresholdSizeInKB = 100000;
                            }
                            //matchedSubWords = S1Subject.Substring(6, S1Subject.Length - 6) + ", " + matchedSubWords;
                        }
                    }
                    OARtr.Close();
                    #endregion

                    System.Windows.Forms.DialogResult mb;
                    if (Attachment.Size / 1024 > iAttachmentThresholdSizeInKB)
                    {

                        mb = MessageBox.Show("Attachment Size is " + Attachment.Size / 1024 + " KB. Do you still want to send it?", "OverSized", MessageBoxButtons.YesNo);
                        switch (mb)
                        {
                            case DialogResult.Yes:
                                Cancel = false;
                                break;
                            case DialogResult.No:
                                Cancel = true;
                                break;
                        }
                    }
                }
                else
                {
                    string fileExtension = string.Empty;
                    int count = myForm.lstbxFileTypes.Items.Count - 1;
                    string []extension=Attachment.FileName.Split('.');
                    for (int counter = 0; counter <= count; counter++)
                    {
                        fileExtension = myForm.lstbxFileTypes.Items[counter].ToString();
                        if(extension[extension.Length-1]==fileExtension)
                        {
                            MessageBox.Show("Files of type " + fileExtension + " are disabled. You can use options window to make changes", "Information");
                            Cancel = true;
                        }
                    }
                }
            }
        }

        void mailItem_AttachmentAdd(Outlook.Attachment attach)
        {
            if (attach.FileName != null)
            {
                iCounter++;
            }
        }

        void oMailItem_BeforeAttachmentAdd(Microsoft.Office.Interop.Outlook.Attachment Attachment, ref bool Cancel)
        {
            if (myForm.chkbxDisallowAttachment.Checked)
            {
                MessageBox.Show("You have restricted sending of attachments. You can change it through the options","Information");
                Cancel = true;
            }
            else
            {
                if (!myForm.chkbxRestrictFileTypes.Checked)
                {
                    System.Windows.Forms.DialogResult mb;

                    #region Code to retrieve Attachment size limit at runtime
                    StreamReader OARtr = new StreamReader(SubDir + "\\" + OARWordListFile);
                    string S1Size = null;
                    while ((S1Size = OARtr.ReadLine()) != null)
                    {
                        //Inluded the below check in Build 0.9.7.3 for Empty string to avoid treating of 
                        //empty lines as a valid string. In the OR condition we have logic where, if the line in the OARWordlist
                        //contains SUB:-: then only we check for matches.
                        if (S1Size == string.Empty || (!S1Size.Contains("SIZE:-:")))
                        {
                            continue;
                        }
                        else
                        {
                            try
                            {
                                iAttachmentThresholdSizeInKB = (Convert.ToInt32(S1Size.Substring(7, S1Size.Length - 7)));
                            }
                            catch
                            {
                                notifyicon.ShowBalloonTip(5000, "Oversized", "Please enter a valid value for Attachment threshold size. Defaulting to 100000 KB", ToolTipIcon.Warning);
                                iAttachmentThresholdSizeInKB = 100000;
                            }
                            //matchedSubWords = S1Subject.Substring(6, S1Subject.Length - 6) + ", " + matchedSubWords;
                        }
                    }
                    OARtr.Close();
                    #endregion

                    if (Attachment.Size / 1024 > iAttachmentThresholdSizeInKB)
                    {

                        mb = MessageBox.Show("Attachment Size is " + Attachment.Size / 1024 + " KB. Do you still want to send it?", "OverSized", MessageBoxButtons.YesNo);
                        switch (mb)
                        {
                            case DialogResult.Yes:
                                Cancel = false;
                                break;
                            case DialogResult.No:
                                Cancel = true;
                                break;
                        }
                    }
                }
                else
                {
                    string fileExtension = string.Empty;
                    int count = myForm.lstbxFileTypes.Items.Count - 1;
                    string[] extension = Attachment.FileName.Split('.');
                    for (int counter = 0; counter < count; counter++)
                    {
                        fileExtension = myForm.lstbxFileTypes.Items[counter].ToString();
                        if (extension[extension.Length - 1] == fileExtension)
                        {
                            MessageBox.Show("Files of type " + fileExtension + " are disabled. You can use options window to make changes", "Information");
                            Cancel = true;
                        }
                    }
                }
            }
        }

        void oMailItem_AttachmentAdd(Microsoft.Office.Interop.Outlook.Attachment Attachment)
        {
            if (Attachment.FileName != null)
            {
                iCounter++;
            }
        }

        void mailItem_AttachmentRemove(Microsoft.Office.Interop.Outlook.Attachment Attachment)
        {
            if(iCounter>0)
            iCounter--;
        }

        void oMailItem_AttachmentRemove(Microsoft.Office.Interop.Outlook.Attachment Attachment)
        {
            if(iCounter>0)
            iCounter--;
        }

        private void OutlookAttachmentReminderAddin_Shutdown(object sender, System.EventArgs e)
        {
            notifyicon.Dispose();
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(OutlookAttachmentReminderAddin_Startup);
            this.Shutdown += new System.EventHandler(OutlookAttachmentReminderAddin_Shutdown);
        }

        #endregion

        private void fnSaveAttachmentByUser()
        {
            string sPath = Interaction.InputBox("Please enter a valid path to save the files. This location should already exist", "Destination Path", "", 0, 0);
            if (sPath == "")
                return;
            string sSender = Interaction.InputBox("Please enter the sender name.. Leave empty for downloading all attachments", "User Name", "", 0, 0);
            sPath = sPath + "\\" + gFolder.Name;

            FileInfo fInfoDir = new FileInfo(sPath);
            if (!fInfoDir.Exists)
            {
                try
                {
                    Directory.CreateDirectory(sPath);
                }
                catch
                {
                    MessageBox.Show("Problems Creating the Sub Directory " + gFolder.Name);
                    return;
                }
            }

            Microsoft.Office.Interop.Outlook.Items items = gFolder.Items;
            
            FileInfo fInfo = new FileInfo(sPath + "\\" + OARSaveAttachLog);

            #region Much Faster search

            const string PR_HAS_ATTACH = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B";
            const string PR_SENDER_NAME = "http://schemas.microsoft.com/mapi/proptag/0x0C1A001E";

            // Create filter
            string filter = "@SQL=" + "\""
            + PR_HAS_ATTACH + "\"" + " = 1" +
            " AND " + "\"" + PR_SENDER_NAME + "\"" + " LIKE '%" + sSender + "%'";

            Outlook.Table table = gFolder.GetTable(filter, Outlook.OlTableContents.olUserItems);

            // Remove default columns
            table.Columns.RemoveAll();
            table.Columns.Add("EntryID");

            int counter=0;
            while (!table.EndOfTable)
            {
                counter++;
                Outlook.Row nextRow = table.GetNextRow();
                String sb = nextRow["EntryID"].ToString();
                Outlook.NameSpace outlookNS = this.Application.GetNamespace("MAPI");
                Outlook.MAPIFolder mFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

                Outlook::MailItem newMail;

                if (outlookNS.GetItemFromID(sb, gFolder.StoreID) is Outlook.MailItem)
                {
                    try
                    {
                        newMail = (Outlook::MailItem)outlookNS.GetItemFromID(sb, gFolder.StoreID);
                    }
                    catch
                    {
                        counter--;
                        continue;
                    }
                }
                else
                {
                    counter--;
                    continue;
                }

                for (int counter1 = 1; counter1 <= newMail.Attachments.Count; counter1++)
                {
                    try
                    {
                        if (!Directory.Exists(sPath + "\\" + newMail.SenderName))
                        {
                            try
                            {
                                Directory.CreateDirectory(sPath + "\\" + newMail.SenderName);
                            }
                            catch (System.Exception ex)
                            {
                                MessageBox.Show(ex.Message + " - " + newMail.CreationTime.ToString());
                                //return;
                            }
                        }

                        Outlook.Attachment attachment = newMail.Attachments[counter1];
                        attachment.SaveAsFile(sPath + "\\" + newMail.SenderName + "\\" + attachment.FileName);
                        
                        notifyicon.ShowBalloonTip(1000, "Save", "Saved - " + attachment.FileName + " ." + (table.GetRowCount() - counter).ToString() + "emails remaining.", ToolTipIcon.Info);

                        if (myForm.chkBxDeleteAttachments.Checked)
                        {
                            attachment.Delete();
                        }

                    }

                    #region Section to write failures to OARSaveAttachments
                    catch (System.Exception exption)
                    {
                        logMessage.fnLogSaveAttachment(ref fInfo, exption);
                    }
                    #endregion
                }

            }

            #endregion

            MessageBox.Show("Attachments saved to " + sPath + " . Log also saved there");

        }

        private void fnSaveAttachmentByDate()
        {
            string sPath = Interaction.InputBox("Please enter a valid path to save the files. This location should already exist", "Destination Path", "", 0, 0);
            DateTime dDateTime = Convert.ToDateTime(Interaction.InputBox("Please enter the date in form of mm/dd/yyyy only.", "User Name", "", 0, 0));
            if (sPath == "")
                return;
            sPath = sPath + "\\" + gFolder.Name;

            FileInfo fInfoDir = new FileInfo(sPath);
            if (!fInfoDir.Exists)
            {
                try
                {
                    Directory.CreateDirectory(sPath);
                }
                catch
                {
                    MessageBox.Show("Problems Creating the Sub Directory " + gFolder.Name);
                    return;
                }
            }

            Microsoft.Office.Interop.Outlook.Items items = gFolder.Items;

            FileInfo fInfo = new FileInfo(sPath + "\\" + OARSaveAttachLog);

            #region Much Faster search

            const string PR_HAS_ATTACH = "http://schemas.microsoft.com/mapi/proptag/0x0E1B000B";
            const string DATE_RECIEVED = "urn:schemas:httpmail:datereceived";

            // Create filter
            string filter = "@SQL=" + "\""
            + PR_HAS_ATTACH + "\"" + " = 1" +
            " AND " + "\"" + DATE_RECIEVED + "\"" + " < '" + dDateTime + "'";

            Outlook.Table table = gFolder.GetTable(filter, Outlook.OlTableContents.olUserItems);

            // Remove default columns
            table.Columns.RemoveAll();
            table.Columns.Add("EntryID");

            int counter=0;
            while (!table.EndOfTable)
            {
                counter++;
                Outlook.Row nextRow = table.GetNextRow();
                String sb = nextRow["EntryID"].ToString();
                Outlook.NameSpace outlookNS = this.Application.GetNamespace("MAPI");
                Outlook.MAPIFolder mFolder = this.Application.Session.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
                Outlook::MailItem newMail;
                if (outlookNS.GetItemFromID(sb, gFolder.StoreID) is Outlook.MailItem)
                {
                    try
                    {
                        newMail = (Outlook::MailItem)outlookNS.GetItemFromID(sb, gFolder.StoreID);
                    }
                    catch
                    {
                        continue;
                    }
                }
                else
                {
                    continue;
                }
                for (int counter1 = 1; counter1 <= newMail.Attachments.Count; counter1++)
                {
                    try
                    {
                        if (!Directory.Exists(sPath + "\\" + newMail.SenderName))
                        {
                            try
                            {
                                Directory.CreateDirectory(sPath + "\\" + newMail.SenderName);
                            }
                            catch (System.Exception ex)
                            {
                                MessageBox.Show(ex.Message + " - " + newMail.SenderName);
                                //return;
                            }
                        }

                        Outlook.Attachment attachment = newMail.Attachments[counter1];
                        attachment.SaveAsFile(sPath + "\\" + newMail.SenderName + "\\" + attachment.FileName);
                        
                        notifyicon.ShowBalloonTip(1000, "Save", "Saved - " + attachment.FileName + " ." + (table.GetRowCount() - counter).ToString() + "emails remaining.", ToolTipIcon.Info);
                        
                        if (myForm.chkBxDeleteAttachments.Checked)
                        {
                            attachment.Delete();
                        }

                    }

                    #region Section to write failures to OARSaveAttachments
                    catch (System.Exception exption)
                    {
                        logMessage.fnLogSaveAttachment(ref fInfo, exption);
                    }
                    #endregion
                }

            }

            #endregion

            MessageBox.Show("Attachments saved to " + sPath + " . Log also saved there");

        }

        private void fnSaveIncomingAttach(object EntryIDCollection)
        {
            Outlook.NameSpace outlookNS = this.Application.GetNamespace("MAPI");
            Outlook.Folder mFolder = (Outlook.Folder)outlookNS.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);

            if (myForm.chkBxAutoSaveIncomingAttachments.Checked && OARRuleCreated)
            {
                char separator = ';';
                string[] subjectRule;
                string[] recipientRule;
                string[] FromRule;

                for (int ctr = 0; ctr < OSFR.sRuleName.Count; ctr++)
                {
                    if (!OSFR.bIsActive[ctr])
                        continue;

                    DirectoryInfo dirInfo = new DirectoryInfo(OSFR.sDestination[ctr] + "\\" + OSFR.sRuleName[ctr]);
                    if (!dirInfo.Exists)
                    {
                        notifyicon.ShowBalloonTip(5000,"OAR Notification - Error 0x002","Looks like a problem while adding rule as the directory with name of the rule is not created and hence manifesting here. Dump Rules to check their validity", ToolTipIcon.Error);
                        
                        continue;
                    }
                    subjectRule = OSFR.sSubjectRule[ctr].Split(separator);
                    recipientRule = OSFR.sToRule[ctr].Split(separator);
                    FromRule = OSFR.sFromRule[ctr].Split(separator);

                    try
                    {
                        if (outlookNS.GetItemFromID(EntryIDCollection.ToString(), mFolder.StoreID) is Outlook.MailItem);
                    }
                    catch
                    {
                        return;
                    }

                    if (outlookNS.GetItemFromID(EntryIDCollection.ToString(), mFolder.StoreID) is Outlook.MailItem)
                    {
                        Outlook::MailItem newMail = (Outlook::MailItem)outlookNS.GetItemFromID(EntryIDCollection.ToString(), mFolder.StoreID);

                        bool bDwnldAttachment = false;

                        if (FromRule.Length == 0)
                            bDwnldAttachment = true;
                        else if (FromRule[0].Length == 0)
                            bDwnldAttachment = true;
                        else
                        {
                            foreach (string s in FromRule)
                            {
                                if (newMail.SenderName.Contains(s))
                                    bDwnldAttachment = true;
                                else
                                {
                                    if (bDwnldAttachment == true)
                                        bDwnldAttachment = true;
                                    else
                                    {
                                        bDwnldAttachment = false;
                                        return;
                                    }
                                }
                            }
                        }

                        if (subjectRule.Length == 0)
                            bDwnldAttachment = true;
                        else if (subjectRule[0].Length == 0)
                            bDwnldAttachment = true;
                        else
                        {
                            foreach (string s in subjectRule)
                            {
                                string tt = newMail.Subject;
                                
                                if (newMail.Subject.Contains(s))
                                    bDwnldAttachment = true;
                                else
                                {
                                    if (bDwnldAttachment == true)
                                        bDwnldAttachment = true;
                                    else
                                    {
                                        bDwnldAttachment = false;
                                        return;
                                    }

                                }
                            }
                        }

                        if (recipientRule.Length == 0)
                            bDwnldAttachment = true;
                        else if (recipientRule[0].Length == 0)
                            bDwnldAttachment = true;
                        else
                        {
                            foreach (string s in recipientRule)
                            {
                                if (newMail.To.Contains(s))
                                    bDwnldAttachment = true;
                                else
                                {
                                    if (bDwnldAttachment == true)
                                        bDwnldAttachment = true;
                                    else
                                    {
                                        bDwnldAttachment = false;
                                        return;
                                    }

                                }
                            }
                        }

                        if (bDwnldAttachment)
                        {
                            Outlook.MailItem tmpMail=newMail;
                            if (tmpMail.Attachments.Count < 1)
                                continue;

                            notifyicon.ShowBalloonTip(5000, "Attachment Count", "There are " + tmpMail.Attachments.Count.ToString() + " attachments", ToolTipIcon.Info);
                            for (int attCtr = 1; attCtr <= tmpMail.Attachments.Count; attCtr++)
                            {
                                if (!OSFR.bOverwriteAttachmentsWSameName[ctr] && File.Exists(OSFR.sDestination[ctr] + "\\" + OSFR.sRuleName[ctr] + "\\" + tmpMail.Attachments[attCtr].FileName))
                                {
                                    try
                                    {
                                        tmpMail.Attachments[attCtr].SaveAsFile(OSFR.sDestination[ctr] + "\\" + OSFR.sRuleName[ctr] + "\\" + (System.DateTime.Now).ToBinary().ToString() + "_" + tmpMail.Attachments[attCtr].FileName);
                                    }
                                    catch(System.Exception ex)
                                    {
                                        notifyicon.ShowBalloonTip(5000, "File Save Error", ex.Message, ToolTipIcon.Error);
                                    }
                                }
                                else
                                {
                                    try
                                    {
                                        tmpMail.Attachments[attCtr].SaveAsFile(OSFR.sDestination[ctr] + "\\" + OSFR.sRuleName[ctr] + "\\" + tmpMail.Attachments[attCtr].FileName);
                                    }
                                    catch (System.Exception ex)
                                    {
                                        notifyicon.ShowBalloonTip(5000, "File Save Error", ex.Message, ToolTipIcon.Error);
                                    }
                                }
                                notifyicon.ShowBalloonTip(5000, "File Saved", "Based on Rule - " + OSFR.sRuleName[ctr] + ", attachement - " + newMail.Attachments[attCtr].FileName + ", is saved at " + OSFR.sDestination[ctr], ToolTipIcon.Info);

                                if (OSFR.bRemoveAttachmentFromMail[ctr])
                                {
                                    newMail.Attachments[attCtr].Delete();
                                }

                            }

                        }

                    }

                }

            }

        }
    }
}
