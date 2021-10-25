using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Cache;
using System.Data.OleDb;
using System.Diagnostics;

namespace RanchArchive
{
    public partial class Form1 : Form
    {   // DEFINE GLOBAL VARIABLES
        // SQL
        string strConnectionString = "Data Source=tcp:sql2k804.discountasp.net;Provider=SQLOLEDB;Initial Catalog=SQL2008R2_905030_ranchsqldb;User ID=SQL2008R2_905030_ranchsqldb_user;Password=CucamongaValley2011;";
        string userDomainName = Environment.UserDomainName;
        // ENVIRONMENT VARIABLES
        string userMachineName = Environment.MachineName;
        string userCurrentUser = Environment.UserName;
        string strDefaultSkyDriveImages = @"C:\Users\" + Environment.UserName + @"\SkyDrive\Image Gallery - Web Edited & Preview Images";
        string strRanchArchiveUpdate = @"C:\Users\" + Environment.UserName + @"\SkyDrive\Application\Desktop Applications\RanchArchive\setup.exe";
        int intNextPanelName = 0;
        
        // SOME CONSTANTS for determining display size
        int intX = 3;
        int intY = 3;
        int intXhome = 3;
        int intPictureWidth = 300;
        int intPictureHeight = 300;
        int intIncrement = 305;
        int intCutlineDepth = 40;
        int intVerticalIncrement = 300 + 40 + 10;
        bool boolInitPhase = true;
        bool boolIsMailAddressList = false;
        //string siOriginalName;
        string gblstrDBLogID;
        string gblstrOriginalName;
        string gblstrDescription;
        string gblstrLocation;
        string gblstrDateofPicture;
        string gblstrPersons;

        // DEFINE THE original WebJournalNames to be used in the UNDO operation
        //int njID = 0;
        string njType, njWebNameID, njWeNamList, njNameSortOrder, njMiddleName, njMaidenName, njOtherNameRef = "";
        string njBirthDate, njBirthPlace, njIntermentLocation, njDateDied, njMother, njFather, njSiblings, njSpouse = "";
        string njChildrenNames, njKnowledgeBase, njYouthPicture, njAdultPicture = "";
        
        public Form1()
        {
            InitializeComponent();
            tsslVersion.Text = "2014.01.17 / 1";
            // BUTTON MANAGEMENT VISIBILITY
            btnANRsaveRecord.Visible = true;
            btnANRupdateRecord.Visible = false;
            btnGoToRecord.Visible = false;

            // TOOL TIPS
            ToolTip ttClearButton = new System.Windows.Forms.ToolTip();
            ttClearButton.ToolTipTitle = "START BUTTON";
            ttClearButton.UseAnimation = true;
            ttClearButton.ShowAlways = true;
            //ttClearButton.ToolTipIcon = ToolTipIcon.Info;
            //ttClearButton.IsBalloon = true;
            ttClearButton.SetToolTip(this.btnANRcancel, "This will erase the boxes above.");

            // CLOSE BUTTON
            ToolTip ttCloseButton = new System.Windows.Forms.ToolTip();
            ttCloseButton.ToolTipTitle = "CLOSE BUTTON";
            ttCloseButton.UseAnimation = true;
            ttCloseButton.ShowAlways = true;
            //ttCloseButton.IsBalloon = true;
            //ttCloseButton.ToolTipIcon = ToolTipIcon.Info;
            ttCloseButton.SetToolTip(this.btnANRclose, "This will close this form and return to the SEARCH form");

            // SAVE RECORD BUTTON
            ToolTip ttSaveButton = new System.Windows.Forms.ToolTip();
            ttSaveButton.ToolTipTitle = "SAVE RECORD BUTTON";
            ttSaveButton.UseAnimation = true;
            ttSaveButton.ShowAlways = true;
            //ttSaveButton.IsBalloon = true;
            //ttSaveButton.ToolTipIcon = ToolTipIcon.Info;
            ttSaveButton.SetToolTip(this.btnANRsaveRecord, "This button will SAVE this as a new record after asking you to confirm your desire.");

            // UPDATE RECORD BUTTON
            ToolTip ttUpdateButton = new System.Windows.Forms.ToolTip();
            ttUpdateButton.ToolTipTitle = "UPDATE RECORD BUTTON";
            ttUpdateButton.UseAnimation = true;
            ttUpdateButton.ShowAlways = true;
            //ttUpdateButton.IsBalloon = true;
            //ttUpdateButton.ToolTipIcon = ToolTipIcon.Info;
            ttUpdateButton.SetToolTip(this.btnANRupdateRecord, "This button will UPDATE the record you chose after asking you to confirm");

            // SELECTION LIST BOX
            ToolTip ttListBox = new System.Windows.Forms.ToolTip();
            //ttListBox.ToolTipTitle = "UPDATE RECORD SELECTION";
            ttListBox.UseAnimation = true;
            ttListBox.ShowAlways = true;
            //ttListBox.InitialDelay = 1000;
            ttListBox.IsBalloon = true;
            //ttListBox.ToolTipIcon = ToolTipIcon.Info;
            //ttListBox.BackColor = Color.Yellow; // DOESN'T SEEM TO WORK SO I TURNED IT OFF SO THINGS WOULD GO FASTER
            ttListBox.SetToolTip(this.lbxANRpreviousAditionsOrUpdates, "Click on a line in this box to edit (UPDATE) the record chosen.\n\rIf you select the wrong row, choose another.");

            // ORIGINAL NAME TEXT BOX
            ToolTip ttOriginalName = new System.Windows.Forms.ToolTip();
            //ttOriginalName.ToolTipTitle = "ORIGINAL NAME";
            ttOriginalName.UseAnimation = true;
            ttOriginalName.ShowAlways = true;
            //ttOriginalName.InitialDelay = 1000;
            //ttOriginalName.IsBalloon = true;
            //ttOriginalName.ToolTipIcon = ToolTipIcon.Info;
            ttOriginalName.SetToolTip(this.tbxANRoriginalName, "The name you give to the scanned image which ends in a number.");

            // DESCRIPTION TEXT BOX
            ToolTip ttDescription = new System.Windows.Forms.ToolTip();
            //ttDescription.ToolTipTitle = "DESCRIPTION";
            ttDescription.UseAnimation = true;
            ttDescription.ShowAlways = true;
            //ttDescription.InitialDelay = 1000;
            //ttDescription.IsBalloon = true;
            //ttDescription.ToolTipIcon = ToolTipIcon.Info;
            ttDescription.SetToolTip(this.tbxANRdescription, "Enter text that describes the event or objects in the picture (not already listed in Location or Persons boxes.");

            // LOCATION COMBO BOX
            ToolTip ttLocation = new System.Windows.Forms.ToolTip();
            //ttLocation.ToolTipTitle = "LOCATION";
            ttLocation.UseAnimation = true;
            ttLocation.ShowAlways = true;
            //ttLocation.InitialDelay = 1000;
            //ttLocation.IsBalloon = true;
            //ttLocation.ToolTipIcon = ToolTipIcon.Info;
            ttLocation.SetToolTip(this.cbxANRlocation, "Choose from the drown down list.");

            // PERSONS TEXT BOX
            ToolTip ttPersons = new System.Windows.Forms.ToolTip();
            //ttPersons.ToolTipTitle = "PERSONS";
            ttPersons.UseAnimation = true;
            ttPersons.ShowAlways = true;
            //ttPersons.InitialDelay = 1000;
            //ttPersons.IsBalloon = true;
            //ttPersons.ToolTipIcon = ToolTipIcon.Info;
            ttPersons.SetToolTip(this.tbxANRpersons, "Enter free form or use the box at the left (above ADD PERSONS) to add names.");

            // DATE OF PICTURE TEXT BOX
            ToolTip ttDateOfPicture = new System.Windows.Forms.ToolTip();
            //ttDateOfPicture.ToolTipTitle = "DATE OF PICTURE";
            ttDateOfPicture.UseAnimation = true;
            ttDateOfPicture.ShowAlways = true;
            //ttDateOfPicture.InitialDelay = 1000;
            //ttDateOfPicture.IsBalloon = true;
            ttDateOfPicture.ToolTipIcon = ToolTipIcon.Info;
            ttDateOfPicture.SetToolTip(this.tbxANRdateOfPicture, "Enter a date in any form you choose that relates to when this picture was taken.");

            // PERSONS COMBO BOX
            ToolTip ttPersonsCB = new System.Windows.Forms.ToolTip();
            //ttPersonsCB.ToolTipTitle = "PERSONS DROP DOWN LIST";
            ttPersonsCB.UseAnimation = true;
            ttPersonsCB.ShowAlways = true;
            //ttPersonsCB.InitialDelay = 1000;
            //ttPersonsCB.IsBalloon = true;
            //ttPersonsCB.ToolTipIcon = ToolTipIcon.Info;
            ttPersonsCB.SetToolTip(this.cbxANRpersons, "Clicking on items in this box will add them to the Persons box.");

            // ADD LOCATION BUTTON
            ToolTip ttLocationButton = new System.Windows.Forms.ToolTip();
            //ttLocationButton.ToolTipTitle = "ADD LOCATION";
            ttLocationButton.UseAnimation = true;
            ttLocationButton.ShowAlways = true;
            ttLocationButton.InitialDelay = 1000;
            //ttLocationButton.IsBalloon = true;
            //ttLocationButton.ToolTipIcon = ToolTipIcon.Info;
            ttLocationButton.SetToolTip(this.btnAddlocation, "This button will allow you to add a new location to the drop down list of locations.");

            // ADD NEW PERSON BUTTON
            ToolTip ttPersonsButton = new System.Windows.Forms.ToolTip();
            //ttPersonsButton.ToolTipTitle = "ADD PERSON";
            ttPersonsButton.UseAnimation = true;
            ttPersonsButton.ShowAlways = true;
            ////ttPersonsButton.InitialDelay = 1000;
            //ttPersonsButton.IsBalloon = true;
            //ttPersonsButton.ToolTipIcon = ToolTipIcon.Info;
            ttPersonsButton.SetToolTip(this.btnANRperson, "This will take you to the KNOWLEDGE BASE where you can add a new person to the list above this button.");

            // 

            // GLOBAL ADDRESS BOOK ARRAY
            string[] globalAddressBookArrayURI = new string[50];     // Make bigger than needed

            ////  DEFINE THE original WebJournalNames to be used in the UNDO operation
            //int njID = 0;
            //string njType, njWebNameID, njWeNamList, njNameSortOrder, njMiddleName, njMaidenName, njOtherNameRef = "";
            //string njBirthDate, njBirthPlace, njIntermentLocation, njDateDied, njMother, njFather, njSiblings, njSpouse = "";
            //string njChildrenNames, njKnowledgeBase, njYouthPicture, njAdultPicture = "";
            //DateTime njDateTimeRecorded = DateTime.Now;
            
            // PANEL SETTINGS
            pnlSettings.Visible = false;            // Only turned on when Settings shown
            pnlEmail.Visible = false;
            pnlAddNewRecord.Visible = false;
            pnlLightbox.Visible = false;
            pnlSearchFilter.Visible = false;
            pnlKnowledgeBase.Visible = false;

            pnlWelcome.Parent = panel1;
            pnlWelcome.BringToFront();
            pnlWelcome.Visible = false;
            //lbxPictureResults.Visible = false;

            // MAKE THE SCREEN MAXIMUM SIZE
            System.Drawing.Rectangle workingRectangle = Screen.PrimaryScreen.WorkingArea;
            int intScreenWidth = workingRectangle.Width;
            int intScreenHeight = workingRectangle.Height;
            this.Size = new System.Drawing.Size(workingRectangle.Width - 10, workingRectangle.Height - 10);
            this.Location = new System.Drawing.Point(3, 3);

            if (userMachineName == "DV-1")
            {
                strDefaultSkyDriveImages = @"F:\" + @"\Hofer SkyDrive\Image Gallery - Web Edited & Preview Images";
                strRanchArchiveUpdate = @"F:\Hofer SkyDrive\Application\Desktop Applications\RanchArchive\setup.ext";
            }

            // FILL THE COMBO BOXES WITH VALUES
            fillComboBox("ImageDate", "ImageDate", "ImageDate", cbxSearchFilterDateOfPicture);
            //fillComboBox("ImageDate", "ImageDate", "ImageDate", cbxANRimageDate);  // DEPRICATED
            fillComboBox("KeyWord", "KeywordSearch", "KeyWord", cbxSearchFilterKeyword);
            fillComboBox("Location","WebLocations","Location",cbxSearchFilterLocation);
            fillComboBox("Location", "WebLocations", "Location", cbxANRlocation);
            fillComboBox("FamilyName", "FamilyGroupID", "FamilyName", cbxSearchFilterFamily);
            fillComboBox("WeNamList", "WebNames", "WeNamList", cbxSearchFilterPerson);
            fillComboBox("WeNamList", "WebNames", "WeNamList", cbxANRpersons);
            fillComboBox("Type", "ImageType", "Type", cbxSearchFilterImageType);
        }

        public void sendEmail (string strTo, string strCC, string strSubject, string strMessage)
        {   
            string formatedTodaysDateTime = string.Format("{0:G}", DateTime.Now);
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Port = 587;
            smtpClient.Host = "smtp.office365.com";
            smtpClient.EnableSsl = true;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential("robert.perry@hoferranch.onmicrosoft.com", "Cucamonga.2011");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.SendCompleted += new SendCompletedEventHandler(SendCompleteEventHandler);
            MailAddress from = new MailAddress("robert.perry@hoferranch.onmicrosoft.com", "RanchArchive Mailer");

            //int intTOindex = cbxEmailTo.SelectedIndex;
            //int intCCindex = cbxEmailCC.SelectedIndex;
            ////string strTO = globalAddressBookArrayURI[intTOindex];
            MailAddress to = new MailAddress(strTo,"");
            MailAddress cc = new MailAddress(strCC,"");
            MailMessage mmMessage = new MailMessage(from, to);
            mmMessage.Subject = strSubject;
            mmMessage.IsBodyHtml = false;
            mmMessage.CC.Add(cc);
            string strMessaegBody = strMessage;
            //strMessaegBody += strMessage;
            mmMessage.Body = strMessaegBody;

            // SEND THE EMAIL
            /// TODO: Somewhere here we must check for an empty TO field and abort sending
            try  // to send the email
            {
                smtpClient.SendAsync(mmMessage, "");
                tsEmailStatus.ForeColor = Color.DarkBlue;
                tsEmailStatus.BackColor = Color.LightBlue;
                tsEmailStatus.Text = "Sending mail to: \'" + strTo + "\'";
                if (strCC != "") { tsEmailStatus.Text += " and to \'" + strCC + "\'"; }
                pnlEmail.Visible = false;   // We have sent the mail, close the panel
            }
            catch (Exception eMail)
            {
                if (eMail.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                else
                { 
                    string message = "Failed to send email. Please record the information and email from your regular EMail program.";
                    message += "\r\nError Number: " + eMail.HResult;
                    message += "\r\n" + eMail.Message;
                    MessageBox.Show(message, "ERROR SENDING EMAIL",MessageBoxButtons.OK);
                    mailProgramError(message);               
                }
            }
        }

        public void fillComboBox(string selectName, string selectTableName, string selectOrderBy, ComboBox cbxName)
        {   // This will take the three arguments passed and read the database table specified and fill the combo box given ordering
            // the combo box list by the field specified.
            string strSQL = "SELECT COUNT(" + selectName + ") AS lastKey FROM " + selectTableName;
            try   // to set up a CONNECTION to the SQL DB (E1)
            {
                OleDbConnection cn = new OleDbConnection();
                cn.ConnectionString = strConnectionString;
                try // to OPEN the connection (E2)
                {
                    cn.Open();
                    try // to create an OLEDB COMMAND (E3)
                    {
                        OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
                        try  // to create a DATA READER (E4)
                        {
                            OleDbDataReader myDataReader;
                            try  // to read the single data item returned, viz, the number of records in the table (E5)
                            {
                                myDataReader = myCommand.ExecuteReader();
                                try // to read the returned value, viz, the number of records in the table (E6)
                                {
                                    myDataReader.Read();
                                    string keyCount = myDataReader["lastKey"].ToString();
                                    int intKeyCount = Convert.ToInt32(keyCount) +1; // Allow for every box having an ANY choice
                                    // GARBAGE COLLECT
                                    myDataReader.Close();
                                    myDataReader.Dispose();
                                    // BUILD AN IN MEMORY ARRAY THAT WILL FILL THE COMBOBOX specified by cbxName
                                    var arrayForCombobox = new string[intKeyCount];
                                    // FILL THE COMBOBOX cbxName with the data from table selectTableName field selectTableName ordered by selectOrderBy
                                    strSQL = "SELECT " + selectName + " FROM " + selectTableName + " ORDER BY " + selectOrderBy;
                                    try  // to fill the combobox
                                    {
                                        OleDbConnection cn2 = new OleDbConnection();
                                        cn2.ConnectionString = strConnectionString;
                                        cn2.Open();
                                        OleDbCommand myCommand2 = new OleDbCommand(strSQL, cn2);
                                        OleDbDataReader myDataReader2;
                                        myDataReader2 = myCommand2.ExecuteReader();
                                        int count = 0;
                                        // ADD THE ANY CHOICE TO EACH BOX
                                        arrayForCombobox[count++] = "Any";

                                        while (myDataReader2.Read())
                                        {
                                            arrayForCombobox[count++] = myDataReader2[selectName.ToString()].ToString().Trim();
                                        }
                                        cbxName.DataSource = arrayForCombobox;
                                    }
                                    catch (Exception e7)
                                    {
                                        if (e7.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                                        else
                                        {
                                            string message = "E7: Failed to fill the combobox \'" + selectTableName + "\'\r\nError #: " + e7.HResult + "\r\n" + e7.Message + "\r\nThis problem will be eMailed to the developer";
                                            MessageBox.Show(message, "FILL COMBO BOX", MessageBoxButtons.OK);
                                            // sendBugReport (strTO, message);
                                            mailProgramError(message);
                                        }
                                    }
                                }
                                catch (Exception e6)
                                {
                                    if (e6.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                                    else
                                    {
                                        string message = "E6: Failed during the READING of the data. \r\nError #: " + e6.HResult + "\r\n" + e6.Message + "\r\nThis problem will be eMailed to the developer";
                                        MessageBox.Show(message, "FILL COMBO BOX", MessageBoxButtons.OK);
                                        mailProgramError(message);
                                    }
                                }
                            }
                            catch (Exception e6)
                            {
                                if (e6.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                                else
                                {
                                 string message = "E5: Failed during EXECUTION of the DATA READER. \r\nError #: " + e6.HResult + "\r\n" + e6.Message + "\r\nThis problem will be eMailed to the developer";
                                MessageBox.Show(message, "FILL COMBO BOX", MessageBoxButtons.OK);
                                mailProgramError(message);
                                }
                            }
                        }
                        catch (Exception e4)
                        {
                            if (e4.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                            else
                            {
                                string message = "E4: Failed during createion of a DATA READER. \r\nError #: " + e4.HResult + "\r\n" + e4.Message + "\r\nThis problem will be eMailed to the developer";
                                MessageBox.Show(message, "FILL COMBO BOX", MessageBoxButtons.OK);
                                mailProgramError(message);
                            }
                        }
                    }
                    catch (Exception e3)
                    {
                        if (e3.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                        else
                        {
                         string message = "E3: Failed during creation of a connection COMMAND. \r\nError #: " + e3.HResult + "\r\n" + e3.Message + "\r\nThis problem will be eMailed to the developer";
                        MessageBox.Show(message, "FILL COMBO BOX", MessageBoxButtons.OK);
                        mailProgramError(message);
                        }
                    }
                }
                catch (Exception e2)
                {
                    if (e2.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                    else
                    {
                        string message = "E2: Failed during connection OPEN. \r\nError #: " + e2.HResult + "\r\n" + e2.Message + "\r\nThis problem will be eMailed to the developer";
                        MessageBox.Show(message, "FILL COMBO BOX", MessageBoxButtons.OK);
                        mailProgramError(message);
                    }
                }
            }
            catch (Exception e1)
            {
                if (e1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                else
                {
                    string message = "E1: Failed during connection setup. \r\nError #: " + e1.HResult + "\r\n" + e1.Message + "\r\nThis problem will be eMailed to the developer";
                    MessageBox.Show(message, "FILL COMBO BOX", MessageBoxButtons.OK);
                    mailProgramError(message);
                }
            }
        }

        private void panel1Size_regionChanged(object sender, EventArgs e)
        {
            MessageBox.Show("Screen size changed","REGION CHANGED",MessageBoxButtons.OK);
        }

        // EXIT THE PROGRAM
        private void eXITToolStripMenuItem_Click(object sender, EventArgs e)
        {   // Exit the program
            this.Close();
            /// TODO: Add any cleanup code that needs to be done, maybe logging the work being done.
        }

        // MENU CHOICE: FILE > SETTINGS
        private void settingsToolStripMenuItem_Click(object sender, EventArgs e)
        {   // SET UP OTHER PANELS
            pnlAddNewRecord.Visible = false;
            pnlKnowledgeBase.Visible = false;
            //pnlWelcome.Visible = false;
            //pnlSettings.Visible = false;
            pnlEmail.Visible = false;
            pnlLightbox.Visible = false;
            pnlSearchFilter.Visible = false;

            // PANEL SETTINGS
            pnlSettings.Parent = panel1;
            pnlSettings.Location = new Point(30, 50);
            pnlSettings.BackColor = Color.OldLace;
            pnlSettings.Visible = true;



            pnlSearchFilter.Visible = false;
            pnlSettings.BringToFront();

            lblSettingsDomainNameValue.Text = userDomainName;
            lblSettingsMachineNameValue.Text = userMachineName;
            lblSettingsUserNameValue.Text = userCurrentUser;
            lblSettingsImageGalleryValue.Text = strDefaultSkyDriveImages;
            boolInitPhase = false;              // Turn this off since we are passed Initialization if we got here
        }

        // SETTINGS PANEL "CLOSE" BUTTON
        private void btnSettingsClose_Click(object sender, EventArgs e)
        {   // CLOSE THE PANEL SHOWING THE SETTINGS INFORMATION
            //pnlWelcome.Visible = false;
            pnlSettings.Visible = false;
            pnlLightbox.Visible = true;
            pnlSearchFilter.Visible = true;
        }

        private void placeNewPictureBox(int intX, int intY, int intIncrement, string strPicture, Panel p, string strCutline, string strDBLogID, bool IsOriginalFilename)
        {
            PictureBox pb = new PictureBox();
            pb.Parent = p;
            pb.Location = new Point(intX, intY);
            pb.Size = new Size(intPictureWidth, intPictureHeight);
            pb.SizeMode = PictureBoxSizeMode.Zoom;
            //pb.Load(strDefaultSkyDriveImages + @"\" + strPicture);
            pb.BorderStyle = BorderStyle.FixedSingle;
            pb.Visible = true;
            pb.Name = "pbx" + Convert.ToString(intNextPanelName);
            intNextPanelName += 1;
            TextBox cutline = new TextBox();
            cutline.Parent = p;
            cutline.Location = new Point(intX, intY + intIncrement);
            cutline.Size = new Size(intPictureHeight, intCutlineDepth);
            cutline.Text = strCutline;
            cutline.Name = "Cutline for " + pb.Name;
            if (IsOriginalFilename) { cutline.BackColor = Color.LightBlue; } else { cutline.BackColor = Color.OldLace; }
            // OldLace IS THE COLOR FOR WebJPGFilename (WEB) and LightBlue IS THE COLOR FOR USE OF OriginalImageName
            cutline.Multiline = true;
            cutline.ScrollBars = ScrollBars.Vertical;
            cutline.Visible = true;
            if (File.Exists(strDefaultSkyDriveImages + @"\" + strPicture))
            {
                pb.Load(strDefaultSkyDriveImages + @"\" + strPicture);
                if (tbxSearchFilterJPGpictureName.Text != "")
                {   // SPECIAL CASE WHERE WE WANT THE USER TO BE ABLE TO CLICK ON THE PICTURE AND GO TO THE RECORD 

                    // REBUILD THE tbxSearchFilterJPGpictureName.Text box with the correct JPG filename.
                    string strJPGname = tbxSearchFilterJPGpictureName.Text.Trim();
                    strJPGname = strJPGname.Replace(".jpg", "").Replace(".tif", "") + ".jpg";
                    //string[] filenameParts = strJPGname.Split('.');
                    //tbxSearchFilterJPGpictureName.Text = filenameParts[0] + ".jpg";

                    btnGoToRecord.Visible = true;
                    btnGoToRecord.Parent = p;
                    btnGoToRecord.Location = new Point(intX, intY+intVerticalIncrement + 10);
                    btnGoToRecord.Tag = strDBLogID + " | | ";     // NEED THIS TO BE A STRING OF THIS SORT "DBLogID | |" because 
                }
            }
            else
            {   // FILE DOESN'T EXIST, BUT WE NEED TO INDICATE THIS
                pb.BackColor = Color.DarkGray;
                // SEND ERROR MESSAGE
                string message = "The file named \'" + strDefaultSkyDriveImages + @"\" + strPicture + "\' does not exist. The cutline contains the DBLogID.\r\n\n";
                message += cutline.Text; // +" DBLOGID: " + cutline.Tag.ToString();
                // FOR DEGUGGING LEAVE THIS OFF: sendEmail("steve.breckner@hoferranch.onmicrosoft.com","robert.perry@hoferranch.onmicrosoft.com","Named file in SQL ScannedImages missing in SkyDrive",message);
                // MAYBE COULD ADD CODE HERE TO SET AN ACTION ITEM FOR THIS AS WELL
            }
            if (cbxAlwaysClearSearchCriteria.Checked)
            {   // CLEAR SEARCH BRITERIAL BOXEX 
                
            }
        }

        private void Form1_Resize(object sender, System.EventArgs e)
        {
            Control control = (Control)sender;
            //toolStripStatusLabel1.Text = "WORKING RECTANGLE: " + control.Size.Height.ToString() + ", " + control.Size.Width.ToString();
            // WELCOME PANEL
            //pnlWelcome.Visible = false;
            // PANEL1 - THE CONTAINER PANEL
            int panel1Width = panel1.Width;
            int panelHeigth = panel1.Width;
            panel1.BackColor = Color.DarkGray;
            panel1.Width = control.Size.Width - 10;  // was - 100, -80
            panel1.Height = control.Size.Height -100;
            panel1.Location = new Point(3, 25);
            panel1.Location = new Point(0, 0);

            // SEARCH PANEL
            pnlSearchFilter.BackColor = Color.OldLace;
            pnlSearchFilter.Location = new Point(3, 30);
            pnlSearchFilter.Size = new Size(210, control.Size.Height);
            
            // LIGHTBOX PANEL
            pnlLightbox.Size = new Size(control.Size.Width-300, control.Size.Height-130);
            pnlLightbox.Location = new Point(214, 30);
            pnlLightbox.AutoScroll = true;
            pnlLightbox.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            pnlLightbox.BackColor = Color.OldLace;

            // RESIZE THE BOX
            if (!boolInitPhase && pnlLightbox.IsAccessible) {findPictures(); }  // Redisplay
        }
        //
        // Find Pictures Worker process
        private void findPicturesWorker()
        {
            int intMaxPanelWidth = pnlLightbox.Width;
            // SECOND THING IS TO REMOVE THE EXISTING PANEL AND ITS CONTENTS
            string strNewDate, strNewKeyword, strNewLocation, strNewFamily, strNewPerson, strNewImageType = "";

            // DISPOSE OF PREVIOUS PICTURES
            foreach (Control pic in pnlLightbox.Controls)
                if (pic.GetType() == typeof(PictureBox))
                    pic.Dispose();
            foreach (Control tbx in pnlLightbox.Controls)       // SEEMS TO BE NECESSARY TO DO THIS MANY TIMES
                if (tbx.GetType() == typeof(TextBox))
                    tbx.Dispose();
            foreach (Control tbx in pnlLightbox.Controls)
                if (tbx.GetType() == typeof(TextBox))
                    tbx.Dispose();
            foreach (Control tbx in pnlLightbox.Controls)
                if (tbx.GetType() == typeof(TextBox))
                    tbx.Dispose();
            foreach (Control tbx in pnlLightbox.Controls)
                if (tbx.GetType() == typeof(TextBox))
                    tbx.Dispose();
            foreach (Control tbx in pnlLightbox.Controls)
                if (tbx.GetType() == typeof(TextBox))
                    tbx.Dispose();

            // AND WE MUST TURN OFF THE GOTO RECORD BUTTON
            btnGoToRecord.Visible = false;

            // MAKE SURE THE LIGHTBOX IS TURNED ON
            pnlLightbox.Parent = panel1;
            pnlLightbox.Visible = true;
            pnlLightbox.BackColor = Color.OldLace;
            pnlLightbox.BringToFront();

            // Now build new set of pictures
            strNewDate = cbxSearchFilterDateOfPicture.Text.Trim().Replace("\n", "");      // Trimming shouldn't be necessary, but for safety will do it anyway
            strNewKeyword = cbxSearchFilterKeyword.Text.Trim().Replace("\n", "");
            strNewLocation = cbxSearchFilterLocation.Text.Trim(); //.Replace("\n","");
            strNewFamily = cbxSearchFilterFamily.Text.Trim().Replace("\n", "");
            strNewPerson = cbxSearchFilterPerson.Text.Trim().Replace("\n", "");
            strNewImageType = cbxSearchFilterImageType.Text.Trim().Replace("\n", "");

            // CHECK FOR 'ANY' AND CONVERT IT TO %
            if (strNewDate == "Any") { strNewDate = "%"; } else { strNewDate = "%" + strNewDate + "%"; }
            if (strNewKeyword == "Any") { strNewKeyword = "%"; } else { strNewKeyword = "%" + strNewKeyword + "%"; }
            if (strNewLocation == "Any") { strNewLocation = "%"; } else { strNewLocation = "%" + strNewLocation + "%"; }
            if (strNewFamily == "Any") { strNewFamily = "%"; } else { strNewFamily = "%" + strNewFamily + "%"; }
            if (strNewPerson == "Any") { strNewPerson = "%"; } else { strNewPerson = "%" + strNewPerson + "%"; }
            if (strNewImageType == "Any") { strNewImageType = "%"; } else { strNewImageType = "%" + strNewImageType + "%"; }

            // BUILD THE SQL QUERY
            string strSQL = "";
            if (tbxSearchFilterJPGpictureName.Text == "")
            {
                strSQL = "SELECT DBLogID, Description, WebDescription, WebJPGFilename, OriginalImageName FROM ScannedImages ";
                strSQL += "WHERE ";
                strSQL += "ImageDate LIKE '" + strNewDate + "' AND ";
                strSQL += "Keywords LIKE '" + strNewKeyword + "' AND ";
                strSQL += "WebLocation LIKE '" + strNewLocation + "' AND ";
                strSQL += "FamilyID LIKE '" + strNewFamily + "' AND ";
                strSQL += "WebPersonID LIKE '" + strNewPerson + "' AND ";
                strSQL += "ImageType LIKE '" + strNewImageType + "' AND ";
                strSQL += "DuplicateImage ='0'";
                strSQL += "ORDER BY DBLogID";
            }
            else     // WE WANT TO FIND THE PICTURE BY JPG FILENAME ONLY
            {   // NOTE: This is a special case where the user has a filename and wants to view just that picture. The scenario
                // is if someone gets a question about a picture and the jpg name of the picture is known, then that will be
                // the easiest way to answer the question.

                // NOTE: the record probably contains the filename with the .tif extension, so we need to strip off all extensions and perform a LIKE check
                //string[] filenameParts = tbxSearchFilterJPGpictureName.Text.Split('.');
                //string strOriginalFilename = filenameParts[0];
                string strNamePart = tbxSearchFilterJPGpictureName.Text.Replace(".jpg", "").Replace(".tif", "");

                // BUILD THE SQL STATEMENT
                strSQL = "SELECT DBLogID, Description, WebDescription, WebJPGFilename, OriginalImageName FROM ScannedImages ";
                strSQL += "WHERE WebJPGFilename = '" + tbxSearchFilterJPGpictureName.Text + "' OR OriginalImageName = '" + strNamePart + ".tif' OR OriginalImageName ='" + strNamePart + ".jpg'";
            }

            // EXECUTE SQL
            //lbxPictureResults.Items.Clear();        // Erase anything already there
            string strDBLogID, strDescription, strWebDescription, strWebJPGFilename, strOriginalImageName = "";
            try  // to set up a connection [E1]
            {
                OleDbConnection cn = new OleDbConnection();
                cn.ConnectionString = strConnectionString;
                cn.Open();
                try // to CREATE A command AND DataReader AND execute
                {
                    OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
                    OleDbDataReader myDataReader;
                    myDataReader = myCommand.ExecuteReader();
                    try  // to read the data
                    {
                        intX = 3;
                        intY = 3;
                        string strCutline = "";
                        int intRecordCount = 0;
                        //toolStripProgressBar1.Maximum = 200;
                        //toolStripProgressBar1.Step = 1;
                        //toolStripProgressBar1.BackColor = Color.SaddleBrown;
                        //toolStripProgressBar1.Enabled = true;
                        bool IsOriginalFilename = false;    // Meaning that is a Web filename
                        while (myDataReader.Read())
                        {
                            //toolStripProgressBar1.PerformStep();

                            strDBLogID = myDataReader["DBLogID"].ToString();
                            strDescription = myDataReader["Description"].ToString().Trim().Replace("\n", "");
                            strWebDescription = myDataReader["WebDescription"].ToString().Trim().Replace("\n", "");
                            strWebJPGFilename = myDataReader["WebJPGFilename"].ToString().Trim();
                            strOriginalImageName = myDataReader["OriginalImageName"].ToString().Trim().Replace("\n", "");
                            if (strWebJPGFilename == null) { strWebJPGFilename = ""; }      // IF RETURNED NAME IS A NULL THEN MAKE IT AN EMPTY STRING
                            if (strWebDescription != "") { strCutline = strWebDescription; }
                            else { strCutline = strDescription; }
                            //strCutline += "\r\n[" + strDBLogID + " : " + strWebJPGFilename + "]";
                            if (strWebJPGFilename != "")
                            {
                                IsOriginalFilename = false;
                                strCutline += "\r\n[" + strDBLogID + " : " + strWebJPGFilename + "]";
                                placeNewPictureBox(intX, intY, intIncrement, strWebJPGFilename, pnlLightbox, strCutline, strDBLogID, IsOriginalFilename);
                                intRecordCount += 1;
                                intX += intIncrement;
                                if (intX + intIncrement > intMaxPanelWidth)
                                {
                                    intX = intXhome;
                                    intY += intVerticalIncrement;
                                }
                            }
                            else   // No Web filename, try the Original file name and check the on machine SkyDrive directory for that.
                            {
                                if (strOriginalImageName != "")  // there is a name in the OriginalImageName field, use it
                                {
                                    IsOriginalFilename = true;
                                    // STRIP OFF THE EXTENSION AND REPLACE IT WITH A .jpg EXTENSION AND TRY AGAIN
                                    string strName = strOriginalImageName.Trim();
                                    //string[] filenameParts = strJPGname.Split('.');
                                    //string strOriginalFilenameJPG = filenameParts[0] + ".jpg";
                                    strName = strName.Replace(".jpg", "").Replace(".tif", "") + ".jpg";
                                    strCutline += "\r\n[" + strDBLogID + " : " + strName + "]";
                                    placeNewPictureBox(intX, intY, intIncrement, strName, pnlLightbox, strCutline, strDBLogID, IsOriginalFilename);
                                    intRecordCount += 1;
                                    intX += intIncrement;
                                    if (intX + intIncrement > intMaxPanelWidth)
                                    {
                                        intX = intXhome;
                                        intY += intVerticalIncrement;
                                    }
                                }
                            }
                        }
                        myDataReader.Close();
                        myCommand.Dispose();
                        cn.Close();
                        cn.Dispose();
                        //lbxPictureResults.Items.Add("RECORDS DISPLAYED: " + intRecordCount.ToString());

                        toolStripStatusLabel1.Text = "RECORDS DISPLAYED: " + intRecordCount.ToString();
                        //toolStripProgressBar1.Enabled = false;      // turn it off
                    }
                    catch (Exception e3)
                    {
                        if (e3.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK);  }
                        else
                        {
                            string message = "[102:E3] Failure to read the picture \'" + tbxSearchFilterJPGpictureName.Text + "' or \'" + strOriginalImageName + "\' as a JPG file.";
                            message += "\r\nError:" + e3.HResult + "\r\n" + e3.Message + "\r\nThis error has been reported via email.";
                            MessageBox.Show(message, "FIND BUTTON REQUEST", MessageBoxButtons.OK);    
                        }
                        
                    }

                }
                catch (Exception e2)
                {
                    if (e2.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                    else
                    {
                        string message = "Failure to create a command, data reader and execute same [102:E2] \r\nError:" + e2.HResult + "\r\n" + e2.Message + "\r\nThis error has been reported via email.";
                        MessageBox.Show(message, "FIND BUTTON REQUEST", MessageBoxButtons.OK);
                        mailProgramError(message);
                    }
                    
                }
            }
            catch (Exception e1)
            {
                if (e1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                else
                {
                    string message = "Error creating a connection to the SQL Database. [102:E1]\r\nError:" + e1.HResult + "\r\n" + e1.Message + "\r\nThis error has been reported via email.";
                    MessageBox.Show(message, "FIND BUTTON REQUEST", MessageBoxButtons.OK);
                    mailProgramError(message);
                }
                
            }
            // CHECK IF WE ARE TO CLEAR THE SELECTION BOXES
            if (cbxAlwaysClearSearchCriteria.Checked) { clearSelectionCriteriaBoxes(); }
        }

        //
        // FIND PICTURES
        // =============
        private void findPictures()
        { // THE WILL LOOK UP PICTURES BASED ON THE SEARCH CRITERIA
            // TURN OFF WELCOME PANEL
            pnlWelcome.Visible = false;


            // FIRST THING IS TO CHECK IF THEY PROVIDE ANY SEARCH CRITERIA

            if (cbxSearchFilterKeyword.Text == "Any" && cbxSearchFilterLocation.Text == "Any" && cbxSearchFilterFamily.Text == "Any" && cbxSearchFilterPerson.Text == "Any" && cbxSearchFilterImageType.Text == "Any" && cbxSearchFilterDateOfPicture.Text == "Any" && tbxSearchFilterJPGpictureName.Text == "")
            {
                //MessageBox.Show("You have chosen not to view all the records in the database.", "SEARCH CRITERIA", MessageBoxButtons.OK);
                DialogResult answer = MessageBox.Show("You entered no criteria.\r\n\nit is recommended that you choose NO and not display all the pictures?", "SEARCH CRITERIA", MessageBoxButtons.YesNo);
                if (answer == DialogResult.Yes)
                {
                    findPicturesWorker();
                }
                else
                {
                    MessageBox.Show("You chose NOT to display ALL THE PICTURES, which you probably didn\'t want to do.","SEARCH FILES", MessageBoxButtons.OK);
                }
            }
            else  // DON'T WANT TO DISPLAY ANYTHING
            {
                findPicturesWorker();
            }
        }

        // BUTON IN SEARCH: FIND PICTURES
        private void btnFindPictures_Click(object sender, EventArgs e)
        {   // FIND AND DISPLAY THE PICTURES
            // REBUILD THE tbxSearchFilterJPGpictureName.Text box with the correct JPG filename.
            if (tbxSearchFilterJPGpictureName.Text != "")       // No need to add .jpg on the end if the field is empty
            { 
                string strJPGname = tbxSearchFilterJPGpictureName.Text.Trim();
                //string[] filenameParts = strJPGname.Split('.');
                //tbxSearchFilterJPGpictureName.Text = filenameParts[0] + ".jpg";            
                strJPGname = strJPGname.Replace(".jpg", "").Replace(".tif", "") + ".jpg";
                tbxSearchFilterJPGpictureName.Text = strJPGname;    // Write the name with .jpg added back into the search field
            }
            findPictures();
        }

        private void clearSelectionCriteriaBoxes()
        { 
            cbxSearchFilterDateOfPicture.Text = "Any";
            cbxSearchFilterDateOfPicture.Text = "Any";
            cbxSearchFilterImageType.Text = "Any";
            cbxSearchFilterKeyword.Text = "Any";
            cbxSearchFilterLocation.Text = "Any";
            cbxSearchFilterPerson.Text = "Any";
            cbxSearchFilterFamily.Text = "Any";
            tbxSearchFilterJPGpictureName.Text = "";        
        }

        // BUTTON IN SEARCH: CLEAR
        private void btnClearSelections_Click(object sender, EventArgs e)
        {   // Set all the comboox fields back to Any
            clearSelectionCriteriaBoxes();
        }

        // MENU ITEM: Scanned Images
        private void addNewRecordToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // DISPLAY THE TOP TRANSCTIONS
            showTopScannedImages_ListBox();
            // SET UP AddNewRecord PANEL
            pnlAddNewRecord.Parent = panel1;
            pnlAddNewRecord.Location = new Point(30, 50);
            pnlAddNewRecord.Visible = true;
          
            // SETUP ANR PANEL AND ITS OTHER BUTTONS, BOXES AND LIST BOX
            pnlAddNewRecord.BringToFront();
            lblANRpreviousAdds.Parent = pnlAddNewRecord;
            lblANRdateOfPicture.Parent = pnlAddNewRecord;
            lblANRdescription.Parent = pnlAddNewRecord;
            lblANRjpgFilename.Parent = pnlAddNewRecord;
            lblANRlocation.Parent = pnlAddNewRecord;
            //lblANRotherSelectors.Parent = pnlAddNewRecord;
            lblANRpersons.Parent = pnlAddNewRecord;
            lblANRrecID.Parent = pnlAddNewRecord;
            lblANRscanName.Parent = pnlAddNewRecord;
            btnANRcancel.Parent = pnlAddNewRecord;
            btnANRclose.Parent = pnlAddNewRecord;
            btnANRperson.Parent = pnlAddNewRecord;
            btnANRsaveRecord.Parent = pnlAddNewRecord;
            btnANRupdateRecord.Parent = pnlAddNewRecord;
            btnANRselectRecIDorJPG.Parent = pnlAddNewRecord;
            tbxANRadultPicture.Parent = pnlAddNewRecord;
            tbxANRdateOfPicture.Parent = pnlAddNewRecord;
            tbxANRdbLogID.Parent = pnlAddNewRecord;
            tbxANRdescription.Parent = pnlAddNewRecord;
            tbxANRjpgFilename.Parent = pnlAddNewRecord;
            tbxANRoriginalName.Parent = pnlAddNewRecord;
            tbxANRpersons.Parent = pnlAddNewRecord;
            tbxANRyouthPicture.Parent = pnlAddNewRecord;
            lblANRpreviousAdds.Parent = pnlAddNewRecord;

            // SET UP OTHER PANELS
            //pnlAddNewRecord.Visible = false;
            pnlKnowledgeBase.Visible = false;
            pnlWelcome.Visible = false;
            pnlSettings.Visible = false;
            pnlEmail.Visible = false;
            pnlLightbox.Visible = false;
            pnlSearchFilter.Visible = false;

            boolInitPhase = false;              // Turn this off since we are passed Initialization if we got here

            // ADJUST BUTTONS ON WAY OUT
            btnANRsaveRecord.Visible = false;       // Nothing to SAVE so don't show it. This get's lit when the START NEW RECORD button clicked
            btnANRupdateRecord.Visible = false;    // Nothing to UPDATE do don't show this either. It will get lit when a record is selected
            btnANRupdateInstruction.Visible = true; // While nothing to update, give direction on how to do so

            btnANRcancel.Text = "START NEW RECORD";
            btnANRcancel.Visible = true;            // We can start a new record however
            tbxANRoriginalName.ReadOnly = false;    // We are starting a new record and the user must be able to enter data.

            btnANRupdateInstruction.Parent = pnlAddNewRecord;
            btnANRupdateInstruction.BringToFront();
            btnANRupdateInstruction.Visible = true;
        }

        private void showSelectedRecord(string strSelectedLine)
        {   /// <summary>
            /// PURPOSE: This method will display the record chosen by the user from the lbxANRpreviousAditionsOrUpdates list box
            /// OUTCOMES:   1) the journal variables will be filled which will be consumed when an UPDATE function is performed (NOTE: not on an ADD)
            ///             2) selected fields will be displayed in the ADD/UPDATE form to allow the user to modify any fields
            ///             3) the ADD NAME will be changed to MODIFY NAME via turning off and on the proper buttons
            /// </summary>          
            //// ENSURE THAT THE PANEL IS VISIBLE AND TO THE FRONT
            //panel1.BringToFront();
            //panel1.Visible = true;
            // MAKE SURE PANEL VISIBLE
            pnlAddNewRecord.Visible = true;
            pnlAddNewRecord.BringToFront();

            // PARSE THE strSelectedLine
            string [] fields = strSelectedLine.Split('|');      // new Char {'|'});
            string strDBLogID = fields[0].Trim();

            // SELECT AND DISPLAY THE CHOSEN RECORD 
            string strSQL = "SELECT DBLogID, OriginalName, Description, Location, DateofPicture, Persons, WebJPGFilename FROM ScannedImages WHERE DBLogID = '" + strDBLogID + "'";
            string strWebJPGFilename = "";
            try
            {
                OleDbConnection cn = new OleDbConnection();
                cn.ConnectionString = strConnectionString;
                cn.Open();
                OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
                OleDbDataReader myDataReader;
                // THERE WILL ONLY BE ONE RECORD (OR IT WILL FAIL)
                myDataReader = myCommand.ExecuteReader();
                // POPULATE THE FIELDS
                myDataReader.Read();
                gblstrDBLogID = myDataReader["DBLogID"].ToString();
                tbxANRoriginalName.Text = myDataReader["OriginalName"].ToString();
                gblstrOriginalName = tbxANRoriginalName.Text;
                tbxANRdescription.Text = myDataReader["Description"].ToString();
                gblstrDescription = tbxANRdescription.Text;
                tbxANRdateOfPicture.Text = myDataReader["DateofPicture"].ToString();
                gblstrDateofPicture = tbxANRdateOfPicture.Text;
                cbxANRlocation.Text = myDataReader["Location"].ToString();
                gblstrLocation = cbxANRlocation.Text;
                tbxANRpersons.Text = myDataReader["Persons"].ToString();
                gblstrPersons = tbxANRpersons.Text;
                strWebJPGFilename = myDataReader["WebJPGFilename"].ToString();
                strWebJPGFilename = strDefaultSkyDriveImages + @"\" + strWebJPGFilename;

                // DISPLAY PICTURE IF THERE IS A WEB PICTURE

                /// TODO: Broad this to also look for picture given original filename (Steve's filename)
                if (File.Exists(strWebJPGFilename))
                {
                    pbxANRwebJPGFilename.WaitOnLoad = true;
                    pbxANRwebJPGFilename.SizeMode = PictureBoxSizeMode.Zoom;
                    pbxANRwebJPGFilename.Load(strWebJPGFilename);
                    pbxANRwebJPGFilename.Visible = true;
                }
                else { pbxANRwebJPGFilename.Visible = false; }

                // CLEAN UP
                myDataReader.Close(); myDataReader.Dispose();
                myCommand.Dispose();
                cn.Close(); cn.Dispose();

                /// TODO: Turn off the ADD NAME and turn on the UPDATE NAME buttons
                // TURN OFF btnANradd
                btnANRsaveRecord.Visible = false;
                btnANRupdateRecord.Visible = true;
                tbxANRoriginalName.ReadOnly = true;         // Need to block this field from update
                btnANRupdateInstruction.Visible = false;
            }
            catch (Exception E1)
            {
                if (E1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                else
                {
                    if (E1.HResult == -2146233079)
                    {   // THE PROVIDED DBLogID number doesn't exist
                        tbxANRdbLogID.Text = "";
                        clearAddNewRecordFields();
                        btnANRupdateInstruction.Visible = true;
                        btnANRupdateRecord.Visible = false;
                    }
                    else
                    {
                        string e1message = "Failed to display the selected entry \'" + strDBLogID + "\'\r\n";
                        e1message += "Error Number: " + E1.HResult + "\r\n" + E1.Message + "\r\n\n";
                        e1message += "An EMail message has been sent.";
                        MessageBox.Show(e1message, "DISPLAY SELECTED ENTRY", MessageBoxButtons.OK);
                        mailProgramError(e1message);
                    }

                }
            }
        }

        private void showTopScannedImages_ListBox()
        {   /// This method will show, in descending order, the last number of Scanned Image records
            /// Because this is a list box, the usr will be able to click on a box and that record will be displayed
            /// --- and --- the SAVE RECORD button will become the UPDATE RECORD button
            /// 
            lbxANRpreviousAditionsOrUpdates.Items.Clear();      // Erase previous data
            try
            {
                string strSQL = "SELECT TOP 200 DBLogID, OriginalName, Description  FROM ScannedImages WHERE CreatedBy = 'Stephanie Cooley' ORDER BY DBLogID DESC";
                OleDbConnection cn = new OleDbConnection();
		        cn.ConnectionString = strConnectionString;
		        cn.Open();
		        OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
		        OleDbDataReader myDataReader;
		        myDataReader = myCommand.ExecuteReader();
                string strDBLogID, strOriginalName, strDescription = "";        // Initialize empty
                while (myDataReader.Read())
                {
                    strDBLogID = myDataReader["DBlogID"].ToString();
                    strOriginalName = myDataReader["OriginalName"].ToString().Trim().Replace("\n","");
                    strDescription = myDataReader["Description"].ToString().Trim().Replace("\n", "");
                    // PAD OR REDUCE THE SIZE OF FIELDS
                    if (strOriginalName.Length < 60) { strOriginalName = strOriginalName.PadRight(60, ' '); } else { strOriginalName = strOriginalName.Substring(0, 56) + " ..."; }
                    if (strDescription.Length < 90) { strDescription = strDescription.PadRight(90, ' '); } else { strDescription = strDescription.Substring(0, 90) + " ..."; }
                    string strListBoxLine = string.Format("{0:5} | {1} | {2}", strDBLogID, strOriginalName, strDescription);
                    lbxANRpreviousAditionsOrUpdates.Items.Add(strListBoxLine);
                }
                // CLEAN UP
                myDataReader.Close(); myDataReader.Dispose();
                myCommand.Dispose();
                cn.Close(); cn.Dispose();
            }
            catch (Exception e1)
            {
                if (e1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                else
                {
                    string e1Message = "Failed to create the list of previous ADDs or UPDATEs.\r\nMessage number: " + e1.HResult;
                    e1Message += "\r\n" + e1.Message;
                    e1Message += "\r\n\nAn EMail message has been sent with this information. No further action required and you may continue";
                    MessageBox.Show(e1Message,"CREATE PREVIOUS ADD/UPDATE LIST",MessageBoxButtons.OK);
                    mailProgramError(e1Message);
                }
            }
        }

        // ANR BUTTON: CLOSE
        private void btnANRclose_Click(object sender, EventArgs e)
        {   // THIS WILL DROP THE ADD-NEW-RECORD PANEL AND LIGHT THE OTHER PANELS
            pnlAddNewRecord.Visible = false;
            //pnlLightbox.Visible = true;
            //pnlSearchFilter.Visible = true;
        }


        private void clearAddNewRecordFields()
        {   // This erases all the fields
            tbxANRdateOfPicture.Text = "";
            cbxANRlocation.Text = "";
            cbxANRpersons.Text = "";
            tbxANRdescription.Text = "";
            tbxANRpersons.Text = "";
            tbxANRoriginalName.Text = "";
            pbxANRwebJPGFilename.Visible = false;
            btnANRupdateInstruction.Visible = false;
            tbxANRjpgFilename.Text = "";
            tbxANRdbLogID.Text = "";
        }

        // ANR BUTTON: Clear Fields
        private void btnANRcancel_Click(object sender, EventArgs e)
        {   // THIS BUTTON WAS ONCE NAMED CANCEL, THEN CLEAR AND NOW START
            clearAddNewRecordFields();
            btnHRupdatePersonRecord.Visible = false;        // Since we are starting a new record, UPDATE is not a choice
            btnANRcancel.Text = "RE-START NEW RECORD";
            btnANRupdateInstruction.Visible = false;
            btnANRupdateRecord.Visible = false;
            btnANRsaveRecord.Visible = true;
            tbxANRoriginalName.ReadOnly = false;            // Since we are starting a new record, the user must be able to enter data here
        }

        // AND BUTTON: SAVE NEW RECORD
        private void btnANRsaveRecord_Click(object sender, EventArgs e)

        {   // SAVE THE RECORD
            if (tbxANRoriginalName.Text == "")  // empty required field
            {
                MessageBox.Show("The Scanned Name field must not be empty.", "ADD NEW PERSON", MessageBoxButtons.OK);
            }
            else   // Critical field has something in it
            {
                string message = "Are you sure you want to add this record?";
                DialogResult answer = MessageBox.Show(message,"ADD RECORD",MessageBoxButtons.YesNo);
                if (answer == DialogResult.Yes)
                {   // WE HAVE THE GO AHEAD TO WRITE THE RECORD 
                    DateTime dtCreatedModifiedDate = DateTime.Now;
                    try  // to save the record
                    {
                        string strSQL = "INSERT INTO ScannedImages (";
                        strSQL += "Created, Modified, OriginalName, Description, "; // 1-4
                        strSQL += "DateScanned, ScanResolution, DateofPicture, ";   // 5-7
                        strSQL += "PictureType, OriginalMedia, Color, ";            // 8-10
                        strSQL += "Persons, Location, ";                           // 11, 12
                        strSQL += "CreatedBy, ModifiedBy, SharepointID, ";          // 13-15
                        strSQL += "Photographer, CameraScannerID, OriginalImageName, SPImageAvailable, NoSPData, "; // 16-20
                        strSQL += "DuplicateImage, DuplicateImageID, ";             // 21,  22
                        strSQL += "ActionRequired, ActionRequiredNotes, ";          // 23, 24
                        strSQL += "EditedImageAvailable, EditedFilename, ";            // 25, 26
                        strSQL += "WebImageAvailable, WebJPGFilename, DateToWeb, "; // 27-29
                        strSQL += "WebDescription, FamilyID, WebPersonID, WebLocation, ";   // 30-33
                        strSQL += "Keywords, PSDFileAvailable, ImageType, SPImage, AccessDBLogID, ActionRequiredBy) VALUES (";  // 34-39
                        strSQL += "'" + dtCreatedModifiedDate.ToString() + "', ";   // 1 Created
                        strSQL += "'" + dtCreatedModifiedDate.ToString() + "', ";   // 2 Modified
                        strSQL += "'" + tbxANRoriginalName.Text.Trim().Replace("'","''") + "', ";  // 3 OriginalName
                        strSQL += "'" + tbxANRdescription.Text.Trim().Replace("'","''") + "', ";   // 4 Description
                        strSQL += "'" + dtCreatedModifiedDate.ToString().Trim().Replace("'","''") + "', ";   // 5 DateScanned
                        strSQL += "'3200', ";                                               // 6 ScanResolution defaults to 3200
                        strSQL += "'" + tbxANRdateOfPicture .Text.Trim().Replace("'", "''") + "', ";                  // 7 DateOfPicture
                        strSQL += "'TIF', ";                                                // 8 PictureType defaults to TIF
                        strSQL += "'Print', ";                                              // 9 OriginalMedia defaults to Print
                        strSQL += "'Black and White', ";                                    // 10 Color defaults to Black and White
                        strSQL += "'" + tbxANRpersons.Text.Trim().Replace("'","''") + "', ";                         // 11 Persons
                        strSQL += "'" + cbxANRlocation.Text.Trim().Replace("'", "''") + "', ";                        // 12 Location
                        strSQL += "'Stephanie Cooley', ";                                   // 13 CreatedBy defaults to Stephanie Cooley
                        strSQL += "'Stephanie Cooley', ";                                   // 14 ModifiedBy defaults to Stephanie Cooley
                        strSQL += "'0', ";                                                  // 15 SharepointID defaults to 0 since this record not created by Sharepoint
                        strSQL += "'', ";                                                   // 16 Photographer defaults to blank since the photographer is probably unknown
                        strSQL += "'Epson V700 Photo Scanner', ";                           // 17 CameraScannerID deafuts to this because it is what the Farm Docent has
                        strSQL += "'', ";                                                   // 18 OriginalImageName defaults to empty-string because Reviewer creates this
                        strSQL += "'0', ";                                                  // 19 SPImageAvailable defaults to false
                        strSQL += "'0', ";                                                  // 20 NoSPData defaults to true, no SP date because this data entered by RanchArchive and not through the Sharepoint program
                        strSQL += "'0', ";                                                  // 21 DuplicateImage defaults to false, ie, this isn't a duplicate image
                        strSQL += "'', ";                                                   // 22 DuplicateImageID defaults to empty-string as this isn't a duplicate
                        strSQL += "'1', ";                                                  // 23 ActionRequired defaults to true because Steve needs to review record
                        strSQL += "'Record added \"" + tbxANRoriginalName.Text.Trim().Replace("'","''") + "\" created on " + dtCreatedModifiedDate.ToShortDateString() + " needs review.', ";                   // 24 ActionRequredNotes defaults to this because the Reviewer needs to create some fields during review
                        strSQL += "'0', ";                                                  // 25 EditedImageAvailable defaults to false since this is a new record
                        strSQL += "'', ";                                                   // 26 EditedFilename defaults to empty-string since this is a new record
                        strSQL += "'0', ";                                                  // 27 WebImageAvailable defaults to false since this is a new record
                        strSQL += "'', ";                                                   // 28 WebJPGFilename defaults to empty-string since this is a new reocrd
                        strSQL += "'1/1/2000', ";                                           // 29 DateToWeb defaults to this bogus date which signifies not on Web
                        strSQL += "'', ";                                                   // 30 WebDescription defaults to empty-string because it is added by Reviewer
                        strSQL += "'', ";                                                   // 31 FamilyID defaults to empty-string because it is added by Reviewer
                        strSQL += "'', ";                                                   // 32 WebPersonID defaults to empty-string because it is added by Reviewer
                        strSQL += "'', ";                                                   // 33 WebLocation defaults to empty-string because it is added by Reviewer
                        strSQL += "'', ";                                                   // 34 Keywords defaults to empty-string because it is added by Reviewer
                        strSQL += "'0', ";                                                  // 35 PSDFileAvailable defaults to false because this is a new record
                        strSQL += "'Photograph', ";                                         // 36 ImageType defaults to Photograph because that is the most likely case
                        strSQL += "'No', ";                                                 // 37 SPImage defaults to No because the field has been refactored (check with Steve)
                        strSQL += "'0', ";                                                  // 38 AccessDBLogID defaults to 0 to signify that this was never in the Access DB
                        strSQL += "'Steven Breckner')";                                      // 39 ActionRequiredBy defaults to Steve Breckner, the Reviewer when this written
                        // EXECUTE THE INSERT
                        OleDbConnection cn = new OleDbConnection();
                        cn.ConnectionString = strConnectionString;
                        cn.Open();
                        OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
                        myCommand.ExecuteNonQuery();
                    
                        // REWRITE THE LIST BOX OF RECENT ENTRIES
                        showTopScannedImages_ListBox();

                        // CLEAR THE FIELDS
                        clearAddNewRecordFields();

                        // ADJUST BUTTONS following successful ADD NEW RECORD (ANR)
                        btnANRupdateRecord.Visible = false;             // Nothing to update
                        btnANRupdateInstruction.Visible = true;         // But show the UPDATE Instruction button
                        btnANRcancel.Text = "START A NEW RECORD";        // We can now CREATE NEW RECORD
                        btnANRcancel.Visible = true;
                    }
                    catch (Exception e1)
                    {
                        if (e1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                        else
                        {
                            string e1message = "[104:E1] Unable to save this record now. Try again later. \r\n\n(" + e1.HResult + ") " + e1.Message;
                            e1message += "\r\nEmail has been sent and your work has been journaled.";
                            MessageBox.Show(e1message,"SAVE NAME",MessageBoxButtons.OK);

                            // JOURNAL THE FAILURE
                            DateTime dtFailureDateTime = DateTime.Now;
                            //string formatedTodaysDateTime = string.Format("{0:G}", todaysDateTime);
                            string strSQL2 = "INSERT INTO ScannedImagesRetryJournal (";
                            strSQL2 += "retryFailedDate, retrySQL, retryMessageBoxShow, retryErrorCode, retryHResult, retryMessage) VALUES (";
                            strSQL2 += "'" + dtFailureDateTime.ToString() + "', ";  // 1 retryFailedDate
                            strSQL2 += "'" + strSQL2.Replace("'", "''") + "', ";             // 2 retrySQL
                            //strSQL2 += "'SQL statement goes here', ";
                            strSQL2 += "'" + e1message.Replace("'", "''") + "', ";          // 3 retryMessageBoxShow
                            //strSQL2 += "'the message box text I created', ";
                            strSQL2 += "'" + "104:E1" + "', ";                              // 4 retryErrorCode
                            strSQL2 += "'" + e1.HResult.ToString() + "', ";                 // 5 retryHResult
                            strSQL2 += "'" + e1.Message.ToString().Replace("'","''") + "')";                   // 6 retryMessage
                            OleDbConnection cn2 = new OleDbConnection();
                            cn2.ConnectionString = strConnectionString;
                            cn2.Open();
                            OleDbCommand myCommand = new OleDbCommand(strSQL2, cn2);
                            myCommand.ExecuteNonQuery();

                            // SEND EMAIL
                            mailProgramError(e1message);

                            // ADJUST BUTTONS following successful ADD NEW RECORD (ANR)
                            btnANRupdateRecord.Visible = false;             // Nothing to update
                            btnANRupdateInstruction.Visible = true;         // But show the UPDATE Instruction button
                            btnANRcancel.Text = "START A NEW RECORD";       // We can now CREATE NEW RECORD
                            btnANRcancel.Visible = true;
                        }
                    }
                }
                else  // NO SELECTED to confirmation message
                {
                }
            }
        }

        // MENU ITEM: SEARCH
        private void searchToolStripMenuItem_Click(object sender, EventArgs e)
        {   // DARKEN panels not needed and LIGHTEN those needed
            // SET UP LIGHTBOX PANEL
            pnlLightbox.Visible = true;
            pnlLightbox.BringToFront();
            pnlLightbox.BackColor = Color.WhiteSmoke;
            pnlLightbox.Parent = panel1;
            pnlSearchFilter.Visible = true;
            pnlSearchFilter.BringToFront();

            // SET UP OTHER PANELS
            pnlAddNewRecord.Visible = false;
            pnlKnowledgeBase.Visible = false;
            pnlWelcome.Visible = false;
            pnlSettings.Visible = false;
            pnlEmail.Visible = false;
            btnANRupdateInstruction.Visible = false;
            //pnlLightbox.Visible = false;
            //pnlSearchFilter.Visible = false;

            //lbxPictureResults.Visible = true;
            boolInitPhase = false;              // Turn this off since we are passed Initialization if we got here
        }

        private void btnAddlocation_Click(object sender, EventArgs e)
        {   // Try for poping up a new panel that is brought forward and then when the SAVE button is pressed drop the panel.
            string message = "Do you still want to add \r\n\n\'" + cbxANRlocation.Text.Trim().ToUpper() + "\'\r\n\n as a NEW LOCATION?";
            DialogResult answer = MessageBox.Show(message, "ADD NEW LOCATION", MessageBoxButtons.YesNo);
            if (answer == DialogResult.Yes)  // yes, add it
            {
                try
                {
                    string strSQL = "INSERT INTO WebLocations (Location) VALUES ('" + cbxANRlocation.Text.Trim() + "')";
                    OleDbConnection cn = new OleDbConnection();
                    cn.ConnectionString = strConnectionString;
                    cn.Open();
                    OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
                    myCommand.ExecuteNonQuery();
                    string successMessge = "You have successfully ADDED \'" + cbxANRlocation.Text.Trim() + "\' to the Location list";
                    MessageBox.Show(successMessge, "ADD NEW LOCATION",MessageBoxButtons.OK);

                    // REFILL THE LOCATION COMBO BOXES
                    fillComboBox("Location", "WebLocations", "Location", cbxSearchFilterLocation);
                    fillComboBox("Location", "WebLocations", "Location", cbxANRlocation);
                    sendEmail("robert.perry@hoferranch.onmicrosoft", "robert.perry@perry.onmicrosoft.com", "Notification of NEW Location", "The new location " + cbxANRlocation.Text.Trim().ToUpper() + " added to the WebLocation table");
                }
                catch (Exception e1)
                {
                    string e1message = "";
                    switch (e1.HResult)
                    {
                        case -21767259: e1message = "You have possibly lost connection to the network. Try again late."; break;
                        default: message = "Failed to enter \'" + cbxANRlocation.Text.Trim().ToUpper() + "\' to the LOCATION list.\r\n";
                                 message += "Error number: " + e1.HResult + "\r\n" + e1.Message;
                                 MessageBox.Show(e1message, "ADD NEW LOCATION",MessageBoxButtons.OK);
                                 break;
                    }
                    mailProgramError(message);
                }
            }
            else    // no, don't add it
            {
                message = "You have chosen NOT to ADD \'" + cbxANRlocation.Text.ToUpper() + "\n as a NEW LOCATION?";
                MessageBox.Show(message, "ADD NEW LOCATION", MessageBoxButtons.OK);
            }
        }

        private void btnANRperson_Click(object sender, EventArgs e)
        {   // Try for poping up a new panel that is brought forward and then when the SAVE button is pressed drop the panel.
            showKnowledgeBase();
        }

        // ANR ROW SELECTED
        private void cbxANRpersons_SelectedIndexChanged(object sender, EventArgs e)
        {   // ADD THE SELECTED PERSON TO THE text box tbxANRpersons
            if (boolInitPhase == false)
            {
                if (tbxANRpersons.Text.Trim() == "") { tbxANRpersons.Text = cbxANRpersons.Text; }
                else
                {
                    tbxANRpersons.Text += "; " + cbxANRpersons.Text;
                }            
            }
        }

        private void lbxANRpreviousAditionsOrUpdates_SelectedIndexChanged(object sender, EventArgs e)
        {   /// <summary>
            ///     PURPOSE: Call the routine to parse the returned value
            /// </summary>
            string strSelectedLine = lbxANRpreviousAditionsOrUpdates.Text;
            showSelectedRecord(strSelectedLine);
            btnANRupdateInstruction.Visible = false;
            btnANRupdateRecord.Visible = true;
            //tbxANRoriginalName.ReadOnly = true;         // This shouldn't be changes. See Bug 86 // NOW DELT WITH IN THE SHOW RECORD METHOD
        }

        // UPDATE BUTTON CLICKED IN ADD OR UPDATE SCANNED ...
        private void btnANRupdateRecord_Click(object sender, EventArgs e)
        {   // WRITE BACK THE RECORD AND JOURNAL
            /// TODO: 1) Display the Challenge Box
            ///         2) If OK then make the journal entry using the jnl variables
            ///         3) Update the reocrd
            ///         4) Turn off the UPDATE RECORD button
            ///         5) Turn on the SAVE RECORD button
            ///         6) Place the cursor in the first field
            // GET CONFIRMATION

            if (tbxANRoriginalName.Text == "")  // empty required field
            {
                MessageBox.Show("The Scanned Name field must not be empty.","ADD NEW PERSON", MessageBoxButtons.OK);
            }
            else
	        {
                DialogResult answer = MessageBox.Show("Do you still want to update this record?","UPDATE RECORD",MessageBoxButtons.YesNo);
                if (answer == DialogResult.Yes)
                { 
                    // JOURNAL THE BEFORE CHANGE INFORMATION
                    try
                    {
                        // UPDATE BUTTON(S)
                        btnANRsaveRecord.Visible = false;
                        btnANRupdateRecord.Visible = true;
                        DateTime dtNow = DateTime.Now;
                        string sqlJournal = "INSERT INTO ScannedImageJournal (";
                        sqlJournal += "sijDBLogID, sijJournalDate, sijOriginalName, sijDescription, sijLocation, sijDateofPicture, sijPersons) VALUES (";
                        sqlJournal += "'" + gblstrDBLogID.ToString() + "', ";
                        sqlJournal += "'" + dtNow.ToString() + "', ";
                        sqlJournal += "'" + gblstrOriginalName.ToString().Trim().Replace("'","''").Replace("\n","") + "', ";
                        sqlJournal += "'" + gblstrDescription.ToString().Trim().Replace("'", "''").Replace("\n", "") + "', ";
                        sqlJournal += "'" + gblstrLocation.ToString().Trim().Replace("'", "''").Replace("\n", "") + "', ";
                        sqlJournal += "'" + gblstrDateofPicture.ToString().Trim().Replace("'", "''").Replace("\n", "") + "', ";
                        sqlJournal += "'" + gblstrPersons.ToString().Trim().Replace("'", "''").Replace("\n", "") + "')";
                        OleDbConnection cn = new OleDbConnection();
                        cn.ConnectionString = strConnectionString;
                        cn.Open();
                        OleDbCommand myCommand = new OleDbCommand(sqlJournal, cn);
                        myCommand.ExecuteNonQuery();
                        try
                        {   // UPDATE THE RECORD
                            string strSQL = "UPDATE ScannedImages SET ";
                            strSQL += "ModifiedBy = 'Stephanie Cooley', ";
                            strSQL += "Modified = '" + dtNow.ToString() + "', ";
                            strSQL += "OriginalName = '" + tbxANRoriginalName.Text.Trim().Replace("'", "''").Replace("\n", "") + "', ";
                            strSQL += "Description = '" + tbxANRdescription.Text.Trim().Replace("'", "''").Replace("\n", "") + "', ";
                            strSQL += "Location = '" + cbxANRlocation.Text.Trim().Replace("'", "''").Replace("\n", "") + "', ";
                            strSQL += "DateofPicture = '" + tbxANRdateOfPicture.Text.Trim().Replace("'", "''").Replace("\n", "") + "', ";
                            strSQL += "Persons = '" + tbxANRpersons.Text.Trim().Replace("'", "''").Replace("\n", "") + "', ";
                            strSQL += "ActionRequired = '1', ";
                            strSQL += "ActionRequiredBy = 'Steven Breckner', ";
                            strSQL += "ActionRequiredNotes = 'Review required because critical fields for DBLogID ''" + gblstrDBLogID + "'' may have changed.' ";
                            strSQL += "WHERE DBLogID = '" + gblstrDBLogID + "'";
                            OleDbConnection cn2 = new OleDbConnection();
                            cn2.ConnectionString = strConnectionString;
                            cn2.Open();
                            OleDbCommand myCommand2 = new OleDbCommand(strSQL, cn2);
                            myCommand2.ExecuteNonQuery();
                            myCommand2.Dispose();
                            cn2.Close(); cn2.Dispose();
                            // CHANGE BUTTON
                            btnANRupdateRecord.Visible = false;
                            btnANRsaveRecord.Visible = true;
                            // CLEAR FIELDS
                            clearAddNewRecordFields();
                            // REFRESH LISTBOX
                            showTopScannedImages_ListBox();
                            statusStrip1.Text = "UPDATE of \'" + tbxANRpersons.Text + "\' record successful";

                            // RESET ANR BUTTONS
                            btnANRsaveRecord.Visible = false;
                            btnANRupdateInstruction.Visible = true;
                            btnANRupdateRecord.Visible = false;
                        }
                        catch (Exception e2)
                        {
                            if (e2.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                            else
                            {
                                string e1message2 = "[103:E2] Failed to UPDATE the record.\r\n\n(" + e2.HResult + ")\r\n" + e2.Message;
                                mailProgramError(e1message2);
                                statusStrip1.Text = "UPDATE of \'" + tbxANRpersons.Text + "\' record failed";
                            }
                        }
                    }
                    catch (Exception e1)
                    {

                        string e1message = "[103:E1] Failed to journal the before change ScannedRecord fields and the record has NOT been UPDATED either.\r\n\n(" + e1.HResult + ")\r\n" + e1.Message;
                        mailProgramError(e1message);
                    }
                }
                else
                {
                    MessageBox.Show("You chose NOT to UPDATE this record.", "UPDATE RECORD", MessageBoxButtons.YesNo);
                }

	        }
        }


        public void sendEmailToolStripMenuItem_Click(object sender, EventArgs e)
        {   // CREATE MAIL MESSAGE, first creating the ADDRESS BOOK if not already created
            // SET UP OTHER PANELS
            pnlKnowledgeBase.Visible = false;
            pnlLightbox.Visible = false;
            pnlSearchFilter.Visible = false;
            pnlWelcome.Visible = false;
            pnlSettings.Visible = false;
            pnlAddNewRecord.Visible = false;
            btnANRupdateInstruction.Visible = false;

            //SET UP ADDRESS
            String[] globalAddressBookArrayURI = new string[50];
            //string globalTest = "Global Test";
            //globalTest = "Global Test";

           // TURN THE PANEL ON
            pnlEmail.Parent = panel1;
            pnlEmail.Location = new Point(30, 50);
            pnlEmail.BackColor = Color.OldLace;
            pnlEmail.Visible = true;
            pnlEmail.BringToFront();

            // CHECK IF THERE IS AN ADDRESS BOOK
            if (!boolIsMailAddressList)
            {   // NO ADDRESS BOOK, BUILD IT
                try  // to read records from the MailAddresses table
                {
                    string strSQL = "SELECT COUNT(mailID) AS addressBookCount FROM MailAddresses";
                    OleDbConnection cn = new OleDbConnection();
                    cn.ConnectionString = strConnectionString;
                    cn.Open();
                    OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
                    OleDbDataReader myDataReader;
                    myDataReader = myCommand.ExecuteReader();
                    myDataReader.Read();
                    int intCount = Convert.ToInt32(myDataReader["addressBookCount"].ToString());
                    myDataReader.Close(); myDataReader.Dispose();
                    myCommand.Dispose();
                    cn.Close(); cn.Dispose();

                    //var comboBoxArray = new string[intCount, 2];
                    var comboBoxArray = new string[intCount];         // moved to global
                    var comboBoxCCarray = new string[intCount];

                    strSQL = "SELECT mailVisualName, mailURI FROM MailAddresses ORDER BY mailVisualName";
                    OleDbConnection cn2 = new OleDbConnection();
                    cn2.ConnectionString = strConnectionString;
                    cn2.Open();
                    OleDbCommand myCommand2 = new OleDbCommand(strSQL, cn2);
                    OleDbDataReader myDataReader2;
                    myDataReader2 = myCommand2.ExecuteReader();
                    comboBoxArray[0] = "";
                    comboBoxCCarray[0] = "";
                    globalAddressBookArrayURI[0] = "";
                    for (int i = 1; i < intCount; i++)
                    {
                        myDataReader2.Read();
                        //comboBoxArray[i] = myDataReader2["mailVisualName"].ToString();
                        comboBoxArray[i] = myDataReader2["mailURI"].ToString();
                        comboBoxCCarray[i] = comboBoxArray[i];
                        globalAddressBookArrayURI[i] = myDataReader2["mailURI"].ToString();
                        
                        //comboBoxArray[i, 0] = myDataReader2["mailVisualName"].ToString();
                        //comboBoxArray[i, 1] = myDataReader2["mailURI"].ToString();
                    }
                    myDataReader2.Close(); myDataReader2.Dispose();
                    myCommand2.Dispose();
                    cn.Close(); cn.Dispose();
                    // Fill the Combo Boxes
                    cbxEmailTo.DataSource = comboBoxArray;
                    cbxEmailCC.DataSource = comboBoxCCarray;
                    boolIsMailAddressList = true;
                }
                catch (Exception e1)
                {
                    if (e1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                    else
                    {
                        string e1Message = "ERROR building Address Book\r\nMessage Number: " + e1.HResult + "\r\n" + e1.Message;
                        MessageBox.Show(e1Message,"BUILD ADDRESS BOOK", MessageBoxButtons.OK);
                        mailProgramError(e1Message);
                    }
                }
            }
        }

        private void btnEmailCancel_Click(object sender, EventArgs e)
        {   // CANCEL EMAIL MESSAGE by clearing the fields and closing the panel
            tbxEmailTo.Text = "";
            tbxEmailCC.Text = "";
            tbxEmailSubject.Text = "";
            rtbEmailBody.Clear();
            pnlEmail.Visible = false;
        }

        private void SendCompleteEventHandler(object sender, EventArgs e)
        {
            string message = "Mail sent.";
            tsEmailStatus.ForeColor = Color.White;
            tsEmailStatus.BackColor = Color.DarkBlue;
            tsEmailStatus.Text = message;
        }

        private void mailProgramError(string strMessageBoxMessage)
        {   // BUILD THE EMAIL
            DateTime todaysDateTime = DateTime.Now;
            string formatedTodaysDateTime = string.Format("{0:G}", todaysDateTime);
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Port = 587;
            smtpClient.Host = "smtp.office365.com";
            smtpClient.EnableSsl = true;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential("robert.perry@hoferranch.onmicrosoft.com", "");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.SendCompleted += new SendCompletedEventHandler(SendCompleteEventHandler);
            MailAddress from = new MailAddress("robert.perry@hoferranch.onmicrosoft.com", "RanchArchive Error Reporting");

            int intTOindex = cbxEmailTo.SelectedIndex;
            int intCCindex = cbxEmailCC.SelectedIndex;
            //string strTO = globalAddressBookArrayURI[intTOindex];
            MailAddress to = new MailAddress("Robert.Perry@Hoferranch.onmicrosoft.com", "Rob Perry");
            MailAddress cc = new MailAddress("SBreck@outlook.com", "Steve Breckner");
            MailMessage mmMessage = new MailMessage(from, to);
            mmMessage.Subject = "RanchArchive has logged an error";
            mmMessage.IsBodyHtml = false;
            mmMessage.CC.Add(cc);
            string strMessaegBody = "Date of Error: " + formatedTodaysDateTime + "\r\n\n";
            strMessaegBody += "MessageBox.Show: " + strMessageBoxMessage + "\r\n\n";
            //strMessaegBody += "This has been journaled.";
            mmMessage.Body = strMessaegBody;

            // SEND THE EMAIL
            /// TODO: Somewhere here we must check for an empty TO field and abort sending
            try  // to send the email
            {
                if (cbxEmailTo.Text != "")  // ONLY SEND IF To IS NOT AN EMPTY STRING
                {
                    smtpClient.SendAsync(mmMessage, "");
                    tsEmailStatus.ForeColor = Color.White;
                    tsEmailStatus.BackColor = Color.DarkBlue;
                    tsEmailStatus.Text = "Sending mail to: \'" + cbxEmailTo.Text + "\'";
                    if (cbxEmailCC.Text != "") { tsEmailStatus.Text += " and to \'" + cbxEmailCC.Text + "\'"; }
                    pnlEmail.Visible = false;   // We have sent the mail, close the panel
                }
                else
                {
                    tsEmailStatus.ForeColor = Color.White;
                    tsEmailStatus.BackColor = Color.DarkBlue;
                    tsEmailStatus.Text = "No email notification sent.";
                }
            }
            catch (Exception eMail)
            {
                if (eMail.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                else
                {
                    string message = "Failed to send email about the error. Please record the information and email from your regular EMail program.";
                    message += "\r\nError Number: " + eMail.HResult;
                    message += "\r\n" + eMail.Message;
                    MessageBox.Show(message, "ERROR SENDING EMAIL",MessageBoxButtons.OK);
                    mailProgramError(message);
                }
            }
        }

        private void btnEmailSend_Click(object sender, EventArgs e)
        {   // BUILD THE EMAIL
            DateTime todaysDateTime = DateTime.Now;
            string formatedTodaysDateTime = string.Format("{0:G}", todaysDateTime);
            SmtpClient smtpClient = new SmtpClient();
            smtpClient.Port = 587;
            smtpClient.Host = "smtp.office365.com";
            smtpClient.EnableSsl = true;
            smtpClient.UseDefaultCredentials = false;
            smtpClient.Credentials = new NetworkCredential("robert.perry@hoferranch.onmicrosoft.com", "Cucamonga.2011");
            smtpClient.DeliveryMethod = SmtpDeliveryMethod.Network;
            smtpClient.SendCompleted += new SendCompletedEventHandler(SendCompleteEventHandler);
            MailAddress from = new MailAddress("robert.perry@hoferranch.onmicrosoft.com", "RanchArchive Mailer");

            int intTOindex = cbxEmailTo.SelectedIndex;
            int intCCindex = cbxEmailCC.SelectedIndex;
            //string strTO = globalAddressBookArrayURI[intTOindex];
            MailAddress to = new MailAddress(cbxEmailTo.Text, "Principal receipent");
            MailAddress cc = new MailAddress(cbxEmailCC.Text, "Carbon Copy");
            MailMessage mmMessage = new MailMessage(from, to);
            mmMessage.Subject = tbxEmailSubject.Text;
            mmMessage.IsBodyHtml = false;
            mmMessage.CC.Add(cc);
            mmMessage.Body = rtbEmailBody.Text;

            // SEND THE EMAIL
            /// TODO: Somewhere here we must check for an empty TO field and abort sending
            try  // to send the email
            {
                smtpClient.SendAsync(mmMessage, "");
                tsEmailStatus.Text = "Sending mail to: \'" + cbxEmailTo.Text + "\'";
                if (cbxEmailCC.Text != "") { tsEmailStatus.Text += " and to \'" + cbxEmailCC.Text + "\'";}
                pnlEmail.Visible = false;   // We have sent the mail, close the panel
            }
            catch (Exception eMail)
            {
                if (eMail.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                else
                {
                    string message = "Failed to send email to \'" + tbxEmailTo.Text +  "\'" ;
                    if (cbxEmailCC.Text != "") { message += " and to \'" + cbxEmailCC.Text + "\'";}
                    message += ".";
                    message += "\r\nError Number: " + eMail.HResult;
                    message += "\r\n" + eMail.Message;
                    mailProgramError(message);
                }
            }
        }

        private void updateProgramToolStripMenuItem_Click(object sender, EventArgs e)
        {   // UPDATE PROGRAM
            DialogResult answer =  MessageBox.Show("Do you wish to continue updating this program?\r\nTHIS PROGRAM WILL STOP AND AUTOMATICALLY RESTART", "PROGRAM UPDATE/REFRESH", MessageBoxButtons.YesNo);
            if (answer == DialogResult.Yes)
            {
                Process p = new Process();
                p.StartInfo.FileName = strRanchArchiveUpdate;
                p.StartInfo.UseShellExecute = false;
                p.StartInfo.RedirectStandardOutput = true;
                p.Start();
                this.Close();
            }
            else
            {
                MessageBox.Show("You chose to not update/refresh the program at this time.", "Program Update/Refresh", MessageBoxButtons.OK);
            }
        }

        //
        // KNOWLEDGE BASE
        // ==============
        private void showKnowledgeBase()
        {   /// <summary>
            ///     ENTRY INTO KNOWLEDGE BASE PANEL
            ///     1. Display the panel (it will be turned off when the CLOSE button is chosen)
            ///     2. If the first time, build the combo box
            ///     3. Adjust the button visibility
            /// </summary>
            // ADJUST pnlKnowledgeBase
            pnlKnowledgeBase.Parent = panel1;           // Associate panel with parent
            pnlKnowledgeBase.Location = new Point(30, 50);
            pnlKnowledgeBase.Visible = true;            // Show ourselves
            pnlKnowledgeBase.BackColor = Color.OldLace;
            pnlKnowledgeBase.BringToFront();

            // SET OTHER PANELS
            pnlSearchFilter.Visible = false;            // Don't show
            pnlLightbox.Visible = false;
            pnlSettings.Visible = false;
            pnlLightbox.Visible = false;
            pnlAddNewRecord.Visible = false;

            // ADJUST WHAT BUTTONS ARE SHOWING
            btnANRupdateInstruction.Visible = false;
            btnHRstartNewRecord.Visible = true;
            btnHRupdatePersonRecord.Visible = false;
            btnHRaddPersonRecord.Visible = false;
            btnHRclose.Visible = true;
            btnHRupdateInstructions.Visible = true;     // Tells how to start update.

            ClearHistoryRecordFields();
            if (userMachineName == "DV-2" || userMachineName == "PC-07" || userMachineName == "Rob")
            {
                tbxANRyouthPicture.Visible = true; tbxANRadultPicture.Visible = true;
            }
            else
            {
                tbxANRyouthPicture.Visible = false; tbxANRadultPicture.Visible = false;
            }
            // CHECK IF THE COMBO BOX HAS BEEN BUILT, AND IF NOT, BUTILD IT
            if (cbxHRweNamList.DataSource == null)
            {   // BUILD THE BOX
                fillComboBox("WeNamList", "WebNames", "WeNamList", cbxHRweNamList);
            }
        }

        // KNOWLEDGE BASE: AN EXISTING PERSON RECORD SELECTED
        private void cbxHRweNamList_SelectedIndexChanged(object sender, EventArgs e)
        {   /// <summary>
            ///     USER HAS CHOSEN A NAME FROM THE cbxHRweNamList COMBO BOX
            ///     1. Display the record
            ///     2. Save the values for possible journaling
            ///     3. Adjust the button visibility
            /// </summary>
            // BUILD OUT THE COMBO BOX  IF NOT ALREADY BUILT
            //btnHRstartNewRecord.Visible = true;
            //btnHRstartNewRecord.Text = "CANCEL UPDATE && START NEW RECORD";
            //btnHRaddPersonRecord.Visible = false;
            //btnHRupdatePersonRecord.Visible = true;
           

            // LOOK UP THE SELECTED RECORD
            try  // to read the selected record
            {
                if (cbxHRweNamList.Text == "Any" || cbxHRweNamList.Text == "")  // then we are ok to proceed
                { }
                else // There must be a valid name in the selected field
                {
                    string strSQL = "SELECT * FROM WebNames WHERE WeNamList ='" + cbxHRweNamList.Text + "'";
                    OleDbConnection cn = new OleDbConnection();
                    cn.ConnectionString = strConnectionString;
                    cn.Open();
                    OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
                    OleDbDataReader myDataReader;
                    myDataReader = myCommand.ExecuteReader();
                    myDataReader.Read();
                    // FILL OUT THE FIELDS
                    njWebNameID = ""; // myDataReader["WebNameID"].ToString();
                    njWeNamList = myDataReader["WeNamList"].ToString();
                    njNameSortOrder = myDataReader["NameSortOrder"].ToString();
                    tbxHRnameSortOrder.Text = njNameSortOrder;
                    njWebNameID = myDataReader["MiddleName"].ToString();
                    tbxHRmiddleName.Text = njWebNameID;
                    njMaidenName = myDataReader["MaidenName"].ToString();
                    tbxHRmaidenName.Text = njMaidenName;
                    njOtherNameRef = myDataReader["OtherNameRef"].ToString();
                    tbxHRotherNamedRef.Text = njOtherNameRef;
                    njBirthDate = myDataReader["BirthDate"].ToString();
                    tbxHRbirthDate.Text = njBirthDate;
                    njBirthPlace = myDataReader["BirthPlace"].ToString();
                    tbxHRbirthPlace.Text = njBirthPlace;
                    njIntermentLocation = myDataReader["IntermentLocation"].ToString();
                    tbxHRintermentLocation.Text = njIntermentLocation;
                    njDateDied = myDataReader["DateDied"].ToString();
                    tbxHRdateDied.Text = njDateDied;
                    njMother = myDataReader["Mother"].ToString();
                    tbxHRmother.Text = njMother;
                    njFather = myDataReader["Father"].ToString();
                    tbxHRfather.Text = njFather;
                    njSiblings = myDataReader["Siblings"].ToString();
                    tbxHRsiblings.Text = njSiblings;
                    njSpouse = myDataReader["Spouse"].ToString();
                    tbxHRspouse.Text = njSpouse;
                    njChildrenNames = myDataReader["ChildrenNames"].ToString();
                    tbxHRchildrenNames.Text = njChildrenNames;

                    njKnowledgeBase = myDataReader["KnowledgeBase"].ToString();
                    rtbHRknowledgeBase.Text = njKnowledgeBase;

                    njYouthPicture = myDataReader["YouthPicture"].ToString().Trim();
                    tbxANRyouthPicture.Text = njYouthPicture;
                    njAdultPicture = myDataReader["AdultPicture"].ToString().Trim();
                    tbxANRadultPicture.Text = njAdultPicture;

                    // SHOULD THE PICTURE CUTLINES BE MADE VISIBLE?
                    if (userMachineName == "DV-2" || userMachineName == "PC-09" || userMachineName == "Rob")
                    {
                        tbxANRyouthPicture.Visible = true;
                        tbxANRadultPicture.Visible = true;
                    }

                    // DISPLAY PICTURE IF THERE ARE ANY TO DISPLAY
                    if (njYouthPicture != "")   // display picture and cutline
                    {
                        string skyYouthJPG = strDefaultSkyDriveImages + @"\" + njYouthPicture;
                        if (File.Exists(skyYouthJPG)) { pbxHRyouthPicture.Load(skyYouthJPG); pbxHRyouthPicture.Visible = true; }
                    }
                    else  // make invisible the picture box and but leave the cutline because it can be modified
                    {
                        pbxHRyouthPicture.Visible = false;
                    }
                    if (njAdultPicture != "")   // display picture and cutline
                    {
                        string skyAdultJPG = strDefaultSkyDriveImages + @"\" + njAdultPicture;
                        if (File.Exists(skyAdultJPG)) { pbxHRadultPicture.Load(skyAdultJPG); pbxHRadultPicture.Visible = true; }
                    }
                    else  // make invisible the picture box but leave the cutline box because it can be modified.
                    {
                        pbxHRadultPicture.Visible = false;
                    }
                    njType = "UPDATE";

                    // ADJUST BUTTONS
                    btnHRstartNewRecord.Visible = true;
                    btnHRupdatePersonRecord.Visible = true;
                    btnHRaddPersonRecord.Visible = false;
                    btnHRclose.Visible = true;
                    btnHRupdateInstructions.Visible = false;        // Show just START NEW and UPDATE
                }
            }
            catch (Exception e1)
            {
                if (e1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                else
                {
                    string e1message = "[106-E1] Failed to read the selected PERSON for UPDATE purposes.\r\n\n(" + e1.HResult + ")\r\n" + e1.Message;
                    MessageBox.Show(e1message, "UPDATE", MessageBoxButtons.OK);
                    mailProgramError(e1message);
                }
            }
        }

        // MENU ITEM: KNOWLEDGE BASE CLICKED
        private void knowledgeBaseToolStripMenuItem_Click(object sender, EventArgs e)
        {   /// <summary>
            ///     CALL showKnowledgeBase since I haven't yet learned how to call a _Clink event from a button
            showKnowledgeBase();
        }

        // ADD PERSON RECORD
        // =================
        private void btnHRaddPersonRecord_Click(object sender, EventArgs e)
        {   /// <summary>
            ///     THE ADD NEW PERSON RECORD BUTTON HAS BEEN PRESSED
            /// </summary>

            // CONFIRM DESIRE TO ADD A NEW PERSON RECORD
            //string messageAddNewPerson = "Do you still wish to ADD \"" + cbxHRweNamList.Text + "\'?";
            //DialogResult answer = MessageBox.Show(messageAddNewPerson,"ADD NEW PERSON RECORD", MessageBoxButtons.YesNo);
            if (true)
            {
                // CHECK MISSING ARGUMENTS
                if (cbxHRweNamList.Text == "" || tbxHRnameSortOrder.Text == "")     // then we don't have some needed information
                {
                    MessageBox.Show("Both of the NAME fields must be filled out, because on or the other or both are not, the UPDATE is cancelled.", "UPDATE NAME CHECK", MessageBoxButtons.OK);
                    btnHRaddPersonRecord.Visible = true;   // TURN THIS BACK ON, the user should not have to reenter everything to get going again.
                }  
                    //    .Visible = true;        
                else // we have enough information to continue
                {    // UPON UPDATE CONFIRMATION, Journal
                    DialogResult answer2 = MessageBox.Show("Do you wish to ADD \'" + cbxHRweNamList.Text + "\'?", "", MessageBoxButtons.YesNo);
                    if (answer2 == DialogResult.Yes)
                    {
                        try
                        {
                            // NOW UPDATE THE WebPerson TABLE WITH THIS RECORD
                            // CHECK rtbHRknowledgeBase.Text FOR SPECIAL CHARACTERS BY CREATING A STRING AND THEN USING IT IN THE SQL
                            string strKnowledgeBase = rtbHRknowledgeBase.Text;
                            strKnowledgeBase = strKnowledgeBase.Trim().Replace("'","''");
                            try
                            {
                                string strSQL = "INSERT INTO WebNames (";
                                strSQL += "WeNamList, ";            // 1
                                strSQL += "NameSortOrder, ";        // 2
                                strSQL += "MiddleName, ";           // 3
                                strSQL += "MaidenName, ";             // 4
                                strSQL += "OtherNameRef, ";           // 5
                                strSQL += "BirthPlace, ";             // 6
                                strSQL += "BirthDate, ";              // 7
                                strSQL += "IntermentLocation, ";      // 8
                                strSQL += "DateDied, ";               // 9
                                strSQL += "Mother, ";                 // 10
                                strSQL += "Father, ";                 // 11
                                strSQL += "Siblings, ";               // 12
                                strSQL += "Spouse, ";                 // 13
                                strSQL += "ChildrenNames, ";          // 14
                                strSQL += "KnowledgeBase, ";          // 15
                                strSQL += "YouthPicture, ";           // 16
                                strSQL += "AdultPicture";           // 17
                                strSQL += ") VALUES (";             // VALUES ()
                                strSQL += "'" + cbxHRweNamList.Text.Trim() + "',";                          // 1
                                strSQL += "'" + tbxHRnameSortOrder.Text.Trim() + "', ";                     // 2
                                strSQL += "'" + tbxHRmiddleName.Text.Trim().Replace("'", "''") + "', ";     // 3
                                strSQL += "'" + tbxHRmaidenName.Text.Trim().Replace("'", "''") + "', ";     // 4
                                strSQL += "'" + tbxHRotherNamedRef.Text.Trim().Replace("'", "''") + "', ";  // 5
                                strSQL += "'" + tbxHRbirthPlace.Text.Trim().Replace("'", "''") + "', ";     // 6
                                strSQL += "'" + tbxHRbirthDate.Text.Trim().Replace("'", "''") + "', ";      // 7
                                strSQL += "'" + tbxHRintermentLocation.Text.Trim().Replace("'", "''") + "', "; // 8
                                strSQL += "'" + tbxHRdateDied.Text.Trim().Replace("'", "''") + "', ";       // 9
                                strSQL += "'" + tbxHRmother.Text.Trim().Replace("'", "''") + "', ";         // 10
                                strSQL += "'" + tbxHRfather.Text.Trim().Replace("'", "''") + "', ";         // 11
                                strSQL += "'" + tbxHRsiblings.Text.Trim().Replace("'", "''") + "', ";       // 12
                                strSQL += "'" + tbxHRspouse.Text.Trim().Replace("'", "''") + "', ";         // 13
                                strSQL += "'" + tbxHRchildrenNames.Text.Trim().Replace("'", "''") + "', ";  // 14
                                strSQL += "'" + strKnowledgeBase + "', ";                                   // 15
                                strSQL += "'" + tbxANRyouthPicture.Text.Trim().Replace("'","''") + "', ";   // 16
                                strSQL += "'" + tbxANRadultPicture.Text.Trim().Replace("'", "''") + "'";    // 17
                                strSQL += ")";

                                OleDbConnection cn2 = new OleDbConnection();
                                cn2.ConnectionString = strConnectionString;
                                cn2.Open();
                                OleDbCommand myCommand2 = new OleDbCommand(strSQL, cn2);
                                myCommand2.ExecuteNonQuery();

                                // UPDATE STATUS
                                tsslUpdateAddStatus.ForeColor = Color.Green;
                                tsslUpdateAddStatus.Text = "\'" + cbxHRweNamList.Text.Trim() + "\' Knowledge Base record ADDED";

                                // SEND EMAIL NOTIFICATIN
                                string strMessage = "The record in the Knowledge Base for \r\n\n\t";
                                strMessage += cbxHRweNamList.Text.Trim() + "\r\n\nwas been ADDED at " + DateTime.Now.ToString() + ".";
                                sendEmail("steve.breckner@hoferranch.onmicrosoft.com", "robert.perry@hoferranch.onmicrosoft.com", "Knowledge Base Record ADD", strMessage);

                                // CLEAR THE FIELDS
                                ClearHistoryRecordFields();

                                // REBUILD THE DROP DOWN COMBO BOX
                                fillComboBox("WeNamList", "WebNames", "WeNamList", cbxHRweNamList);

                                // ADJUST BUTTON VISIBILITY
                                btnHRupdatePersonRecord.Visible = false;
                                btnHRaddPersonRecord.Visible = false;
                                btnHRstartNewRecord.Visible = true;
                                btnHRupdateInstructions.Visible = true;
                                // NOTIFIY USER
                                statusStrip1.Text = "ADD of \'" + cbxHRweNamList.Text + "\' successfull.";
                            }
                            catch (Exception e2)
                            {
                                if (e2.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                                else
                                {
                                    string e2message = "[107:E2] Unable to update the \'" + njWeNamList + "\'/\'" + cbxHRweNamList.Text + "\' record.\r\n\n(" + e2.HResult + ")\r\n" + e2.Message;
                                    MessageBox.Show(e2message, "UPDATE PERSON", MessageBoxButtons.OK);
                                    mailProgramError(e2message);
                                    statusStrip1.Text = "ADD of \'" + cbxHRweNamList.Text + "\' failed.";
                                }
                            }
                        }
                        catch (Exception e1)
                        {
                            if (e1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                            else
                            {
                                string e1message = "[108:E1] Unable to INSERT the name \'" + cbxHRweNamList.Text + "\'.\r\n\n(" + e1.HResult + ")\r\n" + e1.Message;
                                MessageBox.Show(e1message, "UPDATE PERSON RECORD", MessageBoxButtons.OK);
                                mailProgramError(e1message);
                            }
                        }
                    }
                    else
                    {
                        MessageBox.Show("You have chosen not to update the \'" + cbxHRweNamList + "\' person record.", "UPDATE RECORD", MessageBoxButtons.OK);
                    }
                }
            }  // end of IF seeking verification
            //else      // NO LONGER NECESSARY BECAUSE THE IF IS ALWAYS TRUE
            //{
            //    MessageBox.Show("You have chosen NOT to ADD this person.", "ADD NEW PERSON",MessageBoxButtons.OK);
            //}
            // RESET BUTTONS (only the START NEW RECORD and CLOSE buttons should be lit
            //btnHRstartNewRecord.Visible = true;
            //btnHRaddPersonRecord.Visible = false;
            //btnHRupdatePersonRecord.Visible = false;
        }
        //
        // UPDATE PERSON RECORD
        // ====================
        private void btnHRupdatePersonRecord_Click(object sender, EventArgs e)
        {   /// <summary>
            ///     THE UPDATE PERSON RECORD BUTTON HAS BEEN PRESSED
            ///     1. Upon confirmation, journal the before update fields
            ///     2. Execute the SQL UPDATE
            ///     3. Clear the fields
            ///     4. Adjust button visibility
            /// </summary>
            // CHECK MISSING ARGUMENTS
            if (cbxHRweNamList.Text == "" || tbxHRnameSortOrder.Text == "")     // then we don't have some needed information
            {
                MessageBox.Show("Both of the NAME fields must be filled out, because on or the other or both are not, the UPDATE is cancelled.","UPDATE NAME CHECK", MessageBoxButtons.OK);
            }
            else // we have enough information to continue
            {    // UPON UPDATE CONFIRMATION, Journal
                DialogResult answer = MessageBox.Show("Do you wish to UPDATE the information for \r\n\'" + cbxHRweNamList.Text + "\'?", "", MessageBoxButtons.YesNo);
                if (answer == DialogResult.Yes)
                {
                    if (njWeNamList == null) { njWeNamList = ""; }
                    if (njNameSortOrder == null) { njNameSortOrder = ""; }
                    if (njMiddleName == null) { njMiddleName = ""; }
                    if (njMaidenName == null) { njMaidenName = ""; }
                    if (njOtherNameRef == null) { njOtherNameRef = ""; }
                    if (njBirthDate == null) { njBirthDate = ""; }
                    if (njBirthPlace == null) { njBirthPlace = ""; }
                    if (njIntermentLocation == null) { njIntermentLocation = ""; }
                    if (njDateDied == null) { njDateDied = ""; }
                    if (njMother == null) { njMother = ""; }
                    if (njFather == null) { njFather = ""; }
                    if (njSiblings == null) { njSiblings = ""; }
                    if (njChildrenNames == null) { njChildrenNames = ""; }
                    if (njKnowledgeBase == null) { njKnowledgeBase = ""; }
                    if (njYouthPicture == null) { njYouthPicture = ""; }
                    if (njAdultPicture == null) { njAdultPicture = ""; }
                    DateTime njDateTimeRecorded = DateTime.Now;

                    //// JOURNAL THE BEFORE RECORD INFORMATION
                    string strSQLjournal = "INSERT INTO WebNamesJournal (";
                    strSQLjournal += "njType, ";              // 0
                    strSQLjournal += "njWeNamList, ";         // 1
                    strSQLjournal += "njNameSortOrder, ";     // 2
                    strSQLjournal += "njMiddleName, ";        // 3  
                    strSQLjournal += "njMaidenName, ";        // 4
                    strSQLjournal += "njOtherNameRef, ";      // 5
                    strSQLjournal += "njBirthPlace, ";        // 6
                    strSQLjournal += "njBirthDate, ";         // 7
                    strSQLjournal += "njIntermentLocation, "; // 8
                    strSQLjournal += "njDateDied, ";          // 9
                    strSQLjournal += "njMother, ";            // 10
                    strSQLjournal += "njFather, ";            // 11
                    strSQLjournal += "njSiblings, ";          // 12
                    strSQLjournal += "njSpouse, ";            // 13
                    strSQLjournal += "njChildrenNames, ";     // 14
                    strSQLjournal += "njKnowledgeBase, ";     // 15
                    strSQLjournal += "njYouthPicture, ";      // 16
                    strSQLjournal += "njAdultPicure, ";      // 16
                    strSQLjournal += "njDateTimeRecorded)";   // 17
                    strSQLjournal += " VALUES (";
                    strSQLjournal += "'UPDATE', ";                                              // 0 njType
                    strSQLjournal += "'" + njWeNamList.Trim().Replace("'", "''") + "', ";       // 1 njWeNamList
                    strSQLjournal += "'" + njNameSortOrder.Trim().Replace("'", "''") + "', ";   // 2 njNameSortOrder
                    strSQLjournal += "'" + njMiddleName.Trim().Replace("'", "''") + "', ";             // 3 njMiddleName
                    strSQLjournal += "'" + njMaidenName.Trim().Replace("'", "''") + "', ";             // 4 njMaidenName
                    strSQLjournal += "'" + njOtherNameRef.Trim().Replace("'", "''") + "', ";           // 5 njOtherNameRef
                    strSQLjournal += "'" + njBirthPlace.Trim().Replace("'", "''") + "', ";             // 6 njBirthPlace
                    strSQLjournal += "'" + njBirthDate.Trim().Replace("'", "''") + "', ";              // 7 njBirthDate
                    strSQLjournal += "'" + njIntermentLocation.Trim().Replace("'", "''") + "', ";      // 8 njIntermentLocation
                    strSQLjournal += "'" + njDateDied.Trim().Replace("'", "''") + "', ";               // 9 njDateDied
                    strSQLjournal += "'" + njMother.Trim().Replace("'", "''") + "', ";                 // 10 njMother
                    strSQLjournal += "'" + njFather.Trim().Replace("'", "''") + "', ";                 // 11 njFather
                    strSQLjournal += "'" + njSiblings.Trim().Replace("'", "''") + "', ";               // 12 njSiblings
                    strSQLjournal += "'" + njSpouse.Trim().Replace("'", "''") + "', ";                 // 13 njSpouse
                    strSQLjournal += "'" + njChildrenNames.Trim().Replace("'", "''") + "', ";          // 14 njChildrenNames
                    strSQLjournal += "'" + njKnowledgeBase.Trim().Replace("'", "''") + "', ";          // 15 njKnowledgeBase
                    strSQLjournal += "'" + njYouthPicture.Trim().Replace("'", "''") + "', ";           // 16 njYouthPicture
                    strSQLjournal += "'" + njAdultPicture.Trim().Replace("'", "''") + "', ";           // 17 njAdultPicture
                    strSQLjournal += "'" + njDateTimeRecorded + "')";                           // 18 njDateTimeRecorded

                    try
                    {
                        OleDbConnection cn = new OleDbConnection();
                        cn.ConnectionString = strConnectionString;
                        cn.Open();
                        OleDbCommand myCommand = new OleDbCommand(strSQLjournal, cn);
                        myCommand.ExecuteNonQuery();

                        // NOW UPDATE THE RECORD
                        try
                        {
                            string strKnowledgeBase = rtbHRknowledgeBase.Text;
                            strKnowledgeBase = strKnowledgeBase.Trim().Replace("'", "''");
                            string strSQL = "UPDATE WebNames SET ";
                            strSQL += "WeNamList = '" + cbxHRweNamList.Text + "', ";
                            strSQL += "NameSortOrder = '" + tbxHRnameSortOrder.Text + "', ";
                            strSQL += "MiddleName = '" + tbxHRmiddleName.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "MaidenName = '" + tbxHRmaidenName.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "OtherNameRef = '" + tbxHRotherNamedRef.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "BirthPlace = '" + tbxHRbirthPlace.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "BirthDate = '" + tbxHRbirthDate.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "IntermentLocation = '" + tbxHRintermentLocation.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "DateDied = '" + tbxHRdateDied.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "Mother = '" + tbxHRmother.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "Father = '" + tbxHRfather.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "Siblings = '" + tbxHRsiblings.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "Spouse = '" + tbxHRspouse.Text.Trim().Replace("'","''") + "', ";
                            strSQL += "ChildrenNames = '" + tbxHRchildrenNames.Text.Trim().Replace("'", "''") + "', ";
                            strSQL += "KnowledgeBase = '" + strKnowledgeBase + "', ";
                            strSQL += "YouthPicture = '" + tbxANRyouthPicture.Text + "', ";
                            strSQL += "AdultPicture = '" + tbxANRadultPicture.Text + "' ";
                            strSQL += "WHERE WeNamList = '" + cbxHRweNamList.Text + "'";
                            //NO NEED TO WRITE BACK THE YouthPicture and AdultPicture since the user can't change them
                            OleDbConnection cn2 = new OleDbConnection();
                            cn2.ConnectionString = strConnectionString;
                            cn2.Open();
                            OleDbCommand myCommand2 = new OleDbCommand(strSQL, cn2);
                            myCommand2.ExecuteNonQuery();

                            // UPDATE STATUS
                            tsslUpdateAddStatus.ForeColor = Color.Green;
                            tsslUpdateAddStatus.Text = "\'" + cbxHRweNamList.Text.Trim() + "\' record in Knowledge Base has been UPDATED";

                            // SEND EMAIL NOTIFICATIN
                            string strMessage = "The record in the Knowledge Base for \r\n\n\t";
                            strMessage += cbxHRweNamList.Text.Trim() + "\r\n\nwas UPDATED at " + DateTime.Now.ToString() + ".";
                            sendEmail("steve.breckner@hoferranch.onmicrosoft.com", "robert.perry@hoferranch.onmicrosoft.com", "Knowledge Base Record UPDATE Notification", strMessage);

                            // CLEAR THE FIELDS
                            ClearHistoryRecordFields();

                            // ADJUST BUTTON VISIBILITY
                            btnHRupdatePersonRecord.Visible = false;
                            btnHRaddPersonRecord.Visible = false;
                            btnHRstartNewRecord.Visible = true;
                            btnHRupdateInstructions.Visible = true;
                        }
                        catch (Exception e2)
                        {
                            if (e2.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                            else
                            {
                                string e2message = "[107:E2] Unable to update the \'" + njWeNamList + "\' record.\r\n\n(" + e2.HResult + ")\r\n" + e2.Message;
                                MessageBox.Show(e2message,"UPDATE PERSON",MessageBoxButtons.OK);
                                mailProgramError(e2message);
                            }
                        }
                    }
                    catch (Exception e1)
                    {
                        if (e1.HResult == -2147467529) { MessageBox.Show("You may have lost your network connection.", "POSSIBLE NETWORK FAILURE", MessageBoxButtons.OK); }
                        else
                        {
                            string e1message = "[107:E1] Failed to journal the preupdate information.\r\n\n(" + e1.HResult + ")\r\n" + e1.Message;
                            MessageBox.Show(e1message, "UPDATE PERSON RECORD", MessageBoxButtons.OK);
                            mailProgramError(e1message);
                        }
                    }
                }
                else
                {
                    MessageBox.Show("You have chosen not to update the \'" + cbxHRweNamList + "\' person record.", "UPDATE RECORD", MessageBoxButtons.OK);
                }
            }
        }

        // KNOWLEDGE BASE: BUTTON: START NEW RECORD
        private void btnHRstarNewRecord_Click(object sender, EventArgs e)
        {   /// <summary>
            ///     THE START NEW PERSON RECORD BUTTON HAS BEEN PRESSED
            /// </summary>
            ClearHistoryRecordFields();
            btnHRaddPersonRecord.Visible = true;
            btnHRstartNewRecord.Visible = true;
            btnHRupdatePersonRecord.Visible = false;
            btnHRupdateInstructions.Visible = false;
            cbxHRweNamList.Text = "";
            //if (cbxHRweNamList.Text != "Any" || cbxHRweNamList.Text != "") {fillComboBox("WeNamList", "WebNames", "WeNamList", cbxHRweNamList);}
        }

        private void ClearHistoryRecordFields()
        {   /// <summary>
            ///     Clear all the fields in the form
            /// </summary>
            cbxHRweNamList.Text = "";
            tbxHRnameSortOrder.Text = "";
            tbxHRmiddleName.Text = "";
            tbxHRmaidenName.Text = "";
            tbxHRotherNamedRef.Text = ""; 
            tbxHRbirthDate.Text = "";
            tbxHRbirthPlace.Text = "";
            tbxHRintermentLocation.Text = "";
            tbxHRdateDied.Text = "";
            tbxHRmother.Text = "";
            tbxHRfather.Text = "";
            tbxHRsiblings.Text = "";
            tbxHRspouse.Text = "";
            tbxHRchildrenNames.Text = "";
            rtbHRknowledgeBase.Clear();
            pbxHRadultPicture.Visible = false;
            pbxHRyouthPicture.Visible = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {   /// <summary>
            ///     THE CLOSE BUTTON HAS BEEN PRESSED
            ///     1. Close the panel
            /// </summary>
            pnlKnowledgeBase.Visible = false;
        }

        private void btnGoToRecord_Click(object sender, EventArgs e)
        {   /// <summary>
            ///     The .Tag value contiaining the DBLogID of the record to be edited
            /// </summary>
            pnlSearchFilter.Visible = false;
            pnlLightbox.Visible = false;
            pnlAddNewRecord.Visible = true;
            pnlAddNewRecord.BringToFront();
            pnlAddNewRecord.Parent = panel1;
            string strDBLogID = btnGoToRecord.Tag.ToString();
            showSelectedRecord(strDBLogID);
        }

        private void helpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pnlWelcome.Visible = false;
            pnlWelcome.Parent = panel1;
            pnlWelcome.BringToFront();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            // DO NOTHING HERE IF USER CLICKS
            string message = "To UPDATE a record, choose from the list below by scrolling until you find the record you want and DOUBLE-CLICKING to choose it for updating.";
            MessageBox.Show(message,"SELECT RECORD TO UPDATE", MessageBoxButtons.OK);
        }

        private void menuBarHelpToolStripMenuItem_Click(object sender, EventArgs e)
        {   //
            pnlWelcome.Visible = true;

        }

        private void btnCloseMenu_Click(object sender, EventArgs e)
        {
            pnlWelcome.Visible = false;
        }

        private void tbxANRjpgFilename_TextChanged(object sender, EventArgs e)
        {   // TEXT BOX FOR ANR JPGFILENAME

        }

        private void tbxANRdbLogID_TextChanged(object sender, EventArgs e)
        {   // TEXT BOX FOR ANR JPGFILENAME

        }

        private void lookUpRecordByDBLogID(string strJPGFilename)
        {   // GIVEN THE JPG NUMBER, LOOK FOR A RECORD CONTAINING IT
            string strJPGFilenameRegularized = strJPGFilename.Trim().Replace(".jpg","") + ".jpg";
            string strSQL = "SELECT DBLogID FROM ScannedImages WHERE WebJPGFilename ='" + strJPGFilenameRegularized + "'";
            try
            {
                OleDbConnection cn = new OleDbConnection();
		        cn.ConnectionString = strConnectionString;
		        cn.Open();
		        OleDbCommand myCommand = new OleDbCommand(strSQL, cn);
		        OleDbDataReader myDataReader;
		        myDataReader = myCommand.ExecuteReader();
                myDataReader.Read();
                string strDBLogID = myDataReader["DBLogID"].ToString();
                tbxANRjpgFilename.Text = "";
                showSelectedRecord(strDBLogID);
                btnANRupdateInstruction.Visible = false;
                btnANRupdateRecord.Visible = true;
            }
            catch (Exception e1)
            {
                string e1message = "";
                if (e1.HResult != -2146233079)
                { 
                    e1message = "SQL ERROR: " + e1.HResult + "\r\n" + e1.Message;
                    MessageBox.Show(e1message,"SELECT RECORD BY JPG FILENAME", MessageBoxButtons.OK);
                    mailProgramError(e1message);
                }
                tbxANRjpgFilename.Text = "";
                clearAddNewRecordFields();
                btnANRupdateInstruction.Visible = true;
                btnANRupdateRecord.Visible = false;
            }
        }

        private void btnANRselectRecIDorJPG_Click(object sender, EventArgs e)
        {   // ANR SELECT BUTTON FOR DBLogID or WebJPGFilename

            if (tbxANRdbLogID.Text != "")  // CHECK FIRST FOR DBLogID
            {
                showSelectedRecord(tbxANRdbLogID.Text.Trim());
                tbxANRdbLogID.Text = "";        // SHOW IT NO LONGER
            }
            else  // CHECK SECOND FOR WebJPGFilename
            {
                lookUpRecordByDBLogID(tbxANRjpgFilename.Text);
                tbxANRjpgFilename.Text = "";
            }
        }
    }
}
