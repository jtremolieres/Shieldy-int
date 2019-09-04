using NotifySecurity.Properties;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Outlook;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace NotifySecurity
{


    [ComVisible(true)]
    public class Ribbon1 : IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon1()
        {
            StartUp = true;
        }

        public Boolean StartUp = false;
        public String ddlEntityValue = "Company";

        #region IRibbonExtensibility Members


        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion

        public string GetCustomUI(string ribbonID)
        {

            String txtRibbon = GetResourceText("NotifySecurity.Ribbon1.xml");

            return txtRibbon;
        }

        #region Ribbon Callbacks
        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {

            this.ribbon = ribbonUI;
        }

        public Bitmap Btn_GetImage(IRibbonControl control)
        {

            return new Bitmap(Resources.shieldy);
        }
        #endregion


        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

            try
            {
                Thread.CurrentThread.CurrentUICulture = CultureInfo.GetCultureInfo("Company");
            }
            catch (System.Exception)
            {
            }

        }


        public string GetContextMenuLabel(Office.IRibbonControl control)
        {
            return Resources.ContextMenuLabel;
        }

        public string GetGroupLabel(Office.IRibbonControl control)
        {
            return "Shieldy";

        }
        public string GetTabLabel(Office.IRibbonControl control)
        {
            return Resources.TabLabel;

        }


        public string GetSupertipLabel(Office.IRibbonControl control)
        {
            var v = System.Reflection.Assembly.GetAssembly(typeof(Ribbon1)).GetName().Version;
            int revMaj = v.Major;
            int revMin = v.Minor;
            int revBuild = v.Build;
            int revRev = v.Revision;

            return "Shieldy v" + revMaj.ToString() + "." + revMin.ToString() + "." + revBuild.ToString() + "." + revRev.ToString();// versionInfo.ToString();
        }


        public string GetScreentipLabel(Office.IRibbonControl control)
        {
            return Resources.ScreentipLabel;
        }

        public string GetButtonLabel(Office.IRibbonControl control)
        {
            return Resources.ButtonLabel + " " + Properties.Settings.Default.CustomerName;
        }

        public void ShowMessageClick(Office.IRibbonControl control)
        {

            CreateNewMailToSecurityTeam(control);
        }

        public Bitmap GetImage(IRibbonControl control)
        {
            return new Bitmap(Properties.Resources.shieldy);

        }

        private void CreateNewMailToSecurityTeam(IRibbonControl control)
        {

            Selection selection =
                Globals.ThisAddIn.Application.ActiveExplorer().Selection;

            if (selection.Count == 1)   // Check that selection is not empty.
            {
                object selectedItem = selection[1];   // Index is one-based.
                Object mailItemObj = selectedItem as Object;
                MailItem mailItem = null;// selectedItem as MailItem;
                if (selection[1] is Outlook.MailItem)
                {
                    mailItem = selectedItem as MailItem;
                }

                MailItem tosend = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                tosend.Attachments.Add(mailItemObj);

                #region create mail from default
                try
                {

                    tosend.To = Properties.Settings.Default.Security_Team_Mail;
                    tosend.Subject = Resources.EmailSubject;

                    tosend.CC = Properties.Settings.Default.Security_Team_Mail_cc;
                    tosend.BCC = Properties.Settings.Default.Security_Team_Mail_bcc;

                    #region retrieving message header
                    string allHeaders = "";
                    if (selection[1] is Outlook.MailItem)
                    {
                        string[] preparedByArray = mailItem.Headers("X-PreparedBy");
                        string preparedBy;
                        if (preparedByArray.Length == 1)
                            preparedBy = preparedByArray[0];
                        else
                            preparedBy = "";
                        allHeaders = mailItem.HeaderString();
                    }
                    else
                    {
                        string typeFound = "unknown";
                        typeFound = (selection[1] is Outlook.MailItem) ? "MailItem" : typeFound;

                        if (typeFound == "unknown")
                            typeFound = (selection[1] is Outlook.MeetingItem) ? "MeetingItem" : typeFound;

                        if (typeFound == "unknown")
                            typeFound = (selection[1] is Outlook.ContactItem) ? "ContactItem" : typeFound;

                        if (typeFound == "unknown")
                            typeFound = (selection[1] is Outlook.AppointmentItem) ? "AppointmentItem" : typeFound;

                        if (typeFound == "unknown")
                            typeFound = (selection[1] is Outlook.TaskItem) ? "TaskItem" : typeFound;

                        allHeaders = Resources.ItemNotMail;
                    }

                    #endregion

                    string SwordPhishURL = SwordphishObject.SetHeaderIDtoURL(allHeaders);

                    if (SwordPhishURL != SwordphishObject.NoHeaderFound)
                    {
                        string SwordPhishAnswer = SwordphishObject.SendNotification(SwordPhishURL);
                    }
                    else
                    {
                        tosend.Body = Resources.EmailBody_line1; //"Hello, I received the attached email and I think it is suspicious";
                        tosend.Body += "\n";
                        tosend.Body += Resources.EmailBody_line2; //"I think this mail is malicious for the following reasons:";
                        tosend.Body += "\n";
                        tosend.Body += Resources.EmailBody_YourReason; 
                        tosend.Body += "\n";
                        tosend.Body += Resources.EmailBody_line3; //"Please analyze and provide some feedback.";
                        tosend.Body += "\n";
                        tosend.Body += "\n";

                        tosend.Body += GetCurrentUserInfos();

                        tosend.Body += "\n\n"+ Resources.EmailBody_msgHeader+": \n--------------\n" + allHeaders + "\n\n";

                        tosend.Save();
                        tosend.Display();
                    }

                }
                catch (System.Exception ex)
                {
                    MessageBox.Show(Resources.UsingDefaultTemplate + ex.Message, Resources.MsgBox_Title);

                    MailItem mi = (MailItem)Globals.ThisAddIn.Application.CreateItem(OlItemType.olMailItem);
                    mi.To = Properties.Settings.Default.Security_Team_Mail;
                    mi.Subject = Resources.ErrorSubject;
                    String txt = (Resources.ErrorText + ex);
                    mi.Body = txt;
                    mi.Save();
                    mi.Display();
                }
            }
            else if (selection.Count < 1)   // Check that selection is not empty.
            {
                MessageBox.Show(Resources.MsgSelectOneMail, Resources.MsgBox_Title);
            }
            else if (selection.Count > 1)
            {
                MessageBox.Show(Resources.MsgSelectOnlyOneMail, Resources.MsgBox_Title);
            }
            else
            {
                MessageBox.Show(Resources.MsgBadLuck, Resources.MsgBox_Title);
            }


        }
        #endregion


        public String GetCurrentUserInfos()
        {

            String wComputername = System.Environment.MachineName + " (" + System.Environment.OSVersion.ToString() + ")";
            String wUsername = System.Environment.UserDomainName + "\\" + System.Environment.UserName;

            string str = Resources.EmailBody_PossiblyUsefulIInformation;


            Outlook.AddressEntry addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
            if (addrEntry.Type == "EX")
            {
                Outlook.ExchangeUser currentUser =
                    Globals.ThisAddIn.Application.Session.CurrentUser.
                    AddressEntry.GetExchangeUser();
                if (currentUser != null)
                {
                    str += "\n" + Resources.EmailBody_Username + currentUser.Name;
                    str += "\n" + Resources.EmailBody_STMPAddress + currentUser.PrimarySmtpAddress;
                    str += "\n" + Resources.EmailBody_Title + currentUser.JobTitle;
                    str += "\n" + Resources.EmailBody_Department + currentUser.Department;
                    str += "\n" + Resources.EmailBody_Location + currentUser.OfficeLocation;
                    str += "\n" + Resources.EmailBody_BusinessPhone + currentUser.BusinessTelephoneNumber;
                    str += "\n" + Resources.EmailBody_MobilePhone + currentUser.MobileTelephoneNumber;
                    str += "\n" + CultureInfo.CurrentUICulture.Name;
                    str += "\n" + CultureInfo.CurrentUICulture.DisplayName;

                }
            }
            str += "\n" + Resources.EmailBody_WindowsUsername + wUsername;
            str += "\n" + Resources.EmailBody_Computername + wComputername;
            str += "\n";
            return str;
        }

    }

    public static class MailItemExtensions
    {
        private const string HeaderRegex =
            @"^(?<header_key>[-A-Za-z0-9]+)(?<seperator>:[ \t]*)" +
                "(?<header_value>([^\r\n]|\r\n[ \t]+)*)(?<terminator>\r\n)";
        private const string TransportMessageHeadersSchema =
            "http://schemas.microsoft.com/mapi/proptag/0x007D001E";

        public static string[] Headers(this MailItem mailItem, string name)
        {
            var headers = mailItem.HeaderLookup();
            if (headers.Contains(name))
                return headers[name].ToArray();
            return new string[0];
        }

        public static ILookup<string, string> HeaderLookup(this MailItem mailItem)
        {
            var headerString = mailItem.HeaderString();
            var headerMatches = Regex.Matches
                (headerString, HeaderRegex, RegexOptions.Multiline).Cast<Match>();
            return headerMatches.ToLookup(
                h => h.Groups["header_key"].Value,
                h => h.Groups["header_value"].Value);
        }

        public static string HeaderString(this MailItem mailItem)
        {
            return (string)mailItem.PropertyAccessor
                .GetProperty(TransportMessageHeadersSchema);
        }

    }

}
#endregion