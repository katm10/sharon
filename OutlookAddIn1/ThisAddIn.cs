using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace OutlookAddIn1
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.Items items;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            outlookNameSpace = this.Application.GetNamespace("MAPI");
            inbox = outlookNameSpace.GetDefaultFolder(
                    Microsoft.Office.Interop.Outlook.
                    OlDefaultFolders.olFolderInbox);

            items = inbox.Items;
            items.ItemAdd +=
                new Outlook.ItemsEvents_ItemAddEventHandler(items_ItemAdd);
        }

        void items_ItemAdd(object Item)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            if (!SearchforEmail(mail.SenderEmailAddress))
            {
                AddContact(mail.SenderEmailAddress);
            }
            if (Item != null)
            {
                if (mail.MessageClass == "IPM.Note" &&
                           mail.Body.ToLower().Contains("sharon"))
                {
                    sendFirstEmail(Item, 0);
                }
                else if (mail.MessageClass == "IPM.Note" && mail.Subject.ToUpper().Contains("SMS EMAIL"))
                {
                    sendFirstEmail(Item, 1);
                    Outlook.ContactItem contact = mail.Sender.GetContact();
                    contact.Email2Address = mail.Body;
                }
                else if (mail.MessageClass == "IPM.Note" && mail.Subject.ToUpper().Contains("ZIP CODE"))
                {
                    sendFirstEmail(Item, 2);
                    Outlook.ContactItem contact = mail.Sender.GetContact();
                    contact.Email2DisplayName = mail.Body;
                }
                else
                {
                    sendFirstEmail(Item, 3);
                }

            }
    }

        private string sendFirstEmail(object Item, int num)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;
            string subjectEmail = "Welcome to your Sharon Community!";
            string bodyEmail;
            switch (num)
            {
                case 0:
                    bodyEmail = "Please respond to this email with only your SMS Email Address and make the subject line 'SMS Email'. If you don't know your SMS email address, go to http://www.emailtextmessages.com/.";
                    break;
                case 1:
                    bodyEmail = "Please respond to this email with only your zip code and make the subject 'Zip Code'.";
                    break;
                case 2:
                    bodyEmail = "Thank you! You should recieve your first text soon.";
                    break;
                default:
                    bodyEmail = "Sorry, I don't understand that. Please refer to http://sharontherobot.squarespace.com"
                        break;
            }
            
            Outlook.MailItem response = mail.Reply();
            response.Body = bodyEmail;
            response.Subject = subjectEmail;
            response.Send();
            return response.ConversationID;
        }

        private void CreateEmailItem(string subjectEmail,
               string toEmail, string bodyEmail)
        {
            Outlook.MailItem eMail = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            eMail.Body = bodyEmail;
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            ((Outlook._MailItem)eMail).Send();
        }

        private Boolean SearchforEmail(string Address)
        {
            string contactMessage = string.Empty;
            Outlook.MAPIFolder contacts = (Outlook.MAPIFolder)
                this.Application.ActiveExplorer().Session.GetDefaultFolder
                 (Outlook.OlDefaultFolders.olFolderContacts);
            foreach (Outlook.ContactItem foundContact in contacts.Items)
            {
                if (foundContact.Email1Address != null)
                {
                    if (foundContact.Email1Address == Address)
                    {
                        return true;
                    }
                }
            }return false;
        }
        private void AddContact(String mailAddress)
        {
            Outlook.ContactItem newContact = (Outlook.ContactItem)
                this.Application.CreateItem(Outlook.OlItemType.olContactItem);
            newContact.Email1Address = mailAddress;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see http://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
