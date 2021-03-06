﻿using System;
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
            Outlook.MailItem mail;
            if (Item is Outlook.MailItem)
            {
                mail = (Outlook.MailItem)Item;
            } else
            {
                return;
            }

            Outlook.ContactItem sender = AddContact(mail.Sender.Address);
            string[] splitEmail = mail.Sender.Address.Split('@');

            if (mail != null && mail.MessageClass == "IPM.Note")
            {
                if (!IsDigitsOnly(splitEmail[0]))
                {
                    if (mail.Subject != null)
                    {
                        if (mail.Subject.ToUpper().Contains("SMS EMAIL".ToUpper()))
                        {
                            sendFirstEmail(Item, 1);
                            if (sender != null)
                            {
                                string[] sms = mail.Body.Split('_');
                                sender.Email2Address = sms[0];
                                sender.Display(true);
                                sender.Save();
                            }
                        }
                        else if (mail.Subject.ToUpper().Contains("ZIP CODE"))
                        {
                            sendFirstEmail(Item, 2);
                            if (sender != null)
                            {
                                string[] zipSplit = mail.Body.Split('_');
                                sender.FirstName = zipSplit[0];
                                sender.Save();
                                sender.Display(true);
                                SendFirstText(sender);
                            }
                        }
                        else
                        {
                            sendFirstEmail(Item, 0);
                        }
                    } else
                    {
                        sendFirstEmail(Item, 0);
                    }
                }
                else
                {
                    sendText(mail);
                }
            }
        }

        private void sendFirstEmail(object Item, int num)
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
                    bodyEmail = "Sorry, I don't understand that. Please refer to http://sharontherobot.squarespace.com";
                    break;
            }

            Outlook.MailItem response = mail.Reply();
            response.Body = bodyEmail;
            response.Subject = subjectEmail;
            response.Send();
        }

        public void sendText(Outlook.MailItem mail)
        {
            Outlook.MailItem response = mail.Reply();
            string body;
            string[] body2Split = mail.Body.Split(' ');

            if (isUrl(body2Split[0]))
            {
                string zip = AddContact(mail.SenderEmailAddress).FirstName;
                string url = body2Split[0];
                List<Outlook.ContactItem> zipContacts = findZipContacts(zip);
                foreach (Outlook.ContactItem contact in zipContacts)
                { if (contact.Email2Address != null)
                    {
                        this.CreateEmailItem(null, contact.Email2Address, url);
                    }
                }

            }
            else
            {
                if (body2Split[0].ToUpper().Trim() == "STOP")
                {
                    this.CreateEmailItem(null, mail.SenderEmailAddress, "You are no longer part of the Sharon community.");
                    Outlook.ContactItem contact = AddContact(mail.SenderEmailAddress);
                    contact.Delete();
                }
                else
                {
                    switch (body2Split[0].ToUpper().Trim())
                    {
                        case "CHAT":
                            body = "Please go to https://tlk.io/ and make a chat room. Then, respond with only the link.";
                            break;
                        case "POLL":
                            body = "Please go http://www.poll-maker.com/ and make a poll. Then, respond with only the link.";
                            break;
                        case "PETITION":
                            body = "Please go to http://www.change.org and make a petition. Then, respond with only the link.";
                            break;
                        case "CALL":
                            body = "Please go to http://whoismyrepresentative.com/ to learn about how to reach your representatives.";
                            break;
                        default:
                            body = "Sorry, I didn't get that. Please type 'Chat', 'Poll', 'Petition' or 'Call'.";
                            break;
                    }
                    response.Subject = null;
                    response.Body = body;
                    response.Send();
                }
            }
        }
        private Outlook.MailItem CreateEmailItem(string subjectEmail,
               string toEmail, string bodyEmail)
        {
            Outlook.MailItem eMail = (Outlook.MailItem)
                this.Application.CreateItem(Outlook.OlItemType.olMailItem);
            eMail.Subject = subjectEmail;
            eMail.To = toEmail;
            eMail.Body = bodyEmail;
            eMail.Importance = Outlook.OlImportance.olImportanceLow;
            eMail.Send();
            return eMail;
        }

        private Outlook.ContactItem SearchforEmail(string Address)
        {
            Address = Address.Trim();
            string contactMessage = string.Empty;
            Outlook.MAPIFolder contacts =
                this.Application.ActiveExplorer().Session.GetDefaultFolder
                 (Outlook.OlDefaultFolders.olFolderContacts);
            foreach (Outlook.ContactItem foundContact in contacts.Items)
            {
                if (foundContact.Email1Address != null)
                {
                    string email1 = foundContact.Email1Address.Trim();
                    if (foundContact.Email1Address.Trim() == Address)
                    {
                        return foundContact;
                    }
                }
                if (foundContact.Email2Address != null)
                {
                    string email2 = foundContact.Email2Address.Trim();
                    if (email2 == Address)
                    {
                        return foundContact;
                    }

                }
            } return null;
        }

        public List<Outlook.ContactItem> findZipContacts(string zip)
        {
            var listOfContacts = new List<Outlook.ContactItem>();
            string contactMessage = string.Empty;
            Outlook.MAPIFolder contacts =
                this.Application.ActiveExplorer().Session.GetDefaultFolder
                 (Outlook.OlDefaultFolders.olFolderContacts);
            foreach (Outlook.ContactItem foundContact in contacts.Items)
            {
                if (foundContact.FirstName != null)
                {
                    if (foundContact.FirstName == zip)
                    {
                        listOfContacts.Add(foundContact);
                    }
                }
            }
            return listOfContacts;
        }
        private Outlook.ContactItem AddContact(String mailAddress)
        {
            if (SearchforEmail(mailAddress) != null)
            {
                return SearchforEmail(mailAddress);
            }
            Outlook.ContactItem newContact = (Outlook.ContactItem)
                this.Application.CreateItem(Outlook.OlItemType.olContactItem);
            try
            {
                newContact.FirstName = "contact placeholder";
                newContact.Email1Address = mailAddress;
                newContact.Save();
                newContact.Display(true);
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("no worky");
            } return newContact;
        }


        private void SendFirstText(Outlook.ContactItem contact)
        {
            string[] bodyEmail = new string[]
            { $"Hello! Welcome to the Sharon Community! You are part of the {contact.FirstName} community.",
                "If you want to change your zip code, write an email to sharoncommunity@outlook.com with only your zip code and make the subject 'Zip Code'.",
                "If you would like to stop recieving texts, text 'Stop'.",
                "Please text 'Call', 'Chat', 'Poll', or 'Petition'." };
            foreach (string line in bodyEmail)
            {
                this.CreateEmailItem(null, contact.Email2Address, line);
            }
        }
        public bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }

            return true;
        }
    

        public bool isUrl(string uriName) { Uri uriResult;
            bool result = Uri.TryCreate(uriName, UriKind.Absolute, out uriResult)
                && (uriResult.Scheme == Uri.UriSchemeHttp || uriResult.Scheme == Uri.UriSchemeHttps);
            return result;
        }
        private void ReplyWithVoiceMail()
        {
            Outlook.MailItem mail = (Outlook.MailItem)Application.ActiveInspector().CurrentItem;
            Outlook.Action action = mail.Actions.Add();
            action.Name = "Reply";
            action.ReplyStyle = Outlook.OlActionReplyStyle.olUserPreference;
            action.ResponseStyle = Outlook.OlActionResponseStyle.olOpen;
            action.CopyLike = Outlook.OlActionCopyLike.olReply;
           
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
