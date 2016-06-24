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
            }else
            {
                return;
            }

            Outlook.ContactItem sender = AddContact(mail.Sender.Address);

            if (mail != null)
            {
                if (mail.MessageClass == "IPM.Note" && mail.Subject.ToUpper().Contains("SMS EMAIL".ToUpper()))
                {
                    sendFirstEmail(Item, 1);
                    Outlook.ContactItem contact = mail.Sender.GetContact();
                    if (contact != null)
                    { 
                        string sms = mail.Body;
                        contact.Email2Address = sms;
                    }
                }
                else if (mail.MessageClass == "IPM.Note" && mail.Subject.ToUpper().Contains("ZIP CODE"))
                {
                    sendFirstEmail(Item, 2);
                    Outlook.ContactItem contact = mail.Sender.GetContact();
                    if (contact != null)
                    {
                        string zip = mail.Body;
                        contact.Email2DisplayName = zip;
                        SendFirstText(contact);
                    }
                }
                else
                {
                    sendFirstEmail(Item, 0);
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

        private Outlook.ContactItem SearchforEmail(string Address)
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
                        return foundContact;
                    }
                }
            }return null;
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
                newContact.Email1Address = "katm10@live.com";
            }
            catch
            {
                System.Windows.Forms.MessageBox.Show("no worky");
            } return newContact;
        }

        private void SendFirstText(Outlook.ContactItem contact)
        {
            string bodyEmail = "Hello! Welcome to the Sharon Community! You are part of the" + contact.Email2DisplayName + "community. If you want to change your zip code, write an email to sharoncommunity@outlook.com with only your zip code and make the subject 'Zip Code'. If you would like to stop recieving texts, text 'STOP'. ";
            CreateEmailItem(null, contact
                        .Email2Address, bodyEmail);
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
