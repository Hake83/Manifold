/* Outlook addin to pass information to manifold_lister and to create
 * popup notifications for sales e-mails
 * 
 * Copyright(C) 2018  Ryan S. Hake
 *
 * This program is free software: you can redistribute it and/or modify
 * it under the terms of the GNU General Public License as published by
 * the Free Software Foundation, either version 3 of the License, or
 * any later version.
 *
 * This program is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.See the
 * GNU General Public License for more details.
 * 
 * You should have received a copy of the GNU General Public License
 * along with this program.If not, see<http://www.gnu.org/licenses/>.
 */

using System;
using System.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Windows.Forms;
using System.IO;
using System.Text.RegularExpressions;
using System.Diagnostics;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using System.Collections.Generic;
using System.Drawing;


namespace Manifold
{
    public partial class ThisAddIn
    {
        Outlook.NameSpace outlookNameSpace;
        Outlook.MAPIFolder inbox;
        Outlook.MAPIFolder sharedInbox;
        Outlook.Items items;
        Outlook.Items sharedItems;
        Outlook.Recipient recipient;
        

        #region ThisAddIn_Startup
        /// <summary>
        /// A lot of code to monitor the outlook inbox and fire an e-vent for new e-mails in the inbox
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            outlookNameSpace = Application.GetNamespace("MAPI");
            recipient = outlookNameSpace.CreateRecipient("SFEng");    
            recipient.Resolve();
            if (recipient.Resolved)
            {
                sharedInbox = outlookNameSpace.GetSharedDefaultFolder(recipient, Outlook.OlDefaultFolders.olFolderInbox);
                sharedItems = sharedInbox.Items;
                sharedItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(SharedItems_ItemAdd);
            }
            inbox = outlookNameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderInbox);
            items = inbox.Items;
            items.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(Items_ItemAdd);                    //Event that fires from new e-mails into the inbox
        }

        /// <summary>
        /// process new mail from shared inbox
        /// </summary>
        /// <param name="NewItem"></param>
        private void SharedItems_ItemAdd(object NewItem)
        {
            Outlook.MailItem mail = (Outlook.MailItem)NewItem;    //New mail object
            MailParser(mail);                                                                               //process the mail for manifolds.
            /// <summary>
            /// The following code uses the tulpep/notification-popup library from nuget
            /// to display a new e-mail popup on the lower right corner of the screen
            /// as the shared inbox does not provide a new e-mail notification popup 
            /// stock from Outlook.
            /// </summary>
            try
            {
                Size size = new Size(48, 48);
                var popupNotifier = new Tulpep.NotificationWindow.PopupNotifier();
                popupNotifier.TitleText = "New Mail";
                popupNotifier.ContentText = "From " + mail.SenderName + "\n" + mail.Subject;
                popupNotifier.IsRightToLeft = false;
                popupNotifier.Image = image(mail);
                popupNotifier.ImageSize = size;
                popupNotifier.Popup();
                popupNotifier.Click += Popup_Click;

            }
            catch(Exception) { }

            void Popup_Click(object sender, EventArgs e)
            {
                try
                {
                    mail.Display(true);
                }
                catch (Exception)
                { }
            }
        }
        #endregion

        /// <summary>
        /// Function to determine if a pdf attachment is included this will change the icon displayed on the popup
        /// pictures in the signature block are counted as attachments so must iterate through each att to check
        /// if it is a pdf or not
        /// </summary>
        /// <param name=""></param>
        /// <returns></returns>
        private Bitmap image(Outlook.MailItem mailItem)
        {
            if (mailItem.Attachments.Count > 0)
            {
                for (int i = 1; i <= mailItem.Attachments.Count; i++)               
                {
                    if (mailItem.Attachments[i].FileName.ToLower().Contains(".pdf"))        
                    {
                        return Properties.Resources.email_with_pdf;
                    }
                }
                return Properties.Resources.email_icon;
            }
            else
            {
                return Properties.Resources.email_icon;
            }
        }

        /// <summary>
        /// Handle the new e-mail event from main inbox
        /// </summary>
        /// <param name="Item"></param>
        private void Items_ItemAdd(object Item)
        {
            Outlook.MailItem mail = (Outlook.MailItem)Item;     //New mail object
            MailParser(mail);                                   //Process new mail for smartflow stuff
        }

        private static void MailParser(Outlook.MailItem mail)
        {
            #region E-Mail Filter
            /// <summary>
            /// Creates a list read from text file of e-mail addresses of senders, later if the sender is on this list the program will do stuff else it will ignore e-mail.
            /// </summary>
            List<string> filter = new List<string>();                       //New list to hold e-mail addresses to check
            try
            {
                if (File.Exists(@"c:\User Programs\filter.txt"))            //populates the list if the text file exists
                {
                    using (StreamReader sr = new StreamReader(@"c:\User Programs\filter.txt"))  // the streamreader will auto close the text file when done reading.
                    {
                        while (sr.Peek() >= 0)                              //reads until the end of the file, peek checks the next line for content.
                        {
                            filter.Add(sr.ReadLine());
                        }
                    }
                }
                else                                                        //If the text file didn't exist asks the user to create one
                {
                    MessageBox.Show(@"Please define a text file at C:\User Programs\filter.txt insert salesman names you'd like to run manifold program for one on each line");
                }
            }
            catch { }
            #endregion
            bool tester = filter.Any(name => mail.SenderEmailAddress.ToUpper().Contains(name.ToUpper()));       //Create a bool object to check if the mail sender is one of the names in the filter file
            if (tester)                                     //If the sender is one of the filter names then continue more checks, otherwise will ignore the e-mail
            {
                try
                {
                    var attachments = mail.Attachments;     //Pull out the attachments to read pdf sales orders
                    var attachmentMatches = new List<string>();
                    if (attachments.Count > 0)
                    {
                        string tempPdf = System.IO.Path.GetTempFileName();
                        for (int i = 1; i <= mail.Attachments.Count; i++)
                        {
                            if (mail.Attachments[i].FileName.Contains(".pdf"))
                            {
                                mail.Attachments[i].SaveAsFile(tempPdf);
                                List<string> tempAttachments = Attachment(tempPdf);
                                if (tempAttachments.Count > 0)
                                {
                                    attachmentMatches.AddRange(tempAttachments);
                                }
                            }
                        }
                        File.Delete(tempPdf);
                    }
                    // Find manifold PN's in the e-mail body with regular expressions
                    MatchCollection matches = Regex.Matches(mail.Body, @"\d\d?\w+-\d\d?-.+?(?=\s)", RegexOptions.IgnoreCase);


                    // was having problem with passing matches so change to string
                    string[] output = new string[attachmentMatches.Count + matches.Count];
                    if (matches.Count > 0)
                    {
                        for (int i = 0; i < matches.Count; i++)
                        {
                            output[i] = matches[i].ToString();
                        }
                    }
                    if (attachmentMatches.Count > 0)
                    {
                        for (int i = 0; i < attachmentMatches.Count; i++)
                        {
                            output[i + matches.Count] = attachmentMatches[i];
                        }
                    }
                    if (output.Count() > 0)
                    {
                        ProcessStartInfo startInfo = new ProcessStartInfo();            // code to pass arguments to manifoldlister.  Console menu to open manifolds.
                        startInfo.CreateNoWindow = false;
                        startInfo.UseShellExecute = false;
                        startInfo.FileName = "ManifoldLister.exe";
                        const string argsSeparator = " ";
                        string args = string.Join(argsSeparator, output);
                        startInfo.Arguments = args;
                        try
                        {
                            Process.Start(startInfo);
                        }
                        catch (Exception)
                        {
                            MessageBox.Show("Was not able to start ManifoldLister.exe");
                        }
                    }

                    //OpenFolder(output);

                }
                catch (Exception) { }
            }
            mail.Close(Outlook.OlInspectorClose.olDiscard);
            GC.Collect();


        }

        // Function to read pdf attachments to search for manifold looking numbers
        private static List<string> Attachment(string pdfFile)
        {
            PdfReader reader = new PdfReader(pdfFile);
            string data = PdfTextExtractor.GetTextFromPage(reader, 1);
            reader.Close();
            List<string> output = new List<string>();
            try
            {
                MatchCollection m = Regex.Matches(data, @"\d\d?\w+-\d\d?-.+?(?=\s)", RegexOptions.IgnoreCase);
                if (m.Count > 0)
                {
                    for (int i = 0; i < m.Count; i++)
                    {
                        output.Add(m[i].ToString());
                    }
                    return output;
                }

            }
            catch (SystemException)
            {
            }
            return output;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
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
