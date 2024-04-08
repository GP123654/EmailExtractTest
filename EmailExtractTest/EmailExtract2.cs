using System;
using System.IO;
using MailKit.Net.Imap;
using MailKit;
using MimeKit;

namespace EmailExtractTest
{
    /// <summary>
    /// Class used to get information and attachments from an email
    /// </summary>
    public class EmailExtract2
    {
        /// <summary>
        /// Creates a unique file name - can always split at the _ if needed
        /// </summary>
        /// <param name="basePath">Directory of where the file is or will be / file path</param>
        /// <param name="originalName">Name of the file</param>
        /// <returns></returns>
        private static string GetUniqueFilename(string basePath, string originalName)
        {
            //Removes the extention and gets the file name
            string fileNameWithoutExtension = Path.GetFileNameWithoutExtension(originalName);
            //Making it a txt file / but can be original
            string extension = ".txt"; //Path.GetExtension(originalName);
            //Creates the unique file name gets the name without the extension and adds an underscore with the date to make it unique and then adds the extention
            string uniqueFileName = fileNameWithoutExtension + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + extension;
            //Puts the unique file into the location or something like that
            return Path.Combine(basePath, uniqueFileName);
        }

        /// <summary>
        /// Extracts the information and the attachments from the email
        /// </summary>
        public void ExtractDataFromEmail()
        {
            //Imap server for gmail or whichever server you are using
            string host = "imap.gmail.com";
            //Imap port
            int port = 993;

            using (var client = new ImapClient())
            {
                //Connects to the email server
                client.Connect(host, port, true);
                //Authenticates the email of the account that you are getting the emails from / the ones you want to access
                client.Authenticate("greg.postings97@gmail.com", "swnfmrvdruttdaei");
                //Opens the inbox and gives read / write permissions to access the emails
                client.Inbox.Open(FolderAccess.ReadWrite);

                //This gets the newest message sent to the inbox
                var message = client.Inbox.GetMessage(client.Inbox.Count - 1);

                //Can remove this but can be useful for troubleshooting and stuff
                Console.WriteLine("Subject: " + message.Subject);
                Console.WriteLine("From: " + message.From);
                Console.WriteLine("Date: " + message.Date);

                //Path that we save the files to
                string savePath = "D:\\Other Things\\Not for school\\RevRed\\EmailExtractTest\\files"; // Set your desired directory path for saving attachments

                //For each of the attachments in the message
                foreach (var attachment in message.Attachments)
                {
                    string uniqueFileName = GetUniqueFilename(savePath, attachment.ContentDisposition?.FileName ?? "Attachment");

                    using (var stream = File.Create(uniqueFileName))
                    {
                        if (attachment is MessagePart rfc822)
                        {
                            rfc822.Message.WriteTo(stream);
                        }
                        else
                        {
                            var part = (MimePart)attachment;
                            part.Content.DecodeTo(stream);
                        }
                    }
                }

                //Disconnects from the server
                client.Disconnect(true);
            }
        }
    }
}