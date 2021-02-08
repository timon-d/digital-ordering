using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;

namespace DigitalOrdering.Classes
{
    class MailNotification
    {
        private static String server;
        private static String senderAddress;
        private static String senderDisplayName = "Bestellung";
        private static String[] loggingRecipients = {"timon.dages@ipm.fraunhofer.de"};

        private SmtpClient smtpClient;
        private MailMessage notification;


        public MailNotification(String[] recipients, String subject, String body)
        {
            this.smtpClient = new SmtpClient(server);
            this.notification = new MailMessage();
            this.notification.From = new MailAddress(senderAddress, senderDisplayName);
            this.notification.IsBodyHtml = true;
            this.notification.BodyEncoding = System.Text.Encoding.UTF8;
            SetRecipients(recipients);
            this.notification.Subject = subject;
            this.notification.Body = body;
        }

        private void SetRecipients(String[] recipients){
            foreach (String recipient in recipients)
            {
                if (recipient != "")
                {
                    this.notification.To.Add(new MailAddress(recipient));
                }
            }
            foreach (String loggingRecipient in loggingRecipients)
            {
                if (loggingRecipient != "")
                {
                    this.notification.Bcc.Add(new MailAddress(loggingRecipient));
                }
            }
        }

        public void InsertAttachment(byte[] byteArray, String fileName)
        {
            if (byteArray != null && byteArray.Length > 0)
            {
                MemoryStream memoryStream = new MemoryStream(byteArray);
                Attachment attachment = new Attachment(memoryStream, fileName);
                this.notification.Attachments.Add(attachment);
            }
        }

        public void SendMailNotification()
        {
            try
            {
                this.smtpClient.Send(notification);
            } catch (Exception ex)
            {

            }
        }
    }
}
