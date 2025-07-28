using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Net.Mail;
using System.Net;

namespace DailyAlbiExtractor
{
    public class EmailSender
    {
        private readonly string _smtpServer;
        private readonly int _smtpPort;
        private readonly string _smtpUsername;
        private readonly string _smtpPassword;
        private readonly string _fromEmail;
        private readonly string[] _toEmails;

        public EmailSender(string smtpServer, int smtpPort, string smtpUsername, string smtpPassword, string fromEmail, string[] toEmails)
        {
            _smtpServer = smtpServer;
            _smtpPort = smtpPort;
            _smtpUsername = smtpUsername;
            _smtpPassword = smtpPassword;
            _fromEmail = fromEmail;
            _toEmails = toEmails;
        }

        public void SendEmail(string[] attachmentPaths)
        {
            using (var message = new MailMessage())
            {
                message.From = new MailAddress(_fromEmail);
                foreach (var to in _toEmails)
                {
                    message.To.Add(to);
                }
                message.Subject = "Daily API Extract - Changes and Full Data";
                message.Body = "Attached are the full data extract and any changes/additions from the previous day.";

                foreach (var path in attachmentPaths)
                {
                    message.Attachments.Add(new Attachment(path));
                }

                using (var smtp = new SmtpClient(_smtpServer, _smtpPort))
                {
                    smtp.Credentials = new NetworkCredential(_smtpUsername, _smtpPassword);
                    smtp.EnableSsl = true;
                    smtp.Send(message);
                }
            }
        }
    }
}