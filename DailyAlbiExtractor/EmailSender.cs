using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.IO;

namespace DailyAlbiExtractor
{
    public class EmailSender
    {
        private readonly string _smtpServer;
        private readonly int _smtpPort;
        private readonly string _smtpUsername;
        private readonly string _smtpPassword;
        private readonly string _fromEmail;
        private readonly IEnumerable<string> _toEmails;

        public EmailSender(
            string smtpServer,
            int smtpPort,
            string smtpUsername,
            string smtpPassword,
            string fromEmail,
            IEnumerable<string> toEmails)
        {
            _smtpServer = smtpServer;
            _smtpPort = smtpPort;
            _smtpUsername = smtpUsername;
            _smtpPassword = smtpPassword;
            _fromEmail = fromEmail;
            _toEmails = toEmails?.Where(e => !string.IsNullOrWhiteSpace(e)).Select(e => e.Trim())
                        ?? Enumerable.Empty<string>();
        }

        public void SendEmail(string subject, string body, IEnumerable<string> attachmentPaths)
        {
            if (!_toEmails.Any())
                throw new InvalidOperationException("No recipient emails were provided.");

            using (var message = new MailMessage())
            {
                message.From = new MailAddress(_fromEmail);
                foreach (var to in _toEmails)
                    message.To.Add(to);

                message.Subject = subject;
                message.Body = body;

                if (attachmentPaths != null)
                {
                    foreach (var path in attachmentPaths.Where(File.Exists))
                        message.Attachments.Add(new Attachment(path));
                }

                // Modern TLS (in case the host requires it)
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

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
