// EmailSender.cs
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;

namespace DailyAlbiExtractor
{
    public class EmailSender
    {
        private const string Host = "smtp.office365.com";
        private const int Port = 587;

        private readonly string _account;      // your service mailbox (login & From)
        private readonly string _password;     // normal password (no MFA for this mailbox)
        private readonly IEnumerable<string> _to;

        public EmailSender(string accountEmail, string password, IEnumerable<string> recipients)
        {
            _account = (accountEmail ?? "").Trim();
            _password = password ?? "";
            _to = (recipients ?? Enumerable.Empty<string>())
                       .Where(x => !string.IsNullOrWhiteSpace(x))
                       .Select(x => x.Trim());
        }

        public void Send(string subject, string body, IEnumerable<string> attachmentPaths)
        {
            if (string.IsNullOrWhiteSpace(_account))
                throw new InvalidOperationException("Sender account email is required.");
            if (!_to.Any())
                throw new InvalidOperationException("At least one recipient is required.");

            using (var msg = new MailMessage())
            {
                msg.From = new MailAddress(_account);     // MUST match the authenticated mailbox
                foreach (var r in _to) msg.To.Add(r);
                msg.Subject = subject;
                msg.Body = body;
                msg.IsBodyHtml = false;

                if (attachmentPaths != null)
                {
                    foreach (var path in attachmentPaths)
                    {
                        if (!File.Exists(path))
                        {
                            Console.WriteLine($"[WARN] Attachment not found: {path}");
                            continue;
                        }
                        Console.WriteLine($"Attaching: {path}");
                        msg.Attachments.Add(new Attachment(path));
                    }
                }

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12; // TLS 1.2 is enough for EXO

                using (var smtp = new SmtpClient(Host, Port))
                {
                    smtp.EnableSsl = true;                       // STARTTLS
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new NetworkCredential(_account, _password);
                    smtp.Timeout = 30000;

                    smtp.Send(msg);
                }
            }
        }
    }
}
