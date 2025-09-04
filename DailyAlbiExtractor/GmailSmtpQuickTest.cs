//using System;
//using System.Net;
//using System.Net.Mail;

//class GmailSmtpQuickTest
//{
//    static void Main()
//    {
//        Console.WriteLine("=== GMAIL SMTP QUICK TEST ===");
//        Console.Write("Your Gmail address: ");
//        var gmail = Console.ReadLine();

//        Console.Write("Your Gmail App Password (hidden): ");
//        var appPass = ReadPassword();

//        Console.Write("Recipient (try the SAME Gmail for first test): ");
//        var to = Console.ReadLine();

//        try
//        {
//            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls11;

//            using (var msg = new MailMessage())
//            {
//                msg.From = new MailAddress(gmail);   // MUST match authenticated account
//                msg.To.Add(to);
//                msg.Subject = "Gmail SMTP quick test";
//                msg.Body = "If you see this, Gmail SMTP works.";
//                msg.IsBodyHtml = false;

//                using (var smtp = new SmtpClient("smtp.gmail.com", 587))
//                {
//                    smtp.EnableSsl = true;                    // STARTTLS
//                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
