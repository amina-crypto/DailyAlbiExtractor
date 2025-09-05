//using System;
//using System.Net;
//using System.Net.Mail;

//class OutlookSmtpQuickTest
//{
//    static void Main()
//    {
//        Console.WriteLine("=== OUTLOOK / M365 SMTP QUICK TEST ===");
//        Console.Write("Your work email (login & From): ");
//        var email = Console.ReadLine();

//        Console.Write("Password (or App Password if MFA) (hidden): ");
//        var pass = ReadPassword();

//        Console.Write("Recipient (start by sending to yourself): ");
//        var to = Console.ReadLine();

//        try
//        {
//            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 | SecurityProtocolType.Tls13;

//            using (var msg = new MailMessage())
//            {
//                msg.From = new MailAddress(email);   // MUST match login mailbox
//                msg.To.Add(to);
//                msg.Subject = "M365 SMTP quick test";
//                msg.Body = "If you see this, Outlook SMTP works.";
//                msg.IsBodyHtml = false;

//                using (var smtp = new SmtpClient("smtp.office365.com", 587))
//                {
//                    smtp.EnableSsl = true;                    // STARTTLS
//                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
//                    smtp.UseDefaultCredentials = false;
//                    smtp.Credentials = new NetworkCredential(email, pass);
//                    smtp.Timeout = 30000;
//                    Console.WriteLine("Sending...");
//                    smtp.Send(msg);
//                }
//            }

//            Console.WriteLine("✅ Sent. Check Inbox/Junk and Sent.");
//        }
//        catch (SmtpException ex)
//        {
//            Console.WriteLine("❌ SMTP ERROR:\n" + ex);
//            if (ex.InnerException != null) Console.WriteLine("Inner: " + ex.InnerException.Message);
//        }
//        catch (Exception ex)
//        {
//            Console.WriteLine("❌ GENERAL ERROR:\n" + ex);
//        }
//    }

//    private static string ReadPassword()
//    {
//        var sb = new System.Text.StringBuilder();
//        ConsoleKeyInfo k;
//        while ((k = Console.ReadKey(true)).Key != ConsoleKey.Enter)
//        {
//            if (k.Key == ConsoleKey.Backspace) { if (sb.Length > 0) { sb.Length--; Console.Write("\b \b"); } }
//            else if (!char.IsControl(k.KeyChar)) { sb.Append(k.KeyChar); Console.Write("*"); }
//        }
//        Console.WriteLine();
//        return sb.ToString();
//    }
//}
