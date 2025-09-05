// OutlookComSender.cs
using System;
using System.Collections.Generic;
using System.IO;
using Outlook = Microsoft.Office.Interop.Outlook;

public static class OutlookComSender
{
    public static void Send(string subject, string body, IEnumerable<string> attachments, IEnumerable<string> to)
    {
        var app = new Outlook.Application();
        var mail = (Outlook.MailItem)app.CreateItem(Outlook.OlItemType.olMailItem);

        foreach (var r in to) if (!string.IsNullOrWhiteSpace(r)) mail.Recipients.Add(r.Trim());
        mail.Subject = subject;
        mail.Body = body;

        foreach (var p in attachments) if (File.Exists(p)) mail.Attachments.Add(p);

        mail.Recipients.ResolveAll();
        mail.Send(); // goes out via the current Outlook profile
    }
}
