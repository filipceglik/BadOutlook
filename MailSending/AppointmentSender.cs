using Microsoft.Office.Interop.Outlook;
using System;
using System.Xml.Linq;

namespace MailSending
{
    class AppointmentSender : OutlookConn
    {
        public Microsoft.Office.Interop.Outlook.Application oApp = GetOutlookInstance();
        public Microsoft.Office.Interop.Outlook.AppointmentItem appoint;

        public void Reminder(Application app, int days, XDocument doc, string attPath)
        {
            AppointmentItem appoint = app.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olAppointmentItem);
            appoint.Subject = doc.Root.Element("Subject").Value;
            appoint.Location = doc.Root.Element("Location").Value;
            appoint.Body = doc.Root.Element("Body").Value;
            appoint.Attachments.Add(attPath, OlAttachmentType.olByValue, 1);
            appoint.Sensitivity = Microsoft.Office.Interop.Outlook.OlSensitivity.olPrivate;
            appoint.Start = DateTime.Now.AddDays(days);
            appoint.Duration = 120; //minutes
            appoint.ReminderSet = true;
            appoint.ReminderMinutesBeforeStart = 15;
            appoint.Save();
        }

       
        /*public string startTime { get; set; }
        public string endTime { get; set; }
        public string attachment { get; set; }
        MailMessage msg = new MailMessage();*/

    }
}
