using Microsoft.Office.Interop.Outlook;
using System;

namespace MailSending
{
    class AppointmentSender : OutlookConn
    {
        public Application oApp = GetOutlookInstance();

        public void Reminder(Application app, AppointmentDetails appointmentDetails)
        {
            AppointmentItem appoint = app.CreateItem(OlItemType.olAppointmentItem);
            appoint.Subject = appointmentDetails.Subject;
            appoint.Location = appointmentDetails.Location;
            appoint.Body = appointmentDetails.Body;
            appoint.Attachments.Add(appointmentDetails.PathToAttachment, OlAttachmentType.olByValue, 1);
            appoint.Sensitivity = OlSensitivity.olPrivate;
            appoint.Start = DateTime.Now.AddDays(appointmentDetails.DaysToAppointment);
            appoint.Duration = 120; //minutes
            appoint.ReminderSet = true;
            appoint.ReminderMinutesBeforeStart = 15;
            appoint.Save();
        }
    }
}
