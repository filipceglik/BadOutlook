using System;
using System.Xml.Linq;

namespace MailSending
{
    class Program
    {
        static void Main(string[] args)
        {
            var doc = XDocument.Load(Environment.GetEnvironmentVariable("USERPROFILE") + '\\' + args[0]);
            AppointmentSender inv = new AppointmentSender();
            inv.Reminder(inv.oApp, int.Parse(args[1]), doc, Environment.GetEnvironmentVariable("USERPROFILE") +'\\'+ args[2]);
        }
    }
}
