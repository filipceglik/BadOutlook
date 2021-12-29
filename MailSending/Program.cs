using CommandLine;

namespace MailSending
{
    class Program
    {
        static void Main(string[] args)
        {
            Parser.Default.ParseArguments<Options>(args)
                .WithParsed<Options>(o =>
                {
                    AppointmentDetails appointmentDetails = new AppointmentDetails()
                    {
                        Subject = o.Subject,
                        Location = o.Location,
                        Body = o.Body,
                        PathToAttachment = o.Path,
                        DaysToAppointment = o.Days
                    };
                    AppointmentSender inv = new AppointmentSender();
                    inv.Reminder(inv.oApp, appointmentDetails);
                });    
        }

        public class Options
        {
            [Option('s', "subject", Required = true, HelpText = "Subject of the e-mail. ")]
            public string Subject { get; set; }
            [Option('l', "location", Required = true, HelpText = "Appointment location. ")]
            public string Location { get; set; }
            [Option('b', "body", Required = true, HelpText = "Appointment body. ")]
            public string Body { get; set; }
            [Option('p', "path", Required = true, HelpText = "Path to the attachment. ")]
            public string Path { get; set; }
            [Option('d', "days", Required = true, HelpText = "Days until the appointment. ")]
            public int Days { get; set; }
        }
    }
}
