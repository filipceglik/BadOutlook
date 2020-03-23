using System.Runtime.InteropServices;

namespace MailSending
{
    class OutlookConn
    {
        public static Microsoft.Office.Interop.Outlook.Application GetOutlookInstance()
        {
            Microsoft.Office.Interop.Outlook.Application ol = null;
            System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcessesByName("OUTLOOK");
            int collCount = processes.Length;
            if (collCount != 0)
                ol = (Microsoft.Office.Interop.Outlook.Application)Marshal.GetActiveObject("Outlook.Application");
            else
                ol = new Microsoft.Office.Interop.Outlook.Application();

            return ol;
        }
    }
}
