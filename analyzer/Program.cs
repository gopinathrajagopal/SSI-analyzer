namespace analyzer
{
    using System;
    using System.Windows.Forms;

    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            if (!datter.connectionAvailable())
            {
                MessageBox.Show("The Analyzer is unable to establish a connection to the CMT database.  Ensure that you are connected to the network and are on the GSM1900 domain.  If this problem continues, please contact SSI:  SSISupport@T-Mobile.com.", "SSI");
            }
            else
            {
                sessionLog log = new sessionLog();
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                try
                {
                    Application.Run(new main(log));
                }
                catch (Exception exception)
                {
                    errorLog.writeError(exception, log);
                }
            }
        }
    }
}

