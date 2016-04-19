namespace analyzer
{
    using System;
    using System.Collections.Generic;
    using System.Data.SqlClient;
    using System.Windows.Forms;

    public static class errorLog
    {
        public static string filtIllegalXlsChars(string proposedName)
        {
            string str = proposedName;
            List<string> list = new List<string> { 
                "*",
                "/",
                @"\",
                "]",
                "[",
                "?",
                ":"
            };
            foreach (string str2 in list)
            {
                str = str.Replace(str2, string.Empty);
            }
            return str;
        }

        public static void writeError(Exception e, sessionLog sLog)
        {
            MessageBox.Show("The Analyzer has encountered an error and needs to close.  Detailed error information is being delivered to the administrators.  If you have any questions, please contact SSISupport@T-Mobile.com." + Environment.NewLine + Environment.NewLine + "Error message:" + Environment.NewLine + e.Message, "SSI");
            if (datter.connectionAvailable())
            {
                string str = string.Empty;
                if (e.InnerException != null)
                {
                    str = e.InnerException.Message.Replace("'", "''");
                }
                string cmdText = "insert into analyzerErrorLog (sid, tmStmp, message, stackTrace, targetSite, innerException) values (" + sLog.SID.ToString() + ", getdate(), '" + e.Message.Replace("'", "''") + " ', '" + e.StackTrace.Replace("'", "''") + " ', '" + e.TargetSite.Name.Replace("'", "''") + " ', '" + str + " ')";
                SqlConnection connection = new SqlConnection(@"server=psqltdy01\tdy01;database=cmt;uid=CMT_User;pwd=Bonds756");
                SqlCommand command = new SqlCommand(cmdText, connection);
                try
                {
                    connection.Open();
                    command.ExecuteNonQuery();
                    connection.Close();
                    sLog.endSession();
                }
                catch
                {
                    connection.Close();
                }
            }
        }
    }
}

