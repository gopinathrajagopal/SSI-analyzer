namespace analyzer
{
    using System;
    using System.Data;
    using System.Data.SqlClient;

    public static class datter
    {
        public const string CMTConnection = @"server=psqltdy01\tdy01;database=cmt;uid=CMT_User;pwd=Bonds756";
        public const string primaryConFailMsg = "The Analyzer is unable to establish a connection to the CMT database.  Ensure that you are connected to the network and are on the GSM1900 domain.  If this problem continues, please contact SSI:  SSISupport@T-Mobile.com.";
        public const string secondaryConFailMsg = "The Analyzer is unable to establish a conntection to the CMT database at the moment.  Ensure that you are connected to the network and are on the GSM1900 domain.  If you are properly connected but continue to get this message for more than a few minutes, please reach out to the NRP Command Center to report the possible outage at 1-877-792-7286";
        public const string ssiIcmCon = @"server=psqltdy01\tdy01;database=careCallsLog;uid=CMT_User;pwd=Bonds756";

        public static bool connectionAvailable()
        {
            SqlConnection connection = new SqlConnection(@"server=psqltdy01\tdy01;database=cmt;uid=CMT_User;pwd=Bonds756");
            try
            {
                bool flag;
                connection.Open();
                if (connection.State == ConnectionState.Open)
                {
                    flag = true;
                }
                else
                {
                    flag = false;
                }
                connection.Close();
                return flag;
            }
            catch
            {
                connection.Close();
                return false;
            }
        }
    }
}

