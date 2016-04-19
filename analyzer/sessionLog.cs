namespace analyzer
{
    using System;
    using System.Data.SqlClient;

    public class sessionLog
    {
        public string appVer = "4";
        public SqlConnection CMTConnection = new SqlConnection(@"server=psqltdy01\tdy01;database=cmt;uid=CMT_User;pwd=Bonds756");
        public int SID;
        public string userID = Environment.UserName;

        public sessionLog()
        {
            SqlCommand command = new SqlCommand("declare @sid int exec @sid = startAnalyzerSession @appVer = '" + this.appVer + "', @userID = '" + this.userID + "' select 'sid' = @sid", this.CMTConnection);
            this.CMTConnection.Open();
            this.SID = Convert.ToInt32(command.ExecuteScalar());
            this.CMTConnection.Close();
        }

        public void endSession()
        {
            SqlCommand command = new SqlCommand("update analyzersessionlog set timeOut = getdate() where sid = " + this.SID.ToString(), this.CMTConnection);
            this.CMTConnection.Open();
            command.ExecuteNonQuery();
            this.CMTConnection.Close();
        }
    }
}

