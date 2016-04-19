namespace analyzer
{
    using System;
    using System.Data.SqlClient;

    public class usageLog
    {
        public SqlConnection CMTConnection;
        public int renderCnt;
        public int returnRows;
        public int SID;
        public int usageID;
        public string viewNm;

        public usageLog(int sessionID, string viewName, int resultRows)
        {
            this.SID = sessionID;
            this.viewNm = viewName;
            this.returnRows = resultRows;
            this.CMTConnection = new SqlConnection(@"server=psqltdy01\tdy01;database=cmt;uid=CMT_User;pwd=Bonds756");
            SqlCommand command = new SqlCommand("DECLARE @usageID int EXEC @usageID = startAnalyzerUsage @SID = " + this.SID.ToString() + ", @viewName = '" + this.viewNm + "' , @retRows = " + this.returnRows.ToString() + " SELECT 'usageID' = @usageID", this.CMTConnection);
            this.CMTConnection.Open();
            this.usageID = Convert.ToInt32(command.ExecuteScalar());
            this.CMTConnection.Close();
        }

        public void endUsage()
        {
            SqlCommand command = new SqlCommand("update analyzerUsageLog set releaseTime = getdate(), renderCount = " + this.renderCnt.ToString() + " where sid = " + this.SID.ToString() + " and usageID = " + this.usageID.ToString(), this.CMTConnection);
            this.CMTConnection.Open();
            command.ExecuteNonQuery();
            this.CMTConnection.Close();
        }
    }
}

