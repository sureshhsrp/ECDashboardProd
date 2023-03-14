using System;
using System.Collections;
using System.Data;
using System.Configuration;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Web.UI.WebControls.WebParts;
using System.Web.UI.HtmlControls;

using System.Data.SqlClient;
using System.Security.Cryptography;
using System.IO;
using System.Text;

namespace ProductionSheetDashBoard.Utils
{
    public class Utils
    {
        public string strProvider = ConfigurationManager.ConnectionStrings["ConnectionString"].ToString();
        public string sqlText;
        public int CommandTimeOut = 0;
        public SqlConnection objConnection;

        public SqlDataReader GetReader()
        {
            SqlDataReader dr;
            SqlCommand cmd = new SqlCommand();
            cmd.CommandTimeout = 0;
            MakeConnection();
            OpenConnection();
            cmd = GetCommand();
            cmd.CommandTimeout = 0;
            dr = cmd.ExecuteReader();
            return dr;
        }
        SqlCommand GetCommand()
        {
            return new SqlCommand(sqlText, objConnection);
        }
        public void MakeConnection()
        {

            objConnection = new SqlConnection(strProvider);
            // objConnection.= 200000;
        }
        public void OpenConnection()
        {

            objConnection.Open();
        }
        public void CloseConnection()
        {
            objConnection.Close();
        }
        public static string getScalarValue(string SQLString, string CnnString)
        {

            string ReturnValue = string.Empty;
            using (SqlConnection conn = new SqlConnection(CnnString))
            {
                SqlCommand cmd = new SqlCommand(SQLString, conn);
                conn.Open();
                ReturnValue = cmd.ExecuteScalar().ToString(); ;
            }
            return ReturnValue;
        }
        public static int ExecNonQuery(string SQLString, string CnnString)
        {
            try
            {
                int count = 0;
                using (SqlConnection connection = new SqlConnection(CnnString))
                {
                    SqlCommand command = new SqlCommand(SQLString, connection);
                    command.Connection.Open();
                    count = command.ExecuteNonQuery();
                    command.Connection.Close();
                }
                return count;
            }
            catch (SqlException ex)
            {
                if (ex.Number == 2627)
                {
                    return ex.Number;
                }
                throw ex;
            }
            catch (Exception ex)
            {

                throw ex;

            }
        }

     
        public static DataTable GetDataTable(string SQLString, string CnnString)
        {
            Utils dbLink = new Utils();
            dbLink.strProvider = CnnString.ToString();
            dbLink.CommandTimeOut = 0;
            dbLink.sqlText = SQLString.ToString();
            SqlDataReader dr = dbLink.GetReader();
            DataTable tb = new DataTable();
            tb.Load(dr);
            dr.Close();
            dbLink.CloseConnection();
            return tb;
        }

    }
}