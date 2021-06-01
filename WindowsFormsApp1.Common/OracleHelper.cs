using System;
using System.Data;
using Oracle.ManagedDataAccess.Client;

namespace WindowsFormsApp1.Common
{
	internal class OracleHelper
	{
		protected OracleConnection conn;

		protected OracleDataAdapter oda;

		protected OracleCommand cmd;

		public void OpenConn()
		{
			if (conn.State == ConnectionState.Closed)
			{
				conn.Open();
			}
		}

		public DataTable ExcuteSqlReturnDataTable(string strSql, string strConn)
		{
			DataTable dt = new DataTable();
			try
			{
				using (conn = new OracleConnection(strConn))
				{
					conn.Open();
					oda = new OracleDataAdapter(strSql, conn);
					oda.Fill(dt);
					conn.Close();
					conn.Dispose();
				}
			}
			catch (Exception ex)
			{
				DataRow row = dt.NewRow();
				dt.Columns.Add("MSG");
				row[0] = ex.Message;
				dt.Rows.Add(row);
			}
			return dt;
		}
	}
}
