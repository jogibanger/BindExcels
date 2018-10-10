using System;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace DAL
{
    public class sqlDML
    {
        private SqlDataAdapter myAdapter;
        private static SqlConnection conn = null;
       
        private string DataBaseName = string.Empty;
        private string UserID = string.Empty;
        private string Password = string.Empty;
        private string ServerName = string.Empty;

        // Initialise Connection
        public sqlDML()
        {
            try
            {
                myAdapter = new SqlDataAdapter();
            }
            catch (Exception)
            {

                throw;
            }

        }

        // Open Database Connection if Closed or Broken
        private static SqlConnection openConnection()
        {
            try
            {

                string connection = "Data Source=45.249.111.172;Initial Catalog=AEPS;Persist Security Info=True;User ID=sa;Password=Cbw8@WudpJN4c$";// ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString;
               
                conn = new SqlConnection(connection);
           
                if (conn.State == ConnectionState.Closed || conn.State == ConnectionState.Broken)
                {
                    conn.Open();
                }
                return conn;

            }
            catch (Exception)
            {

                throw;
            }

        }

        // Close Connection....
        private static SqlConnection CloseConnection()
        {
            try
            {
                if (conn.State == ConnectionState.Open || conn.State == ConnectionState.Broken)
                {
                    conn.Close();
                }
                return conn;

            }
            catch (Exception)
            {

                throw;
            }

        }

        // Insert data through Text/Procedure with sql parameters
        public static int ExecuteNonquery(string query, SqlParameter[] sqlParameter, CommandType commandType)
        {
            int TempValue = 0;
            SqlCommand sqlCommand = new SqlCommand();
            try
            {
                sqlCommand.Connection = openConnection();
                sqlCommand.CommandText = query;
                sqlCommand.CommandType = commandType;
                sqlCommand.CommandTimeout = 5000;

                if (sqlParameter != null)
                    sqlCommand.Parameters.AddRange(sqlParameter);

                TempValue = sqlCommand.ExecuteNonQuery();
            }
            catch (SqlException)
            {
                CloseConnection();
                throw;
            }
            finally
            {
                CloseConnection();
            }
            return TempValue;
        }

        // To get Single Record
        public static string GetSingleRecord(string query, SqlParameter[] sqlParameter, CommandType commandType)
        {
            string TempValue = string.Empty;
            SqlCommand sqlCommand = new SqlCommand();
            try
            {
                sqlCommand.Connection = openConnection();
                sqlCommand.CommandText = query;
                sqlCommand.CommandType = commandType;

                if (sqlParameter != null)
                    sqlCommand.Parameters.AddRange(sqlParameter);

                TempValue = Convert.ToString(sqlCommand.ExecuteScalar());

            }
            catch (SqlException)
            {
                CloseConnection();
                throw;
            }
            finally
            {
                CloseConnection();
            }
            return TempValue;
        }

        /// Get Single Records using query and command type...
        /// </summary>
        /// <param name="query"></param>
        /// <param name="commandType"></param>
        /// <returns></returns>

        public string GetSingleRecord(string query, CommandType commandType)
        {
            string TempValue = string.Empty;
            SqlCommand sqlCommand = new SqlCommand();
            try
            {
                sqlCommand.Connection = openConnection();
                sqlCommand.CommandText = query;
                sqlCommand.CommandType = commandType;
                TempValue = Convert.ToString(sqlCommand.ExecuteScalar());

            }
            catch (SqlException)
            {
                CloseConnection();
                throw;
            }
            finally
            {
                CloseConnection();
            }
            return TempValue;
        }
        /// Get Multiple Records
        /// </summary>
        /// <param name="query">Query is user for passing the complete query or store procedure name</param>
        /// <param name="sqlParameter"></param>
        /// <param name="commandType"></param>
        /// <returns></returns>
        public static DataTable GetRecords(string query, SqlParameter[] sqlParameter, CommandType commandType)
        {
            SqlCommand sqlCommand = new SqlCommand();
            DataTable records = null;
            try
            {
                records = new DataTable();

                sqlCommand.Connection = openConnection();
                sqlCommand.CommandText = query;
                sqlCommand.CommandType = commandType;

                if (sqlParameter != null)
                    sqlCommand.Parameters.AddRange(sqlParameter);

                records.Load(sqlCommand.ExecuteReader());

            }
            catch (SqlException)
            {
                CloseConnection();
                throw;
            }
            finally
            {
                CloseConnection();
            }
            return records;
        }
        /// Get The Records with Query only and passed alos please command Type
        /// </summary>
        /// <param name="query"></param>
        /// <param name="commandType"></param>
        /// <returns></returns>
        public DataTable GetRecords(string query, CommandType commandType)
        {

            SqlCommand sqlCommand = new SqlCommand();
            DataTable records = null;
            try
            {
                records = new DataTable();
                sqlCommand.Connection = openConnection();
                sqlCommand.CommandText = query;
                sqlCommand.CommandType = commandType;
                records.Load(sqlCommand.ExecuteReader());

            }
            catch (SqlException)
            {
                CloseConnection();
                throw;
            }
            finally
            {
                CloseConnection();
            }
            return records;
        }

        public DataTable GetSingleRecord(SqlCommand Cmd, CommandType commandType)
        {

            string Result = string.Empty;
            DataTable dt = new DataTable();
            try
            {


                Cmd.Connection = openConnection();
                Cmd.CommandType = commandType;
                dt.Load(Cmd.ExecuteReader());





            }
            catch (SqlException)
            {
                CloseConnection();
                throw;
            }
            finally
            {
                CloseConnection();
            }
            return dt;

        }

    }
}
