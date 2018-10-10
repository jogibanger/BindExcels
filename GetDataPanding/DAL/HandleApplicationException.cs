using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

namespace DAL
{
    public class HandleApplicationException
    {

        public static int WriteException(Exception ex)
        {
            string StoreProcedureName = "InsertErroLog";

            SqlParameter[] sqlParameter = {
                    new SqlParameter("HelpLink",ex.HelpLink),
                    new SqlParameter("InnerException",ex.InnerException),
                    new SqlParameter("Source",ex.Source),
                    new SqlParameter("Message",ex.Message),
                    new SqlParameter("StackTrace",ex.Message),

                };
         
            return sqlDML.ExecuteNonquery(StoreProcedureName, sqlParameter, CommandType.StoredProcedure);


        }
    }
}