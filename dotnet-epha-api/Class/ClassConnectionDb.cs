using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using Xceed.Document.NET;

namespace Class
{
    public class ClassConnectionDb
    {
        string[] sMonth = ("JAN,FEB,MAR,APR,MAY,JUN,JUL,AUG,SEP,OCT,NOV,DEC").Split(',');
        string[] sMonths = ("Jan,Feb,Mar,Apr,May,Jun,Jul,Aug,Sep,Oct,Nov,Dec").Split(',');

        String ConnStrSQL = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionConfig")["ConnString"] ?? "";
        static public string ConnectionString()
        {
            return new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionConfig")["ConnString"] ?? "";
        }
        static public DataTable ExecuteAdapterSQLDataTable(string sqlStatement, List<SqlParameter>? parameters, string tableName = "Table1", bool isStoredProcedure = false)
        {
            // Load connection string from configuration
            string connStrSQL = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionConfig")["ConnString"] ?? "";
            if (string.IsNullOrEmpty(connStrSQL))
            {
                throw new ApplicationException("Connection string cannot be null or empty.");
            }

            DataSet dssql = new DataSet();
            using (SqlConnection connsql = new SqlConnection(connStrSQL))
            {
                connsql.Open();
                try
                {
                    using (SqlCommand cmd = new SqlCommand(sqlStatement, connsql))
                    {
                        // Set command type based on whether it is a stored procedure
                        cmd.CommandType = isStoredProcedure ? CommandType.StoredProcedure : CommandType.Text;

                        // Add parameters to the command
                        if (parameters != null)
                        {
                            foreach (var param in parameters)
                            {
                                // Check if parameter is not null and not already added
                                if (param != null && !cmd.Parameters.Contains(param.ParameterName))
                                {
                                    cmd.Parameters.Add(param);
                                }
                            }
                        }

                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        da.Fill(dssql);
                    }
                }
                catch (Exception ex)
                {
                    // Handle exceptions (log or rethrow as needed)
                    throw new ApplicationException("An error occurred while executing the SQL command.", ex);
                }
            }

            if (dssql.Tables.Count > 0)
            {
                foreach (DataColumn column in dssql.Tables[0].Columns)
                {
                    column.ColumnName = column.ColumnName.ToLower();
                }
                dssql.Tables[0].TableName = tableName;
                return dssql.Tables[0];
            }
            else
            {
                // Return an empty DataTable if no data is retrieved
                return new DataTable(tableName);
            }
        }
        static public string ExecuteNonQuerySQL(string sqlStatement, List<SqlParameter>? parameters)
        {
            // Load connection string from configuration
            string connStrSQL = new ConfigurationBuilder().AddJsonFile("appsettings.json").Build().GetSection("ConnectionConfig")["ConnString"] ?? "";
            if (string.IsNullOrEmpty(connStrSQL))
            {
                throw new ApplicationException("Connection string cannot be null or empty.");
            }

            using (SqlConnection connsql = new SqlConnection(connStrSQL))
            {
                connsql.Open();
                try
                {
                    using (SqlCommand cmd = new SqlCommand(sqlStatement, connsql))
                    {
                        cmd.CommandTimeout = 300; // 5 minutes timeout
                        if (parameters != null)
                        {
                            foreach (var param in parameters)
                            {
                                if (param != null && !cmd.Parameters.Contains(param.ParameterName))
                                {
                                    cmd.Parameters.Add(param);
                                }
                            }
                        }

                        try
                        {
                            cmd.ExecuteNonQuery();
                            return "true";
                        }
                        catch (SqlException ex)
                        {
                            // Log the exception (consider using a logging framework)
                            return $"error: Database operation failed.{ex.Message.ToString()}";
                        }
                        catch (Exception ex)
                        {
                            // Log the exception (consider using a logging framework)
                            return $"error: An unexpected error occurred.{ex.Message.ToString()}";
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log the exception (consider using a logging framework)
                    throw new ApplicationException("An error occurred while executing the SQL command.", ex);
                }
            }
        }

        public string ExecuteNonQuerySQLTrans(string sqlStatement, List<SqlParameter> parameters, SqlConnection conn, SqlTransaction? trans = null)
        {
            if (trans == null)
            {
                using (SqlCommand cmd = new SqlCommand(sqlStatement, conn))
                {
                    cmd.CommandTimeout = 300; // 5 minutes timeout
                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            if (param != null && !cmd.Parameters.Contains(param.ParameterName))
                            {
                                cmd.Parameters.Add(param);
                            }
                        }
                    }

                    if (trans != null)
                    {
                        cmd.Transaction = trans;
                    }

                    try
                    {
                        cmd.ExecuteNonQuery();
                        return "true";
                    }
                    catch (Exception ex)
                    {
                        return "error: " + ex.Message;
                    }
                }
            }
            else
            {
                using (SqlCommand cmd = new SqlCommand(sqlStatement, conn, trans))
                {
                    cmd.CommandTimeout = 300; // 5 minutes timeout
                    if (parameters != null)
                    {
                        foreach (var param in parameters)
                        {
                            if (param != null && !cmd.Parameters.Contains(param.ParameterName))
                            {
                                cmd.Parameters.Add(param);
                            }
                        }
                    }

                    try
                    {
                        cmd.ExecuteNonQuery();
                        return "true";
                    }
                    catch (Exception ex)
                    {
                        return "error: " + ex.Message;
                    }
                }
            }
        }

        static private void ChangeColumnNamesToLowerCase(DataTable table)
        {
            foreach (DataColumn column in table.Columns)
            {
                column.ColumnName = column.ColumnName.ToLower();
            }
        }
        public string CleanInput(string input, int maxLength)
        {
            if (input == null) return string.Empty;

            // ลบอักขระที่ไม่พึงประสงค์
            input = input.Replace("'", string.Empty);

            // ตรวจสอบและจำกัดความยาวของค่า
            if (input.Length > maxLength)
            {
                input = input.Substring(0, maxLength);
            }

            return input;
        }


        public static SqlParameter CreateSqlParameter(string parameterName, SqlDbType dbType, object value, int length = 0)
        {
            if (value == null)
            {
                return new SqlParameter(parameterName, dbType) { Value = DBNull.Value };
            }

            if (dbType == SqlDbType.Int)
            {
                if (int.TryParse(value.ToString(), out int intValue))
                {
                    return new SqlParameter(parameterName, dbType) { Value = intValue };
                }
                else
                {
                    return new SqlParameter(parameterName, dbType) { Value = DBNull.Value };
                }
            }
            else if (dbType == SqlDbType.Decimal)
            {
                if (decimal.TryParse(value.ToString(), out decimal decimalValue))
                {
                    return new SqlParameter(parameterName, dbType) { Value = decimalValue };
                }
                else
                {
                    return new SqlParameter(parameterName, dbType) { Value = DBNull.Value };
                }
            }
            else if (dbType == SqlDbType.DateTime)
            {
                if (DateTime.TryParse(value.ToString(), out DateTime dateTimeValue))
                {
                    return new SqlParameter(parameterName, dbType) { Value = dateTimeValue };
                }
                else
                {
                    return new SqlParameter(parameterName, dbType) { Value = DBNull.Value };
                }
            }
            else if (dbType == SqlDbType.VarChar)
            {
                if (length > 0)
                {
                    return new SqlParameter(parameterName, dbType, length) { Value = value != null ? value.ToString() : DBNull.Value };
                }
                else
                {
                    return new SqlParameter(parameterName, dbType) { Value = value != null ? value.ToString() : DBNull.Value };
                }
            }
            else
            {
                return new SqlParameter(parameterName, dbType) { Value = value != null ? value.ToString() : DBNull.Value };
            }
        } 
    }

}
