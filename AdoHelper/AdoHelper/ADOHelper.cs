using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;

namespace AdoHelper
{
    class ADOHelper
    {
        private static string connectionString = string.Empty;
        public ADOHelper(string cs)
        {
            connectionString = cs;
        }


        //Call this method to insert list of data 
        public bool BulkInsertData<T>(List<T> itemList, string command, string paramName)
        {
            bool ret = false;
            try
            {
                DataTable dataTable = listToDataTable(itemList);
                using (SqlConnection con = new SqlConnection(connectionString))
                {
                    con.Open();
                    using (SqlCommand cmd = new SqlCommand(command, con))
                    {
                        cmd.CommandType = CommandType.StoredProcedure;
                        SqlParameter parm = new SqlParameter();
                        parm.ParameterName = paramName;
                        parm.SqlDbType = SqlDbType.Structured;
                        parm.Value = dataTable;
                        cmd.Parameters.Add(parm);
                        cmd.ExecuteNonQuery();
                        ret = true;
                    }
                }
            }
            catch (Exception ex)
            {

                throw new Exception("bulk insert exception " + ex);
            }
            return ret;
        }


        //Call this method to execute some query or store procedure with parameters
        public bool ExecuteNonQuery(string query, SqlParameter[] parameters)
        {
            try
            {
                using (SqlConnection _sqlConnection = new SqlConnection(connectionString))
                {
                    using (SqlCommand _sqlCommand = new SqlCommand(query, _sqlConnection))
                    {
                        _sqlCommand.CommandType = CommandType.StoredProcedure;
                        foreach (SqlParameter p in parameters)
                        {
                            _sqlCommand.Parameters.Add(p);
                        }
                        _sqlConnection.Open();
                        _sqlCommand.ExecuteNonQuery();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        //Call this method to ExecuteNonQuery
        public bool ExecuteNonQuery(string query)
        {
            try
            {
                using (SqlConnection _sqlConnection = new SqlConnection(connectionString))
                {
                    using (SqlCommand _sqlCommand = new SqlCommand(query, _sqlConnection))
                    {
                        _sqlCommand.CommandType = CommandType.Text;
                        _sqlConnection.Open();
                        _sqlCommand.ExecuteNonQuery();
                    }
                }
                return true;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        
        //Call this method to Execute a query and get result as datatable
        public DataTable ExecuteDataTable(string query)
        {
            DataTable _dataTable = new DataTable();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand cmd = new SqlCommand(query, con))
                {
                    using (SqlDataAdapter _sqlDataAdapter = new SqlDataAdapter(cmd))
                    {
                        _sqlDataAdapter.Fill(_dataTable);
                    }
                }
            }
            return _dataTable;
        }

        //Call this method to Execute a query and get result as datatable
        public DataTable ExecuteDataTable(string query, SqlParameter[] parameters)
        {
            DataTable _dataTable = new DataTable();
            using (SqlConnection con = new SqlConnection(connectionString))
            {
                con.Open();
                using (SqlCommand _sqlCommand = new SqlCommand(query, con))
                {
                    _sqlCommand.CommandType = CommandType.StoredProcedure;
                    foreach (SqlParameter p in parameters)
                    {
                        _sqlCommand.Parameters.Add(p);
                    }
                    using (SqlDataAdapter _sqlDataAdapter = new SqlDataAdapter(_sqlCommand))
                    {
                        _sqlDataAdapter.Fill(_dataTable);
                    }
                }
            }
            return _dataTable;
        }

        private DataTable listToDataTable<T>(IList<T> data)
        {
            DataTable dt = new DataTable();
            if (typeof(T).IsValueType || typeof(T).Equals(typeof(string)))
            {
                DataColumn dc = new DataColumn("Value");
                dt.Columns.Add(dc);
                foreach (T item in data)
                {
                    DataRow dr = dt.NewRow();
                    dr[0] = item;
                    dt.Rows.Add(dr);
                }
            }
            else
            {
                PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));
                foreach (PropertyDescriptor prop in properties)
                {
                    dt.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);
                }
                foreach (T item in data)
                {
                    DataRow row = dt.NewRow();
                    foreach (PropertyDescriptor prop in properties)
                    {
                        try
                        {
                            row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;
                        }
                        catch (Exception ex)
                        {
                            row[prop.Name] = DBNull.Value;
                        }
                    }
                    dt.Rows.Add(row);
                }
            }
            return dt;

        }
    }
}
