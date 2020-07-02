using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SODReview
{
    class DataRal
    {
        SqlConnection conn = null;
        public DataRal(string connStr)
        {
            conn = new SqlConnection(connStr);
        }

        public DataTable ExecutePS(string storeName)
        {
            try
            {
                conn.Open();
                SqlCommand comm = new SqlCommand(storeName, conn);
                comm.CommandText = storeName;
                
                DataTable dt = new DataTable();
                using (SqlDataAdapter da = new SqlDataAdapter(comm))
                {
                    da.Fill(dt);
                    return dt;
                }

            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
            finally
            {
                conn.Close();
                
            }
            return null;
        }
    }
}
