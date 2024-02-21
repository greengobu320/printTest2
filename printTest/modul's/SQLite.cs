using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SQLite;
using System.Data;
using System.Windows.Controls;
using System.Windows.Forms;

namespace printTest.modul_s
{
    internal class SQLite
    {
        public DataTable FillGrid(string path, string cmd) {
            DataTable dt = new DataTable();
            string connectionString = $@"Data Source={path}";
            using (var connection = new SQLiteConnection(connectionString))
            {
                connection.Open();
                DataSet ds = new DataSet();
                SQLiteCommand c = new SQLiteCommand(cmd, connection);
                SQLiteDataAdapter da = new SQLiteDataAdapter(c);
                da.Fill(ds, "temDt");
                dt.Merge(ds.Tables[0]);
            }
                return dt;
        }
    }
}
