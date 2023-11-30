using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SQLite;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Kurs
{
    public partial class AddPlan : Form
    {
        static string Connect()
        {

            string databaseName = "fitness.db";
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string databasePath = System.IO.Path.Combine(executablePath, databaseName);

            string connectionString = $"Data Source={databasePath};Version=3;";

            return connectionString;
        }
        public AddPlan()
        {
            InitializeComponent();
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            string query = $"INSERT INTO Plans (plan_name, DatePlan, iduser) VALUES ('{kryptonTextBox1.Text}', '{kryptonDateTimePicker1.Value.ToString()}', '{BaseData.id}')";
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();
            SQLiteCommand comm = new SQLiteCommand(query, conn);
            comm.ExecuteNonQuery();
            MessageBox.Show("План создан");
            conn.Close();
            Form3 frm3 = new Form3();
            frm3.Show();
            this.Close();
        }
    }
}
