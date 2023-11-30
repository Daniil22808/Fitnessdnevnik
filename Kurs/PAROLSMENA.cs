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
    public partial class PAROLSMENA : Form
    {
        public PAROLSMENA()
        {
            InitializeComponent();
        }
        static string Connect()
        {

            string databaseName = "fitness.db";
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string databasePath = System.IO.Path.Combine(executablePath, databaseName);

            string connectionString = $"Data Source={databasePath};Version=3;";

            return connectionString;
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            Form3 frm3 = new Form3();
            frm3.Show();
            this.Close();
        }

        private void kryptonButton8_Click(object sender, EventArgs e)
        {
            if (kryptonTextBox10.Text != "" && kryptonTextBox9.Text != "" && kryptonTextBox6.Text != "" && kryptonTextBox9.Text == kryptonTextBox6.Text)
            {
                string query = $"SELECT Parol FROM Users WHERE user_id = '{BaseData.id}'";

                SQLiteConnection conn = new SQLiteConnection(Connect());

                conn.Open();
                SQLiteCommand command = new SQLiteCommand(query, conn);
                object parol = command.ExecuteScalar();
                string parol2 = Convert.ToString(parol);
                if (parol2 == kryptonTextBox10.Text)
                {
                    string query2 = $"UPDATE Users SET Parol = '{kryptonTextBox6.Text}' WHERE user_id = '{BaseData.id}'";
                    SQLiteCommand command2 = new SQLiteCommand(query2, conn);
                    command2.ExecuteNonQuery();
                }
                else
                {
                    MessageBox.Show("Ошибка");
                }
            }
            else
            {
                MessageBox.Show("Ошибка");
            }
        }
    }
}
