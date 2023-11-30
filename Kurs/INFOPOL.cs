using ComponentFactory.Krypton.Toolkit;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ListView;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

namespace Kurs
{
    public partial class INFOPOL : KryptonForm
    {
        public INFOPOL()
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

        private void kryptonPanel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            string query = $"UPDATE Users SET Login = '{kryptonTextBox1.Text}', email = '{kryptonTextBox2.Text}', age = '{kryptonTextBox3.Text}', weight = '{kryptonTextBox4.Text}', height = '{kryptonTextBox5.Text}', " +
                $"gender = '{kryptonTextBox27.Text}' , name = '{kryptonTextBox7.Text}' , surname = '{kryptonTextBox8.Text}' , number = '{kryptonTextBox9.Text}' , date = '{kryptonTextBox6.Text}' WHERE user_id = '{BaseData.id}'";

            SQLiteConnection conn = new SQLiteConnection(Connect());
            // Команды на взятие данных пользователя
            conn.Open();
            SQLiteCommand command = new SQLiteCommand(query, conn);
            command.ExecuteNonQuery();
            MessageBox.Show($"Успех, '{BaseData.id}'");
            conn.Close();
            Form3 frm3 = new Form3();
            frm3.Show();
            this.Close();
        }
    }
}
