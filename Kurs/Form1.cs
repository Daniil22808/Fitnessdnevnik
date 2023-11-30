using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;
using System.Data.SQLite;


namespace Kurs
{
  
   
    public partial class Form1 : KryptonForm
    {
        static string Connect()
        {
            
            string databaseName = "fitness.db";
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string databasePath = System.IO.Path.Combine(executablePath, databaseName);

            string connectionString = $"Data Source={databasePath};Version=3;";

            return connectionString;
        }
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            Form2 formachka2 = new Form2();
            formachka2.Show();
            this.Hide();
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;


        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void kryptonButton1_Click(object sender, EventArgs e)
        {   
            if(kryptonTextBox1.Text == "" && kryptonTextBox2.Text == "")
            {
                MessageBox.Show("Пустые поля, Ошибка");
            }
            else
            {
                string query1 = "SELECT user_id FROM Users WHERE Login=@Login AND Parol = @Parol";
                SQLiteConnection conn = new SQLiteConnection(Connect());
                conn.Open();
                SQLiteCommand command1 = new SQLiteCommand(query1,conn);
                command1.Parameters.AddWithValue("@Login", kryptonTextBox1.Text);
                command1.Parameters.AddWithValue("@Parol", kryptonTextBox2.Text);
                SQLiteDataAdapter adapter = new SQLiteDataAdapter(command1);
                DataTable Users = new DataTable();
                adapter.Fill(Users);
                object result = command1.ExecuteScalar();
                int user_id = Convert.ToInt32(result);

                if (Users.Rows.Count > 0)
                {
                    BaseData.id = user_id;
                    Form3 frm3 = new Form3();
                    frm3.Show();
                    this.Hide();
                }
                else
                {
                    MessageBox.Show("Неверный логин или пароль");
                }
            }
            
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            kryptonTextBox2.PasswordChar = '\0';
            pictureBox4.Visible = true;
            pictureBox3.Visible = false;
        }

        private void pictureBox4_Click(object sender, EventArgs e)
        {   
            kryptonTextBox2.PasswordChar = '*';
            pictureBox3.Visible = true;
            pictureBox4.Visible = false;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            kryptonTextBox2.Text = "";
            kryptonTextBox1.Text = "";
        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
