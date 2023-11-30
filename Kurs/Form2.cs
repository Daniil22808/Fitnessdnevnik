using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ComponentFactory.Krypton.Toolkit;
using System.Data.SqlClient;
using System.Data.SQLite;
using System.Text.RegularExpressions;

namespace Kurs
{
    public partial class Form2 : KryptonForm
    {

        static string Connect()
        {

            string databaseName = "fitness.db";
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string databasePath = System.IO.Path.Combine(executablePath, databaseName);

            string connectionString = $"Data Source={databasePath};Version=3;";

            return connectionString;
        }
        public Form2()
        {
            InitializeComponent();
        }
       
        private void Form2_Load(object sender, EventArgs e)
        {

        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
           
        }

        // Кнопка Регистрации
        private void kryptonButton1_Click(object sender, EventArgs e)
        {
            var regUser = kryptonTextBox1.Text;
            var regParol = kryptonTextBox2.Text;
            var Name = kryptonTextBox3.Text;
            var Surname = kryptonTextBox4.Text;
            var Age = kryptonTextBox7.Text;
            var symbol1 = regParol[0];
            var Email = kryptonTextBox11.Text;
            var Pol = Convert.ToString(kryptonComboBox1.SelectedItem);
            var Weight = kryptonTextBox6.Text;
            var height = kryptonTextBox8.Text;
            var Date = Convert.ToString(kryptonDateTimePicker1.Value);
            var Number = kryptonTextBox10.Text;
            var hasNumber = new Regex(@"[0-9]+");

            if (regParol == "" || regParol == "" || Name == "" || Surname == "" || Age == "" || Email == "" || Pol == "" || Weight == "" || height == "" || Date == "" || Number == "")
            {
                MessageBox.Show("Пустые поля, Ошибка");
            }
            else if (regParol.Length < 5)
            {
                MessageBox.Show("Минимум 6 символов");
            } else if (symbol1 != Char.ToUpper(symbol1))
            {
                MessageBox.Show("1 буква заглавная");
            }else if (hasNumber.IsMatch(regParol) == false)
            {
                MessageBox.Show("Необходима хотя бы одна цифра");
            } else
            {
                SQLiteConnection conn = new SQLiteConnection(Connect());
                conn.Open();
                string query1 = $"INSERT INTO Users (Login, Parol, name ,surname, age, gender, weight, height, date, email, number) VALUES ('{regUser}', '{regParol}', '{Name}', '{Surname}', '{Age}', '{Pol}', '{Weight}', '{height}', '{Date}', '{Email}', '{Number}');";
                SQLiteCommand command1 = new SQLiteCommand(query1, conn);
                command1.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Success");
                Application.OpenForms[0].Show();
                this.Close();

            }


        }
       

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
          
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {

        }

        private void pictureBox6_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

        private void pictureBox5_Click(object sender, EventArgs e)
        {
            this.Close();
            Application.OpenForms[0].Show();
        }

        private void kryptonLabel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel4_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel4_Click(object sender, EventArgs e)
        {
            Application.OpenForms[0].Show();
            this.Close();
        }

        private void kryptonTextBox9_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
