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
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Xml.Linq;
using System.ComponentModel.Design;
using System.Collections;
using System.Data.SqlTypes;
using System.Security.Cryptography.X509Certificates;
using LiveCharts;
using LiveCharts.Wpf;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using System.Globalization;
using System.Windows.Controls.Primitives;
using System.IO;
using System.Windows.Media.Media3D;

namespace Kurs
{
    public partial class Form3 : KryptonForm
    {
        SQLiteConnection conn = null;
        SQLiteDataAdapter adapter = null;
        DataTable table = null;
        string value;
        static string Connect()
        {

            string databaseName = "fitness.db";
            string executablePath = AppDomain.CurrentDomain.BaseDirectory;
            string databasePath = System.IO.Path.Combine(executablePath, databaseName);

            string connectionString = $"Data Source={databasePath};Version=3;";

            return connectionString;
        }
        public Form3()
        {
            InitializeComponent();

        }


        private void updatexer()
        {
            string query = $"SELECT id, kolvo AS 'Кол-во повторений', time AS 'Продолжительность', weight AS 'Вес', nameexer AS 'Название упражнения', description AS 'Описание', Myshca AS 'Мышца', idtraining FROM exer WHERE iduser = '{BaseData.id}'";

            SQLiteConnection connection = new SQLiteConnection(Connect());
            connection.Open();
            SQLiteDataAdapter da = new SQLiteDataAdapter(query,connection);
            DataTable table = new DataTable();
            da.Fill(table);
            kryptonDataGridView1.DataSource = table;
            kryptonDataGridView1.Columns[7].Visible = false;
            connection.Close();
        }




        private void update3()
        {
            SQLiteConnection con = new SQLiteConnection(Connect());
            con.Open();

            SQLiteDataAdapter da = new SQLiteDataAdapter($"SELECT workout_id, date AS 'Дата тренировки', prodol AS 'Время тренировки', nametraining AS 'Название тренировки', idplan FROM Training WHERE user_id = '{BaseData.id}' ", con);
            DataTable table = new DataTable();
            da.Fill(table);
            kryptonDataGridView3.DataSource = table;
            kryptonDataGridView3.Columns[0].Visible = false;
            kryptonDataGridView3.Columns[4].Visible = false;
            con.Close();
        }

        private void Updategrid2()
        {
            SQLiteConnection con = new SQLiteConnection(Connect());
            con.Open();

            SQLiteDataAdapter da = new SQLiteDataAdapter($"SELECT id, plan_name AS 'Название плана', DatePlan AS 'Дата Создания плана' FROM Plans WHERE iduser = '{BaseData.id}' ", con);
            DataTable table = new DataTable();
            da.Fill(table);
            kryptonDataGridView2.DataSource = table;
            kryptonDataGridView2.Columns[0].Visible = false;
            con.Close();
        }

        private void updateLichka()
        {
            kryptonLabel37.Text = "Логин: " + DATAUSER("Login");
            kryptonLabel38.Text = "Адрес Электронной почты: " + DATAUSER("email");
            kryptonLabel25.Text = "Имя: " + DATAUSER("name");
            kryptonLabel26.Text = "Фамилия: " + DATAUSER("surname");
            kryptonLabel2.Text = "Номер телефона: " + DATAUSER("number");
            kryptonLabel24.Text = "Дата рождения: " + DATAUSER("date");
            kryptonLabel40.Text = "Вес: " + DATAUSER("weight");
            kryptonLabel41.Text = "Рост: " + DATAUSER("height");
            kryptonLabel43.Text = "Пол: " + DATAUSER("gender");
            kryptonLabel39.Text = "Возраст: " + DATAUSER("age");
        }
        private void Form3_Load(object sender, EventArgs e)
        {
            update3();
            updateLichka();
            Updategrid2();
            updatexer();
            zamery();
        }

        private void kryptonButton2_Click(object sender, EventArgs e)
        {
            kryptonPanel8.Visible = false;
            kryptonPanel9.Visible = false;
            kryptonPanel7.Visible = false;
            kryptonPanel6.Visible = false;
            kryptonPanel5.Visible = true;
            kryptonPanel1.Visible = true;
        }

        private void kryptonButton1_Click_2(object sender, EventArgs e)
        {
            kryptonPanel5.Visible = false;
            kryptonPanel6.Visible = false;
            kryptonPanel7.Visible = false;
            kryptonPanel9.Visible = false;
            kryptonPanel8.Visible = false;
            kryptonPanel1.Visible = true;

        }

        private void kryptonButton6_Click(object sender, EventArgs e)
        {
            BaseData.ClearData();
            Application.OpenForms[0].Show();
            this.Close();

        }

        private void kryptonLabel8_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton9_Click(object sender, EventArgs e)
        {
            Update();
        }

        private void kryptonButton3_Click(object sender, EventArgs e)
        {
            kryptonPanel9.Visible = false;
            kryptonPanel8.Visible = false;
            kryptonPanel7.Visible = true;
            kryptonPanel6.Visible = true;
            kryptonPanel5.Visible = true;
            kryptonPanel1.Visible = true;
            
        }

        private void kryptonPanel7_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton4_Click(object sender, EventArgs e)
        {
            kryptonPanel8.Visible = false;
            kryptonPanel9.Visible = true;
            kryptonPanel7.Visible = true;
            kryptonPanel6.Visible = true;
            kryptonPanel5.Visible = true;
            kryptonPanel1.Visible = true;
            

        }

        private void kryptonButton5_Click(object sender, EventArgs e)
        {
            try
            {
                // Запускаем нужный файл
                System.Diagnostics.Process.Start("\\pups\\Kurs\\Справка.pdf");
            }
            catch (Exception exp)
            {
                MessageBox.Show(exp.Message);
            }
        }

        private void kryptonPanel12_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonLabel40_Paint(object sender, PaintEventArgs e)
        {

        }

        private string DATAUSER(string Commanda)
        {
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();
            string query = $"SELECT {Commanda} FROM Users WHERE user_id = '{BaseData.id}'";
            SQLiteCommand comm = new SQLiteCommand(query, conn);
            object res = comm.ExecuteScalar();
            conn.Close();
            return Convert.ToString(res);
        }

        private void kryptonLabel37_Paint(object sender, PaintEventArgs e)
        {

        }
        private void kryptonPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonPanel14_Paint(object sender, PaintEventArgs e)
        {
            updateLichka();
        }

        private void kryptonButton7_Click(object sender, EventArgs e)
        {
            BaseData.secretid = BaseData.id;
            INFOPOL pol = new INFOPOL();
            pol.Show();
            this.Close();
        }

        private void kryptonButton8_Click_1(object sender, EventArgs e)
        {
            PAROLSMENA pol = new PAROLSMENA();
            pol.Show();
            this.Close();
        }

        private void kryptonPanel5_Paint(object sender, PaintEventArgs e)
        {

        }
        //Статистика
        private void diagramm()
        {
            SQLiteConnection conn = new SQLiteConnection(Connect());
            string query = $"SELECT date, prodol FROM Training";
            conn.Open();

            SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn);
            DataSet dataset = new DataSet();
            adapter.Fill(dataset, "Training");
            table = dataset.Tables["Training"];
            cartesianChart1.LegendLocation = LegendLocation.Bottom;
        }

        private void move1()
        {
            diagramm();
            DateTime databegin = new DateTime();
            DateTime dataEnd = new DateTime();
            databegin = kryptonDateTimePicker2.Value;
            dataEnd = kryptonDateTimePicker3.Value;


            LiveCharts.SeriesCollection series = new LiveCharts.SeriesCollection();
            ChartValues<int> KolvoT = new ChartValues<int>();
            ChartValues<DateTime> dateT = new ChartValues<DateTime>();
            List<string> list = new List<string>();
            List<string> listdate = new List<string>();
            foreach (DataRow row in table.Rows)
            {
                DateTime currentDate = Convert.ToDateTime(row["date"]);
                if (currentDate >= databegin && currentDate <= dataEnd)
                {
                    KolvoT.Add(Convert.ToInt32(row["prodol"]));
                    dateT.Add(Convert.ToDateTime(row["date"]));
                }
            }

            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();
            cartesianChart1.AxisX.Add(new Axis()
            {
                Title = "Дaты",
                Labels = listdate
            });

            cartesianChart1.AxisY.Add(new Axis()
            {
                Title = "Продолжительность",
                Labels = list
            });

            LineSeries line = new LineSeries();
            line.Title = "Продолжительность";
            line.Values = KolvoT;

            LineSeries lined = new LineSeries();
            lined.Title = "Даты";
            lined.Values = dateT;

            series.Add(line);

            cartesianChart1.Series = series;
        }
        //Метод ответа2
        private void move2()
        {
            diagramm();
            DateTime databegin = new DateTime();
            DateTime dataEnd = new DateTime();
            databegin = kryptonDateTimePicker2.Value;
            dataEnd = kryptonDateTimePicker3.Value;

            LiveCharts.SeriesCollection series = new LiveCharts.SeriesCollection();
            ChartValues<int> KolvoT = new ChartValues<int>();
            ChartValues<DateTime> dateT = new ChartValues<DateTime>();

            Dictionary<DateTime, int> countByDate = new Dictionary<DateTime, int>();
            foreach (DataRow row in table.Rows)
            {
                DateTime currentDate = Convert.ToDateTime(row["date"]);
                if (currentDate >= databegin && currentDate <= dataEnd)
                {
                    if (countByDate.ContainsKey(currentDate))
                    {
                        countByDate[currentDate]++;
                    }
                    else
                    {
                        countByDate[currentDate] = 1;
                    }
                }
            }

            // Добавить значения count и currentDate в списки KolvoT и dateT
            foreach (var pair in countByDate)
            {
                KolvoT.Add(pair.Value);
                dateT.Add(pair.Key);
                MessageBox.Show($"{pair}");
            }


            cartesianChart1.AxisX.Clear();
            cartesianChart1.AxisY.Clear();
            cartesianChart1.AxisX.Add(new Axis()
            {
                Title = "Дaты",

            });

            cartesianChart1.AxisY.Add(new Axis()
            {
                Title = "Количество",

            });

            LineSeries line = new LineSeries();
            line.Title = $"Количество тренировок";
            line.Values = KolvoT;

            LineSeries lined = new LineSeries();
            lined.Title = "Даты";
            lined.Values = dateT;

            series.Add(line);

            cartesianChart1.Series = series;
        }


        private void kryptonPanel9_Paint(object sender, PaintEventArgs e)
        {

        }
        // Фильтр с периодом
        private void kryptonButton12_Click(object sender, EventArgs e)
        {
            string otvet1 = "Продолжительность тренировок";
            string otvet2 = "Количество";
            if (otvet1.Equals(kryptonComboBox1.SelectedItem.ToString(), StringComparison.OrdinalIgnoreCase))
            {
                move1();
              
            }
            else
            {
                move2();
            }
        }

        //Поиск Определенных данных
        private void kryptonButton13_Click(object sender, EventArgs e)
        {
            /*
             1) Средняя продолжительность тренировки
             2) Общее количество тренировок
             3) Суммарное время тренировок
             */
            DateTime S = kryptonDateTimePicker4.Value;
            DateTime Po = kryptonDateTimePicker5.Value;
            string ot1 = kryptonComboBox2.SelectedItem.ToString();
            if (ot1 == "Средняя продолжительность тренировки")
            {
                string query = $"SELECT date, prodol FROM Training";
                SQLiteConnection conn = new SQLiteConnection(Connect());
                conn.Open();

                SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset, "Training");
                DataTable table = dataset.Tables["Training"];
                List<int> list = new List<int>();

                foreach (DataRow row in table.Rows)
                {
                    DateTime date = Convert.ToDateTime(row["date"]); // Получаем дату из строки
                    int prodol = Convert.ToInt32(row["prodol"]); // Получаем значение prodol из строки

                    // Проверяем, находится ли дата в промежутке между S и Po
                    if (date >= S && date <= Po)
                    {
                        list.Add(prodol); // Добавляем значение prodol в список
                    }
                }

                conn.Close(); // Закрываем соединение с базой данных

                double res = list.Average();


            } else if (ot1 == "Общее количество тренировок")
            {
                string query = $"SELECT * FROM Training";
                SQLiteConnection conn = new SQLiteConnection(Connect());
                conn.Open();

                SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset, "Training");
                DataTable table = dataset.Tables["Training"];
                List<int> list = new List<int>();

                foreach (DataRow row in table.Rows)
                {
                    DateTime date = Convert.ToDateTime(row["date"]); // Получаем дату из строки
                    int prodol = Convert.ToInt32(row["prodol"]); // Получаем значение prodol из строки

                    // Проверяем, находится ли дата в промежутке между S и Po
                    if (date >= S && date <= Po)
                    {
                        list.Add(prodol); // Добавляем значение prodol в список
                    }
                }
                MessageBox.Show($"Общее кол-во тренировок = {list.Count}");
            } else
            {
                string query = $"SELECT * FROM Training";
                SQLiteConnection conn = new SQLiteConnection(Connect());
                conn.Open();

                SQLiteDataAdapter adapter = new SQLiteDataAdapter(query, conn);
                DataSet dataset = new DataSet();
                adapter.Fill(dataset, "Training");
                DataTable table = dataset.Tables["Training"];
                List<int> list = new List<int>();

                foreach (DataRow row in table.Rows)
                {
                    DateTime date = Convert.ToDateTime(row["date"]); // Получаем дату из строки
                    int prodol = Convert.ToInt32(row["prodol"]); // Получаем значение prodol из строки

                    // Проверяем, находится ли дата в промежутке между S и Po
                    if (date >= S && date <= Po)
                    {
                        list.Add(prodol); // Добавляем значение prodol в список
                    }
                }
                MessageBox.Show($"Общее время тренировок = {list.Sum()}");
            }
        }

        private void kryptonButton16_Click(object sender, EventArgs e)
        {
        }
        private void flowLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void kryptonButton9_Click_1(object sender, EventArgs e)
        {
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();

            // Проверяем существование плана
            string planQuery = $"SELECT workout_id FROM Training WHERE nametraining = '{kryptonTextBox9.Text}'";
            SQLiteCommand planCommand = new SQLiteCommand(planQuery, conn);
            long planCount = Convert.ToInt64(planCommand.ExecuteScalar());

            if (planCount > 0)
            {
                // План существует, выполняем запрос на создание тренировки
                string insertQuery = $"INSERT INTO exer (kolvo, time, weight, nameexer, iduser, description, Myshca, idtraining) VALUES ('{kryptonTextBox8.Text}','{kryptonTextBox10.Text}','{kryptonTextBox7.Text}','{kryptonTextBox1.Text}','{BaseData.id}', '{kryptonTextBox2.Text}', '{kryptonTextBox3.Text}', '{planCount}')";
                SQLiteCommand insertCommand = new SQLiteCommand(insertQuery, conn);
                insertCommand.ExecuteNonQuery();

                MessageBox.Show("Упражнение успешно создано");
                updatexer();
            }
            else
            {
                MessageBox.Show("Тренировка с указанным идентификатором не найден.");
            }

            conn.Close();
        }

        private void kryptonButton11_Click_2(object sender, EventArgs e)
        {
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();

            if (kryptonDataGridView1.SelectedRows.Count > 0)
            {
                int selectedIndex = kryptonDataGridView1.SelectedRows[0].Index;
                int rowId = int.Parse(kryptonDataGridView1[0, selectedIndex].Value.ToString());

                string deleteQuery = "DELETE FROM exer WHERE id = @id";
                SQLiteCommand command = new SQLiteCommand(deleteQuery, conn);
                command.Parameters.AddWithValue("@id", rowId);
                command.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("упражнение успешно удалено");
                updatexer();
            }
        }

        private void kryptonTextBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void kryptonButton15_Click(object sender, EventArgs e)
        {
            AddPlan add = new AddPlan();
            add.Show();
            this.Close();
        }

        private void kryptonDataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }


        private void kryptonDataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {   
            
            AddExercise addexer = new AddExercise();
            addexer.Show();
            this.Close();
        }

        private void kryptonDataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            BaseData.Idtraining = kryptonDataGridView1.CurrentRow.Cells[0].Value.ToString();
        }

        private void zamery()
        {
            string query = $"SELECT * FROM Condition";
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();
            SQLiteDataAdapter da = new SQLiteDataAdapter(query, conn);
            DataTable table = new DataTable();
            da.Fill(table);
            kryptonDataGridView4.DataSource = table;
            conn.Close();
            
        }

        private void kryptonButton21_Click(object sender, EventArgs e)
        {
            string date, biceps, yagodici, plechi, predpleche, height, weight, bedra, grud, taliya;
            date = kryptonDateTimePicker1.Value.ToString();
            biceps = kryptonTextBox13.Text;
            plechi = kryptonTextBox16.Text;
            yagodici = kryptonTextBox20.Text;
            predpleche = kryptonTextBox18.Text;
            height = kryptonTextBox14.Text;
            weight = kryptonTextBox12.Text;
            bedra = kryptonTextBox17.Text;
            grud = kryptonTextBox15.Text;
            taliya = kryptonTextBox11.Text;

            string query = $"INSERT INTO Condition (id_user, Taliya, weight, height, Biceps, Predplech, Bedro, Plechi, Grud, Yagodici, date) VALUES ('{BaseData.id}', '{taliya}', '{weight}', '{height}', '{biceps}', '{predpleche}', '{bedra}', '{plechi}', '{grud}', '{yagodici}', '{date}')";
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();
            SQLiteCommand comm = new SQLiteCommand(query,conn);
            comm.ExecuteNonQuery();
            conn.Close();
            zamery();
        }

        private void kryptonButton18_Click(object sender, EventArgs e)
        {
            kryptonPanel8.Visible = false;
            kryptonPanel9.Visible = false;
            kryptonPanel6.Visible = true;
            kryptonPanel5.Visible = true;
            kryptonPanel1.Visible = true;
        }

        private void kryptonButton10_Click(object sender, EventArgs e)
        {
            kryptonPanel9.Visible = true;
            kryptonPanel6.Visible = true;
            kryptonPanel5.Visible = true;
            kryptonPanel7.Visible = true;
            kryptonPanel1.Visible = true;
            kryptonPanel8.Visible = true;
        }

        private void kryptonButton22_Click(object sender, EventArgs e)
        {
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();

            if (kryptonDataGridView4.SelectedRows.Count > 0)
            {
                int selectedIndex = kryptonDataGridView4.SelectedRows[0].Index;
                int rowId = int.Parse(kryptonDataGridView4[0, selectedIndex].Value.ToString());

                string deleteQuery = "DELETE FROM Condition WHERE id = @id";
                SQLiteCommand command = new SQLiteCommand(deleteQuery, conn);
                command.Parameters.AddWithValue("@id", rowId);
                command.ExecuteNonQuery();
                conn.Close();
                zamery();
            }
        }

        private void kryptonButton19_Click(object sender, EventArgs e)
        {
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();

            // Проверяем существование плана
            string planQuery = $"SELECT id FROM Plans WHERE plan_name = '{kryptonTextBox5.Text}'";
            SQLiteCommand planCommand = new SQLiteCommand(planQuery, conn);
            long planCount = Convert.ToInt64(planCommand.ExecuteScalar());

            if (planCount > 0)
            {
                // План существует, выполняем запрос на создание тренировки
                string insertQuery = $"INSERT INTO Training (date, prodol, nametraining, idplan, user_id) VALUES ('{kryptonDateTimePicker6.Value.ToString()}','{kryptonDateTimePicker7.Value.ToString()}','{kryptonTextBox4.Text}','{planCount}','{BaseData.id}')";
                SQLiteCommand insertCommand = new SQLiteCommand(insertQuery, conn);
                insertCommand.ExecuteNonQuery();

                MessageBox.Show("Тренировка успешно создана.");
            }
            else
            {
                MessageBox.Show("План с указанным идентификатором не найден.");
            }

            conn.Close();
        }

        private void kryptonButton17_Click(object sender, EventArgs e)
        {
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();

            if (kryptonDataGridView2.SelectedRows.Count > 0)
            {
                int selectedIndex = kryptonDataGridView2.SelectedRows[0].Index;
                int rowId = int.Parse(kryptonDataGridView2[0, selectedIndex].Value.ToString());

                string deleteQuery = "DELETE FROM Plans WHERE id = @id";
                SQLiteCommand command = new SQLiteCommand(deleteQuery, conn);
                command.Parameters.AddWithValue("@id", rowId);
                command.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("План успешно удален");
                Updategrid2();
            }
        }

        private void kryptonButton20_Click(object sender, EventArgs e)
        {
            SQLiteConnection conn = new SQLiteConnection(Connect());
            conn.Open();

            if (kryptonDataGridView3.SelectedRows.Count > 0)
            {
                int selectedIndex = kryptonDataGridView3.SelectedRows[0].Index;
                int rowId = int.Parse(kryptonDataGridView3[0, selectedIndex].Value.ToString());

                string deleteQuery = "DELETE FROM Training WHERE workout_id = @id";
                SQLiteCommand command = new SQLiteCommand(deleteQuery, conn);
                command.Parameters.AddWithValue("@id", rowId);
                command.ExecuteNonQuery();
                conn.Close();
                MessageBox.Show("Тренировка успешно удалена");
                update3();
            }
        }

        // Печать
        private void kryptonButton14_Click(object sender, EventArgs e)
        {
            int width = 725;
            int height = 583;
            Bitmap bitmap = new Bitmap(width, height);


            // Создание графики на основе изображения
            Graphics graphics = Graphics.FromImage(bitmap);

            // Отрисовка CartesianChart на графике
            this.cartesianChart1.DrawToBitmap(bitmap, new Rectangle(0, 0, width, height));

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "JPEG файлы (*.jpeg)|*.jpeg";
            saveFileDialog.Title = "Сохранить изображение";
            saveFileDialog.InitialDirectory = @"E:\";
            saveFileDialog.FileName = "test.jpeg";

            // Если пользователь нажал "ОК" в SaveFileDialog, сохранить изображение
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Указание пути сохранения изображения
                string filePath = saveFileDialog.FileName;

                // Сохранение изображения
                bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Jpeg);
            }
        }

        private void kryptonButton16_Click_1(object sender, EventArgs e)
        {
            if (kryptonDataGridView4.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);
                for(int i = 1; i < kryptonDataGridView4.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = kryptonDataGridView4.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < kryptonDataGridView4.Rows.Count; i++)
                {   
                    for(int j = 0; j < kryptonDataGridView4.Columns.Count; j++)
                    {
                            xcelApp.Cells[i + 2, j + 1] = kryptonDataGridView4.Rows[i].Cells[j].Value.ToString();
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
        }

        private void kryptonButton23_Click(object sender, EventArgs e)
        {
            if (kryptonDataGridView2.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);
                for (int i = 1; i < kryptonDataGridView2.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = kryptonDataGridView2.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < kryptonDataGridView2.Rows.Count; i++)
                {
                    for (int j = 0; j < kryptonDataGridView2.Columns.Count; j++)
                    {
xcelApp.Cells[i + 2, j + 1] = kryptonDataGridView2.Rows[i].Cells[j].Value.ToString();

                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
        }

        private void kryptonButton24_Click(object sender, EventArgs e)
        {
            if (kryptonDataGridView3.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);
                for (int i = 1; i < kryptonDataGridView3.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = kryptonDataGridView3.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < kryptonDataGridView3.Rows.Count; i++)
                {
                    for (int j = 0; j < kryptonDataGridView3.Columns.Count; j++)
                    {
                    
                            xcelApp.Cells[i + 2, j + 1] = kryptonDataGridView3.Rows[i].Cells[j].Value.ToString();
              
                  
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
        }

        private void kryptonButton25_Click(object sender, EventArgs e)
        {
            if (kryptonDataGridView1.Rows.Count > 0)
            {
                Microsoft.Office.Interop.Excel.Application xcelApp = new Microsoft.Office.Interop.Excel.Application();
                xcelApp.Application.Workbooks.Add(Type.Missing);
                for (int i = 1; i < kryptonDataGridView1.Columns.Count + 1; i++)
                {
                    xcelApp.Cells[1, i] = kryptonDataGridView1.Columns[i - 1].HeaderText;
                }
                for (int i = 0; i < kryptonDataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < kryptonDataGridView1.Columns.Count; j++)
                    {
                       
                            xcelApp.Cells[i + 2, j + 1] = kryptonDataGridView1.Rows[i].Cells[j].Value.ToString();
                     
                    }
                    xcelApp.Columns.AutoFit();
                    xcelApp.Visible = true;
                }
            }
        }
    } 
}
