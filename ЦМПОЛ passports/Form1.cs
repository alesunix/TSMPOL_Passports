using System;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Excell = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data.SqlClient;
using System.Drawing.Printing;
using Word = Microsoft.Office.Interop.Word;
using System.Deployment.Application;
using System.Reflection;
using System.Threading;

namespace ЦМПОЛ_passports
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=192.168.99.4;Initial Catalog=CmpolBase;Persist Security Info=True;User ID=Lan;Password=Samsung0");
        Form2 form2 = new Form2();

        public Form1()
        {
            InitializeComponent();
            comboBox5.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            comboBox1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox5.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            textBox7.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button6_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            comboBox6.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button6_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            comboBox2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button6_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            dateTimePicker1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button6_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
            dateTimePicker2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button6_Click(new object(), new EventArgs()); };//Нажатие кнопки "OK" с клавиатуры
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.cena = textBox11.Text; // Записываем содержимое
            Properties.Settings.Default.Save(); // Сохраняем переменные.
            MessageBox.Show("Цена сохранена", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }
        //-----------------------------------------------------------------------//
        private void Form1_Load(object sender, EventArgs e)//Загрузка формы
        {
            //form2.ShowDialog();
            form2.AccessF2.Text = form2.Data;
            form2.AccessF2.Text = Clipboard.GetText();//Считать текст из буфера обмена 
            if (form2.AccessF2.Text == "medium")
            {
                comboBox4.Enabled = false;
                textBox4.Enabled = false;
                comboBox3.Enabled = false;
                button5.Enabled = false;
            }
            if (form2.AccessF2.Text == "low")
            {
                tabPage2.Enabled = false;
            }
            dateTimePicker1.Value = DateTime.Today;
            dateTimePicker2.Value = DateTime.Today;
            dateTimePicker3.Value = DateTime.Today;

            //-----------------Окраска Гридов-------------------//
            DataGridViewRow row1 = this.dataGridView1.RowTemplate;
            row1.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row1.Height = 5;
            row1.MinimumHeight = 17;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            //dataGridView1.Columns[0].Width = 5;//Ширина столбца
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;//цвет заголовка
            dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            DataGridViewRow row2 = this.dataGridView2.RowTemplate;
            row2.DefaultCellStyle.BackColor = Color.AliceBlue;//цвет строк
            row2.Height = 5;
            row2.MinimumHeight = 17;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            dataGridView2.EnableHeadersVisualStyles = false;
            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.LightSlateGray;//цвет заголовка
            dataGridView2.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            DataGridViewRow row3 = this.dataGridView3.RowTemplate;
            row3.DefaultCellStyle.BackColor = Color.LightSkyBlue;
            row3.Height = 5;
            row3.MinimumHeight = 17;
            dataGridView3.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            dataGridView3.EnableHeadersVisualStyles = false;
            dataGridView3.ColumnHeadersDefaultCellStyle.BackColor = Color.LightSlateGray;//цвет заголовка
            dataGridView3.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            //----------------Окраска Гридов--------------------//             
            users_select();
            punkt_select();
            disp_data();
            Marshruts();
            Podschet();
            itog();
            disp_data();
            dataGridView1.Visible = true;
            dataGridView2.Visible = false;
            dataGridView3.Visible = false;
        }
        public void itog()
        {
            //con.Open();//открыть соединение
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            //{
            //    SqlCommand c = new SqlCommand("UPDATE [Table_pass] SET itogo = (id_id + ip_ip), reis = @reis WHERE id = @id", con);
            //    c.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));//первая строка в гриде
            //    c.Parameters.AddWithValue("@reis", 0);
            //    c.ExecuteNonQuery();
            //}
            //con.Close();//закрыть соединение
            //con.Open();//открыть соединение
            //for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            //{
            //    SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET itogo = (id_id + ip_ip), reis = @reis WHERE id = @id", con);
            //    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));//первая строка в гриде
            //    if (Convert.ToInt32(dataGridView1.Rows[i].Cells[3].Value) != 0 || Convert.ToInt32(dataGridView1.Rows[i].Cells[4].Value) != 0)
            //    {
            //        cmd.Parameters.AddWithValue("@reis", 0);
            //    }
            //    else cmd.Parameters.AddWithValue("@reis", 0);
            //    cmd.ExecuteNonQuery();
            //}
            //con.Close();//закрыть соединение
        }
        public void disp_data()//Select
        {
            //con.Open();//Открываем соединение
            //SqlCommand cmd = con.CreateCommand();
            //cmd.CommandType = CommandType.Text;
            ////cmd.CommandText = "SELECT TOP 1000 * FROM [Table_pass] ORDER BY date DESC";//последние 1000 записей
            //cmd.CommandText = "SELECT * FROM [Table_pass] WHERE (date BETWEEN @StartDate AND @EndDate) ORDER BY date DESC";
            //cmd.Parameters.AddWithValue("@StartDate", DateTime.Today.AddMonths(-6));
            //cmd.Parameters.AddWithValue("@EndDate", DateTime.Today);
            //cmd.ExecuteNonQuery();
            //DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            //SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            //dt.Clear();//чистим DataTable, если он был не пуст
            //da.Fill(dt);//заполняем данными созданный DataTable
            //dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            //con.Close();//Закрываем соединение 

            //con.Open();//открыть соединение
            //SqlCommand cmd = con.CreateCommand();
            //cmd.CommandType = CommandType.Text;
            //cmd.CommandText = "SELECT max(id) AS Id, date, MAX(reis) AS Рейсы, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого, MIN(primecanie) AS Примечание," +
            //    " MIN(type) AS Тип, MIN(punkt) AS Пункт, MIN(processing) AS Processing, MIN(date_processing) AS Дата_обработки, MIN(akt) AS Акт FROM [Table_pass]" +
            //    " WHERE date BETWEEN @StartDate AND @EndDate GROUP BY date ORDER BY date";
            //cmd.Parameters.AddWithValue("@StartDate", DateTime.Today.AddMonths(-12));
            //cmd.Parameters.AddWithValue("@EndDate", DateTime.Today);
            //cmd.ExecuteNonQuery();
            //DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            //SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            //dt.Clear();//чистим DataTable, если он был не пуст
            //da.Fill(dt);//заполняем данными созданный DataTable
            //dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            //con.Close();//закрыть соединение    

            con.Open();//открыть соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT id AS Id, date AS Дата, reis AS Рейсы, id_id AS eID, ip_ip AS ОГП, itogo AS Итого, primecanie AS Примечание," +
                " type AS Тип, punkt AS ОПРН, akt AS Акт, marshrut AS Маршруты FROM [Table_pass]" +
                " WHERE date BETWEEN @StartDate AND @EndDate ORDER BY date";//WHERE reis NOT IN ('0') чтобы не отображать нулевые
            cmd.Parameters.AddWithValue("@StartDate", DateTime.Today.AddMonths(-1));
            cmd.Parameters.AddWithValue("@EndDate", DateTime.Today);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение            
        }
        public void punkt_select()//Вывод Пункт в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT punkt FROM [Table_punkts]";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            //comboBox1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            //comboBox2.DataSource = dt;
            foreach (DataRow row in dt.Rows)
            {
                comboBox1.Items.Add(row[0].ToString());
                comboBox2.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение
        }
        public void users_select()//Вывод пользователей в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT users FROM [Table_users]";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            //comboBox4.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            foreach (DataRow row in dt.Rows)
            {
                comboBox4.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение

            con.Open();//Открываем соединение
            SqlCommand cmd1 = con.CreateCommand();
            cmd1.CommandType = CommandType.Text;
            cmd1.CommandText = "SELECT access FROM [Table_users]";
            cmd1.ExecuteNonQuery();
            DataTable dt1 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da1 = new SqlDataAdapter(cmd1);//создаем экземпляр класса SqlDataAdapter
            dt1.Clear();//чистим DataTable, если он был не пуст
            da1.Fill(dt1);//заполняем данными созданный DataTable
            //comboBox3.DataSource = dt1;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            foreach (DataRow row in dt1.Rows)
            {
                comboBox3.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение
        }
        public void Select_tupe()//(Для выдачи реестров)Выборка по статусу и сортировка по номеру реестра от больших значений к меньшим.
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT * FROM [Table_pass] WHERE type = @type ORDER BY akt DESC";
            if (Convert.ToString(dataGridView2.Rows[0].Cells[10].Value) == "Обычный")
            {
                cmd.Parameters.AddWithValue("@type", "Обычный");
            }
            else if (Convert.ToString(dataGridView2.Rows[0].Cells[10].Value) == "Срочный")
            {
                cmd.Parameters.AddWithValue("@type", "Срочный");
            }
            else if (Convert.ToString(dataGridView2.Rows[0].Cells[10].Value) == "На дом")
            {
                cmd.Parameters.AddWithValue("@type", "На дом");
            }
            else if (Convert.ToString(dataGridView2.Rows[0].Cells[10].Value) == "МИД")
            {
                cmd.Parameters.AddWithValue("@type", "МИД");
            }
            else MessageBox.Show("Ошибка строка 110", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение
        }
        public void Podschet()//Произвести подсчет dataGridView1 и dataGridView2
        {
            if (dataGridView1.Visible == true)
            {
                //Итого 
                double summa = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom);
                    summa += incom;
                }
                textBox6.Visible = true;
                textBox6.Text = summa.ToString() + " штук";
                //ID 
                double ID = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                    ID += incom;
                }
                textBox8.Visible = true;
                textBox8.Text = ID.ToString();
                //ОГП 
                double ogp = 0;
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    double incom;
                    double.TryParse((row.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom);
                    ogp += incom;
                }
                textBox9.Visible = true;
                textBox9.Text = ogp.ToString();
                    //Рейсы 
                    double reis = 0;
                    foreach (DataGridViewRow row in dataGridView1.Rows)
                    {
                        double incom;
                        double.TryParse((row.Cells[2].Value ?? "0").ToString().Replace(".", ","), out incom);
                        reis += incom;
                    }
                    textBox3.Visible = true;
                    textBox3.Text = reis.ToString();
            }
            //----------------------------------------------------------------------------------------------//
                //Подсчет количества строк (не учитывая пустые строки и колонки)
                int count = 0;
                for (int j = 0; j < dataGridView2.RowCount; j++)
                {
                    for (int i = 0; i < dataGridView2.ColumnCount; i++)
                    {
                        if (dataGridView2[i, j].Value != null)
                        {
                            textBox12.Text = Convert.ToString(dataGridView2.Rows.Count - 1) + " ";// -1 это нижняя пустая строка
                            count++;
                            break;
                        }
                    }
                }
        }
        public void Vyborka()
        {
            if (comboBox2.Text != "" & comboBox6.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT date AS Дата, MAX(marshrut) AS Маршруты, MAX(reis) AS Рейсы, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого, MIN(primecanie) AS Примечание," +
                    " MIN(punkt) AS Пункт, MIN(akt) AS Акт, MIN(processing) AS Processing, MIN(type) AS Тип, MIN(id) AS id FROM [Table_pass]" +
                    " WHERE punkt = @punkt AND type = @type AND date BETWEEN @StartDate AND @EndDate GROUP BY date ORDER BY date", con);
                cmd.Parameters.AddWithValue("@punkt", comboBox2.Text);
                cmd.Parameters.AddWithValue("@type", comboBox6.Text);
                cmd.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
                cmd.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                //con.Close();//закрыть соединение
                //con.Open();//открыть соединение
                for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
                {
                    SqlCommand cmd1 = new SqlCommand("UPDATE [Table_pass] SET itogo = (id_id + ip_ip), reis = @reis WHERE id = @id", con);
                    cmd1.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[i].Cells[11].Value));//первая строка в гриде
                    if (Convert.ToInt32(dataGridView2.Rows[i].Cells[3].Value) != 0 || Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value) != 0)
                    {
                        cmd1.Parameters.AddWithValue("@reis", 1);
                    }
                    else cmd1.Parameters.AddWithValue("@reis", 0);
                    cmd1.ExecuteNonQuery();
                }
                SqlCommand cmd2 = new SqlCommand("SELECT date AS Дата, MAX(marshrut) AS Маршруты, MAX(reis) AS Рейсы, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого, MIN(primecanie) AS Примечание," +
                    " MIN(punkt) AS Пункт, MIN(akt) AS Акт, MIN(processing) AS Processing, MIN(type) AS Тип, MIN(id) AS id FROM [Table_pass]" +
                    " WHERE punkt = @punkt AND type = @type AND date BETWEEN @StartDate AND @EndDate GROUP BY date ORDER BY date", con);
                cmd2.Parameters.AddWithValue("@punkt", comboBox2.Text);
                cmd2.Parameters.AddWithValue("@type", comboBox6.Text);
                cmd2.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
                cmd2.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
                cmd2.ExecuteNonQuery();
                DataTable dt2 = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da2 = new SqlDataAdapter(cmd2);//создаем экземпляр класса SqlDataAdapter
                dt2.Clear();//чистим DataTable, если он был не пуст
                da2.Fill(dt2);//заполняем данными созданный DataTable
                dataGridView2.DataSource = dt2;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение            
            }
            else if (textBox7.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("SELECT date AS Дата, MAX(marshrut) AS Маршруты, MAX(reis) AS Рейсы, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого, MIN(primecanie) AS Примечание," +
                    " MIN(punkt) AS Пункт, MIN(akt) AS Акт, MIN(processing) AS Processing, MIN(type) AS Тип, MIN(id) AS id FROM [Table_pass]" +
                "WHERE akt = @akt GROUP BY date ORDER BY date", con);
                cmd.Parameters.AddWithValue("@akt", Convert.ToString(textBox7.Text));
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dataGridView2.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//закрыть соединение 
                if (textBox7.Text == "")//если поле очищено, отобразить базу
                {
                    disp_data();
                }
            }          
            else MessageBox.Show("Необходима Выборка данных", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);           
            disp_data();
            //Рейсы
            double reis = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[2].Value ?? "0").ToString().Replace(".", ","), out incom);
                reis += incom;
            }
            textBox3.Text = reis.ToString();
            //ID 
            double ID = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                ID += incom;
            }
            textBox8.Text = ID.ToString();
            //ОГП 
            double ogp = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom);
                ogp += incom;
            }
            textBox9.Text = ogp.ToString();
            //Итого 
            double summa = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom);
                summa += incom;
            }
            textBox6.Text = summa.ToString();
        }
        public void Vyborka_itog_Srochnyi()
        {
            //------------------------------------Обнуляем рейсы -------------------------------------//
            con.Open();//открыть соединение
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                SqlCommand c = new SqlCommand("UPDATE [Table_pass] SET reis = @reis WHERE id = @id AND date BETWEEN @StartDate AND @EndDate", con);
                c.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
                c.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
                c.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));//первая строка в гриде
                c.Parameters.AddWithValue("@reis", 0);
                c.ExecuteNonQuery();
            }
            con.Close();//закрыть соединение
            //-------------------------------------Выборка первого грида с групировкой по дате и маршруту------------------------//
            con.Open();//открыть соединение
            SqlCommand cm = new SqlCommand("SELECT MIN(id) AS id, date AS Дата, MIN(reis) AS Рейсы, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого, max(primecanie) AS Примечание," +
                " MIN(type) AS Тип, max(punkt) AS ОПРН, MIN(akt) AS Акт, marshrut AS Маршруты FROM [Table_pass]" +
                " WHERE type=@type AND date BETWEEN @StartDate AND @EndDate GROUP BY date, marshrut ORDER BY date", con);
            cm.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cm.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cm.Parameters.AddWithValue("@type", "Срочный");
            cm.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cm);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение           
            //--------------------------------------Ставим рейсы по выборке-------------------------------------------------------//
            con.Open();//открыть соединение
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                SqlCommand c = new SqlCommand("UPDATE [Table_pass] SET reis = @reis WHERE id = @id AND date BETWEEN @StartDate AND @EndDate", con);
                c.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
                c.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
                c.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));//первая строка в гриде
                c.Parameters.AddWithValue("@reis", 1);
                c.ExecuteNonQuery();
            }
            con.Close();//закрыть соединение
            //--------------------------------------Окончательная выборка для выгрузки в WORD-----------------------------------------------------------//
            con.Open();//открыть соединение
            SqlCommand cmd2 = new SqlCommand("SELECT max(date) AS Дата, marshrut AS 'Наименование маршрутов', SUM(reis) AS 'Количество рейсов', MAX(stoimost) AS 'Стоимость услуги за 1 рейс, сом/тый', SUM(allsumm) AS 'Общая сумма за оказанные услуги, сом/тый без учета налогов', SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого," +
                " MIN(akt) AS Акт, MIN(processing) AS Processing, MIN(type) AS Тип, MIN(id) AS id FROM [Table_pass]" +
                " WHERE type = @type AND date BETWEEN @StartDate AND @EndDate GROUP BY marshrut ORDER BY marshrut", con);
            cmd2.Parameters.AddWithValue("@type", "Срочный");
            cmd2.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cmd2.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cmd2.ExecuteNonQuery();
            DataTable dt2 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da2 = new SqlDataAdapter(cmd2);//создаем экземпляр класса SqlDataAdapter
            dt2.Clear();//чистим DataTable, если он был не пуст
            da2.Fill(dt2);//заполняем данными созданный DataTable
            dataGridView2.DataSource = dt2;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            //---------------------------------------рейсы умножаем на стоимость----------------------------------------------//
            for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
            {
                dataGridView2.Rows[i].Cells[4].Value = Convert.ToInt32(dataGridView2.Rows[i].Cells[2].Value) * Convert.ToInt32(dataGridView2.Rows[i].Cells[3].Value);
            }
            //---------------------------------------------------------------------------------------------------------------//
            //Рейсы
            double reis = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[2].Value ?? "0").ToString().Replace(".", ","), out incom);
                reis += incom;
            }
            textBox3.Text = reis.ToString();
            //Сумма услуг
            double summuslug = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                summuslug += incom;
            }
            textBox17.Text = summuslug.ToString();
            //Сумма 
            double summ = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom);
                summ += incom;
            }
            textBox10.Text = summ.ToString();
            //ID 
            double ID = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom);
                ID += incom;
            }
            textBox8.Text = ID.ToString();
            //ОГП 
            double ogp = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[6].Value ?? "0").ToString().Replace(".", ","), out incom);
                ogp += incom;
            }
            textBox9.Text = ogp.ToString();
            //Итого 
            double summa = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[7].Value ?? "0").ToString().Replace(".", ","), out incom);
                summa += incom;
            }
            textBox6.Text = summa.ToString();                     
        }
        public void Vyborka_itog_Obychnyi()
        {          
            con.Open();//Открываем соединение
            SqlCommand cm = con.CreateCommand();
            cm.CommandType = CommandType.Text;
            cm.CommandText = "SELECT punkt FROM [Table_pass] WHERE date BETWEEN @StartDate AND @EndDate GROUP BY punkt ORDER BY punkt";
            cm.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cm.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cm.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cm);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView4.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение
            //--------------------------------------Вставка пустых значений в базу если их нет-----------------------------------------------------------//
            con.Open();//открыть соединение
            for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                        SqlCommand c = new SqlCommand("INSERT INTO [Table_pass] (punkt, id_id, ip_ip, type, date, akt, processing,itogo) VALUES (@punkt, @id_id, @ip_ip, @type, @date, @akt, @processing,@itogo)", con);
                        c.Parameters.AddWithValue("@punkt", Convert.ToString(dataGridView4.Rows[i].Cells[0].Value));
                        c.Parameters.AddWithValue("@processing", "Не обработано");
                        c.Parameters.AddWithValue("@id_id", 0);
                        c.Parameters.AddWithValue("@ip_ip", 0);
                        c.Parameters.AddWithValue("@date", dateTimePicker1.Value);
                        c.Parameters.AddWithValue("@akt", 0);
                        c.Parameters.AddWithValue("@itogo", 0);
                        c.Parameters.AddWithValue("@type", "Обычный");
                        c.ExecuteNonQuery();                    
            }
            con.Close();//закрыть соединение
            //--------------------------------------Вставка пустых значений в базу если их нет-----------------------------------------------------------//          
            con.Open();//открыть соединение
                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                        SqlCommand c = new SqlCommand("INSERT INTO [Table_pass] (punkt, id_id, ip_ip, type, date, akt, processing,itogo) VALUES (@punkt, @id_id, @ip_ip, @type, @date, @akt, @processing,@itogo)", con);
                        c.Parameters.AddWithValue("@punkt", Convert.ToString(dataGridView4.Rows[i].Cells[0].Value));
                        c.Parameters.AddWithValue("@processing", "Не обработано");
                        c.Parameters.AddWithValue("@id_id", 0);
                        c.Parameters.AddWithValue("@ip_ip", 0);
                        c.Parameters.AddWithValue("@date", dateTimePicker1.Value);
                        c.Parameters.AddWithValue("@akt", 0);
                        c.Parameters.AddWithValue("@itogo", 0);
                        c.Parameters.AddWithValue("@type", "На дом");
                        c.ExecuteNonQuery();
                }
            con.Close();//закрыть соединение 
            //--------------------------------------Вставка пустых значений в базу если их нет-----------------------------------------------------------//
            con.Open();//открыть соединение
                for (int i = 0; i < dataGridView4.Rows.Count; i++)
                {
                        SqlCommand c = new SqlCommand("INSERT INTO [Table_pass] (punkt, id_id, ip_ip, type, date, akt, processing,itogo) VALUES (@punkt, @id_id, @ip_ip, @type, @date, @akt, @processing,@itogo)", con);
                        c.Parameters.AddWithValue("@punkt", Convert.ToString(dataGridView4.Rows[i].Cells[0].Value));
                        c.Parameters.AddWithValue("@processing", "Не обработано");
                        c.Parameters.AddWithValue("@id_id", 0);
                        c.Parameters.AddWithValue("@ip_ip", 0);
                        c.Parameters.AddWithValue("@date", dateTimePicker1.Value);
                        c.Parameters.AddWithValue("@akt", 0);
                        c.Parameters.AddWithValue("@itogo", 0);
                        c.Parameters.AddWithValue("@type", "МИД");
                        c.ExecuteNonQuery();
                }
            con.Close();//закрыть соединение
            disp_data();
            Podschet();
            disp_data();
            //--------------------------------------Выборка по типу Обычный-----------------------------------------------------------//
            con.Open();//открыть соединение
            SqlCommand cmd4 = new SqlCommand("SELECT MAX(date) AS Дата, punkt AS 'Наименование территориальных ОПРН', SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого" +
                " FROM [Table_pass]" +
                " WHERE type = @type AND date BETWEEN @StartDate AND @EndDate GROUP BY punkt ORDER BY punkt", con);
            cmd4.Parameters.AddWithValue("@type", "Обычный");
            cmd4.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cmd4.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cmd4.ExecuteNonQuery();
            DataTable dt4 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da4 = new SqlDataAdapter(cmd4);//создаем экземпляр класса SqlDataAdapter
            dt4.Clear();//чистим DataTable, если он не был пуст
            da4.Fill(dt4);//заполняем данными созданный DataTable
            dataGridView2.DataSource = dt4;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            //--------------------------------------Выборка по типу На дом-----------------------------------------------------------//
            con.Open();//открыть соединение
            SqlCommand cmd5 = new SqlCommand("SELECT punkt AS ОПРН, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП" +
                " FROM [Table_pass]" +
                " WHERE type = @type AND date BETWEEN @StartDate AND @EndDate GROUP BY punkt ORDER BY punkt", con);
            cmd5.Parameters.AddWithValue("@type", "На дом");
            cmd5.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cmd5.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cmd5.ExecuteNonQuery();
            DataTable dt5 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da5 = new SqlDataAdapter(cmd5);//создаем экземпляр класса SqlDataAdapter
            dt5.Clear();//чистим DataTable, если он не был пуст
            da5.Fill(dt5);//заполняем данными созданный DataTable
            dataGridView3.DataSource = dt5;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            //--------------------------------------Выборка по типу МИД-----------------------------------------------------------//
            con.Open();//открыть соединение
            SqlCommand cmd6 = new SqlCommand("SELECT punkt AS ОПРН, SUM(ip_ip) AS ОГП" +
                " FROM [Table_pass]" +
                " WHERE type = @type AND date BETWEEN @StartDate AND @EndDate GROUP BY punkt ORDER BY punkt", con);
            cmd6.Parameters.AddWithValue("@type", "МИД");
            cmd6.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cmd6.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cmd6.ExecuteNonQuery();
            DataTable dt6 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da6 = new SqlDataAdapter(cmd6);//создаем экземпляр класса SqlDataAdapter
            dt6.Clear();//чистим DataTable, если он не был пуст
            da6.Fill(dt6);//заполняем данными созданный DataTable
            dataGridView4.DataSource = dt6;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            //------------------------------------------ создаём колонки в гриде---------------------------------------//
            DataGridViewTextBoxColumn ID = new DataGridViewTextBoxColumn();
            ID.Name = "ID";
            DataGridViewTextBoxColumn ogp = new DataGridViewTextBoxColumn();
            ogp.Name = "ОГП";
            DataGridViewTextBoxColumn ogp2 = new DataGridViewTextBoxColumn();
            ogp2.Name = "ОГП";
            DataGridViewTextBoxColumn itogo = new DataGridViewTextBoxColumn();
            itogo.Name = "Итого";
            DataGridViewTextBoxColumn vsego = new DataGridViewTextBoxColumn();
            vsego.Name = "Всего количество шт";
            DataGridViewTextBoxColumn cena = new DataGridViewTextBoxColumn();
            cena.Name = "Цена за 1 ед. паспорта, сом/тый без учета налогов";
            DataGridViewTextBoxColumn summa = new DataGridViewTextBoxColumn();
            summa.Name = "Общая сумма, сом/тый";
            DataGridViewTextBoxColumn primechanie = new DataGridViewTextBoxColumn();
            primechanie.Name = "Примечание";
            //добавляем колонки
            dataGridView2.Columns.AddRange(new DataGridViewColumn[] { ID, ogp, ogp2, itogo, vsego, cena, summa, primechanie });  
            //-----------------------------------Вставка из одного грида в другой по циклу--------------------------------//
            //Подсчет количества строк (не учитывая пустые строки и колонки)
            int count = 0;
            for (int j = 0; j < dataGridView2.RowCount; j++)
            {
                for (int i = 0; i < dataGridView2.ColumnCount; i++)
                {
                    if (dataGridView2[i, j].Value != null)
                    {
                        textBox12.Text = Convert.ToString(dataGridView2.Rows.Count-1) + " ";// -1 это нижняя пустая строка
                        count++;
                        break;
                    }
                }
            }
            for (int i = 0; i < dataGridView3.Rows.Count-1; i++)
            {
                dataGridView2.Rows[i].Cells[5].Value = Convert.ToString(dataGridView3.Rows[i].Cells[1].Value);
                dataGridView2.Rows[i].Cells[6].Value = Convert.ToString(dataGridView3.Rows[i].Cells[2].Value);
            }
            for (int i = 0; i < dataGridView4.Rows.Count-1; i++)
            {
                dataGridView2.Rows[i].Cells[7].Value = Convert.ToString(dataGridView4.Rows[i].Cells[1].Value);
            }
            //dataGridView2.AllowUserToAddRows = false;//Удаление пустой строки в DataGridView
            for (int i = 0; i < dataGridView2.Rows.Count-1; i++)
            {
                dataGridView2.Rows[i].Cells[8].Value = Convert.ToInt32(dataGridView2.Rows[i].Cells[5].Value) + Convert.ToInt32(dataGridView2.Rows[i].Cells[6].Value) + Convert.ToInt32(dataGridView2.Rows[i].Cells[7].Value);                
                dataGridView2.Rows[i].Cells[9].Value = Convert.ToInt32(dataGridView2.Rows[i].Cells[4].Value) + Convert.ToInt32(dataGridView2.Rows[i].Cells[8].Value);
                dataGridView2.Rows[i].Cells[10].Value = 8.6;
                dataGridView2.Rows[i].Cells[11].Value = Convert.ToInt32(dataGridView2.Rows[i].Cells[9].Value) * Convert.ToDouble(dataGridView2.Rows[i].Cells[10].Value);
            }
            //ID 
            double idsum = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[2].Value ?? "0").ToString().Replace(".", ","), out incom);
                idsum += incom;
            }
            textBox8.Text = idsum.ToString();
            //ОГП 
            double ogpsum = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[3].Value ?? "0").ToString().Replace(".", ","), out incom);
                ogpsum += incom;
            }
            textBox9.Text = ogpsum.ToString();
            //Итого 
            double itogosum = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[4].Value ?? "0").ToString().Replace(".", ","), out incom);
                itogosum += incom;
            }
            textBox6.Text = itogosum.ToString();
            //ID на дом 
            double id_na_dom = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[5].Value ?? "0").ToString().Replace(".", ","), out incom);
                id_na_dom += incom;
            }
            textBox14.Text = id_na_dom.ToString();
            //ОГП на дом
            double ogp_na_dom = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[6].Value ?? "0").ToString().Replace(".", ","), out incom);
                ogp_na_dom += incom;
            }
            textBox15.Text = ogp_na_dom.ToString();
            //МИД
            double mid = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[7].Value ?? "0").ToString().Replace(".", ","), out incom);
                mid += incom;
            }
            textBox16.Text = mid.ToString();
            //ИТОГО сумма
            double itogosum2 = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[8].Value ?? "0").ToString().Replace(".", ","), out incom);
                itogosum2 += incom;
            }
            textBox17.Text = itogosum2.ToString();
            //Всего шт
            double vsegosht = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[9].Value ?? "0").ToString().Replace(".", ","), out incom);
                vsegosht += incom;
            }
            textBox18.Text = vsegosht.ToString();
            //Сумма общая
            double obsumm = 0;
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                double incom;
                double.TryParse((row.Cells[11].Value ?? "0").ToString().Replace(".", ","), out incom);
                obsumm += incom;
            }
            textBox13.Text = obsumm.ToString();
        }
        private void Marshruts()
        {
            //---------------------------------Присвоение маршрута-----------------------------//
            con.Open();//открыть соединение          
            
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                string punkt = Convert.ToString(dataGridView1.Rows[i].Cells[8].Value);
                    if (punkt.Contains("ПЕРВОМАЙСКОГО") | punkt.Contains("ОКТЯБРЬСКОГО"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ЦОН-1");
                        cmd.Parameters.AddWithValue("@stoimost", 510);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("АЛАМУДУНСКОГО") | punkt.Contains("ЛЕБЕДИНОВКА") | punkt.Contains("СВЕРДЛОВСКОГО") | punkt.Contains("ВЕРХНЯЯ ЗОНА ТАШ-ТОБО") | punkt.Contains("НИЖНЯЯ ЗОНА ОКТЯБРЬСКИЙ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ЦОН-2");
                        cmd.Parameters.AddWithValue("@stoimost", 396);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ЛЕНИНСКОГО"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ЦОН-3");
                        cmd.Parameters.AddWithValue("@stoimost", 728);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ОШ") | punkt.Contains("АМИР-ТЕМИР") | punkt.Contains("ТОЛОЙКОН"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ЦОН 1 г. ОШ");
                        cmd.Parameters.AddWithValue("@stoimost", 2507);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                if (punkt.Contains("МАНАС-АТА"))
                {
                    SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@marshrut", "ЦОН 2 г. ОШ");
                    cmd.Parameters.AddWithValue("@stoimost", 2507);
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                    cmd.ExecuteNonQuery();
                }
                if (punkt.Contains("ЖАЛАЛ-АБАД"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ЦОН 2 г. Жалал-Абад");
                        cmd.Parameters.AddWithValue("@stoimost", 500);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("КЫЗЫЛ-КИ") | punkt.Contains("КАДАМЖАЙ") | punkt.Contains("БАТКЕН") | punkt.Contains("НООКАТ") | punkt.Contains("КОК-ЖАР") | punkt.Contains("БОЗ-АДЫР") |
                        punkt.Contains("ЖАНЫ-БАЗАР") | punkt.Contains("МАСАЛИЕВ") | punkt.Contains("АЙДАРКЕН") | punkt.Contains("КАЙТПАС") | punkt.Contains("МАРКАЗ") | punkt.Contains("САМАРКАНДЕК"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН г. Кызылкия, Кадамжайского, Баткенского и Ноокатского районов");
                        cmd.Parameters.AddWithValue("@stoimost", 2900);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ИССЫК-КУЛЬ") | punkt.Contains("БОКОНБАЕВО") | punkt.Contains("ЖЕТИ-ОГУЗ") | punkt.Contains("БАЛЫКЧ") | punkt.Contains("ТОН") | punkt.Contains("ТЮП") |
                        punkt.Contains("АНАНЬЕВО") | punkt.Contains("АКСУ") | punkt.Contains("КАРАКОЛ") | punkt.Contains("ЧОЛПОН-АТА") | punkt.Contains("ТАМЧИ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН по Иссык-Кульской области");
                        cmd.Parameters.AddWithValue("@stoimost", 1200);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("НАРЫН") | punkt.Contains("КОЧКОР") | punkt.Contains("ЫССЫК-АТИНСКОГО") | punkt.Contains("КАНТ") | punkt.Contains("ТОКМОК") |
                        punkt.Contains("КЕМИН") | punkt.Contains("ИВАНОВКА") | punkt.Contains("ТАШ-ДОБО") | punkt.Contains("ОКТЯБРЬСКОЕ") | punkt.Contains("ВАСИЛЬЕВСКОЕ") | punkt.Contains("ЧУЙ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Чуйской областипо восточной части, Кочкорского и Нарынского районов");
                        cmd.Parameters.AddWithValue("@stoimost", 1100);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ЖАЙЫЛСКОГО") | punkt.Contains("КАРА-БУУРИНС") | punkt.Contains("МАНАСКОГО") | punkt.Contains("СОКУЛУКСКОГО") | punkt.Contains("МОСКОВСКОГО") |
                        punkt.Contains("БАКАЙ-АТ") | punkt.Contains("ТАЛАС") | punkt.Contains("СУУСАМЫР") | punkt.Contains("ПАНФИЛОВСКОГО"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Чуйской области по западной части и все районы Таласской областей");
                        cmd.Parameters.AddWithValue("@stoimost", 1360);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("АКТАЛ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Ак-Талинского района");
                        cmd.Parameters.AddWithValue("@stoimost", 920);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ЖУМГАЛ") | punkt.Contains("МИН-КУШ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Жумгальского района");
                        cmd.Parameters.AddWithValue("@stoimost", 830);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("КАРА-КУЛЬ") | punkt.Contains("МАЙЛУУ-СУУ") | punkt.Contains("ТАШ-КУМЫР") | punkt.Contains("УЧТЕРЕК") | punkt.Contains("ТЕРЕК-СУУ") | punkt.Contains("ТОЛУК") |
                        punkt.Contains("ШАМАЛДЫСАЙ") | punkt.Contains("НООКЕН") | punkt.Contains("АРСЛАНБАБ") | punkt.Contains("СУЗАК") | punkt.Contains("ОКТЯБРЬ") | punkt.Contains("КЫЗЫЛ-ТУУ") |
                        punkt.Contains("КОК-АРТ") | punkt.Contains("БАРПЫ") | punkt.Contains("КОК-ЖАНГАК") | punkt.Contains("АТАБЕКОВ") | punkt.Contains("КАРА-ДАРЬЯ") | punkt.Contains("ТОКТОГУЛ") |
                        punkt.Contains("БАЗАР-КОРГОН") | punkt.Contains("БУРГОНДУ") | punkt.Contains("КОЧКОР-АТА") | punkt.Contains("ОЗГОРУШ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "Все районы Жалал-Абадской области, кроме Аксы, Ала-Бука");
                        cmd.Parameters.AddWithValue("@stoimost", 2350);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("АЛА-БУК") | punkt.Contains("1 МАЙ") | punkt.Contains("АК-КОРГОН"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Ала-Букинского района");
                        cmd.Parameters.AddWithValue("@stoimost", 1050);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ЧАТКАЛ") | punkt.Contains("СУМСАР"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Чаткальского района");
                        cmd.Parameters.AddWithValue("@stoimost", 1100);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("АКСЫ") | punkt.Contains("КЫЗЫЛ-ЖАР") | punkt.Contains("КАРА-ЖЫГАЧ") | punkt.Contains("ЖАНЫ-ЖОЛ") | punkt.Contains("НАЗАРАЛИЕВ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Аксыйского района");
                        cmd.Parameters.AddWithValue("@stoimost", 620);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ТОГУЗ-ТОРО") | punkt.Contains("КАЗАРМАН"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Тогуз-Тороузкого района");
                        cmd.Parameters.AddWithValue("@stoimost", 2250);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ЛЕЙЛЕК") | punkt.Contains("ИНТЕРНАЦИОНАЛ") | punkt.Contains("КОРГОН") | punkt.Contains("АРКА"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Лейлекского района");
                        cmd.Parameters.AddWithValue("@stoimost", 1460);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("СУЛЮКТ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН г. Сулукты");
                        cmd.Parameters.AddWithValue("@stoimost", 190);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("КАРА-КУЛЬЖ") | punkt.Contains("КЫЗЫЛЖАР") | punkt.Contains("УЗГЕН") | punkt.Contains("МЫРЗАКИ") | punkt.Contains("ЖЫЛАЛДЫ") | punkt.Contains("КУРШАБ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Кара-Кульджинского и Узгенского районов");
                        cmd.Parameters.AddWithValue("@stoimost", 570);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("КАРА-СУЙСК") | punkt.Contains("НАРИМАН") | punkt.Contains("МАДЫ") | punkt.Contains("ОТУЗ-АДЫР") | punkt.Contains("ПАПАН") | punkt.Contains("ТОЛЕЙКЕН"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Кара-Суйского района");
                        cmd.Parameters.AddWithValue("@stoimost", 190);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("АРАВАН") | punkt.Contains("ТОО-МОЮН"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Араванского района");
                        cmd.Parameters.AddWithValue("@stoimost", 190);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("АЛАЙ") | punkt.Contains("ГУЛЬЧА") | punkt.Contains("САРЫТАШ") | punkt.Contains("СОПУ КОРГОН"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Алайского района");
                        cmd.Parameters.AddWithValue("@stoimost", 365);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("ЧОН-АЛАЙ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Чон-Алайского района");
                        cmd.Parameters.AddWithValue("@stoimost", 1370);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
                    if (punkt.Contains("АТ-БАШИ") | punkt.Contains("ТОРУГАРТ"))
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET marshrut = @marshrut, stoimost = @stoimost WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@marshrut", "ОПРН Ат-Башинского района");
                        cmd.Parameters.AddWithValue("@stoimost", 340);
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));
                        cmd.ExecuteNonQuery();
                    }
            }          
            con.Close();//закрыть соединение              
        }

        private void button6_Click(object sender, EventArgs e)//АКТ и Обработка
        {
            label16.Visible = true;
            label16.Text = "Ожидайте идет обработка!";
            button6.Enabled = false;
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
            dataGridView3.Visible = false;
            if (checkBox1.Checked == true & comboBox6.Text == "Обычный")
            {              
                Vyborka_itog_Obychnyi();   
            }
            else if (checkBox1.Checked == true & comboBox6.Text == "Срочный")
            {
                Vyborka_itog_Srochnyi();  
            }
            else Vyborka();

            if (dataGridView2.Rows.Count > 1/*& Convert.ToString(dataGridView2.Rows[0].Cells[8].Value) != "Обработано"*/)
            {
                //select_tupe();//Выборка по типу и сортировка по номеру Акта от больших значений к меньшим.
                if (MessageBox.Show("Вы хотите обработать эти записи?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)//Цикл
                    {
                        SqlCommand cmd = new SqlCommand("UPDATE [Table_pass] SET processing = @processing, date_processing = @date_processing, akt = @akt WHERE id = @id", con);
                        cmd.Parameters.AddWithValue("@processing", "Обработано");
                        cmd.Parameters.AddWithValue("@date_processing", DateTime.Today.AddDays(0));
                        cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView2.Rows[i].Cells[11].Value));
                        cmd.Parameters.AddWithValue("@akt", Convert.ToInt32(dataGridView1.Rows[0].Cells[9].Value) + 1);
                        cmd.ExecuteNonQuery();
                    }
                    con.Close();//закрыть соединение 
                    MessageBox.Show("Обработка выполнена / Присвоен № Акта!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    //Vyborka();
                }
                //Выдача рееста в WORD
                int nomer = Convert.ToInt32(dataGridView1.Rows[0].Cells[9].Value) + 1;//№
                string type = comboBox6.Text;
                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Filter = "Word Documents (*.docx)|*.docx";
                sfd.FileName = "Акт № " + nomer + " " + type + ".docx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    if (comboBox6.Text == "Обычный" & checkBox1.Checked == false)
                    {
                        Obychnye(dataGridView2, sfd.FileName);
                    }
                    else if (comboBox6.Text == "Срочный" & checkBox1.Checked == false)
                    {
                        Srochnye(dataGridView2, sfd.FileName);
                    }
                    else if (comboBox6.Text == "На дом" & checkBox1.Checked == false)
                    {
                        Na_dom(dataGridView2, sfd.FileName);
                    }
                    else if (comboBox6.Text == "МИД" & checkBox1.Checked == false)
                    {
                        MID(dataGridView2, sfd.FileName);
                    }
                    else if (comboBox6.Text == "Обычный" & checkBox1.Checked == true)
                    {
                        Itog_Obychnye(dataGridView2, sfd.FileName);
                    }
                    else if (comboBox6.Text == "Срочный" & checkBox1.Checked == true)
                    {
                        Itog_Srochnye(dataGridView2, sfd.FileName);
                    }
                }
            }
            //else if (dataGridView2.Rows.Count > 1 & Convert.ToString(dataGridView2.Rows[0].Cells[8].Value) == "Обработано")
            //{
            //    if (MessageBox.Show("Вы хотите открыть этот Акт?", "Внимание! Эти данные уже обработаны!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            //    {
            //        button6.Enabled = false;
            //        button6.Text = "Ожидайте!";
            //        //Выдача рееста в WORD
            //        int nomer = Convert.ToInt32(dataGridView2.Rows[0].Cells[7].Value);//№
            //        string type = Convert.ToString(dataGridView2.Rows[0].Cells[9].Value);
            //        SaveFileDialog sfd = new SaveFileDialog();
            //        sfd.Filter = "Word Documents (*.docx)|*.docx";
            //        sfd.FileName = "Акт № " + nomer + " " + type + ".docx";
            //        if (sfd.ShowDialog() == DialogResult.OK)
            //        {
            //            if (Convert.ToString(dataGridView2.Rows[0].Cells[9].Value) == "Обычный")
            //            {
            //                Obychnye(dataGridView2, sfd.FileName);
            //            }
            //            else if (Convert.ToString(dataGridView2.Rows[0].Cells[9].Value) == "Срочный")
            //            {
            //                Srochnye(dataGridView2, sfd.FileName);
            //            }
            //            else if (Convert.ToString(dataGridView2.Rows[0].Cells[9].Value) == "На дом")
            //            {
            //                Na_dom(dataGridView2, sfd.FileName);
            //            }
            //            else if (Convert.ToString(dataGridView2.Rows[0].Cells[9].Value) == "МИД")
            //            {
            //                MID(dataGridView2, sfd.FileName);
            //            }
            //        }
            //    }
            //}
            else MessageBox.Show("Данные не найдены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            dataGridView1.Visible = true;
            dataGridView2.Visible = false;
            dataGridView3.Visible = false;
            label16.Visible = false;
            button6.Enabled = true;
            disp_data();       
            itog();
            Podschet();
            disp_data();
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;
            textBox7.Text = "";
            //Datagridview удаление всех столбцов
            int sum = dataGridView2.Columns.Count;
            for (int i = 0; i < sum; i++)
            {
                dataGridView2.Columns.RemoveAt(0);
            }
        }
        //------------------------------------WORD-----------------------------------------------//
        public void Obychnye(DataGridView dataGridView2, string filename)//Метод экспорта в Word
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Акт");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 12;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 10; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView2.Rows.Count != 0)
            {
                //this.dataGridView1.Columns.RemoveAt(4);//удаление столбца

                string kol_vo = Convert.ToString(textBox3.Text);//кол-во
                string id = Convert.ToString(textBox8.Text);//ID
                string ogp = Convert.ToString(textBox9.Text);//ОГП
                string itogo = Convert.ToString(textBox6.Text);//Итого
                string punkt = Convert.ToString(dataGridView2.Rows[0].Cells[7].Value);//Пункт
                DateTime month = Convert.ToDateTime(dataGridView2.Rows[0].Cells[0].Value);
                int RowCount = dataGridView2.Rows.Count;
                int ColumnCount = dataGridView2.Columns.Count - 5;// столбцы в гриде (-5 последних)             
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r + 1, c] = dataGridView2.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                  //Добавление текста в документ
               
                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = "     Всего за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года доставлено " + itogo + " шт." +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "Ответственный сотрудник                                                                 Подпись руководителя и" +
                Environment.NewLine + "___________________ОПРН/ПО                                                      главного бухгалтера" +
                Environment.NewLine + "                                                                                                              ГП 'Кыргыз почтасы'" +
                Environment.NewLine + "                                                                                                              при ГРС при ПКР" +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "                 МП                                                                                                  МП" +
                Environment.NewLine;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблицы
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //стиль строки заголовка
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = "";
                    oDoc.Application.Selection.Tables[1].Cell(2, c + 1).Range.Text = dataGridView2.Columns[c].HeaderText;
                }             
                //стиль таблицы   
                oDoc.Application.Selection.Tables[1].Columns[2].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Columns[2].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 100;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 140;//ширина столбца
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -0;//Установка отступа слева
                oDoc.Application.Selection.Tables[1].Cell(1, 2).Range.Text = "Количество доставленных обычных паспортов, шт";
                oDoc.Application.Selection.Tables[1].Cell(1, 2).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 4));//Объединение
                //Добавить текст в последнюю строку в таблице
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[1].Range.Text = "Итого";
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[2].Range.Text = id;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[3].Range.Text = ogp;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[4].Range.Text = itogo;
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;
                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    headerRange.Text =
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine + "Акт выполненных работ" +
                    Environment.NewLine + "Между ГП 'Кыргыз почтасы' и ОПРН/ПО " + punkt + " района" +
                    Environment.NewLine + "по доставке идентификационных карт-паспортов гражданина образца 2017 года и" +
                    Environment.NewLine + "общегражданских паспортов, изготовленных в порядке 'Обычные' " +
                    Environment.NewLine + "по " + punkt +
                    Environment.NewLine + "за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года" +
                    Environment.NewLine;

                    headerRange.Font.Size = 13;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    //footerRange.Text = "ЦМПОЛ       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Srochnye(DataGridView dataGridView2, string filename)//Метод экспорта в Word
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Акт");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 12;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 10; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView2.Rows.Count != 0)
            {
                //удаление столбца
                //this.dataGridView1.Columns.RemoveAt(4);//дата записи

                string reis = Convert.ToString(textBox3.Text);//кол-во рейсов
                string id = Convert.ToString(textBox8.Text);//ID
                string ogp = Convert.ToString(textBox9.Text);//ОГП
                string itogo = Convert.ToString(textBox6.Text);//Итого
                string punkt = Convert.ToString(dataGridView2.Rows[0].Cells[7].Value);//Пункт
                DateTime month = Convert.ToDateTime(dataGridView2.Rows[0].Cells[0].Value);
                int RowCount = dataGridView2.Rows.Count;
                int ColumnCount = dataGridView2.Columns.Count - 6;// столбцы в гриде (-6 последних)             
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r+1, c] = dataGridView2.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                  //Добавление текста в документ

                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = "     Всего за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года оказаны услуги " + reis + " раз " +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "Ответственный сотрудник                                                                 Подпись руководителя и" +
                Environment.NewLine + "ОПРН/ПО                                                                                            главного бухгалтера" +
                Environment.NewLine + "                                                                                                              ГП 'Кыргыз почтасы'" +
                Environment.NewLine + "                                                                                                              при ГРС при ПКР" +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "                 МП                                                                                                  МП" +
                Environment.NewLine;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблицы
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //стиль строки заголовка
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = "";
                    oDoc.Application.Selection.Tables[1].Cell(2, c + 1).Range.Text = dataGridView2.Columns[c].HeaderText;
                }
                //стиль таблицы
                oDoc.Application.Selection.Tables[1].Columns[2].Delete();//Удалить столбец     
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 100;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 100;//ширина столбца
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -0;//Установка отступа слева
                oDoc.Application.Selection.Tables[1].Cell(1, 3).Range.Text = "Количество доставленных срочных паспортов, шт.";
                oDoc.Application.Selection.Tables[1].Cell(1, 3).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 5));//Объединение
                //Добавить текст в последнюю строку в таблице
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[1].Range.Text = "Итого";
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[2].Range.Text = reis;             
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[3].Range.Text = id;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[4].Range.Text = ogp;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[5].Range.Text = itogo;

                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;

                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул

                    headerRange.Text =
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine + "Акт выполненных работ" +
                    Environment.NewLine + "Между ГП 'Кыргыз почтасы' и ОПРН/ПО " + punkt + " района" +
                    Environment.NewLine + "по доставке идентификационных карт-паспортов гражданина образца 2017 года и" +
                    Environment.NewLine + "общегражданских паспортов, изготовленных в 'Срочном режиме'" +
                    Environment.NewLine + "по " + punkt + 
                    Environment.NewLine + "за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года" +
                    Environment.NewLine;

                    headerRange.Font.Size = 13;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    //footerRange.Text = "ЦМПОЛ       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Na_dom(DataGridView dataGridView2, string filename)//Метод экспорта в Word
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Акт");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 12;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 10; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView2.Rows.Count != 0)
            {
                //удаление столбца
                //this.dataGridView1.Columns.RemoveAt(4);//дата записи
                string kol_vo = Convert.ToString(textBox3.Text);//кол-во
                string id = Convert.ToString(textBox8.Text);//ID
                string ogp = Convert.ToString(textBox9.Text);//ОГП
                string itogo = Convert.ToString(textBox6.Text);//Итого
                string punkt = Convert.ToString(dataGridView2.Rows[0].Cells[7].Value);//Пункт
                DateTime month = Convert.ToDateTime(dataGridView2.Rows[0].Cells[0].Value);
                int RowCount = dataGridView2.Rows.Count;
                int ColumnCount = dataGridView2.Columns.Count - 5;// столбцы в гриде (-5 последних)             
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r+1, c] = dataGridView2.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                  //Добавление текста в документ

                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = "     Всего за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года принято для доставки " + itogo + " шт." +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "Подпись руководителя ОАДГ                                                       Подпись руководителя и" +
                Environment.NewLine + "ДРНАГС при ГРС при ПКР                                                           главного бухгалтера" +
                Environment.NewLine + "                                                                                                          ГП 'Кыргыз почтасы'" +
                Environment.NewLine + "                                                                                                          при ГРС при ПКР" +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "                 МП                                                                                                МП" +
                Environment.NewLine;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблицы
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //стиль строки заголовка
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = "";
                    oDoc.Application.Selection.Tables[1].Cell(2, c + 1).Range.Text = dataGridView2.Columns[c].HeaderText;
                }
                //стиль таблицы     
                oDoc.Application.Selection.Tables[1].Columns[2].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Columns[2].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 100;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 80;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 140;//ширина столбца
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -0;//Установка отступа слева
                oDoc.Application.Selection.Tables[1].Cell(1, 2).Range.Text = "Количество принятых обычных паспортов, шт.";
                oDoc.Application.Selection.Tables[1].Cell(1, 2).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 4));//Объединение
                //Добавить текст в последнюю строку в таблице
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[1].Range.Text = "Итого";
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[2].Range.Text = id;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[3].Range.Text = ogp;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[4].Range.Text = itogo;
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;

                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    headerRange.Text =
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine + "Акт выполненных работ" +
                    Environment.NewLine + "Между ГП 'Кыргыз почтасы' и ОАДГ ДРНАГС " + punkt + " района" +
                    Environment.NewLine + "по приему идентификационных карт-паспортов гражданина образца 2017 года и" +
                    Environment.NewLine + "общегражданских паспортов с 'Доставкой на дом' " +
                    Environment.NewLine + "по " + punkt +
                    Environment.NewLine + "за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года" +
                    Environment.NewLine;
                    headerRange.Font.Size = 13;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    //footerRange.Text = "ЦМПОЛ       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void MID(DataGridView dataGridView2, string filename)//Метод экспорта в Word
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Акт");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 12;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 10; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView2.Rows.Count != 0)
            {
                //удаление столбца
                //this.dataGridView1.Columns.RemoveAt(4);//дата записи
                string kol_vo = Convert.ToString(textBox3.Text);//кол-во
                string id = Convert.ToString(textBox8.Text);//ID
                string ogp = Convert.ToString(textBox9.Text);//ОГП
                string itogo = Convert.ToString(textBox6.Text);//Итого
                string punkt = Convert.ToString(dataGridView2.Rows[0].Cells[7].Value);//Пункт
                DateTime month = Convert.ToDateTime(dataGridView2.Rows[0].Cells[0].Value);
                int RowCount = dataGridView2.Rows.Count;
                int ColumnCount = dataGridView2.Columns.Count - 5;// столбцы в гриде (-5 последних)             
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r+1, c] = dataGridView2.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                  //Добавление текста в документ

                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = "     Всего за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года принято для доставки " + itogo + " шт." +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "Подпись руководителя ОАДГ                                                       Подпись руководителя и" +
                Environment.NewLine + "ДРНАГС при ГРС при ПКР                                                           главного бухгалтера" +
                Environment.NewLine + "                                                                                                          ГП 'Кыргыз почтасы'" +
                Environment.NewLine + "                                                                                                          при ГРС при ПКР" +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "                 МП                                                                                                МП" +
                Environment.NewLine;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблицы
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //стиль строки заголовка
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = "";
                    oDoc.Application.Selection.Tables[1].Cell(2, c + 1).Range.Text = dataGridView2.Columns[c].HeaderText;
                }
                //стиль таблицы     
                oDoc.Application.Selection.Tables[1].Columns[2].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Columns[2].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Columns[2].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Columns[3].Delete();//Удалить столбец
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 100;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 100;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 100;//ширина столбца
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -0;//Установка отступа слева             
                oDoc.Application.Selection.Tables[1].Cell(1, 2).Range.Text = "Количество принятых паспортов для доставки в Дип. службу МИД КР, шт.";
                oDoc.Application.Selection.Tables[1].Cell(1, 2).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 3));//Объединение
                //Добавить текст в последнюю строку в таблице
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[1].Range.Text = "Итого";
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[2].Range.Text = ogp;
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;

                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул
                    headerRange.Text =
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine + "Акт выполненных работ" +
                    Environment.NewLine + "Между ГП 'Кыргыз почтасы' и ОАДГ ДРНАГС " + punkt + " района" +
                    Environment.NewLine + "по приему идентификационных карт-паспортов гражданина образца 2017 года и" +
                    Environment.NewLine + "общегражданских паспортов в адрес Дипломатической консульской службы Министерство иностранных дел Кыргызской Республики " +
                    Environment.NewLine + "по МИД ОПРН/ПО" +
                    Environment.NewLine + "за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года" +
                    Environment.NewLine;
                    headerRange.Font.Size = 13;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    //footerRange.Text = "ЦМПОЛ       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        //------------------------------------ИТОГОВЫЕ---------------------------------------------------//
        public void Itog_Obychnye(DataGridView dataGridView2, string filename)//Метод экспорта в Word
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientLandscape;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Акт");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 12;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 10; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView2.Rows.Count != 0)
            {
                //удаление столбца
                //this.dataGridView1.Columns.RemoveAt(4);//дата записи
                string id = Convert.ToString(textBox8.Text);//ID
                string ogp = Convert.ToString(textBox9.Text);//ОГП
                string itogo = Convert.ToString(textBox6.Text);//Итого
                
                string id_na_dom = Convert.ToString(textBox14.Text);//Итого
                string ogp_na_dom = Convert.ToString(textBox15.Text);//Итого
                string mid = Convert.ToString(textBox16.Text);//Итого
                string itogo2 = Convert.ToString(textBox17.Text);//Итого
                string vsegosht = Convert.ToString(textBox18.Text);//Итого

                string obsumm = Convert.ToString(textBox13.Text);//сумма
                string summ = Convert.ToString(textBox10.Text);//сумма
                string reis = Convert.ToString(textBox3.Text);//кол-во рейсов
                
                string punkt = Convert.ToString(dataGridView2.Rows[0].Cells[6].Value);//Пункт
                DateTime month = Convert.ToDateTime(dataGridView2.Rows[0].Cells[0].Value);
                int RowCount = dataGridView2.Rows.Count;
                int ColumnCount = dataGridView2.Columns.Count - 0;// столбцы в гриде (-6 последних)         
                Object[,] DataArray = new object[RowCount + 2, ColumnCount + 2];
                // добавить строки              
                int r = 0;
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    for (r = 0; r <= RowCount - 1; r++)
                    {
                        DataArray[r + 1, c] = dataGridView2.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                  //Добавление текста в документ

                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = "     Всего за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года оказаны услуги на сумму " + obsumm + " сом/тыйын." +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "Подпись руководителя и                                                     Подпись руководителя и                                       Подпись руководителя и " +
                Environment.NewLine + "ответственного сотрудника                                                 ответственного сотрудника                                  ответственного сотрудника" +
                Environment.NewLine + "ДРНиАГС ПКР ГРС ПКР                                                    ГЦП ГП 'Инфоком' при ГРС при ПКР                 ГП 'Кыргыз почтасы' при ГРС при ПКР" +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "                 МП                                                                                                 МП                                                                        МП" +
                Environment.NewLine;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблицы
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //стиль строки заголовка
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Range.Bold = 1;//нижняя строка
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {                   
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = "";
                    oDoc.Application.Selection.Tables[1].Cell(2, c + 1).Range.Text = dataGridView2.Columns[c].HeaderText;                  
                }
                //стиль таблицы
                oDoc.Application.Selection.Tables[1].Columns[1].Delete();//Удалить столбец          
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 130;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 45;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 45;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 50;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 45;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[6].Width = 45;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[7].Width = 55;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[8].Width = 60;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[9].Width = 60;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[10].Width = 60;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[11].Width = 70;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[12].Width = 72;//ширина столбца
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -0;//Установка отступа слева
                oDoc.Application.Selection.Tables[1].Cell(1, 2).Range.Text = "Количество принятых обычных паспортов от ГЦП ГП Инфоком, шт.";
                oDoc.Application.Selection.Tables[1].Cell(1, 2).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 4));//Объединение
                oDoc.Application.Selection.Tables[1].Cell(1, 3).Range.Text = "Количество принятых обычных паспортов от ОАДГ, шт.";
                oDoc.Application.Selection.Tables[1].Cell(1, 3).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 6));//Объединение
                oDoc.Application.Selection.Tables[1].Cell(2, 6).Range.Text = "";//Удалить содержимое ячейки
                oDoc.Application.Selection.Tables[1].Cell(2, 5).Range.Text = "С доставкой на дом                    ______________    ID          ОГП ";
                oDoc.Application.Selection.Tables[1].Cell(2, 5).Merge(oDoc.Application.Selection.Tables[1].Cell(2, 6));//Объединение
                oDoc.Application.Selection.Tables[1].Cell(2, 6).Range.Text = "Дип. служба МИД КР      ОГП";              
                //Добавить текст в последнюю строку в таблице
                //oDoc.Application.Selection.Tables[1].Rows.Add();//добавить пустую строку
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[1].Range.Text = "Итого";
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[2].Range.Text = id;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[3].Range.Text = ogp;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[4].Range.Text = itogo;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[5].Range.Text = id_na_dom;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[6].Range.Text = ogp_na_dom;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[7].Range.Text = mid;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[8].Range.Text = itogo2;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[9].Range.Text = vsegosht;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[11].Range.Text = obsumm;
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;

                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул

                    headerRange.Text =
                    Environment.NewLine + "Акт выполненных работ по принятым обычным паспортам для доставки в адрес ОПРН/ПО и ДКС МИД КР за " + Convert.ToString(month.ToString("MMMM yyyy")) + " г." +
                    Environment.NewLine;

                    headerRange.Font.Size = 13;
                    headerRange.Font.Name = "Times New Roman";
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    //footerRange.Text = "ЦМПОЛ       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }
        public void Itog_Srochnye(DataGridView dataGridView2, string filename)//Метод экспорта в Word
        {
            Word.Document oDoc = new Word.Document();
            oDoc.Application.Visible = true;
            //ориентация страницы
            oDoc.PageSetup.Orientation = Word.WdOrientation.wdOrientPortrait;
            // Стиль текста.
            object start = 0, end = 0;
            Word.Range rng = oDoc.Range(ref start, ref end);
            rng.InsertBefore("Акт");//Заголовок
            rng.Font.Name = "Times New Roman";
            rng.Font.Size = 12;
            rng.InsertParagraphAfter();
            rng.InsertParagraphAfter();
            rng.SetRange(rng.End, rng.End);
            oDoc.Content.ParagraphFormat.LeftIndent = oDoc.Content.Application.CentimetersToPoints(0);  // отступ слева
            oDoc.Paragraphs.Format.FirstLineIndent = 0; //Отступ первой строки
            oDoc.Paragraphs.Format.LineSpacing = 10; //межстрочный интервал в первом абзаце.(высота строк)
            oDoc.Paragraphs.Format.SpaceBefore = 3; //межстрочный интервал перед первым абзацем.
            oDoc.Paragraphs.Format.SpaceAfter = 1; //межстрочный интервал после первого абзаца.

            if (dataGridView2.Rows.Count != 0)
            {
                //удаление столбца
                //this.dataGridView1.Columns.RemoveAt(4);//дата записи
                string summuslug = Convert.ToString(textBox17.Text);//сумма
                string summ = Convert.ToString(textBox10.Text);//сумма
                string reis = Convert.ToString(textBox3.Text);//кол-во рейсов
                string id = Convert.ToString(textBox8.Text);//ID
                string ogp = Convert.ToString(textBox9.Text);//ОГП
                string itogo = Convert.ToString(textBox6.Text);//Итого
                string punkt = Convert.ToString(dataGridView2.Rows[0].Cells[6].Value);//Пункт
                DateTime month = Convert.ToDateTime(dataGridView2.Rows[0].Cells[0].Value);
                int RowCount = dataGridView2.Rows.Count;
                int ColumnCount = dataGridView2.Columns.Count - 4;// столбцы в гриде (-6 последних)             
                Object[,] DataArray = new object[RowCount + 1, ColumnCount + 1];
                // добавить строки
                int r = 0;
                for (int c = 0; c <= ColumnCount-1; c++)
                {
                    for (r = 0; r <= RowCount-1; r++)
                    {
                        DataArray[r+1, c] = dataGridView2.Rows[r].Cells[c].Value;
                    } //Конец цикла строки
                } //конец петли колонки
                  //Добавление текста в документ

                oDoc.Content.SetRange(0, 0);// для текстовых строк
                oDoc.Content.Text = "     Всего за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года оказаны услуги на сумму " + summ + " сом." +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "Подпись руководителя ГЦП       Подпись руководителя и           Подпись руководителя и" +
                Environment.NewLine + "ГП 'Инфоком'                                ответственного сотрудника      ответственного сотрудника" +
                Environment.NewLine + "при ГРС при ПКР                         ГРС при ПКР                              ГП 'Кыргыз почтасы'" +
                Environment.NewLine + "                                                                                                             при ГРС при ПКР" +
                Environment.NewLine +
                Environment.NewLine +
                Environment.NewLine + "                 МП                                              МП                                             МП" +
                Environment.NewLine;

                dynamic oRange = oDoc.Content.Application.Selection.Range;
                string oTemp = "";
                for (r = 0; r <= RowCount - 1; r++)
                {
                    for (int c = 0; c <= ColumnCount - 1; c++)
                    {
                        oTemp = oTemp + DataArray[r, c] + "\t";
                    }
                }
                //формат таблицы
                oRange.Text = oTemp;
                object Separator = Word.WdTableFieldSeparator.wdSeparateByTabs;
                object ApplyBorders = true;
                object AutoFit = true;
                object AutoFitBehavior = Word.WdAutoFitBehavior.wdAutoFitContent;

                oRange.ConvertToTable(ref Separator, ref RowCount, ref ColumnCount,
                                      Type.Missing, Type.Missing, ref ApplyBorders,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, Type.Missing, Type.Missing,
                                      Type.Missing, ref AutoFit, ref AutoFitBehavior, Type.Missing);

                oRange.Select();
                oDoc.Application.Selection.Tables[1].Select();
                oDoc.Application.Selection.Tables[1].Rows.AllowBreakAcrossPages = 0;
                oDoc.Application.Selection.Tables[1].Rows.Alignment = 0;
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.InsertRowsAbove(1);
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                //стиль строки заголовка
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[2].Range.Bold = 1;
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Name = "Times New Roman";
                oDoc.Application.Selection.Tables[1].Rows[1].Range.Font.Size = 9;
                //добавить строку заголовка вручную
                for (int c = 0; c <= ColumnCount - 1; c++)
                {
                    oDoc.Application.Selection.Tables[1].Cell(1, c + 1).Range.Text = "";
                    oDoc.Application.Selection.Tables[1].Cell(2, c + 1).Range.Text = dataGridView2.Columns[c].HeaderText;
                }
                //стиль таблицы
                oDoc.Application.Selection.Tables[1].Columns[1].Delete();//Удалить столбец     
                oDoc.Application.Selection.Tables[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;//Выравнивание текста в таблице по центру           
                oDoc.Application.Selection.Tables[1].Rows.Borders.Enable = 1;//borders              
                oDoc.Application.Selection.Tables[1].Rows[1].Select();
                oDoc.Application.Selection.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                oDoc.Application.Selection.Tables[1].Columns[1].Width = 180;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[2].Width = 50;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[3].Width = 70;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[4].Width = 70;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[5].Width = 35;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[6].Width = 35;//ширина столбца
                oDoc.Application.Selection.Tables[1].Columns[7].Width = 45;//ширина столбца
                oDoc.Application.Selection.Tables[1].LeftPadding = 1;//отступ с лева полей ячеек
                oDoc.Application.Selection.Tables[1].RightPadding = 1;//отступ с права полей ячеек
                oDoc.Application.Selection.Tables[1].Rows.LeftIndent = -30;//Установка отступа слева
                oDoc.Application.Selection.Tables[1].Cell(1, 5).Range.Text = "Количество принятых паспортов, шт.";
                oDoc.Application.Selection.Tables[1].Cell(1, 5).Merge(oDoc.Application.Selection.Tables[1].Cell(1, 7));//Объединение
                //Добавить текст в последнюю строку в таблице
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[1].Range.Text = "Итого";
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[2].Range.Text = reis;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[3].Range.Text = summuslug;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[4].Range.Text = summ;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[5].Range.Text = id;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[6].Range.Text = ogp;
                oDoc.Application.Selection.Tables[1].Rows[r + 2].Cells[7].Range.Text = itogo;
                //текст заголовка
                foreach (Word.Section section in oDoc.Application.ActiveDocument.Sections)
                {//Верхний колонтитул
                    DateTime Now = DateTime.Now;

                    Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage].Range;
                    headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                    section.PageSetup.DifferentFirstPageHeaderFooter = -1;//Включить особый колонтитул

                    headerRange.Text =
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine +
                    Environment.NewLine + "Акт выполненных работ" +
                    Environment.NewLine + "Между Государственной регистрационной службой" +
                    Environment.NewLine + "при Правительстве Кыргызской Республики и" +
                    Environment.NewLine + "Государственным предприятием 'Кыргыз почтасы' при ГРС при ПКР" +
                    Environment.NewLine + "по доставке идентификационных карт-паспортов гражданина образца 2017 года и" +
                    Environment.NewLine + "общегражданских паспортов, изготовленных в 'Срочном' режиме." +
                    Environment.NewLine + "за " + Convert.ToString(month.ToString("MMMM yyyy")) + " года" +
                    Environment.NewLine;

                    headerRange.Font.Size = 13;
                    headerRange.Font.Name = "Times New Roman";
                    //headerRange.Font.Bold = 1;
                    headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                    //Нижний колонтитул
                    Word.Range footerRange = section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                    footerRange.Fields.Add(footerRange, Word.WdFieldType.wdFieldPage);
                    //footerRange.Text = "ЦМПОЛ       " + Convert.ToString(Now.ToString("dd.MM.yyyy"));
                    footerRange.Font.Size = 9;
                    footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft;
                }
                //сохранить файл
                oDoc.SaveAs(filename);
            }
        }



        //-----------------------------Админпанель---------------------------//
        private void button3_Click(object sender, EventArgs e)//Добавить
        {
            dataGridView3.Visible = true;
            dataGridView2.Visible = false;
            dataGridView1.Visible = false;
            if (comboBox4.Text != "" & textBox4.Text != "" & comboBox3.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_users] (users, password, access) VALUES (@users, @password, @access)", con);
                cmd.Parameters.AddWithValue("@users", comboBox4.Text);
                cmd.Parameters.AddWithValue("@password", textBox4.Text);
                cmd.Parameters.AddWithValue("@access", comboBox3.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                comboBox3.Text = "";//очистка текстовых полей
                textBox4.Text = "";
                comboBox4.Text = "";
                comboBox4.Select();//Установка курсора
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (textBox5.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_punkts] (punkt) VALUES (@punkt)", con);
                cmd.Parameters.AddWithValue("@punkt", textBox5.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                textBox5.Text = "";
                textBox5.Select();//Установка курсора
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void button4_Click(object sender, EventArgs e)//Изменить
        {
            dataGridView3.Visible = true;
            dataGridView2.Visible = false;
            dataGridView1.Visible = false;
            if (comboBox4.Text != "" & textBox4.Text != "" & comboBox3.Text != "" & dataGridView3.Rows.Count == 1)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("UPDATE [Table_users] SET users = @users, password = @password, access = @access WHERE id = @id", con);
                cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView3.Rows[0].Cells[0].Value));//первая строка в гриде
                cmd.Parameters.AddWithValue("@users", comboBox4.Text);
                cmd.Parameters.AddWithValue("@password", textBox4.Text);
                cmd.Parameters.AddWithValue("@access", comboBox3.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                MessageBox.Show("Готово", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                comboBox4.Select();//Установка курсора
            }
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            comboBox3.Text = "";//очистка текстовых полей
            textBox4.Text = "";
            comboBox4.Text = "";
        }
        private void button5_Click(object sender, EventArgs e)//Удалить
        {
            dataGridView3.Visible = true;
            dataGridView2.Visible = false;
            dataGridView1.Visible = false;
            if (comboBox4.Text != "" & dataGridView3.Rows.Count == 1)
            {
                if (MessageBox.Show("Вы хотите удалить эту запись?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("DELETE FROM [Table_users] WHERE id = @id", con);
                    cmd.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView3.Rows[0].Cells[0].Value));//первая строка в гриде
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                    MessageBox.Show("Запись успешно удалена!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    disp_data();
                    comboBox4.Select();//Установка курсора
                }
                else
                {
                    disp_data();
                    comboBox4.Select();//Установка курсора
                }
            }
            else if (dataGridView3.Rows.Count != 1)
            {
                MessageBox.Show("Произведите поиск", "Внимание! Чтобы удалить запись из базы данных", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            else if (dataGridView3.Rows.Count <= 0)
            {
                MessageBox.Show("В базе не найдено", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            comboBox3.Text = "";//очистка текстовых полей
            textBox4.Text = "";
            comboBox4.Text = "";
        }

        private void button1_Click(object sender, EventArgs e)//Ввод
        {
            if (textBox1.Text != "" & textBox2.Text != "" & comboBox1.Text != "" & comboBox5.Text != "")
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_pass] (punkt, id_id, ip_ip, type, date, akt, processing) VALUES (@punkt, @id_id, @ip_ip, @type, @date, @akt, @processing)", con);
                cmd.Parameters.AddWithValue("@punkt", comboBox1.Text);
                cmd.Parameters.AddWithValue("@processing", "Не обработано");
                cmd.Parameters.AddWithValue("@id_id", textBox1.Text);
                cmd.Parameters.AddWithValue("@ip_ip", textBox2.Text);
                cmd.Parameters.AddWithValue("@date", dateTimePicker3.Value);
                cmd.Parameters.AddWithValue("@akt", 0);
                if (comboBox5.Text == "Обычные") { cmd.Parameters.AddWithValue("@type", "Обычные"); }
                else if (comboBox5.Text == "Срочные") { cmd.Parameters.AddWithValue("@type", "Срочные"); }
                else if (comboBox5.Text == "На дом") { cmd.Parameters.AddWithValue("@type", "На дом"); }
                else if (comboBox5.Text == "МИД") { cmd.Parameters.AddWithValue("@type", "МИД"); }
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                textBox1.Text = "";//очистка текстовых полей
                textBox2.Text = "";
                comboBox1.Text = "";
                comboBox1.Select();//Установка курсора
                MessageBox.Show("Вы успешно добавили запись!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else
            {
                MessageBox.Show("Не все поля заполнены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            disp_data();
            Podschet();
        }

        private void comboBox4_TextChanged(object sender, EventArgs e)//Поиск пользователя
        {
            dataGridView3.Visible = true;
            dataGridView2.Visible = false;
            dataGridView1.Visible = false;
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT id, users, password, access FROM [Table_users]" +
                "WHERE users = @users", con);
            cmd.Parameters.AddWithValue("@users", Convert.ToString(comboBox4.Text));
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView3.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение 
            if (comboBox4.Text == "")//если поле очищено, отобразить базу
            {
                disp_data();
            }
        }
        private void textBox5_TextChanged(object sender, EventArgs e)//Поиск
        {
            dataGridView3.Visible = true;
            dataGridView2.Visible = false;
            dataGridView1.Visible = false;
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT id_sourse FROM [Table_pass]" +
                "WHERE punkt LIKE '%" + Convert.ToString(textBox5.Text) + "%'", con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView3.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            Podschet();//произвести подсчет по методу
            if (textBox5.Text == "")//если поле очищено, отобразить базу
            {
                disp_data();
            }
        }

        private void button7_Click(object sender, EventArgs e)//Отобразить за период
        {
            button7.Text = "Ожидайте";
            button7.Enabled = false;
            
            dataGridView1.Visible = true;
            dataGridView2.Visible = false;
            dataGridView3.Visible = false;
            
            con.Open();//открыть соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT id AS Id, date AS Дата, reis AS Рейсы, id_id AS eID, ip_ip AS ОГП, itogo AS Итого, primecanie AS Примечание," +
                " type AS Тип, punkt AS ОПРН, akt AS Акт, marshrut AS Маршруты FROM [Table_pass]" +
                " WHERE date BETWEEN @StartDate AND @EndDate ORDER BY date";//WHERE reis NOT IN ('0') чтобы не отображать нулевые
            cmd.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cmd.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение 
            
            Marshruts();
            itog();
            Podschet();
                       
            button7.Enabled = true;
            button7.Text = "Отобразить за период";

            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            comboBox5.SelectedIndex = -1;
            comboBox6.SelectedIndex = -1;           
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)//Окраска
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                string S = Convert.ToString(dataGridView1.Rows[i].Cells[7].Value);//Тип
                string V = Convert.ToString("Обычный");
                string L = Convert.ToString("Срочный");
                string K = Convert.ToString("На дом");
                string W = Convert.ToString("МИД");
                if (S == V)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightYellow;//
                }
                else if (S == L)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightSkyBlue;//                
                }
                else if (S == K)
                {
                    dataGridView1.Rows[i].DefaultCellStyle.BackColor = Color.LightSteelBlue;//            
                }
            }
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)//Закрытие формы Выход
        {
            Application.Exit();
        }

        private void button8_Click(object sender, EventArgs e)//test
        {
            con.Open();//открыть соединение
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                SqlCommand c = new SqlCommand("UPDATE [Table_pass] SET reis = @reis WHERE id = @id AND date BETWEEN @StartDate AND @EndDate", con);
                c.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
                c.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
                c.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));//первая строка в гриде
                c.Parameters.AddWithValue("@reis", 0);
                c.ExecuteNonQuery();
            }
            con.Close();//закрыть соединение              
            con.Open();//открыть соединение
            SqlCommand cm = new SqlCommand("SELECT MIN(id) AS id, date AS Дата, MIN(reis) AS Рейсы, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого, max(primecanie) AS Примечание," +
                " MIN(type) AS Тип, max(punkt) AS ОПРН, MIN(akt) AS Акт, marshrut AS Маршруты FROM [Table_pass]" +
                " WHERE type=@type AND date BETWEEN @StartDate AND @EndDate GROUP BY date, marshrut ORDER BY date", con);
            cm.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cm.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cm.Parameters.AddWithValue("@type", "Срочный");
            cm.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cm);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dataGridView1.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение          
            //--------------------------------------Ставим рейсы по выборке-------------------------------------------------------//
            con.Open();//открыть соединение
            for (int i = 0; i < dataGridView1.Rows.Count; i++)//Цикл
            {
                {
                    SqlCommand c = new SqlCommand("UPDATE [Table_pass] SET reis = @reis WHERE id = @id AND date BETWEEN @StartDate AND @EndDate", con);
                    c.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
                    c.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
                    c.Parameters.AddWithValue("@id", Convert.ToInt32(dataGridView1.Rows[i].Cells[0].Value));//первая строка в гриде
                    c.Parameters.AddWithValue("@reis", 1);
                    c.ExecuteNonQuery();
                }
            }
            con.Close();//закрыть соединение

            con.Open();//открыть соединение
            SqlCommand cm4 = new SqlCommand("SELECT MIN(id) AS id, date AS Дата, MAX(reis) AS Рейсы, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого, max(primecanie) AS Примечание," +
                " MIN(type) AS Тип, max(punkt) AS ОПРН, MIN(akt) AS Акт, marshrut AS Маршруты FROM [Table_pass]" +
                " WHERE type=@type AND date BETWEEN @StartDate AND @EndDate GROUP BY date, marshrut ORDER BY date", con);
            cm4.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cm4.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cm4.Parameters.AddWithValue("@type", "Срочный");
            cm4.ExecuteNonQuery();
            DataTable dt4 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da4 = new SqlDataAdapter(cm4);//создаем экземпляр класса SqlDataAdapter
            dt4.Clear();//чистим DataTable, если он был не пуст
            da4.Fill(dt4);//заполняем данными созданный DataTable
            dataGridView1.DataSource = dt4;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            //--------------------------------------Окончательная выборка для выгрузки в WORD-----------------------------------------------------------//
            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("SELECT max(date) AS Дата, marshrut AS Маршруты, SUM(reis) AS Рейсы, MAX(stoimost) AS Стоимость, SUM(allsumm) AS Сумма, SUM(id_id) AS ID, SUM(ip_ip) AS ОГП, SUM(itogo) AS Итого," +
                " MIN(akt) AS Акт, MIN(processing) AS Processing, MIN(type) AS Тип, MIN(id) AS id FROM [Table_pass]" +
                " WHERE type = @type AND date BETWEEN @StartDate AND @EndDate GROUP BY marshrut ORDER BY marshrut", con);
            cmd.Parameters.AddWithValue("@type", "Срочный");
            cmd.Parameters.AddWithValue("@StartDate", dateTimePicker1.Value);
            cmd.Parameters.AddWithValue("@EndDate", dateTimePicker2.Value);
            cmd.ExecuteNonQuery();
            DataTable dt1 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da1 = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt1.Clear();//чистим DataTable, если он был не пуст
            da1.Fill(dt1);//заполняем данными созданный DataTable
            dataGridView2.DataSource = dt1;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//закрыть соединение
            dataGridView1.Visible = false;
            dataGridView2.Visible = true;
        }
    }
}
