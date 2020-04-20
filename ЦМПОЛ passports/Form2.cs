using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Deployment.Application;
using System.Reflection;
using System.Data.SqlClient;

namespace ЦМПОЛ_passports
{
    public partial class Form2 : Form
    {

        SqlConnection con = new SqlConnection(@"Data Source=192.168.99.4;Initial Catalog=CmpolBase;Persist Security Info=True;User ID=Lan;Password=Samsung0");
        public Form2()
        {
            InitializeComponent();
            textBox1.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "Войти" с клавиатуры
            comboBoxF2.KeyDown += (s, e) => { if (e.KeyCode == Keys.Enter) button1_Click(new object(), new EventArgs()); };//Нажатие кнопки "Войти" с клавиатуры
        }

        public void users_select()//Вывод пользователей в Combobox
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = con.CreateCommand();
            cmd.CommandType = CommandType.Text;
            cmd.CommandText = "SELECT users FROM [Table_users] ORDER BY users DESC";
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            foreach (DataRow row in dt.Rows)
            {
                comboBoxF2.Items.Add(row[0].ToString());
            }
            con.Close();//Закрываем соединение
        }
        private void Form2_Load(object sender, EventArgs e)//Загрузка формы
        {
            users_select();
            comboBoxF2.SelectedIndex = 1;//пользователь по умолчанию
            textBox1.Select();//Установка курсора          
        }
        private void button1_Click(object sender, EventArgs e)//Войти
        {
            con.Open();//Открываем соединение
            SqlCommand cmd1 = con.CreateCommand();
            cmd1.CommandType = CommandType.Text;
            cmd1.CommandText = "SELECT * FROM [Table_users] WHERE users = @users";
            cmd1.Parameters.AddWithValue("@users", comboBoxF2.Text);
            cmd1.ExecuteNonQuery();
            DataTable dt1 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da1 = new SqlDataAdapter(cmd1);//создаем экземпляр класса SqlDataAdapter
            dt1.Clear();//чистим DataTable, если он был не пуст
            da1.Fill(dt1);//заполняем данными созданный DataTable
            con.Close();//Закрываем соединение
            Dostup.Access = dt1.Rows[0][3].ToString();//Доступ
            Dostup.Login = dt1.Rows[0][1].ToString();//Логин

            if (textBox1.Text != "" & dt1.Rows[0][2].ToString() == textBox1.Text)
            {
                Form1 Form1 = new Form1();
                //P.label1.Text = "Добро пожаловать! " + comboBoxF2.Text;
                Form1.Show();
                this.Hide();
            }
            else if(dt1.Rows[0][2].ToString() != textBox1.Text)
            {
                MessageBox.Show("Неверный пароль", "Внимание!");
            }
            else
            {
                MessageBox.Show("Введите пароль", "Внимание!");
            }
        }
        private void Form2_FormClosed(object sender, FormClosedEventArgs e)//Закрытие формы Выход
        {
            Application.Exit();
        }
        private void comboBoxF2_SelectedIndexChanged(object sender, EventArgs e)//Установить курсор после выбора
        {
            textBox1.Select();
        }
    }
}
