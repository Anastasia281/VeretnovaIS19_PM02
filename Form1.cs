using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace VeretnovaIS19_ПМ02
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        int price = 0; //переменная хранящая в себе стоимость одного квадратного метра выбраного товара
        Regex proverka = new Regex(@"\d"); //проверка на введение только цифр
        double ItogCost = 0; //переменная хранит итоговую сумму покупки

        private void button1_Click(object sender, EventArgs e)
        {
            double width = Convert.ToInt32(textBox1.Text);
            double height = Convert.ToInt32(textBox2.Text);
            double area = (width / 100) * (height / 100);
            ItogCost = price * area;
            if (comboBox1.SelectedItem.Equals("Окна"))
            {
                if (radioButton1.Checked == true) ItogCost += 1000;
                if (radioButton2.Checked == true) ItogCost += 3400.5;
                if (radioButton3.Checked == true) ItogCost += 2560;
                if (radioButton4.Checked == true) ItogCost += 7900.9;
                if (radioButton5.Checked == true) ItogCost += 6210.5;
            }
            textBox3.Text = "Стоимость 1 кв метра = " + price + "; " +
                "\n Итоговая стоимость = " + ItogCost.ToString() + " руб. ";
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.SelectedItem.Equals("Окна") == true)
            {
                RadioButton(true);
                price = 2500;
            }
            else if (comboBox1.SelectedItem.Equals("Двери") == true)
            {
                RadioButton(false);
                price = 1900;
                pictureBox1.Image = Image.FromFile(Application.StartupPath + @"\Picture\7.jpg");

            }
            else if (comboBox1.SelectedItem.Equals("Балконы") == true)
            {
                RadioButton(false);
                price = 3400;
                pictureBox1.Image = Image.FromFile(Application.StartupPath + @"\Picture\6.jpg");
            }
        }

        public void RadioButton(bool a) //метод включающий и выключающий элементы radioButton
        {
            radioButton1.Enabled = a;
            radioButton2.Enabled = a;
            radioButton3.Enabled = a;
            radioButton4.Enabled = a;
            radioButton5.Enabled = a;
        }

        private void textBox2_Validating(object sender, CancelEventArgs e)
        {
            if (proverka.IsMatch(textBox2.Text) == false)
            {
                MessageBox.Show("Возможен ввод только цифр!");
            }
            if (textBox2.Text == "0" || textBox2.Text.Length == 0)
            {
                MessageBox.Show("Поле не может быть пустое или равное 0!");
            }
        }

        private void textBox1_Validating(object sender, CancelEventArgs e)
        {
            if (proverka.IsMatch(textBox1.Text) == false)
            {
                MessageBox.Show("Возможен ввод только цифр!");
            }
            if (textBox1.Text == "0" || textBox1.Text.Length == 0)
            {
                MessageBox.Show("Поле не может быть пустое или равное 0!");
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Word.Document doc = null;
            try
            {
                // Создаём объект приложения
                Word.Application app = new Word.Application();
                // Открываем шаблон
                doc = app.Documents.Open(Application.StartupPath + @"\ШаблонЧека.docx");
                Random rnd = new Random();
                int nomer = rnd.Next(1, 1000);
                // Добавляем информацию в шаблон
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                Word.Range wRange;
                int i = 0;
                DateTime date = DateTime.Now;
                string tovar = comboBox1.SelectedItem.ToString() + "(" + textBox1.Text + " * " + textBox2.Text + ")";
                string[] data = new string[4] { date.ToString(), ItogCost.ToString() + " руб.", nomer.ToString(), tovar};
                foreach (Word.Bookmark mark in wBookmarks)
                {

                    wRange = mark.Range;
                    wRange.Text = data[i];
                    i++;
                }
                // Сохраняем шаблон как новый файл
                doc.SaveAs2(Application.StartupPath + @"\Чеки\Чек №" + nomer);
                doc.Close();
                MessageBox.Show("Чек сохранен!");
            }
            catch (Exception ex)
            {
                // Если произошла ошибка, то
                // закрываем документ и выводим информацию
                doc.Close();
                doc = null;
                MessageBox.Show("Во время выполнения произошла ошибка!");
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox1.Image = Image.FromFile(Application.StartupPath + @"\Picture\1.jpg");
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox1.Image = Image.FromFile(Application.StartupPath + @"\Picture\2.jpg");
        }

        private void radioButton3_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox1.Image = Image.FromFile(Application.StartupPath + @"\Picture\3.jpg");
        }

        private void radioButton4_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox1.Image = Image.FromFile(Application.StartupPath + @"\Picture\4.jpg");
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            pictureBox1.Image = Image.FromFile(Application.StartupPath + @"\Picture\5.jpg");
        }
    }
}
