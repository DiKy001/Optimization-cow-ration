using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Optimization
{
    public partial class Economy : Form
    {
        private TableBase table;    // объект данных
        private bool ChangeData;    // для проверки на изменение данных 
        private Form saving = new Saving(); // всплывающее окно сохранения данных в Excel 
        private List<string[]> exp;
        private double sExp;

        public Economy(TableBase table, bool ChangeData)
        {
            this.table = table;
            this.ChangeData = ChangeData;
            InitializeComponent();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void hScrollBar1_Scroll(object sender, ScrollEventArgs e)
        {
            if ((hScrollBar1.Value + hScrollBar2.Value + hScrollBar3.Value) > 100)
                e.NewValue = 100 - (hScrollBar2.Value + hScrollBar3.Value);
            label5.Text = e.NewValue.ToString() + "%";
        }

        private void hScrollBar2_Scroll(object sender, ScrollEventArgs e)
        {
            if ((hScrollBar1.Value + hScrollBar2.Value + hScrollBar3.Value) > 100)
                e.NewValue = 100 - (hScrollBar1.Value + hScrollBar3.Value);
            label6.Text = e.NewValue.ToString() + "%";
        }

        private void hScrollBar3_Scroll(object sender, ScrollEventArgs e)
        {
            if ((hScrollBar1.Value + hScrollBar2.Value + hScrollBar3.Value) > 100)
                e.NewValue = 100 - (hScrollBar1.Value + hScrollBar2.Value);
            label7.Text = e.NewValue.ToString() + "%";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ChangeData) // проверка на изменение данных
            {
                // создание диалогового окна выхода
                DialogResult dialogresult = MessageBox.Show("Данные были изменены.\nВы хотите их сохранить перед выходом?", "Выход", MessageBoxButtons.YesNoCancel);
                if (dialogresult == DialogResult.Yes) // при нажатии на кнопку "Да" в диалоговом окне
                {
                    table.SaveData();  // сохранение данных в программный файл
                    Close();    // закрытие программы
                }
                if (dialogresult == DialogResult.No) // при нажатии на кнопку "Нет" в диалоговом окне
                {
                    Close(); // закрытие программы
                }
            }
            // при не измененных данных
            else
                Close(); // закрытие программы
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form form = new Menu(table, ChangeData);
            Hide();
            form.ShowDialog();
            Close();
        }

        private void Economy_Load(object sender, EventArgs e)
        {
            exp = new List<string[]>();
            StreamReader file = new StreamReader("Data\\Expenses.txt");
            string read;
            while (!string.IsNullOrEmpty(read = file.ReadLine()))
            {
                exp.Add(read.Split('\t'));
            }
            file.Close();

            sExp = 0;

            for (int i = 0; i < exp.Count; i++)
            {
                sExp += double.Parse(exp[i][1]);
            }
            sExp += table.Sum;

            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[0].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            dataGridView1.Columns[2].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[0].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[1].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns[2].DefaultCellStyle.Format = "0.00%";

            for (int i = 0; i < exp.Count; i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Cells[0].Value = exp[i][0];
                if (exp[i][0].Contains("корм"))
                {
                    dataGridView1.Rows[i].Cells[1].Value = Math.Round(table.Sum, 2);
                    dataGridView1.Rows[i].Cells[2].Value = Math.Round(table.Sum / sExp, 2) ;
                }
                else
                {
                    dataGridView1.Rows[i].Cells[1].Value = Math.Round(double.Parse(exp[i][1]), 2);
                    dataGridView1.Rows[i].Cells[2].Value = Math.Round(double.Parse(exp[i][1]) / sExp, 2);
                }    
            }
            dataGridView1.Rows.Add();
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = "Итого затрат";
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value = Math.Round(sExp, 2);
            dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[2].Value = 1;

            textBox2.Text = table.Milk.ToString();


        }

        private void button3_Click(object sender, EventArgs e)
        {
            double proceeds = table.Milk * ((double)hScrollBar1.Value / 100 * double.Parse(textBox5.Text) +
                (double)hScrollBar2.Value / 100 * double.Parse(textBox6.Text) + (double)hScrollBar3.Value / 100 * double.Parse(textBox7.Text));
            textBox3.Text = Math.Round(proceeds, 2).ToString();
            textBox1.Text = Math.Round(proceeds - sExp, 2).ToString();
            textBox4.Text = Math.Round((proceeds - sExp) / table.Sum * 100, 2).ToString() + "%";
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form form = new Optimiz(table, ChangeData);
            Hide();
            form.ShowDialog();
            Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFile = new SaveFileDialog(); // создание диалогового окна сохранения
            saveFile.DefaultExt = "*.xlsx|*.txt|All files(*.*)|*.*";    // фильтр отображения типов файлов в диалоговом окне
            saveFile.Filter = "Excel files(*.xlsx)|*.xlsx|Text files(*.txt)|*.txt|All files(*.*)|*.*";  // выбор типов файлов для сохраниния
            if (saveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFile.FileName.Length > 0)  // открытие диалогового окна
            {
                if (saveFile.FileName.Substring(saveFile.FileName.Length - 5) == ".xlsx")   // сохрание в Excel
                {
                    if (File.Exists(saveFile.FileName)) // удаление второго файла при соглашении на замену его в диалоговом окне,
                        File.Delete(saveFile.FileName); // чтобы Microsoft Excel повторно не спрашивал про замену
                    saving.Show();  // открытие всплывающего окна сохранения
                    SaveData(saveFile.FileName);  // сохранение в Excel
                    saving.Close(); // закрытие всплывающего окна сохранения
                }
                else
                    SaveData(new StreamWriter(saveFile.FileName));    // сохранение в текстовый файл
            }
        }

        public void SaveData(StreamWriter file) // сохранение данных в текстовый файл
        {
            file.WriteLine("Характеристики молока\t%\tцена");
            file.WriteLine("Первый сорт\t" + hScrollBar1.Value + "%\t" + textBox5.Text);
            file.WriteLine("Первый сорт\t" + hScrollBar2.Value + "%\t" + textBox6.Text);
            file.WriteLine("Первый сорт\t" + hScrollBar3.Value + "%\t" + textBox7.Text);

            file.WriteLine();
            file.WriteLine("Cуточный объем удоя молока\t" + textBox2.Text + " л.");
            file.WriteLine();

            file.WriteLine("Статья затрат\tСумма\tДоля");
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    file.Write(dataGridView1.Rows[i].Cells[j].Value + "\t");
                file.WriteLine();
            }
            file.WriteLine("\n" + "Прибыль в сутки\t" + textBox1.Text);
            file.WriteLine("\n" + "Выручка в сутки\t" + textBox3.Text);
            file.WriteLine("\n" + "Рентабельность в сутки\t" + textBox4.Text);
            file.Close();
        }

        public void SaveData(string savefilename)   // (перегрузка) сохранение данных в файл Excel
        {
            // открытие Microsoft Excel
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // создание книги
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // создание листа
            Microsoft.Office.Interop.Excel._Worksheet worksheet = workbook.ActiveSheet;
            // подпись листа
            worksheet.Name = "Экономика";

            worksheet.Cells[1, 1] = "Характеристики молока"; worksheet.Cells[1, 2] = "%"; worksheet.Cells[1, 3] = "Цена";
            worksheet.Cells[2, 1] = "Первый сорт"; worksheet.Cells[2, 2] = hScrollBar1.Value + "%"; worksheet.Cells[2, 3] = double.Parse(textBox5.Text);
            worksheet.Cells[3, 1] = "Второй сорт"; worksheet.Cells[3, 2] = hScrollBar2.Value + "%"; worksheet.Cells[3, 3] = double.Parse(textBox6.Text);
            worksheet.Cells[4, 1] = "Экстра сорт"; worksheet.Cells[4, 2] = hScrollBar3.Value + "%"; worksheet.Cells[4, 3] = double.Parse(textBox7.Text);

            worksheet.Cells[6, 1] = "Суточный удой молока";
            worksheet.Cells[6, 2] = double.Parse(textBox2.Text); worksheet.Cells[6, 3] = "л.";


            // заполнение файла Excel из таблицы результатов
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[8, i] = dataGridView1.Columns[i - 1].HeaderText;    // ввод названий колонок
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)   // ввод результатов
                {
                    worksheet.Cells[i + 9, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }

            worksheet.Cells[dataGridView1.Rows.Count + 9, 1] = "Прибыль в сутки";
            worksheet.Cells[dataGridView1.Rows.Count + 9, 2] = double.Parse(textBox1.Text);
            worksheet.Cells[dataGridView1.Rows.Count + 10, 1] = "Выручка в сутки";
            worksheet.Cells[dataGridView1.Rows.Count + 10, 2] = double.Parse(textBox3.Text);
            worksheet.Cells[dataGridView1.Rows.Count + 11, 1] = "Рентабельность в сутки";
            worksheet.Cells[dataGridView1.Rows.Count + 11, 2] = double.Parse(textBox4.Text.Replace("%", ""));
            worksheet.Cells[dataGridView1.Rows.Count + 11, 3] = "%";
            // выравнивание колонок по ширине содержания
            worksheet.Columns.AutoFit();
            // сохранение файла Excel
            workbook.SaveAs(savefilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // выход их Microsoft Excel
            app.Quit();
        }

    }
}
