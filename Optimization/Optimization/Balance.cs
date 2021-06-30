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
    public partial class Balance : Form
    {
        private TableBase table;    // объект данных
        private bool ChangeData;    // для проверки на изменение данных 
        private Form saving = new Saving(); // всплывающее окно сохранения данных в Excel 

        public Balance(TableBase table, bool ChangeData)
        {
            this.table = table;
            this.ChangeData = ChangeData;
            InitializeComponent();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }

        private void Balance_Load(object sender, EventArgs e)
        {
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[0].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            dataGridView1.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

            int k = 0, days = -1, d;

            for (int i = 0; i < table.Vars.Length; i++)
                if (table.Result[i] != 0)
                {
                    dataGridView1.Rows.Add();
                    dataGridView1.Rows[k].Cells[0].Value = table.Vars[i][0];
                    d = (int)(double.Parse(table.Vars[i][table.Vars[i].Length - 1]) / table.Result[i]);
                    dataGridView1.Rows[k++].Cells[1].Value = d;
                    if (days < 0 || d < days)
                        days = d;
                }
            textBox1.Text = days.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form form = new Menu(table, ChangeData);
            Hide();
            form.ShowDialog();
            Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (ChangeData) // проверка на изменение данных
            {
                // создание диалогового окна выхода
                DialogResult dialogresult = MessageBox.Show("Данные были изменены.\nВы хотите их сохранить перед выходом?", "Выход", MessageBoxButtons.YesNoCancel);
                if (dialogresult == DialogResult.Yes) // при нажатии на кнопку "Да" в диалоговом окне
                {
                    table.SaveData(); // сохранение данных в программный файл
                    Close(); // закрытие программы
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
            // построчное заполнение текстового файла
            file.WriteLine("Вид корма\tКоличество дней");
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    file.Write(dataGridView1.Rows[i].Cells[j].Value + "\t");
                file.WriteLine();
            }
            file.WriteLine("\n" + "Корма будет достаточно на " + textBox1.Text + " дн.");
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
            worksheet.Name = "Баланс";
            // заполнение файла Excel из таблицы результатов
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;    // ввод названий колонок
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Columns.Count; j++)   // ввод результатов
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            // ввод целевой функции
            worksheet.Cells[dataGridView1.Rows.Count + 2, 1] = "Корма будет достаточно на";
            worksheet.Cells[dataGridView1.Rows.Count + 2, 2] = double.Parse(textBox1.Text);
            // выравнивание колонок по ширине содержания
            worksheet.Columns.AutoFit();
            // сохранение файла Excel
            workbook.SaveAs(savefilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // выход их Microsoft Excel
            app.Quit();
        }
    }
}
