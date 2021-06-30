using System;
using System.IO;
using System.Windows.Forms;

namespace Optimization
{
    public partial class Optimiz : Form
    {
        private TableBase table;    // объект данных
        private double[] result,    // результаты рассчета 
            concentrate;   // концентрация по группам переменных
        private double sum, // значение целевой функции по результатам
            extra;
        private bool ChangeData;    // для проверки на изменение данных 
        private Form saving = new Saving(); // всплывающее окно сохранения данных в Excel 
        public Optimiz(TableBase table, bool ChangeData)    // конструктор
        {
            this.table = table;
            this.ChangeData = ChangeData;
            InitializeComponent();
        }

        private void Optimiz_Load(object sender, EventArgs e)   // при загрузке проверка на заполненность входных данных
        {
            numericUpDown3.DecimalPlaces = 1;
            numericUpDown4.DecimalPlaces = 1;
            /*label1.Visible = false;
            if (table.VarTable.Count == 0)
                label1.Visible = true;*/
        }

        private void button8_Click(object sender, EventArgs e) // переход в главное меню
        {
            Form form = new Menu(table, ChangeData);
            Hide();
            form.ShowDialog();
            Close();
        }

        private void button1_Click(object sender, EventArgs e)  // переход во вкладку "Входные данные"
        {
           Form form = new Data(table, ChangeData);
            Hide();
            form.ShowDialog();
            Close();
        }

        private void button2_Click(object sender, EventArgs e) // выход из программы с всплывающим окном при изменении данных
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

        private void Concentrate()  // рассчет концентрации по группам
        {
            // расчет суммы значений переменных
            double amount = 0;
            for (int i = 0; i < result.Length; i++)
                amount += result[i];
            // расчет суммы значений переменных каждой группы
            for (int i = 0; i < table.Vars.GetLength(0); i++)
                for (int j = 0; j < table.SternType.Length; j++)
                    if (table.Vars[i][1] == table.SternType[j])
                        concentrate[j] += result[i];
            // рассчет концетрации каждой группы
            for (int j = 0; j < concentrate.Length; j++)
            {
                concentrate[j] /= amount;
            }
        }

        private void button4_Click(object sender, EventArgs e)  //расчет значений и занесение из в таблицу
        {
            dataGridView1.Rows.Clear();
            
            double[,] t = table.TableBuild(int.Parse(numericUpDown1.Value.ToString()), int.Parse(numericUpDown2.Value.ToString()), double.Parse(numericUpDown3.Value.ToString()), double.Parse(numericUpDown4.Value.ToString()));

            Simplex method = new Simplex(t);
            // нахождение результатов оптимизационной задачи
            result = method.Calculate();
            for (int i = 0; i < result.Length; i++)
                result[i] *= double.Parse(textBox1.Text);
            concentrate = new double[table.Groups.Length];
            Concentrate();
            int k = 0;
            for (int i = 0; i < table.SternType.Length; i++)
            {
                dataGridView1.Rows.Add(); // добавление строки
                dataGridView1.Rows[k].Cells[0].Value = table.SternType[i]; // добавление номера группы
                dataGridView1.Columns[1].DefaultCellStyle.Format = "0.00%"; // процентный тип данных для концентрации группы
                dataGridView1.Rows[k++].Cells[1].Value = Math.Round(concentrate[i], 4); ; // вывод в таблицу концентрацию группы
                for (int j = 0; j < table.Vars.GetLength(0); j++)
                {
                    if (table.Vars[j][1] == table.SternType[i])
                    {
                        dataGridView1.Rows.Add();   // добавление строки
                        dataGridView1.Rows[k].Cells[2].Value = table.Vars[j][0];
                        dataGridView1.Rows[k].Cells[3].Value = Math.Round(result[j], 2);
                        if (double.Parse(table.Vars[j][table.Vars[j].Length - 1]) != -1)
                            if ((double.Parse(table.Vars[j][table.Vars[j].Length - 1]) - result[j]) < 0)
                                dataGridView1.Rows[k++].Cells[4].Value = -Math.Round(double.Parse(table.Vars[j][table.Vars[j].Length - 1]) - result[j], 2);
                            else
                                dataGridView1.Rows[k++].Cells[4].Value = 0;
                        else
                            dataGridView1.Rows[k++].Cells[4].Value = 0;

                    }
                }
            }
            // расчет значения целевой функции
            sum = 0;
            extra = 0;
            for (int i = 0; i < result.Length; i++)
            {
                sum += double.Parse(table.Vars[i][table.Vars[i].Length - 2]) * result[i];
                if (result[i] != 0 && (result[i] - double.Parse(table.Vars[i][table.Vars[i].Length - 1])) > 0)
                    extra += double.Parse(table.Vars[i][table.Vars[i].Length - 2]) * (result[i] - double.Parse(table.Vars[i][table.Vars[i].Length - 1]));
            }
                
            // вывод целевой функции
            textBox2.Text = Math.Round(sum, 2).ToString();
            textBox3.Text = Math.Round(extra, 2).ToString();
            table.Result = result;
            table.Sum = sum;
            table.Milk = int.Parse(numericUpDown2.Value.ToString()) * int.Parse(textBox1.Text);
        }

        private void button3_Click(object sender, EventArgs e)  // сохранение данных в файл
        {
            if (dataGridView1.Rows.Count > 1) // проверка заполненности таблицы
            {
                SaveFileDialog saveFile = new SaveFileDialog(); // создание диалогового окна сохранения
                saveFile.DefaultExt = "*.xlsx|*.txt|All files(*.*)|*.*"; // фильтр отображения типов файлов в диалоговом окне
                saveFile.Filter = "Excel files(*.xlsx)|*.xlsx|Text files(*.txt)|*.txt|All files(*.*)|*.*"; // выбор типов файлов для сохраниния
                if (saveFile.ShowDialog() == System.Windows.Forms.DialogResult.OK && saveFile.FileName.Length > 0) // открытие диалогового окна
                {
                    if (saveFile.FileName.Substring(saveFile.FileName.Length - 5) == ".xlsx") // сохрание в Excel
                    {
                        if (File.Exists(saveFile.FileName)) // удаление второго файла при соглашении на замену его в диалоговом окне,
                            File.Delete(saveFile.FileName); // чтобы Microsoft Excel повторно не спрашивал про замену
                        saving.Show(); // открытие всплывающего окна сохранения
                        SaveResult(saveFile.FileName); // сохранение в Excel
                        saving.Close(); // закрытие всплывающего окна сохранения
                    }
                    else
                        SaveResult(new StreamWriter(saveFile.FileName)); // сохранение в текстовый файл
                }
            }
        }

        private void SaveResult(StreamWriter file)  // сохранение результатов в текстовый файл
        {
            // построчное заполнение текстового файла
            file.WriteLine("Группа\t\tКонцентрация\t\tКорм\t\tЗначение\t\tНе хватает");
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView1.Rows[i].Cells.Count; j++)
                    file.Write(dataGridView1.Rows[i].Cells[j].Value + "\t\t");
                file.WriteLine();
            } 
            file.WriteLine("\n" + "Стоимость рациона = " + Math.Round(sum, 2));
            file.WriteLine("\n" + "Стоимость закупки недостающего корма = " + Math.Round(sum, 2));
            file.Close();
        }     
        
        private void button14_Click(object sender, EventArgs e) // кнопка свернуть окно
        {
            WindowState = FormWindowState.Minimized;
        }

        private void numericUpDown3_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDown4.Value < (double)numericUpDown3.Value / 1.5)
                numericUpDown4.Value = (decimal)((double)numericUpDown3.Value / 1.5);
            if ((double)numericUpDown4.Value > (double)numericUpDown3.Value / 1.1)
                numericUpDown4.Value = (decimal)((double)numericUpDown3.Value / 1.1);
        }

        private void numericUpDown4_ValueChanged(object sender, EventArgs e)
        {
            if ((double)numericUpDown4.Value < (double)numericUpDown3.Value / 1.5)
                numericUpDown3.Value = (decimal)((double)numericUpDown4.Value * 1.5);
            if ((double)numericUpDown4.Value > (double)numericUpDown3.Value / 1.1)
                numericUpDown3.Value = (decimal)((double)numericUpDown4.Value * 1.1);
        }

        private void SaveResult(string savefilename)    // (перегрузка) сохранение результатов в файл Excel
        {
            // открытие Microsoft Excel
            Microsoft.Office.Interop.Excel._Application app = new Microsoft.Office.Interop.Excel.Application();
            // создание книги
            Microsoft.Office.Interop.Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
            // создание листа
            Microsoft.Office.Interop.Excel._Worksheet worksheet = workbook.ActiveSheet; 
            // подпись листа
            worksheet.Name = "Результаты";
            // заполнение файла Excel из таблицы результатов
            for (int i = 1; i < dataGridView1.Columns.Count + 1; i++)
            {
                worksheet.Cells[1, i] = dataGridView1.Columns[i - 1].HeaderText;    // ввод названий колонок
            }
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                worksheet.Range["B" + (i + 2)].NumberFormat = "0.00%";  // ввод концентрации групп в процентном типе
                for (int j = 0; j < dataGridView1.Columns.Count; j++)   // ввод результатов
                {
                    worksheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value;
                }
            }
            // ввод целевой функции
            worksheet.Cells[dataGridView1.Rows.Count + 2, 3] = "Стоимость рациона";
            worksheet.Cells[dataGridView1.Rows.Count + 2, 4] = double.Parse(textBox2.Text);
            worksheet.Cells[dataGridView1.Rows.Count + 3, 3] = "Стоимость закупки недостающего корма";
            worksheet.Cells[dataGridView1.Rows.Count + 3, 4] = double.Parse(textBox3.Text);
            // выравнивание колонок по ширине содержания
            worksheet.Columns.AutoFit();
            // сохранение файла Excel
            workbook.SaveAs(savefilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            // выход их Microsoft Excel
            app.Quit();
        }
    }
}
