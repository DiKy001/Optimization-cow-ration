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
    public partial class Data : Form
    {
        private string varlim;  // для значений ограничений переменных при добавлении переменной
        private TableBase table;    // объект данных 
        private bool ChangeData;    // для проверки на изменения данных
        private Form loading = new Loading();   // всплывающее окно загрузки данных из Excel
        private Form saving = new Saving(); // всплывающее окно сохранения данных в Excel
        public Data(TableBase table, bool ChangeData)   // конструктор
        {
            this.table = table;
            this.ChangeData = ChangeData;
            InitializeComponent();
        }

        private void loadData()
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[1].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[2].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[3].HeaderCell.Style.Font = new Font("Tahoma", 12, FontStyle.Bold);
            dataGridView1.Columns[0].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            dataGridView1.Columns[1].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            dataGridView1.Columns[2].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            dataGridView1.Columns[3].DefaultCellStyle.Font = new Font("Tahoma", 12, FontStyle.Regular);
            for (int i = 0; i < table.Vars.GetLength(0); i++)
            {
                dataGridView1.Rows.Add();
                dataGridView1.Rows[i].Height = 30;
                dataGridView1.Rows[i].Cells[0].Value = table.Vars[i][0];
                DataGridViewComboBoxCell cb = (DataGridViewComboBoxCell)dataGridView1.Rows[i].Cells[1];
                cb.DataSource = table.SternType;
                for (int j = 0; j < table.SternType.Length; j++)
                    if (table.SternType[j].Contains(table.Vars[i][1].Trim()))
                        cb.Value = cb.Items[j];
                dataGridView1.Rows[i].Cells[2].Value = table.Vars[i][table.Vars[i].Length - 2];
                dataGridView1.Rows[i].Cells[3].Value = table.Vars[i][table.Vars[i].Length - 1];
            }

            comboBox1.Items.Clear();
            for (int i = 0; i < table.Norms.GetLength(0); i++)
                if (table.Norms[i].Length == 1)
                    comboBox1.Items.Add(table.Norms[i][0]);
        }
        private void Data_Load(object sender, EventArgs e)
        {
            loadData();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            dataGridView2.Columns.Clear();
            for (int i = 0; i < table.Norms.GetLength(0); i++)
                if (table.Norms[i].Length == 1 && table.Norms[i][0] == double.Parse(comboBox1.SelectedItem.ToString()))
                {
                    dataGridView2.Columns.Add("Column" + 1, "12");
                    dataGridView2.Columns[0].HeaderCell.Style.Font = new Font("Tahoma", 10, FontStyle.Bold);
                    for (int k = 1; k < table.Norms[i + 1].Length; k++)
                    {
                        dataGridView2.Rows.Add();
                        dataGridView2.Rows[k - 1].HeaderCell.Value = table.Limints[k - 1];
                        dataGridView2.Rows[k - 1].HeaderCell.Style.Font = new Font("Tahoma", 10, FontStyle.Bold);
                        dataGridView2.Rows[k - 1].Cells[0].Value = table.Norms[i + 1][k];
                    }
                    int j = i + 2;
                    while (j != table.Norms.Length && table.Norms[j].Length != 1)
                    {
                        dataGridView2.Columns.Add("Column" + (j - i), (12 + (j - i - 1) * 2).ToString());
                        dataGridView2.Columns[j - i - 1].HeaderCell.Style.Font = new Font("Tahoma", 10, FontStyle.Bold);
                        for (int k = 1; k < table.Norms[j].Length; k++)
                        {
                            dataGridView2.Rows[k - 1].Cells[j - i - 1].Value = table.Norms[j][k];
                        }
                        j++;
                    }
                }

        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (ChangeData) // проверка на изменение данных
            {
                // создание диалогового окна выхода
                DialogResult dialogresult = MessageBox.Show("Данные были изменены.\nВы хотите их сохранить перед выходом?", "Выход", MessageBoxButtons.YesNoCancel);
                if (dialogresult == DialogResult.Yes)   // при нажатии на кнопку "Да" в диалоговом окне
                {
                    table.SaveData();  // сохранение данных в программный файл
                    Close();    // закрытие программы
                }
                if (dialogresult == DialogResult.No)    // при нажатии на кнопку "Нет" в диалоговом окне
                {
                    Close();    // закрытие программы
                }
            }
            // при не измененных данных
            else
                Close();    // закрытие программы
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
                    table.SaveData(saveFile.FileName);  // сохранение в Excel
                    saving.Close(); // закрытие всплывающего окна сохранения
                }
                else
                    table.SaveData(new StreamWriter(saveFile.FileName));    // сохранение в текстовый файл
            }
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
            loadData();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            string[,] data = new string[dataGridView1.Rows.Count - 1, dataGridView1.Columns.Count];
            for (int i = 0; i < data.GetLength(0); i++)
                for (int j = 0; j < data.GetLength(1); j++)
                    data[i, j] = dataGridView1.Rows[i].Cells[j].Value.ToString();
            table.ChangeData(data);
            table.SaveData();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFile = new OpenFileDialog(); //  диалогового окно
            openFile.DefaultExt = "*.txt|All files(*.*)|*.*";    // фильтры для отображения типов данных
            openFile.Filter = "Text files(*.txt)|*.txt|All files(*.*)|*.*";  // фильтры для типов данных
            if (openFile.ShowDialog() == System.Windows.Forms.DialogResult.OK && openFile.FileName.Length > 0)  // открытие диалогового окна
            {
                {
                    table.FileInput(new StreamReader(openFile.FileName));
                }
            }
            loadData();
        }

        private void button14_Click(object sender, EventArgs e)
        {
            WindowState = FormWindowState.Minimized;
        }
    }
}
