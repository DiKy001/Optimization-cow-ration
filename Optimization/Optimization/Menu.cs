using System;
using System.IO;
using System.Windows.Forms;

namespace Optimization
{
    public partial class Menu : Form
    {
        private TableBase table;    // объект данных
        private bool ChangeData = false;    // для проверки на изменение данных

        public Menu(TableBase table, bool ChangeData)   // конструктор
        {
            this.table = table;
            this.ChangeData = ChangeData;
            InitializeComponent();
        }

        private void button3_Click(object sender, EventArgs e)  // выход из программы с всплывающим окном при изменении данных
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

        private void button2_Click(object sender, EventArgs e)  // переход во вкладку "Входные данные"
        {
            Form form = new Data(table, ChangeData);
            Hide();
            form.ShowDialog();
            Close();
        }

        private void button1_Click(object sender, EventArgs e) // переход во вкладку "Расчёт значений"
        {
            Form form = new Optimiz(table, ChangeData);
            Hide();
            form.ShowDialog();
            Close();
        }

        private void button14_Click(object sender, EventArgs e) // кнопка свернуть окно
        {
            WindowState = FormWindowState.Minimized;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form form = new AboutBox1();
            form.ShowDialog();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (table.Result != null)
            {
                Form form = new Balance(table, ChangeData);
                Hide();
                form.ShowDialog();
                Close();
            }
            else
            {
                MessageBox.Show("Сначала необходимо рассчитать рацион", "Сообщение");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (table.Result != null)
            {
                Form form = new Economy(table, ChangeData);
                Hide();
                form.ShowDialog();
                Close();
            }
            else
            {
                MessageBox.Show("Сначала необходимо рассчитать рацион", "Сообщение");
            }
        }
    }
}
