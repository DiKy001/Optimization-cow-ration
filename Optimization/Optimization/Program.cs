using System;
using System.IO;
using System.Windows.Forms;
namespace Optimization
{
    class Program
    {
        [STAThread]
        static void Main(string[] args) // создание объекта данных при загрузке программы
        {
            TableBase table;    // создание объекта данных
            if (File.Exists("Data\\Stern.txt") && File.Exists("Data\\Norms.txt")) // проверка существование программного файла данных
            {
                table = new TableBase(true); // при существовании, данные берутся из файла
            }
            else
                table = new TableBase(); // при осутствии файла, создается пустой объект
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Menu(table, false)); 
        }
    }
}
