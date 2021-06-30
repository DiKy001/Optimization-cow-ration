using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;

namespace Optimization
{
    public class TableBase
    {
        private List<List<double>> vartable = new List<List<double>>(); // таблица набора значений для переменных
        private List<string[]> vars = new List<string[]>(); // переменные
        private List<double[]> norms = new List<double[]>(); // нормы
        private string[] limitsName;   // массив ограничений
        private double[] groups;   // разбиение переменных на группы в процентном соотношении
        private string[] sternType;      // типы кормов
        private List<double> limitsamount = new List<double>();  // массив ограничений размеров переменных
        private int filelenth = 22;     // максимальная длина файла загрузки данных Excel
        private double bottomError = 0.5;
        private double[] result;
        private double sum;
        private int mLit;

        public int Milk
        {
            get { return mLit; }
            set { mLit = value; }
        }

        public double Sum
        {
            get { return sum; }
            set { sum = value; }
        }

        public double[] Result
        {
            get { return result; }
            set { result = value; }
        }
        public List<List<double>> VarTable    // возвращение массива значений для переменных
        {
            get { return vartable; }
        }
        public string[][] Vars    // возвращение массива имен переменных
        {
            get { return vars.ToArray(); }
        }

        public double[][] Norms    // возвращение массива имен переменных
        {
            get { return norms.ToArray(); }
        }
        public string[] Limints    // возвращение массива ограничений
        {
            get { return limitsName; }
        }
       /* public List<double> LimintsList    // возвращение массива ограничений
        {
            get { return limits; }
        }*/
        public double[] Groups  // возврат массива групп разбиения
        {
            get {return groups; }
        }
        public string[] SternType  // возврат координат групп разбиения
        {
            get { return sternType; }
        }
        public double[] LimitsAmount    // возврат массива ограничений размеров переменных
        {
            get { return limitsamount.ToArray(); }
        }

        public TableBase()   // конструктор, при не созданном файле входных данных
        {
        }

        public TableBase(bool a) // (перегрузка) конструктор для ввода значений из програмного файла
        {
            vars = new List<string[]>(); 
            norms = new List<double[]>(); 
            StreamReader st = new StreamReader("Data\\Stern.txt");
            string read;
            while (!string.IsNullOrEmpty(read = st.ReadLine()))
            {
                string name = read;
                string type = st.ReadLine();
                string[] values = st.ReadLine().Split('\t');
                string amount = st.ReadLine().Split(' ')[1];
                vars.Add(new string[values.Length + 3]);
                int j = vars.Count - 1;
                vars[j][0] = name;
                vars[j][1] = type;
                for (int i = 2; i < vars[j].Length - 1; i++)
                    vars[j][i] = values[i - 2];
                vars[j][vars[j].Length - 1] = amount;
                st.ReadLine();
            }
            st.Close();

            StreamReader nm = new StreamReader("Data\\Norms.txt");
            limitsName = nm.ReadLine().Split('\t').ToArray();
            nm.ReadLine();
            while (!string.IsNullOrEmpty(read = nm.ReadLine()))
            {
                norms.Add(new double[] { double.Parse(read) });
                nm.ReadLine();
                while (!string.IsNullOrEmpty(read = nm.ReadLine()))
                {
                    string[] n1 = read.Split('\t');
                    double[] n = new double[n1.Length];
                    for (int i = 0; i < n.Length; i++)
                        n[i] = double.Parse(n1[i]);
                    norms.Add(n);
                }
            }
            nm.Close();

            groups = new double[] { 0.18, 0.63, 0.17, 0.02};
            sternType = new string[] {"грубый", "сочный", "концентрат", "минеральная добавка"};
            result = null;
        }

         public void FileInput(StreamReader st) // ввод данных из файла Excel
        {
            vars = new List<string[]>();
            norms = new List<double[]>();
            string read;
            while (!string.IsNullOrEmpty(read = st.ReadLine()))
            {
                string name = read;
                string type = st.ReadLine();
                string[] values = st.ReadLine().Split('\t');
                string amount = st.ReadLine().Split(' ')[1];
                vars.Add(new string[values.Length + 3]);
                int j = vars.Count - 1;
                vars[j][0] = name;
                vars[j][1] = type;
                for (int i = 2; i < vars[j].Length - 1; i++)
                    vars[j][i] = values[i - 2];
                vars[j][vars[j].Length - 1] = amount;
                st.ReadLine();
            }

            limitsName = st.ReadLine().Split('\t').ToArray();
            st.ReadLine();
            while (!string.IsNullOrEmpty(read = st.ReadLine()))
            {
                norms.Add(new double[] { double.Parse(read) });
                st.ReadLine();
                while (!string.IsNullOrEmpty(read = st.ReadLine()))
                {
                    string[] n1 = read.Split('\t');
                    double[] n = new double[n1.Length];
                    for (int i = 0; i < n.Length; i++)
                        n[i] = double.Parse(n1[i]);
                    norms.Add(n);
                }
            }
            st.Close();
        }

        /*// добавление переменной без группы (помещается в первую группу)
        public void AddVar(string varname, string varvalues, string limitamount) 
        {
            varnames.Add(varname); // добавление имени переменной в список
            vartable.Add(varvalues.Split('\t').Select(double.Parse).ToList());  // добавление значений для ограничений в список
            limitsamount.Add(double.Parse(limitamount)); // добавление ограничения по размеру в список
            posgroups[posgroups.Count - 1]++; // добавление переменной в последнюю группу
        }
        // (перегрузка) добавление переменной без ограничения по размеру переменной
        public void AddVar(string varname, string varvalues, int group)
        {
            varnames.Insert(posgroups[group], varname); // добавление имени переменной в список
            vartable.Insert(posgroups[group], varvalues.Split('\t').Select(double.Parse).ToList()); // добавление значений для ограничений в список
            limitsamount.Insert(posgroups[group], -1);  // добавление ограничения по размеру в список, -1 - неограничено
            for (int i = group - 1; i < posgroups.Count; i++) // добавление переменной в группу
                posgroups[i]++;
        }
        // (перегрузка) добавление переменной без ограничения по размеру переменной и без группы (помещается в первую группу)
        public void AddVar(string varname, string varvalues)
        {
            varnames.Add(varname); // добавление имени переменной в список
            vartable.Add(varvalues.Split('\t').Select(double.Parse).ToList()); // добавление значений для ограничений в список
            limitsamount.Add(-1); // добавление ограничения по размеру в список, -1 - неограничено
            posgroups[posgroups.Count - 1]++; // добавление переменной в последнюю группу
        }
        // (перегрузка) добавление переменной со всеми значениями
        public void AddVar(string varname, string varvalues, string limitamount, int group)
        {
            varnames.Insert(posgroups[group], varname); // добавление имени переменной в список
            vartable.Insert(posgroups[group], varvalues.Split('\t').Select(double.Parse).ToList()); // добавление значений для ограничений в список
            limitsamount.Insert(posgroups[group], double.Parse(limitamount)); // добавление ограничения по размеру в список
            for (int i = group - 1; i < posgroups.Count; i++) // добавление переменной в группу
                posgroups[i]++;
        }

        // добавление ограничения
        public void AddLimit(string limit)
        {
            limits.Add(double.Parse(limit));    // добавление ограничения
            for (int i = 0; i < vartable.Count; i++)    // добавление значения по ограничению для каждой переменной со значением 0
                vartable[i].Insert(vartable[i].Count - 1, 0);
        }

        // изменить значение для переменной
        public void ChangeValue(string value, int i, int j)
        {
            vartable[i][j] = double.Parse(value);
        }

        // изменить значение ограничения
        public void ChangeLimit(string value, int i)
        {
            limits[i] = double.Parse(value);
        }

        // удаление значения ограничения
        public void DeleteLimit(int i)
        {
            limits.RemoveAt(i);
            for (int j = 0; j < vartable.Count; j++)    // удаление значения по ограничению для каждой переменной
                vartable[j].RemoveAt(i);
        }

        // удаление переменной
        public void DeleteVar(int i)
        {
            varnames.RemoveAt(i);   // удаление имени переменной
            vartable.RemoveAt(i);   // удаление значений по ограничениям для переменной
            limitsamount.RemoveAt(i);   // удаление ограничения по размеру для переменной
            for (int j = 0; j < posgroups.Count; j++)   // удаление переменной из группы
                if (posgroups[j] >= i)
                    posgroups[j]--;
        }*/

        // построение таблицы, пригодной для использования в решении оптимизационной задачи
        public double[,] TableBuild(int mass, int milk, double jir, double bel)
        {
            // При ограничениях с знаком <= знак не меняется, при >= знак меняется на противоположный
            // При условии минимизации знак меняется на противоположный, при максимизации знак остаётся прежним
            double[,] table = new double[norms[2].Length + groups.Length, vars.Count + 1];
            // занесение величин рассходов переменных
            for (int i = 0; i < vars.Count; i++)
                for (int j = 2; j < vars[i].Length - 2; j++)
                {
                    table[j - 2, i + 1] = -double.Parse(vars[i][j]);
                }
            // заполнение первого столбца ограничениями
            for (int j = 0; j < norms.Count; j++)
            {
                if (norms[j].Length == 1 && norms[j][0] == mass)
                {
                    do
                    {
                        j++;
                    } while (norms[j][0] != milk);

                    for (int i = 0; i < limitsName.Length; i++)
                    {
                        table[i, 0] = -norms[j][i + 1];
                    }
                }
            }
            
            // занесение целевой функции в таблицу
            for (int j = 0; j < table.GetLength(1) - 1; j++)
                table[table.GetLength(0) - 1, j + 1] = -double.Parse(vars[j][vars[j].Length - 2]);
            // занесение ограничений по концентрации групп
           
            double[] ngroups = new double[groups.Length];
            ngroups[0] = groups[0] * jir / 3.8;
            for (int i = 1; i < ngroups.Length; i++)
                ngroups[i] = groups[i] + (ngroups[0] - groups[0]) / 3;
            double gr = ngroups[2];
            ngroups[2] *= bel / 3.1;
            for (int i = 0; i < ngroups.Length; i++)
                if (i != 2)
                    ngroups[i] += (ngroups[2] - gr) / 3;

            if (jir > 3.8 && bel > 3.1)
            {
                for (int j = table.GetLength(0) - 5; j < table.GetLength(0); j++)
                    table[j, 0] = 0;
                for (int j = 0; j < vars.Count; j++)
                    if (vars[j][1].Contains("грубый"))
                        table[table.GetLength(0) - 5, j + 1] = 1 - ngroups[0];
                    else
                        table[table.GetLength(0) - 5, j + 1] = -ngroups[0];
                for (int j = 0; j < vars.Count; j++)
                    if (vars[j][1].Contains("сочный"))
                        table[table.GetLength(0) - 4, j + 1] = 1 - ngroups[1];
                    else
                        table[table.GetLength(0) - 4, j + 1] = -ngroups[1];
                for (int j = 0; j < vars.Count; j++)
                    if (vars[j][1].Contains("концентрат"))
                        table[table.GetLength(0) - 3, j + 1] = 1 - ngroups[2];
                    else
                        table[table.GetLength(0) - 3, j + 1] = -ngroups[2];
                for (int j = 0; j < vars.Count; j++)
                    if (vars[j][1].Contains("минеральная добавка"))
                        table[table.GetLength(0) - 2, j + 1] = 1 - ngroups[3];
                    else
                        table[table.GetLength(0) - 2, j + 1] = -ngroups[3];
            }
            else
            {
                for (int j = table.GetLength(0) - 5; j < table.GetLength(0); j++)
                    table[j, 0] = 0;
                for (int j = 0; j < vars.Count; j++)
                    if (vars[j][1].Contains("грубый"))
                        table[table.GetLength(0) - 5, j + 1] = -1 + ngroups[0];
                    else
                        table[table.GetLength(0) - 5, j + 1] = ngroups[0];
                for (int j = 0; j < vars.Count; j++)
                    if (vars[j][1].Contains("сочный"))
                        table[table.GetLength(0) - 4, j + 1] = -1 + ngroups[1];
                    else
                        table[table.GetLength(0) - 4, j + 1] = ngroups[1];
                for (int j = 0; j < vars.Count; j++)
                    if (vars[j][1].Contains("концентрат"))
                        table[table.GetLength(0) - 3, j + 1] = -1 + ngroups[2];
                    else
                        table[table.GetLength(0) - 3, j + 1] = ngroups[2];
                for (int j = 0; j < vars.Count; j++)
                    if (vars[j][1].Contains("минеральная добавка"))
                        table[table.GetLength(0) - 2, j + 1] = -1 + ngroups[3];
                    else
                        table[table.GetLength(0) - 2, j + 1] = ngroups[3];
            }
            return table;
        }
        
        public void SaveData() // сохранение данных в текстовый файл
        {
            StreamWriter st = new StreamWriter(File.Create("Data\\Stern.txt"));
            // запись ограничений в файл
            for (int i = 0; i < vars.Count; i++)
            {
                st.WriteLine(vars[i][0]);
                st.WriteLine(vars[i][1]);
                for (int j = 2; j < vars[i].Length - 2; j++)
                    st.Write(vars[i][j] + "\t");
                st.WriteLine(vars[i][vars[i].Length - 2]);
                st.WriteLine("Количество " + vars[i][vars[i].Length - 1]);
                st.WriteLine();
            }
            st.Close();

            StreamWriter nr = new StreamWriter(File.Create("Data\\Norms.txt"));

            for (int i = 0; i < limitsName.Length - 1; i++)
                nr.Write(limitsName[i] + "\t");
            nr.WriteLine(limitsName[limitsName.Length - 1]);
            for (int i = 0; i < norms.Count; i++)
            {
                if (norms[i].Length == 1)
                {
                    nr.WriteLine();
                    nr.WriteLine(norms[i][0]);
                    nr.WriteLine();
                }
                else
                {
                    for (int j = 0; j < norms[i].Length - 1; j++)
                        nr.Write(norms[i][j] + "\t");
                    nr.WriteLine(norms[i][norms[i].Length - 1]);
                }
            }
            nr.Close();
        }

        public void ChangeData(string[,] data)
        {
            for (int i = 0; i < vars.Count; i++)
            {
                vars[i][0] = data[i, 0];
                vars[i][1] = data[i, 1];
                vars[i][vars[i].Length - 2] = data[i, 2];
                vars[i][vars[i].Length - 1] = data[i, 3];
            }
        }

        public void SaveData(StreamWriter st) // сохранение данных в текстовый файл
        {
            // запись ограничений в файл
            for (int i = 0; i < vars.Count; i++)
            {
                st.WriteLine(vars[i][0]);
                st.WriteLine(vars[i][1]);
                for (int j = 2; j < vars[i].Length - 2; j++)
                    st.Write(vars[i][j] + "\t");
                st.WriteLine(vars[i][vars[i].Length - 2]);
                st.WriteLine("Количество " + vars[i][vars[i].Length - 1]);
                st.WriteLine();
            }

            st.WriteLine();
            for (int i = 0; i < limitsName.Length - 1; i++)
                st.Write(limitsName[i] + "\t");
            st.WriteLine(limitsName[limitsName.Length - 1]);
            for (int i = 0; i < norms.Count; i++)
            {
                if (norms[i].Length == 1)
                {
                    st.WriteLine();
                    st.WriteLine(norms[i][0]);
                    st.WriteLine();
                }
                else
                {
                    for (int j = 0; j < norms[i].Length - 1; j++)
                        st.Write(norms[i][j] + "\t");
                    st.WriteLine(norms[i][norms[i].Length - 1]);
                } 
            }
            st.Close();
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
            worksheet.Name = "Данные";
            // заполнение файла Excel из атрибутов класса
            int k = 1;
            for (int i = 0; i < vars.Count; i++)
            {
                worksheet.Cells[k++, 1] = vars[i][0];
                worksheet.Cells[k++, 1] = vars[i][1];
                for (int j = 2; j < vars[i].Length - 1; j++)
                    worksheet.Cells[k, j - 1] = double.Parse(vars[i][j]);
                worksheet.Cells[++k, 1] = "Количество";
                worksheet.Cells[k, 2] = vars[i][vars[i].Length - 1];
                k += 2;
            }

            for (int i = 0; i < limitsName.Length; i++)
                worksheet.Cells[k, i + 1] = limitsName[i];
            for (int i = 0; i < norms.Count; i++)
            {
                if (norms[i].Length == 1)
                {
                    k += 2;
                    worksheet.Cells[k, 1] = norms[i][0];
                    k += 2;
                }
                else
                {
                    for (int j = 0; j < norms[i].Length; j++)
                        worksheet.Cells[k, j + 1] = norms[i][j];
                    k++;
                }
            }
            worksheet.Columns.AutoFit(); // выравнивание колонок по ширине содержания
            // сохранение файла Excel
            workbook.SaveAs(savefilename, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            app.Quit(); // выход их Microsoft Excel
        }
    }
}
