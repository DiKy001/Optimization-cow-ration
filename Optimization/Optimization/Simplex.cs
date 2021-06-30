using System.Collections.Generic;

namespace Optimization
{
    class Simplex
    {
        private double[,] tablebase; // исходная таблица
        private List<int> basis;     // список базисных переменных
        private int m, n;            // размеры симплекс таблицы
        private double[] result;    // массивы результатов


        public Simplex(double[,] tablebase)  // конструктор
        {
            this.tablebase = tablebase;
            m = tablebase.GetLength(0);
            n = tablebase.GetLength(1) + m - 1;
            result = new double[tablebase.GetLength(1) - 1];
        }

        private double[,] SimplexTable()    // преобразование исходных данных в симплекс-таблицу
        {
            double[,] table = new double[m, tablebase.GetLength(1) + m - 1];
            for (int i = 0; i < m; i++)
            {
                // вся базовая таблица переносится, а остальные значения заполняются нулями
                for (int j = 0; j < table.GetLength(1); j++)
                {
                    if (j < tablebase.GetLength(1))
                        table[i, j] = tablebase[i, j];
                    else
                        table[i, j] = 0;
                }
                //выставляем коэффициент 1 перед базисной переменной в строке
                if ((tablebase.GetLength(1) + i) < table.GetLength(1))
                    table[i, tablebase.GetLength(1) + i] = 1;
            }
            return table;
        }

        public double[] Calculate() // выполнение оптимизационной задачи
        {
            // расчет первичных результатов для каждой группы
            double[,] table = SimplexTable();
            result = TableCalculate(table);
            return result;
        }

        // выполнение математического рассчета
        private double[] TableCalculate(double[,] table)
        {
            // создание списка базисных переменных
            basis = new List<int>();
            for (int i = 0; i < m; i++)
                for (int j = 0; j < n; j++)
                    if ((tablebase.GetLength(0) + i) < n)
                        basis.Add(n + i);
            double[] result = new double[tablebase.GetLength(1) - 1]; //массив результатов
            int mainCol, mainRow; // ведущие столбец и строка

            while (Continue(table))
            {
                mainRow = FindMainRow(table); // выбор исключаемой переменной (строка)
                mainCol = FindMainCol(table, mainRow);  // выбор включаемой переменной (столбец)
                // замена базисной переменной
                basis[mainRow] = mainCol;
                // преобразование таблицы по математическим прицыпам элементарных преобразований матрицы
                double[,] newtable = new double[m, n];
                for (int j = 0; j < n; j++)
                    newtable[mainRow, j] = table[mainRow, j] / table[mainRow, mainCol];

                for (int i = 0; i < m; i++)
                {
                    if (i == mainRow)
                        continue;

                    for (int j = 0; j < n; j++)
                        newtable[i, j] = table[i, j] - table[i, mainCol] * newtable[mainRow, j];
                }

                table = newtable;
            }

            //заносим в result найденные значения X
            for (int i = 0; i < result.Length; i++)
            {
                int k = basis.IndexOf(i + 1);
                if (k != -1)
                    result[i] = table[k, 0];
                else
                    result[i] = 0;
            }
            return result;
        }

        // Условие продолжения преобразования симплекс-таблицы
        private bool Continue(double[,] table)
        {
            // в первом столбце должно присутствовать отрицательное значение
            for (int i = 1; i < m; i++)
                if (table[i, 0] < 0)
                    // и в данном строке должно присутствовать отрицательное значение
                    for (int j = 1; j < n; j++)
                        if (table[i, j] < 0)
                            return true;
            // отсутствие отрицательных значений в первом столбце свидетельствует о успершном завершении расчетов
            // отсутствие отрицательных значений в проверяемой строке (в первом столбце присутствуют отрицательные значения)
            // свидетельствует о невозможности решения задачи
            return false;
        }

        // Нахождение исключаемой переменной
        private int FindMainRow(double[,] table)
        {
            int mainRow = 0;
            // выбирается наибольшая по абсолютной величине отрицательная базисная переменная
            for (int i = 1; i < m - 1; i++)
                if (table[i, 0] < table[mainRow, 0])
                    for (int j = 1; j < n; j++)
                        if (table[i, j] < 0)
                            mainRow = i;
            return mainRow;
        }

        // Нахождение включаемой переменной
        private int FindMainCol(double[,] table, int mainRow)
        {
            int mainCol = 1;
            // выбирается наименьшее из отношений отрицательных коэфициентов выбранной базисной переменной  
            // с соответстующими ей отрицательных коэфициентов целовой функции
            for (int j = 1; j < n; j++)
                if (table[mainRow, j] < 0)
                {
                    mainCol = j;
                    break;
                }
            double value = table[m - 1, mainCol] / table[mainRow, mainCol];

            for (int j = mainCol + 1; j < n; j++)
                if ((table[mainRow, j] < 0) && ((table[m - 1, j] / table[mainRow, j]) < value))
                {
                    mainCol = j;
                    value = table[m - 1, j] / table[mainRow, j];
                }

            return mainCol;
        }
    }
}
