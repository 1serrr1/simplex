using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace simplex
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
     
         private void OnSolveClicked(object sender, RoutedEventArgs e)
        {
            try
            {
                // Ввод целевой функции
                var objectiveFunction = ParseInput(ObjectiveFunctionTextBox.Text); // "6, 5, 7"

                // Ввод ограничений
                var constraints = ParseMatrix(ConstraintsTextBox.Text); // "2, 1, 3; 6, 0, 3; 5, 1, 0; 1, 4, 2; 3, 3, 0"

                // Ввод правой части (фонды)
                var rightSide = ParseInput(RightSideTextBox.Text); // "60, 80, 80, 50, 56"

                // Решение задачи с помощью симплексного метода
                var result = SolveSimplex(objectiveFunction, constraints, rightSide);

                // Выводим результат на экран
                ResultTextBlock.Text = $"Оптимальное решение: {string.Join(", ", result.Item1)}\nЦелевая функция: {result.Item2}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }

        // Парсинг строки в массив чисел
        private int[] ParseInput(string input)
        {
            return input.Split(',').Select(int.Parse).ToArray();
        }

        // Парсинг строки в двумерный массив (матрица ограничений)
        private int[,] ParseMatrix(string input)
        {
            var rows = input.Split(';');
            var matrix = new int[rows.Length, rows[0].Split(',').Length];

            for (int i = 0; i < rows.Length; i++)
            {
                var cols = rows[i].Split(',');
                for (int j = 0; j < cols.Length; j++)
                {
                    matrix[i, j] = int.Parse(cols[j]);
                }
            }

            return matrix;
        }

        // Симплексный метод
        private Tuple<double[], double> SolveSimplex(int[] objectiveFunction, int[,] constraints, int[] rightSide)
        {
            int m = constraints.GetLength(0); // Количество ограничений
            int n = constraints.GetLength(1); // Количество переменных

            double[,] simplexTable = new double[m + 1, n + m + 1];

            // Заполняем симплексную таблицу:
            for (int j = 0; j < n; j++)
            {
                simplexTable[m, j] = objectiveFunction[j];
            }

            // Ограничения
            for (int i = 0; i < m; i++)
            {
                for (int j = 0; j < n; j++)
                {
                    simplexTable[i, j] = constraints[i, j];
                }
                simplexTable[i, n + i] = 1;
                simplexTable[i, n + m] = rightSide[i];
            }

            while (true)
            {
                // Шаг 1: Проверяем на положительные элементы в последней строке
                int pivotColumn = -1;
                for (int j = 0; j < n + m; j++)
                {
                    if (simplexTable[m, j] > 0)  // Ищем положительные элементы
                    {
                        pivotColumn = j;
                        break;
                    }
                }

                // Если все элементы не положительные, мы нашли оптимум
                if (pivotColumn == -1)
                {
                    break;
                }

                // Шаг 2: Находим строку для обмена (пивот)
                int pivotRow = -1;
                double minRatio = double.MaxValue;
                for (int i = 0; i < m; i++)
                {
                    if (simplexTable[i, pivotColumn] > 0)
                    {
                        double ratio = simplexTable[i, n + m] / simplexTable[i, pivotColumn];
                        if (ratio < minRatio)
                        {
                            minRatio = ratio;
                            pivotRow = i;
                        }
                    }
                }

                // Шаг 3: Обновляем симплексную таблицу
                double pivotValue = simplexTable[pivotRow, pivotColumn];
                for (int j = 0; j < n + m + 1; j++)
                {
                    simplexTable[pivotRow, j] /= pivotValue;
                }

                // Обновляем другие строки
                for (int i = 0; i < m + 1; i++)
                {
                    if (i != pivotRow)
                    {
                        double factor = simplexTable[i, pivotColumn];
                        for (int j = 0; j < n + m + 1; j++)
                        {
                            simplexTable[i, j] -= factor * simplexTable[pivotRow, j];
                        }
                    }
                }
            }

            // Результат: оптимальные значения переменных
            double[] solution = new double[n];
            for (int j = 0; j < n; j++)
            {
                solution[j] = 0;
                for (int i = 0; i < m; i++)
                {
                    if (simplexTable[i, j] == 1)
                    {
                        solution[j] = simplexTable[i, n + m];
                        break;
                    }
                }
            }

            // Целевая функция: если значение отрицательно, приравниваем его к нулю
            double objectiveValue = -simplexTable[m, n + m];
            if (objectiveValue < 0)
            {
                objectiveValue = 0;
            }

            // Возвращаем результат: оптимальные значения переменных и значение целевой функции
            return new Tuple<double[], double>(solution, objectiveValue);
        }
    }
}