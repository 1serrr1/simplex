using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
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
using Excel = Microsoft.Office.Interop.Excel;

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
                var objectiveFunction = ParseInput(ObjectiveFunctionTextBox.Text);
                var constraints = ParseMatrix(ConstraintsTextBox.Text);
                var rightSide = ParseInput(RightSideTextBox.Text);

                var result = SolveSimplex(objectiveFunction, constraints, rightSide);

                // Форматируем результат для отображения
                StringBuilder resultText = new StringBuilder();
                resultText.AppendLine("Оптимальное решение:");
                for (int i = 0; i < result.Item1.Length; i++)
                {
                    resultText.AppendLine($"x{i + 1} = {result.Item1[i]}");
                }
                resultText.AppendLine($"Значение целевой функции: {result.Item2}");

                ResultTextBlock.Text = resultText.ToString();

                // Сохраняем результаты для экспорта
                App.Current.Properties["SimplexResult"] = result;
                App.Current.Properties["ObjectiveFunction"] = objectiveFunction;

                MessageBox.Show("Решение найдено! Перейдите во вкладку 'Результат' для просмотра.",
                              "Готово", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка: {ex.Message}");
            }
        }
        // "6, 5, 7"
        // "2, 1, 3; 6, 0, 3; 5, 1, 0; 1, 4, 2; 3, 3, 0"
        // "60, 80, 80, 50, 56"
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

        private void ExportToTxtButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(App.Current.Properties["SimplexResult"] is Tuple<double[], double> result))
                {
                    MessageBox.Show("Нет данных для экспорта. Сначала решите задачу.");
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Текстовые файлы (*.txt)|*.txt",
                    DefaultExt = "txt",
                    FileName = "SimplexResult_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    using (StreamWriter writer = new StreamWriter(saveFileDialog.FileName))
                    {
                        writer.WriteLine("Результаты решения задачи симплекс-методом");
                        writer.WriteLine("-------------------------------------------");
                        writer.WriteLine();

                        writer.WriteLine("Целевая функция:");
                        writer.WriteLine(string.Join(", ", (int[])App.Current.Properties["ObjectiveFunction"]));
                        writer.WriteLine();

                        writer.WriteLine("Оптимальное решение:");
                        for (int i = 0; i < result.Item1.Length; i++)
                        {
                            writer.WriteLine($"x{i + 1} = {result.Item1[i]}");
                        }
                        writer.WriteLine();

                        writer.WriteLine($"Значение целевой функции: {result.Item2}");
                    }

                    MessageBox.Show("Результаты успешно экспортированы в текстовый файл!",
                                  "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в TXT: {ex.Message}");
            }
        }

        private void ExportToExcelButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!(App.Current.Properties["SimplexResult"] is Tuple<double[], double> result))
                {
                    MessageBox.Show("Нет данных для экспорта. Сначала решите задачу.");
                    return;
                }

                SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                    Filter = "Excel файлы (*.xlsx)|*.xlsx",
                    DefaultExt = "xlsx",
                    FileName = "SimplexResult_" + DateTime.Now.ToString("yyyyMMdd_HHmmss")
                };

                if (saveFileDialog.ShowDialog() == true)
                {
                    Excel.Application excelApp = new Excel.Application();
                    Excel.Workbook workbook = excelApp.Workbooks.Add();
                    Excel.Worksheet worksheet = workbook.ActiveSheet;

                    try
                    {
                        // Заголовок
                        worksheet.Cells[1, 1] = "Результаты решения задачи симплекс-методом";
                        Excel.Range header = worksheet.Range["A1"];
                        header.Font.Bold = true;
                        header.Font.Size = 14;

                        // Целевая функция
                        worksheet.Cells[3, 1] = "Целевая функция:";
                        worksheet.Cells[3, 2] = string.Join(", ", (int[])App.Current.Properties["ObjectiveFunction"]);

                        // Оптимальное решение
                        worksheet.Cells[5, 1] = "Оптимальное решение:";
                        for (int i = 0; i < result.Item1.Length; i++)
                        {
                            worksheet.Cells[6 + i, 1] = $"x{i + 1}";
                            worksheet.Cells[6 + i, 2] = result.Item1[i];
                        }

                        // Значение целевой функции
                        worksheet.Cells[6 + result.Item1.Length, 1] = "Значение целевой функции:";
                        worksheet.Cells[6 + result.Item1.Length, 2] = result.Item2;

                        // Форматирование
                        worksheet.Columns.AutoFit();
                        workbook.SaveAs(saveFileDialog.FileName);

                        MessageBox.Show("Результаты успешно экспортированы в Excel файл!",
                                      "Успех", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    finally
                    {
                        workbook.Close();
                        excelApp.Quit();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Ошибка при экспорте в Excel: {ex.Message}");
            }
        }
    }
}