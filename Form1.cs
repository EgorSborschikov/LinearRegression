using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace AttestetionWork
{
    public partial class Form1 : Form
    {
        private double[] dataX;
        private double[] dataY;
        private string series_name_data_default = "Данные";
        private string series_name_data;

        private double[] dataXe;

        private double[] data_line;
        private double coef_line_a, coef_line_b;

        private double stepX;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            int RangeCb, RangeCe, RangeRb, RangeRe;//Диапазоны оси Абцисс
            int RangeDCb, RangeDCe, RangeDRb, RangeDRe;//Диапазоны оси Ординат
            try
            {
                RangeCb = Convert.ToInt32(textBox1.Text);
                RangeCe = Convert.ToInt32(textBox5.Text);
                RangeRb = Convert.ToInt32(textBox3.Text);
                RangeRe = Convert.ToInt32(textBox7.Text);
                RangeDCb = Convert.ToInt32(textBox2.Text);
                RangeDCe = Convert.ToInt32(textBox6.Text);
                RangeDRb = Convert.ToInt32(textBox4.Text);
                RangeDRe = Convert.ToInt32(textBox8.Text);
            }
            catch
            {
                MessageBox.Show("Проблема с текстом в полях диапазона, должны быть числа", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //Проверяем диапазн, истина == представление данных строкой иначе столбцом
            bool RangeByRows = RangeCb != RangeCe;
            bool RangeDByRows = RangeDCb != RangeDCe; //Оси Ординат
            int CountX, CountY;

            if (RangeByRows && RangeRb != RangeRe)
            {
                MessageBox.Show("Проблема с диапазоном ячеек Оси Абцисс!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //Вычисление числа элементов X
            if (RangeByRows)
                CountX = RangeCe - RangeCb + 1;
            else
                CountX = RangeRe - RangeRb + 1;

            //Вычисление числа элементов Y
            if (RangeDByRows)
                CountY = RangeDCe - RangeDCb + 1;
            else
                CountY = RangeDRe - RangeDRb + 1;

            if (CountX != CountY)
            {
                MessageBox.Show("Число элементов Оси Абцисс и оси Ординат не сошлось", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            dataX = new double[CountX];
            dataY = new double[CountY];

            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            var result = openDialog.ShowDialog();
            if (result != DialogResult.OK)
            {
                MessageBox.Show("Файл не выбран!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            string fileName = System.IO.Path.GetFileName(openDialog.FileName);

            Microsoft.Office.Interop.Excel.Application ExcelApp;
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorksheet;

            try
            {
                ExcelApp = new Microsoft.Office.Interop.Excel.Application();
                ExcelWorkbook = ExcelApp.Workbooks.Open(openDialog.FileName);
                ExcelWorksheet = ExcelWorkbook.Sheets[numericUpDown1.Value];
            }
            catch (Exception exe)
            {
                /*
                 * Возможная утечка ресурсов
                 */
                MessageBox.Show(exe.ToString(), "Возникла ошибка при открытии файла Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            
            try
            {
                //Загрузка оси Х
                if (RangeByRows)//если данные хранятся в файле по строкам
                {
                    int p = 0;
                    for (int i = RangeCb - 1; i < RangeCe; i++)
                    {
                        int j = RangeRb;
                        dataX[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j, i + 1].Text.ToString());
                    }
                }
                else//если данные хранятся в файле по столбцам
                {
                    int p = 0;
                    for (int j = RangeRb - 1; j < RangeRe; j++)
                    {
                        int i = RangeCb;
                        dataX[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j + 1, i].Text.ToString());
                    }
                }

                //Загрузка Y
                if (RangeDByRows)//если данные хранятся в файле в виде строки
                {
                    int p = 0;
                    for (int i = RangeDCb - 1; i < RangeDCe; i++)
                    {
                        int j = RangeDRb;
                        dataY[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j, i + 1].Text.ToString());
                    }
                    if (checkBox1.Checked && RangeDCb > 1)//Загрузка подписи данных
                        series_name_data = ExcelWorksheet.Cells[RangeDRb, RangeDCb - 1].Text.ToString();
                    else
                        series_name_data = series_name_data_default;
                }
                else
                {
                    int p = 0;
                    for (int j = RangeDRb - 1; j < RangeDRe; j++)
                    {
                        int i = RangeDCb;
                        dataY[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j + 1, i].Text.ToString());
                    }
                    if (checkBox1.Checked && RangeDRb > 1)//Загрузка подписи данных
                        series_name_data = ExcelWorksheet.Cells[RangeDRb - 1, RangeDCb].Text.ToString();
                    else
                        series_name_data = series_name_data_default;
                }
            }
            catch
            {
                MessageBox.Show("Ошибка в файле Excel либо неверный диаппазон");
            }
            ExcelWorkbook.Close(false, Type.Missing, Type.Missing); // закрыть файл не сохраняя
            ExcelApp.Quit(); // Закрыть экземпляр Excel
            GC.Collect();   //Инициировать сборщик мусора

            //Вычисление среднего шага оси Х
            stepX = (dataX.Last() - dataX.First()) / (dataX.Length - 1);

            //Вычисления, построение графиков
            calculate();
            set_chart_series();
            ToggleLinearRegressionLine(true);// Проверка выводимых линий трендов

            update_labels_coef();

            update_predict_table();
            
        }

        private void numericUpDownPredictNum_ValueChanged(object sender, EventArgs e)
        {
            //Вычисления и построение графиков
            calculate();
            set_chart_series();
        }

        private void update_labels_coef()
        {
            label10.Text = "y = " + coef_line_a.ToString("N") + " + " + coef_line_b.ToString("N") + " * x";//вид регрессии
            
            //параметры линейной функции
            label11.Text = (coef_line_b * Math.Sqrt(sigma2(dataX)) / Math.Sqrt(sigma2(dataY))).ToString("N");//расчёт Rxy
            label12.Text = (approx_error(dataY, data_line)).ToString("N") + " %";// расчёт A
            label13.Text = (1.0 - sigma2(data_line) / sigma2(dataY)).ToString("N");// расчёт R2

            //Коэффициент корреляции
            label14.Text = correlation(dataX, dataY).ToString("N");
        }


        private void update_predict_table()
        {       
            System.Data.DataTable dt = new System.Data.DataTable();
            dt.Columns.Add("сдвиг");
            dt.Columns.Add("Год");
            dt.Columns.Add("Линейная функция");

            int offset = -2;
            double year = dataX[dataX.Length + offset];
            for (int i = 0; i < 20; i++)
            {
                double new_year = year + stepX * i;
                DataRow r = dt.NewRow();
                r["сдвиг"] = offset + i;
                r["Год"] = new_year;
                r["Линейная функция"] = f_line(new_year).ToString("N");
                dt.Rows.Add(r);
            }
            dataGridView1.DataSource = dt;
        }

        private void ToggleLinearRegressionLine(bool showLine)
        {
            //серия данных для линейной регрессии имеет имя "Линейная регрессия"
            chart1.Series["Линейная регрессия"].Enabled = showLine;
        }

        private void set_chart_series()
        {
            chart1.Series.Clear();

            // Настройка осей
            chart1.ChartAreas[0].AxisX.Minimum = dataX.Min();
            chart1.ChartAreas[0].AxisX.Maximum = dataX.Max() + stepX * (double)numericUpDown2.Value;
            chart1.ChartAreas[0].AxisX.MajorGrid.Interval = stepX;

            // Добавление серии данных для точек из файла
            chart1.Series.Add("Данные");
            chart1.Series["Данные"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Point;
            chart1.Series["Данные"].ChartArea = "ChartArea1";
            chart1.Series["Данные"].Points.DataBindXY(dataX, dataY);

            // Добавление серии данных для линейной регрессии
            chart1.Series.Add("Линейная регрессия");
            chart1.Series["Линейная регрессия"].ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Line;
            chart1.Series["Линейная регрессия"].ChartArea = "ChartArea1";

            // Расчет линейной регрессии и привязка данных к серии
            double sumX = 0, sumY = 0, sumXY = 0, sumX2 = 0;
            int n = dataX.Length;
            for (int i = 0; i < n; i++)
            {
                sumX += dataX[i];
                sumY += dataY[i];
                sumXY += dataX[i] * dataY[i];
                sumX2 += dataX[i] * dataX[i];
            }
            double b = (n * sumXY - sumX * sumY) / (n * sumX2 - sumX * sumX);
            double a = (sumY - b * sumX) / n;

            double minX = dataX.Min();
            double maxX = dataX.Max() + stepX * (double)numericUpDown2.Value;
            for (double x = minX; x <= maxX; x += stepX)
            {
                double y = a + b * x;
                chart1.Series["Линейная регрессия"].Points.AddXY(x, y);
            }
        }

        private double f_line(double x)
        {
            return coef_line_a + coef_line_b * x;
        }

        private void make_dataXe(int predict_num)
        {
            dataXe = new double[dataX.Length + predict_num];
            for (int i = 0; i < dataXe.Length; i++)
            {
                if (i < dataX.Length)
                    dataXe[i] = dataX[i];
                else
                    dataXe[i] = dataXe[i - 1] + stepX;
            }
        }

        private void calculate()
        {
            make_dataXe((int)numericUpDown2.Value);

            calc_line();
        }

        private double mean(double[] x)
        {
            double res = 0;
            int n = x.Length;
            for (int i = 0; i < n; i++)
            {
                res += x[i];
            }
            return res / n;
        }

        private double sigma2(double[] x)
        {
            double res = 0;
            double meanx = mean(x);
            int n = x.Length;
            for (int i = 0; i < n; i++)
            {
                double tmp = (x[i] - meanx);
                res += tmp * tmp;
            }
            return res / n;
        }

        private double mean_product(double[] x, double[] y)
        {
            double res = 0;
            int n = x.Length;
            for (int i = 0; i < n; i++)
            {
                res += x[i] * y[i];
            }
            return res / n;
        }

        private void calc_line()
        {
            coef_line_b = (mean_product(dataX, dataY) - mean(dataX) * mean(dataY)) / sigma2(dataX);
            coef_line_a = mean(dataY) - coef_line_b * mean(dataX);
            data_line = new double[dataXe.Length];
            for (int i = 0; i < data_line.Length; i++)
            {
                data_line[i] = f_line(dataXe[i]);
            }
        }
        private double approx_error(double[] Y, double[] f)
        {
            double res = 0;
            int n = Y.Length;
            for (int i = 0; i < n; i++)
            {
                res += Math.Abs((Y[i] - f[i]) / Y[i]);
            }
            return res / n * 100;
        }

        private double correlation(double[] X, double[] Y)
        {
            double res = 0;
            int n = X.Length;
            for (int i = 0; i < n; i++)
            {
                res += (X[i] - mean(X)) * (Y[i] - mean(Y));
            }
            return res / (Math.Sqrt(sigma2(X)) * Math.Sqrt(sigma2(Y)) * n);
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string text = "";

            // Открытие файла
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Text files (*.txt)|*.txt";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;

                // Чтение содержимого файла
                try
                {
                    text = File.ReadAllText(filePath);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error reading file: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                // Если CheckBox2 отмечен, выполняем выравнивание по столбцам
                if (checkBox2.Checked)
                {
                    text = AlignTextByColumns(text);
                }

                // Импортируем данные и используем их так же, как в методе для Button1
                ImportData(text);
                calculate();
                set_chart_series();
                ToggleLinearRegressionLine(true);
                update_labels_coef();
                update_predict_table();
            }
            else
            {
                MessageBox.Show("No file selected.", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private string AlignTextByColumns(string text)
        {
            // Разделение текста на строки
            string[] lines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            // Определение максимальной длины строки
            int maxLength = lines.Max(line => line.Length);

            // Выравнивание текста по столбцам
            StringBuilder sb = new StringBuilder();
            foreach (string line in lines)
            {
                sb.AppendFormat("{0,-" + maxLength + "}", line);
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private void ImportData(string data)
        {// Разделение текста на строки
            string[] lines = data.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

            // Определение числа элементов X и Y
            int countX = lines[0].Split(',').Length;
            int countY = lines.Length;

            // Создание массивов для данных X и Y
            dataX = new double[countX];
            dataY = new double[countY - 1];

            // Заполнение массивов данными из файла
            for (int i = 0; i < countY; i++)
            {
                string[] values = lines[i].Split(',');
                if (values.Length != countX)
                {
                    MessageBox.Show($"Неверный формат данных в строке {i + 1}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }

                if (i == 0)
                {
                    for (int j = 0; j < countX; j++)
                    {
                        if (!double.TryParse(values[j], out dataX[j]))
                        {
                            MessageBox.Show($"Неверный формат данных в столбце {j + 1} строки {i + 1}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                else
                {
                    for (int j = 0; j < countX; j++)
                    {
                        if (!double.TryParse(values[j], out dataY[i - 1]))
                        {
                            MessageBox.Show($"Неверный формат данных в столбце {j + 1} строки {i + 1}", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }

            // Вычисление среднего шага оси X
            stepX = (dataX.Last() - dataX.First()) / (dataX.Length - 1);

            // Вызов других методов для вычислений и построения графиков
            calculate();
            set_chart_series();
            ToggleLinearRegressionLine(true);
            update_labels_coef();
            update_predict_table();

        }
    }
}
