using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using System.Windows.Forms;

namespace AttestetionWork
{
    public partial class Form1 : Form
    {
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
            //dataX = new double[CountX];
            //dataY = new double[CountY];

            OpenFileDialog openDialog = new OpenFileDialog();
            openDialog.Filter = "Файл Excel|*.XLSX;*.XLS";
            var result = openDialog.ShowDialog();
            if (result != DialogResult.OK)
            {
                MessageBox.Show("Файл не выбран!", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            //string fileName = System.IO.Path.GetFileName(openDialog.FileName);

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
            /*Microsoft.Office.Interop.Excel.Range r =;
            int x = r.Column;*/
            try
            {
                //Загрузка оси Х
                if (RangeByRows)//если данные хранятся в файле по строкам
                {
                    int p = 0;
                    for (int i = RangeCb - 1; i < RangeCe; i++)
                    {
                        int j = RangeRb;
                        //dataX[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j, i + 1].Text.ToString());
                    }
                }
                else//если данные хранятся в файле по столбцам
                {
                    int p = 0;
                    for (int j = RangeRb - 1; j < RangeRe; j++)
                    {
                        int i = RangeCb;
                        //dataX[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j + 1, i].Text.ToString());
                    }
                }

                //Загрузка Y
                /*if (RangeDByRows)//если данные хранятся в файле в виде строки
                {
                    int p = 0;
                    for (int i = RangeDCb - 1; i < RangeDCe; i++)
                    {
                        int j = RangeDRb;
                        dataY[p++] = Convert.ToDouble(ExcelWorksheet.Cells[j, i + 1].Text.ToString());
                    }
                    if (checkBoxCaptionY.Checked && RangeDCb > 1)//Загрузка подписи данных
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
                    if (checkBoxCaptionY.Checked && RangeDRb > 1)//Загрузка подписи данных
                        series_name_data = ExcelWorksheet.Cells[RangeDRb - 1, RangeDCb].Text.ToString();
                    else
                        series_name_data = series_name_data_default;
                }*/
            }
            catch
            {
                MessageBox.Show("Ошибка в файле Excel либо неверный диаппазон");
            }
            ExcelWorkbook.Close(false, Type.Missing, Type.Missing); // закрыть файл не сохраняя
            ExcelApp.Quit(); // Закрыть экземпляр Excel
            GC.Collect();   //Инициировать сборщик мусора

            //Вычисление среднего шага оси Х
            /*stepX = (dataX.Last() - dataX.First()) / (dataX.Length - 1);

            //Вычисления, построение графиков
            calculate();
            set_chart_series();
            checkBoxSeries_CheckedChanged(null, null);// Проверка выводимых линий трендов

            update_labels_coef();

            update_predict_table();
            update_acorrel_table();*/
        }

        private void numericUpDownPredictNum_ValueChanged(object sender, EventArgs e)
        {
            //Вычисления и построение графиков
            //calculate();
            //set_chart_series();
        }

        /*private void update_labels_coef()
        {
            labelCoefLine.Text = "y = " + coef_line_a.ToString("N") + " + " + coef_line_b.ToString("N") + " * x";
            
            //параметры линейной функции
            labelRxyLine.Text = (coef_line_b * Math.Sqrt(sigma2(dataX)) / Math.Sqrt(sigma2(dataY))).ToString("N");
            labelAperLine.Text = (approx_error(dataY, data_line)).ToString("N") + " %";
            labelR2Line.Text = (1.0 - sigma2(data_line) / sigma2(dataY)).ToString("N");

            //Коэффициент корреляции
            labelCorrel.Text = correlation(dataX, dataY).ToString("N");
        }*/
    }
}
