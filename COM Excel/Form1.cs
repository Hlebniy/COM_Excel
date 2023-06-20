using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace COM_Excel
{
    public partial class Form1 : Form
    {
        private Excel.Application excelApp;
        private Excel.Workbook workbook;
        private Excel.Worksheet worksheet;
        private int currentRow = 2;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // Создание экземпляра приложения Excel
            excelApp = new Excel.Application();
            workbook = excelApp.Workbooks.Add();
            worksheet = workbook.ActiveSheet;

            // Задание заголовков таблицы
            string[] headers = { "Имя", "Возраст", "Email" };
            for (int i = 0; i < headers.Length; i++)
            {
                worksheet.Cells[1, i + 1] = headers[i];
            }
        }

        private void ReleaseComObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Ошибка при освобождении COM-объекта: " + ex.Message);
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            string filePath = "data.xlsx";

            try
            {
                // Открываем существующий файл
                excelApp = new Excel.Application();
                workbook = excelApp.Workbooks.Open(filePath);
                worksheet = workbook.ActiveSheet;

                // Находим последнюю заполненную строку
                int lastRow = worksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell).Row;

                // Записываем новые данные в таблицу начиная со следующей строки
                worksheet.Cells[lastRow + 1, 1] = textBox1.Text;
                worksheet.Cells[lastRow + 1, 2] = textBox2.Text;
                worksheet.Cells[lastRow + 1, 3] = textBox3.Text;

                // Сохраняем изменения
                workbook.Save();

                MessageBox.Show("Данные успешно добавлены в таблицу Excel.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка при добавлении данных в файл: " + ex.Message);
            }
            finally
            {
                // Закрываем приложение Excel
                workbook.Close();
                excelApp.Quit();
                ReleaseComObject(worksheet);
                ReleaseComObject(workbook);
                ReleaseComObject(excelApp);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string name = textBox1.Text;
            string age = textBox2.Text;
            string email = textBox3.Text;

            // Запись данных в таблицу
            worksheet.Cells[currentRow, 1] = name;
            worksheet.Cells[currentRow, 2] = age;
            worksheet.Cells[currentRow, 3] = email;

            currentRow++;
            MessageBox.Show("Данные успешно добавлены в таблицу Excel.");
        }
    }
}