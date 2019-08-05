using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Windows.Forms;
using System.IO;

namespace spreadsheet_autogen
{
   public class GenerateSheet
    {
        /// <summary>
        /// Импортирует пакет для работы с Excel
        /// </summary>
        /// <returns>ExcelPackage</returns>
        public ExcelPackage ImportPackage()
        {
            ExcelPackage excelPackage = new ExcelPackage();
            return excelPackage;
        }
        /// <summary>
        /// Создает новую таблицу
        /// </summary>
        /// <param name="excelPackage"></param>
        /// <returns>ExcelWorksheet</returns>
        public ExcelWorksheet CreateWorksheet(ExcelPackage excelPackage)
        {
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("TestSheet");
            return worksheet;
        }

        /// <summary>
        /// Заполняет таблицу случайными числами
        /// </summary>
        /// <param name="RBeforeParse"></param>
        /// <param name="CBeforeParse"></param>
        /// <param name="MinValue"></param>
        /// <param name="MaxValue"></param>
        public void CellsRandomNumbers(string RBeforeParse, string CBeforeParse, string MinValue, string MaxValue)
        {
            var excelPackage = this.ImportPackage();
            var worksheet = this.CreateWorksheet(excelPackage);
            Random random = new Random();
            var GetValue = new MainWindow();
            long Rows = long.Parse(RBeforeParse);
            long Columns = long.Parse(CBeforeParse);
     
                    for (int i = 1; i <= Rows; i++)
                    {
                        for (int ii = 1; ii <= Columns; ii++)
                        {
                            long randomValue = random.NextLong(long.Parse(MinValue), long.Parse(MaxValue));
                            worksheet.Cells[i, ii].Value = randomValue;
                        }
                    }
            SaveSheet(excelPackage);
        }

        /// <summary>
        /// Заполняет таблицы случайными строками
        /// </summary>
        /// <param name="RandomString"></param>
        /// <param name="RBeforeParse"></param>
        /// <param name="CBeforeParse"></param>
        /// <param name="CharLength"></param>
        public void CellRandomString(string RandomString, string RBeforeParse, string CBeforeParse, string CharLength)
        {
            var excelPackage = this.ImportPackage();
            var worksheet = this.CreateWorksheet(excelPackage);
            long Rows = long.Parse(RBeforeParse);
            long Columns = long.Parse(CBeforeParse);
            Random random = new Random();


            for (int i = 1; i <= Rows; i++)
            {
                for (int ii = 1; ii <= Columns; ii++)
                {
                    string randomValue = random.RandomString(int.Parse(CharLength));
                    worksheet.Cells[i, ii].Value = randomValue;
                }
            }
            SaveSheet(excelPackage);
        }

        /// <summary>
        /// Заполняет таблицу пользовательскими данными
        /// </summary>
        /// <param name="RBeforeParse"></param>
        /// <param name="CBeforeParse"></param>
        /// <param name="value"></param>
        public void CellUserValue(string RBeforeParse, string CBeforeParse, string value)
        {
            var excelPackage = this.ImportPackage();
            var worksheet = this.CreateWorksheet(excelPackage);
            long Rows = long.Parse(RBeforeParse);
            long Columns = long.Parse(CBeforeParse);

            for (int i = 1; i <= Rows; i++)
            {
                for (int ii = 1; ii <= Columns; ii++)
                {
                    worksheet.Cells[i, ii].Value = value;
                }
            }
            SaveSheet(excelPackage);
        }
        
        /// <summary>
        /// Сохраняет xlsx файл в указанное место
        /// </summary>
        /// <param name="excelPackage"></param>
        public void SaveSheet(ExcelPackage excelPackage)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog
            {
                Title = "Сохранить",
                Filter = "Excel files|*.xlsx|All files|*.*",
                FileName = "Test_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx"
            };

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                FileInfo file = new FileInfo(saveFileDialog.FileName);
                excelPackage.SaveAs(file);
                
                MessageBox.Show("Сохранено", "Success");
            }
        }


    }
    }

