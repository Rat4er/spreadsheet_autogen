using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.Windows.Forms;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace spreadsheet_autogen
{
    public class GenerateSheet : IGenerateSheet
    {
        public Excel.Application excelApp => getExcelApp();
        internal Excel.Application getExcelApp()
        {
            try
            {
                Excel.Application excelApp = new Excel.Application();
            }
            catch(System.Runtime.InteropServices.COMException exc)
            {
                MessageBox.Show($"Some error has been occured {exc.InnerException} ");
                return null;
            }
            return excelApp;
        }
        public Excel.Workbook CreateWorkbookLegacy()
        {  
            Excel.Workbook workBook;
            workBook = excelApp.Workbooks.Add();
            return workBook;
        }
        /// <summary>
        /// Создает книгу и лист Excel используя Legacy библиотеку
        /// </summary>
        /// <returns>workSheet</returns>
        public Excel.Worksheet CreateWorksheetUseLegacy(Excel.Workbook workBook)
        {
            Excel.Worksheet workSheet;
            workSheet = (Excel.Worksheet)workBook.Sheets.Add(Type.Missing, workBook.Sheets[1], Type.Missing, Type.Missing);
            workSheet.Name = "Test";
            return workSheet;
        }
        /// <summary>
        /// Заполняет таблицу случайными числами используя Legacy библиотеку
        /// </summary>
        /// <param name="RBeforeParse"></param>
        /// <param name="CBeforeParse"></param>
        /// <param name="MinValue"></param>
        /// <param name="MaxValue"></param>
        public void CellRandomNumbersUseLegacy(Excel.Workbook workBook, Excel.Worksheet workSheet, string RBeforeParse, string CBeforeParse, string MinValue, string MaxValue)
        {
            Random random = new Random();
            var GetValue = new MainWindow();
                        
            
            long Rows = long.Parse(RBeforeParse);
            long Columns = long.Parse(CBeforeParse);

            for (int i = 1; i <= Rows; i++)
            {
                for (int ii = 1; ii <= Columns; ii++)
                {
                    long randomValue = random.NextLong(long.Parse(MinValue), long.Parse(MaxValue));
                    workSheet.Cells[i, ii] = randomValue;
                }
            }

            workBook.SaveAs("C:\\TestFile\\Test.xls", Excel.XlFileFormat.xlWorkbookNormal,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(Missing.Value, Missing.Value, Missing.Value);
        }

        /// <summary>
        /// Заполняет таблицу случайными строками используя Legacy библиотеку
        /// </summary>
        /// <param name="RandomString"></param>
        /// <param name="RBeforeParse"></param>
        /// <param name="CBeforeParse"></param>
        /// <param name="CharLength"></param>
        public void CellRandomStringUseLegacy(Excel.Workbook workBook, Excel.Worksheet workSheet, string RandomString, string RBeforeParse, string CBeforeParse, string CharLength)
        {
            long Rows = long.Parse(RBeforeParse);
            long Columns = long.Parse(CBeforeParse);
            Random random = new Random();

            for (int i = 1; i <= Rows; i++)
            {
                for (int ii = 1; ii <= Columns; ii++)
                {
                    string randomValue = random.RandomString(int.Parse(CharLength));
                    workSheet.Cells[i, ii] = randomValue;
                }
            }
            workBook.SaveAs("C:\\TestFile\\Test.xls", Excel.XlFileFormat.xlWorkbookNormal,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(Missing.Value, Missing.Value, Missing.Value);
        }

        /// <summary>
        /// Заполняет таблицу пользовательскими данными используя Legacy библиотеку
        /// </summary>
        /// <param name="workBook"></param>
        /// <param name="workSheet"></param>>
        /// <param name="RBeforeParse"></param>
        /// <param name="CBeforeParse"></param>
        /// <param name="value"></param>
        public void CellUserValueUseLegacy(Excel.Workbook workBook, Excel.Worksheet workSheet, string RBeforeParse, string CBeforeParse, string value)
        {
            long Rows = long.Parse(RBeforeParse);
            long Columns = long.Parse(CBeforeParse);


            for (int i = 1; i <= Rows; i++)
            {
                for (int ii = 1; ii <= Columns; ii++)
                {
                    workSheet.Cells[i, ii].Value = value;
                }
            }
            workBook.SaveAs("C:\\TestFile\\Test.xls", Excel.XlFileFormat.xlWorkbookNormal,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workBook.Close(Missing.Value, Missing.Value, Missing.Value);
        }

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

