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
        ////public ExcelWorksheet worksheet
        //    public ExcelWorksheet CreateWorksheet()
        //{
        //    ExcelPackage excelPackage = new ExcelPackage();
            
        //    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("TestSheet");
        //    return worksheet;
        //}

        public void CellsRandomNumbers(string RBeforeParse, string CBeforeParse, string MinValue, string MaxValue)
        {
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("TestSheet");
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
        public void SaveSheet(ExcelPackage excelPackage)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "Сохранить";
            saveFileDialog.Filter = "Excel files|*.xlsx|All files|*.*";
            saveFileDialog.FileName = "Test_" + DateTime.Now.ToString("dd-MM-yyyy") + ".xlsx";

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                //Get the FileInfo
                FileInfo fi = new FileInfo(saveFileDialog.FileName);
                //write the file to the disk
                excelPackage.SaveAs(fi);
            }
        }
    }
}
