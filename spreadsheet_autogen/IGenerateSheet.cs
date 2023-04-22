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
    public interface IGenerateSheet
    {
        Excel.Workbook CreateWorkbookLegacy();
        Excel.Worksheet CreateWorksheetUseLegacy(Excel.Workbook workBook);
        void CellRandomNumbersUseLegacy(Excel.Workbook workBook, Excel.Worksheet workSheet, string RBeforeParse, string CBeforeParse, string MinValue, string MaxValue);
        void CellRandomStringUseLegacy(Excel.Workbook workBook, Excel.Worksheet workSheet, string RandomString, string RBeforeParse, string CBeforeParse, string CharLength);
        void CellUserValueUseLegacy(Excel.Workbook workBook, Excel.Worksheet workSheet, string RBeforeParse, string CBeforeParse, string value);
        ExcelPackage ImportPackage();
        ExcelWorksheet CreateWorksheet(ExcelPackage excelPackage);
        void CellsRandomNumbers(string RBeforeParse, string CBeforeParse, string MinValue, string MaxValue);
        void CellRandomString(string RandomString, string RBeforeParse, string CBeforeParse, string CharLength);
        void CellUserValue(string RBeforeParse, string CBeforeParse, string value);
        void SaveSheet(ExcelPackage excelPackage);
    }
}
