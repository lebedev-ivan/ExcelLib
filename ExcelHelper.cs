using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public class ExcelHelper
    {
        public Microsoft.Office.Interop.Excel.Workbook OpenExcelFile(string filePath)
        {
            // Создаём экземпляр Excel
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            if (excelApp == null)
            {
                throw new Exception("Не удалось создать объект Excel.");
            }

            // Открываем файл
            Microsoft.Office.Interop.Excel.Workbook workbook = excelApp.Workbooks.Open(filePath);
            Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1]; // Выбираем первый лист

            return workbook;
        }

        public void CloseWorkbook(Microsoft.Office.Interop.Excel.Workbook workbook)
{
    workbook.Close(false);
    workbook.Application.Quit();
    System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook);
}

    }
}
