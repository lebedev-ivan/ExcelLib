using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public class ExcelSorter
    {
        private Microsoft.Office.Interop.Excel.Worksheet _worksheet;

        public ExcelSorter(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            _worksheet = worksheet;
        }

        /// <summary>
        /// Сортирует по колонке "Номер" (по возрастанию/убыванию)
        /// </summary>
        public void SortByNumber(bool ascending = true)
        {
            Excel.Range usedRange = _worksheet.UsedRange;
            if (usedRange.Rows.Count <= 1) return; // Проверяем, есть ли данные

            Excel.Range dataRange = _worksheet.Range["A1"].CurrentRegion;
            dataRange.Sort(
                _worksheet.Range["A2"],
                ascending ? Excel.XlSortOrder.xlAscending : Excel.XlSortOrder.xlDescending
            );
        }

        /// <summary>
        /// Сортирует по колонке "Имена" (по алфавиту)
        /// </summary>
        public void SortByNames()
        {
            Excel.Range usedRange = _worksheet.UsedRange;
            if (usedRange.Rows.Count <= 1) return; // Проверяем, есть ли данные

            Excel.Range dataRange = _worksheet.Range["A1"].CurrentRegion;
            dataRange.Sort(
                _worksheet.Range["B2"],
                Excel.XlSortOrder.xlAscending
            );
        }
    }
}
