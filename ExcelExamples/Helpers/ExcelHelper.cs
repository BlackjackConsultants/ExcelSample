using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.Helpers {
    /// <summary>
    /// Wrapper class to get excel information.
    /// </summary>
    public static class ExcelHelper {
        /// <summary>
        /// returns the value of a cell.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="sheetName"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static string GetCellValue<T>(SpreadsheetDocument spreadsheetDocument, string sheetName, string range){
            string cellValue = null;
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault(n => n.LocalName == sheetName);
            if (sheetData != null){
                var firstOrDefault = sheetData.Elements<Cell>().FirstOrDefault(c => c.CellReference == range);
                if (firstOrDefault != null)
                    cellValue = firstOrDefault.ToString();
            }
            return cellValue;
        }

        /// <summary>
        /// returns true if the cell is highlighted.
        /// </summary>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="sheetName"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static bool CellIsHighlighted(SpreadsheetDocument spreadsheetDocument, string sheetName, string range) {
            WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
            SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault(n => n.LocalName == sheetName);
            if (sheetData != null) {
                var cell = sheetData.Elements<Cell>().FirstOrDefault(c => c.CellReference == range);
                if (cell != null)
                    return cell.StyleIndex != 0;
            }
            return false;
        }
    }
}
