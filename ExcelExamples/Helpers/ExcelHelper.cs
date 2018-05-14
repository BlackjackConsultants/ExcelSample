using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelExamples.Helpers {

    /// <summary>
    /// Wrapper class to get excel information.
    /// </summary>
    public static class ExcelHelper {
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

        /// <summary>
        /// loads a spreadsheet file.
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public static Stream LoadSpreadSheet(string filename, bool isEditable) {
            Stream stream = File.Open(filename, FileMode.Open);
            return stream;
        }

        /// <summary>
        /// gets a sheet data.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="sheetName"></param>
        /// <returns></returns>
        public static SheetData GetSheetData(Stream stream, string sheetName) {
            // Open a SpreadsheetDocument based on a stream.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            WorksheetPart wsPart = spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;

            if (wsPart != null) {
                return wsPart.Worksheet.GetFirstChild<SheetData>();
            }
            return null;
        }
        
        /// <summary>
        /// returns the cell value.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="sheetName"></param>
        /// <param name="cellReference"></param>
        /// <returns></returns>
        public static string GetCellValue(Stream stream, string sheetName, string cellReference) {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            WorksheetPart wsPart = spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
            string cellValue = string.Empty;

            if (wsPart != null) {
                Worksheet worksheet = wsPart.Worksheet;
                Cell cell = GetCell(worksheet, "A", 1);
                if (cell.DataType != null) {
                    if (cell.DataType == CellValues.SharedString) {
                        int id = -1;

                        if (Int32.TryParse(cell.InnerText, out id)) {
                            SharedStringItem item = GetSharedStringItemById(spreadsheetDocument.WorkbookPart, id);

                            if (item.Text != null) {
                                cellValue = item.Text.Text;
                            } else if (item.InnerText != null) {
                                cellValue = item.InnerText;
                            } else if (item.InnerXml != null) {
                                cellValue = item.InnerXml;
                            }
                        }
                    }
                }
            }
            return cellValue;
        }

        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex) {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            return row.Elements<Cell>().Where(c => string.Compare
                      (c.CellReference.Value, columnName +
                      rowIndex, true) == 0).First();
        }

        private static Row GetRow(Worksheet worksheet, uint rowIndex) {
            return worksheet.GetFirstChild<SheetData>().
                  Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id) {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
    }
}
