using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualStudio.TestTools.UnitTesting;

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
        public static string GetCellValue<T>(SheetData sheetData, string reference){
            string rowIndex = Regex.Match(reference, @"\d+").Value;
            string cellValue = null;
            if (sheetData != null){
                var row = sheetData.Elements<Row>().Where(r => r.RowIndex.Value == int.Parse(rowIndex)).FirstOrDefault();
                if (row != null) {
                    var cell = row.Elements<Cell>().Where(c => c.CellReference.Value == reference).FirstOrDefault();
                    if (cell != null) {
                        return cell.CellValue.ToString();
                    }
                }
            }
            return null;
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
            ////// Open a SpreadsheetDocument based on a stream.
            ////SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
            ////////var testing = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(1).Name;

            ////int sheetIndex = 0;
            ////foreach (WorksheetPart worksheetpart in spreadsheetDocument.WorkbookPart.WorksheetParts) {
            ////    Worksheet worksheet = worksheetpart.Worksheet;

            ////    // Grab the sheet name each time through your loop
            ////    string sheetname = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().ElementAt(sheetIndex).Name;

            ////    foreach (SheetData sheetData in worksheet.Elements<SheetData>()) {

            ////    }
            ////    sheetIndex++;
            ////}









            ////foreach (Sheet sheet in spreadsheetDocument.WorkbookPart.Workbook.Sheets) {

            ////}
            ////var sheet = spreadsheetDocument.WorkbookPart.Workbook.Sheets.Where(s => s.Name == sheetName);
            ////// Add a new worksheet.
            ////foreach (WorksheetPart worksheetPart in spreadsheetDocument.WorkbookPart.WorksheetParts) {
            ////    var name = worksheetPart.Uri.ToString().Split('/').Where(n => n == sheetName + ".xml").FirstOrDefault();
            ////    if (name != null) {
            ////        return worksheetPart.Worksheet.Elements<SheetData>().First();
            ////    }
            ////}
            return null;
        }

        public static SheetData ExcelDocTest(Stream stream, string sheetName) {
            Debug.WriteLine("Running through sheet.");
            int rowsComplete = 0;

            using (SpreadsheetDocument spreadsheetDocument =
                            SpreadsheetDocument.Open(@"path\to\Spreadsheet.xlsx", false)) {
                WorkbookPart workBookPart = spreadsheetDocument.WorkbookPart;

                foreach (Sheet s in workBookPart.Workbook.Descendants<Sheet>()) {
                    WorksheetPart wsPart = workBookPart.GetPartById(s.Id) as WorksheetPart;
                    Debug.WriteLine("Worksheet {1}:{2} - id({0}) {3}", s.Id, s.SheetId, s.Name,
                        wsPart == null ? "NOT FOUND!" : "found.");

                    if (wsPart == null) {
                        continue;
                    }

                    Row[] rows = wsPart.Worksheet.Descendants<Row>().ToArray();

                    //assumes the first row contains column names 
                    foreach (Row row in wsPart.Worksheet.Descendants<Row>()) {
                        rowsComplete++;

                        bool emptyRow = true;
                        List<object> rowData = new List<object>();
                        string value;

                        foreach (Cell c in row.Elements<Cell>()) {
                            value = GetCellValue(c);
                            emptyRow = emptyRow && string.IsNullOrWhiteSpace(value);
                            rowData.Add(value);
                        }

                        Debug.WriteLine("Row {0}: {1}", row,
                            emptyRow ? "EMPTY!" : string.Join(", ", rowData));
                    }
                }

            }
            Debug.WriteLine("Done, processed {0} rows.", rowsComplete);
        }

        public static string GetCellValue(Cell cell) {
            if (cell == null)
                return null;
            if (cell.DataType == null)
                return cell.InnerText;

            string value = cell.InnerText;
            switch (cell.DataType.Value) {
                case CellValues.SharedString:
                    // For shared strings, look up the value in the shared strings table.
                    // Get worksheet from cell
                    OpenXmlElement parent = cell.Parent;
                    while (parent.Parent != null && parent.Parent != parent
                            && string.Compare(parent.LocalName, "worksheet", true) != 0) {
                        parent = parent.Parent;
                    }
                    if (string.Compare(parent.LocalName, "worksheet", true) != 0) {
                        throw new Exception("Unable to find parent worksheet.");
                    }

                    Worksheet ws = parent as Worksheet;
                    SpreadsheetDocument ssDoc = ws.WorksheetPart.OpenXmlPackage as SpreadsheetDocument;
                    SharedStringTablePart sstPart = ssDoc.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();

                    // lookup value in shared string table
                    if (sstPart != null && sstPart.SharedStringTable != null) {
                        value = sstPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                    }
                    break;

                //this case within a case is copied from msdn. 
                case CellValues.Boolean:
                    switch (value) {
                        case "0":
                            value = "FALSE";
                            break;
                        default:
                            value = "TRUE";
                            break;
                    }
                    break;
            }
            return value;
        }
    }
}
