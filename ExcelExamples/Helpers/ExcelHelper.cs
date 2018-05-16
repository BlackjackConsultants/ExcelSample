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
using ExcelExamples.Extension;

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
        public static bool CellIsHighlighted(SpreadsheetDocument spreadsheetDocument, string sheetName, string cellReference, int withIndex) {
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            WorksheetPart wsPart = spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
            string cellValue = string.Empty;
            string cellRefLetter = cellReference.Substring(0, cellReference.FirstDigitIndex());
            uint cellRefNumber = cellReference.GetNumericValue();

            if (wsPart != null) {
                Worksheet worksheet = wsPart.Worksheet;
                Cell cell = GetCell(worksheet, cellRefLetter, cellRefNumber);
                return withIndex == cell.StyleIndex;
            }
            return false;
        }

        public static int AddFontStyle(SpreadsheetDocument spreadsheetDocument, int fontSize, string fontName, string fontColor) {
            var styleSheet = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;
            var font = new Font(
                new FontSize() { Val = fontSize },
                new Color() { Rgb = new HexBinaryValue() { Value = fontColor } },
                new FontName() { Val = fontName });
            styleSheet.Fonts.Append(font);
            return Convert.ToInt32(styleSheet.Fonts.Elements<Font>().ToList().Count);
        }

        public static int AddFillStyle(SpreadsheetDocument spreadsheetDocument, string fillBackgroundColorName, string fillForeColorName) {
            var styleSheet = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;
            var fillForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue() { Value = fillForeColorName } };
            var fillBackgroundColor = new BackgroundColor() { Rgb = new HexBinaryValue() { Value = fillBackgroundColorName } };
            var fill = new Fill(new PatternFill() {
                PatternType = PatternValues.None,
                BackgroundColor = fillBackgroundColor,
                ForegroundColor = fillForegroundColor
            });
            styleSheet.Fills.Append(fill);
            return Convert.ToInt32(styleSheet.Fills.Elements<Fill>().ToList().Count);
        }

        public static int AddBorderStyle(SpreadsheetDocument spreadsheetDocument) {
            var styleSheet = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;
            var border = new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder());
            styleSheet.Borders.Append(border);
            return Convert.ToInt32(styleSheet.Borders.Elements<Border>().ToList().Count);
        }

        public static int AddCellFormatStyle(SpreadsheetDocument spreadsheetDocument, int fontId, int fillId, int borderId) {
            var styleSheet = spreadsheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet;
            var cellFormat = new CellFormat() { FontId = Convert.ToUInt32(fontId), FillId = Convert.ToUInt32(fillId), BorderId = Convert.ToUInt32(borderId) };
            styleSheet.CellFormats.Append(cellFormat);
            return Convert.ToInt32(styleSheet.CellFormats.Elements<CellFormat>().ToList().Count);
        }

        private static Stylesheet GenerateStyleSheet(int fontSize, string fontName, string fontColor, string fillBackgroundColorName, string fillForeColorName) {
            var fillForegroundColor = new ForegroundColor() { Rgb = new HexBinaryValue() { Value = fillForeColorName } };
            var fillBackgroundColor = new BackgroundColor() { Rgb = new HexBinaryValue() { Value = fillBackgroundColorName } };

            Stylesheet styleSheet = new Stylesheet(
                    new Fonts(
                        new Font(
                            new FontSize() { Val = 11 },
                            new Color() { Rgb = new HexBinaryValue() { Value = fontColor } },
                            new FontName() { Val = fontName }
                         )
                    ),
                    new Fills(
                        new Fill(
                            new PatternFill() { PatternType = PatternValues.None, BackgroundColor = fillBackgroundColor, ForegroundColor = fillForegroundColor }
                        )
                     ),
                    new Borders(                
                        new Border(
                            new LeftBorder(),
                            new RightBorder(),
                            new TopBorder(),
                            new BottomBorder(),
                            new DiagonalBorder()
                        )
                    ),
                    new CellFormats(
                    )
             );
            return styleSheet;
        }

        /// <summary>
        /// associates an existing style to the cell.
        /// </summary>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="sheetName"></param>
        /// <param name="range"></param>
        /// <returns></returns>
        public static void HighlightCell(SpreadsheetDocument spreadsheetDocument, string sheetName, string cellReference, uint withIndex) {
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            WorksheetPart wsPart = spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
            string cellValue = string.Empty;
            string cellRefLetter = cellReference.Substring(0, cellReference.FirstDigitIndex());
            uint cellRefNumber = cellReference.GetNumericValue();

            if (wsPart != null) {
                Worksheet worksheet = wsPart.Worksheet;
                Cell cell = GetCell(worksheet, cellRefLetter, cellRefNumber);
                cell.StyleIndex.Value = withIndex;
            }
        }

        /// <summary>
        /// returns the cell value from spreadsheet.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="sheetName"></param>
        /// <param name="cellReference"></param>
        /// <returns></returns>
        public static int GetCellStyleIndex(SpreadsheetDocument spreadsheetDocument, string sheetName, string cellReference) {
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            WorksheetPart wsPart = spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
            string cellValue = string.Empty;
            string cellRefLetter = cellReference.Substring(0, cellReference.FirstDigitIndex());
            uint cellRefNumber = cellReference.GetNumericValue();

            if (wsPart != null) {
                Worksheet worksheet = wsPart.Worksheet;
                Cell cell = GetCell(worksheet, cellRefLetter, cellRefNumber);
                if (cell.DataType != null) {
                    return int.Parse(cell.StyleIndex);
                }
            }
            return 0;
        }

        /// <summary>
        /// returns the cell value from stream.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="sheetName"></param>
        /// <param name="cellReference"></param>
        /// <returns></returns>
        public static string GetCellValue(string filename, string sheetName, string cellReference) {
            Stream stream = File.Open(filename, FileMode.Open);
            return GetCellValue(stream, sheetName, cellReference);
        }

        /// <summary>
        /// returns the cell value from stream.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="sheetName"></param>
        /// <param name="cellReference"></param>
        /// <returns></returns>
        public static string GetCellValue(Stream stream, string sheetName, string cellReference) {
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false);
            return GetCellValue(spreadsheetDocument, sheetName, cellReference);
        }

        /// <summary>
        /// returns the cell value from spreadsheet.
        /// </summary>
        /// <param name="stream"></param>
        /// <param name="sheetName"></param>
        /// <param name="cellReference"></param>
        /// <returns></returns>
        public static string GetCellValue(SpreadsheetDocument spreadsheetDocument, string sheetName, string cellReference) {
            var sheet = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();
            WorksheetPart wsPart = spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id) as WorksheetPart;
            string cellValue = string.Empty;
            string cellRefLetter = cellReference.Substring(0, cellReference.FirstDigitIndex());
            uint cellRefNumber = cellReference.GetNumericValue();

            if (wsPart != null) {
                Worksheet worksheet = wsPart.Worksheet;
                Cell cell = GetCell(worksheet, cellRefLetter, cellRefNumber);
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

        /// <summary>
        /// loads a spreadsheetdocument.  use this so that you dont have to load streams each time. is faster.
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="isEditable"></param>
        /// <returns></returns>
        public static SpreadsheetDocument LoadSpreadSheetDocument(string filename, bool isEditable) {
            Stream stream = File.Open(filename, FileMode.Open);
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, isEditable);
            return spreadsheetDocument;
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

        private static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id) {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }
    }
}
