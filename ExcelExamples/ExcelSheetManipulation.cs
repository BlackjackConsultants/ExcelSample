using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExamples.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ExcelExamples {
    [TestClass]
    public class ExcelSheetManipulation {
        const string tempfilename = "c:\\temp\\testing.xlsx";
        const string filename = "ExcelFileExtract\\test.xlsx";

        [TestMethod]
        public void GetCellValueFromCellReference() {
            var spreadSheetDocument = ExcelHelper.LoadSpreadsheetDocument("ExcelFileExtract\\test.xlsx", true);
            string value1 = ExcelHelper.GetCellValue(spreadSheetDocument, "testing", "A1");
            Assert.AreEqual(value1, "sssd");
            string value2 = ExcelHelper.GetCellValue(spreadSheetDocument, "testing", "B2");
            Assert.AreEqual(value2, "fdfdf");
        }

        [TestMethod]
        public void CellIsHighlightedTest() {
            var spreadSheetDocument = ExcelHelper.LoadSpreadsheetDocument("ExcelFileExtract\\test.xlsx", true);
            var value1 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "A1", 1);
            Assert.IsTrue(value1);
            var value2 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "B1", 2);
            Assert.IsTrue(value2);
        }

        [TestMethod]
        public void HighlightCell() {
            var spreadSheetDocument = ExcelHelper.LoadSpreadsheetDocument(filename, true);
            var value1 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "A1", 1);
            Assert.IsTrue(value1);
            ExcelHelper.HighlightCell(spreadSheetDocument, "testing", "B1", 2);
            var value2 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "B1", 2);
            ExcelHelper.HighlightCell(spreadSheetDocument, "testing", "B3", 2);
            var value3 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "B3", 2);
            ExcelHelper.HighlightCell(spreadSheetDocument, "testing", "B4", 3);
            var value4 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "B4", 3);
            spreadSheetDocument.SaveAs(tempfilename);
            System.Diagnostics.Process.Start(tempfilename);
            Assert.IsTrue(value2);
        }


        [TestMethod]
        public void CreateStyleAndAssociateIt() {
            var spreadsheet = ExcelHelper.CreateNewSpreadSheetDocument("output.xlsx", true);
            spreadsheet.AddWorkbookPart();
            spreadsheet.WorkbookPart.Workbook = new Workbook();
            var wsPart = spreadsheet.WorkbookPart.AddNewPart<WorksheetPart>();
            wsPart.Worksheet = new Worksheet();
            var stylesPart = spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet();

            var stylesSheet = spreadsheet.WorkbookPart.WorkbookStylesPart.Stylesheet;
            // add font
            var f1 = ExcelHelper.AddFontStyle(spreadsheet, 14, "Arial", "FFFFFFF0");
            var f2 = ExcelHelper.AddFontStyle(spreadsheet, 12, "Arial", "F0F0F0F0");
            Assert.AreNotEqual(f1, f2);

            // add fill
            // create a solid red fill
            var solidRed = new PatternFill() { PatternType = PatternValues.Solid };
            var fg = new ForegroundColor { Rgb = HexBinaryValue.FromString("FFFF0000") }; // red fill
            var bg = new BackgroundColor { Indexed = 64 };

            var fill1 = ExcelHelper.AddFillStyle(spreadsheet, null, null, PatternValues.None);
            var fill2 = ExcelHelper.AddFillStyle(spreadsheet, null, null, PatternValues.Gray125);
            var fill3 = ExcelHelper.AddFillStyle(spreadsheet, bg, fg, PatternValues.Solid);
            Assert.AreNotEqual(fill1, fill2);
            Assert.AreNotEqual(fill1, fill3);

            // add border
            var b1 = ExcelHelper.AddBorderStyle(spreadsheet);
            var b2 = ExcelHelper.AddBorderStyle(spreadsheet);
            Assert.AreNotEqual(b1, b2);

            // add cellformat
            var c1 = ExcelHelper.AddCellFormatStyle(spreadsheet);
            var c2 = ExcelHelper.AddCellFormatStyle(spreadsheet, 0, 0, 0, 2, true);
            Assert.AreNotEqual(c1, c2);
            stylesSheet.Save();

            // add sheet data
            var sheetData = wsPart.Worksheet.AppendChild(new SheetData());
            // add row data
            var row = sheetData.AppendChild(new Row());
            row.AppendChild(new Cell() { CellValue = new CellValue("This"), DataType = CellValues.String });
            row.AppendChild(new Cell() { CellValue = new CellValue("is"), DataType = CellValues.String });
            row.AppendChild(new Cell() { CellValue = new CellValue("a"), DataType = CellValues.String });
            row.AppendChild(new Cell() { CellValue = new CellValue("test."), DataType = CellValues.String });
            sheetData.AppendChild(new Row());
            // add row data
            row = sheetData.AppendChild(new Row());
            row.AppendChild(new Cell() { CellValue = new CellValue("Value:"), DataType = CellValues.String });
            row.AppendChild(new Cell() { CellValue = new CellValue("123"), DataType = CellValues.Number });
            row.AppendChild(new Cell() { CellValue = new CellValue("Formula:"), DataType = CellValues.String });
            // style index = 1, i.e. point at our fill format
            row.AppendChild(new Cell() { CellFormula = new CellFormula("B3"), DataType = CellValues.Number, StyleIndex = 1 });

            // save worksheet
            wsPart.Worksheet.Save();

            var sheets = spreadsheet.WorkbookPart.Workbook.AppendChild(new Sheets());
            sheets.AppendChild(new Sheet() { Id = spreadsheet.WorkbookPart.GetIdOfPart(wsPart), SheetId = 1, Name = "Test" });

            spreadsheet.WorkbookPart.Workbook.Save();
            spreadsheet.Close();
        }
    }
}
