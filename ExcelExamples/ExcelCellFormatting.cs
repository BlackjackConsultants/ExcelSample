using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExamples.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace ExcelExamples {
    [TestClass]
    public class ExcelCellFormatting {
        [TestMethod]
        public void ChangeBorderColor() {
            // Open the document for editing.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("ExcelFile\\Sample.xlsx", false)) {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                foreach (Row r in sheetData.Elements<Row>()) {
                    foreach (Cell c in r.Elements<Cell>()) {
                        text = c.CellValue.Text;
                        Console.Write(text + " 111111111");
                    }
                }
            }
        }

        /// <summary>
        /// https://stackoverflow.com/questions/15791732/openxml-sdk-having-borders-for-cell?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa
        /// </summary>
        [TestMethod]
        public void ChangeCellBorder() {
            // Open the document for editing.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("ExcelFile\\Sample.xlsx", true)) {
                // Code removed here.
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                foreach (Row r in sheetData.Elements<Row>()) {
                    foreach (Cell c in r.Elements<Cell>()) {
                        c.CellValue.Text = c.CellValue.Text + " 111111111";
                        ExcelHelper.SetCellBorder(worksheetPart, workbookPart, c);
                    }
                }
                spreadsheetDocument.Save();
            }
        }

        [TestMethod]
        public void ChangeCellValueAndSave() {
            // Open the document for editing.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("ExcelFile\\Sample.xlsx", true)) {
                // Code removed here.
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();
                string text;
                foreach (Row r in sheetData.Elements<Row>()) {
                    foreach (Cell c in r.Elements<Cell>()) {
                        c.CellValue.Text = c.CellValue.Text + " 111111111";
                    }
                }
                spreadsheetDocument.Save();
            }
        }

        [TestMethod]
        public void LoadExcelFile() {
            // Open the document for editing.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open("ExcelFile\\Sample.xlsx", false)) {
                // Code removed here.
                Assert.IsNotNull(spreadsheetDocument);
            }
        }
    }
}
