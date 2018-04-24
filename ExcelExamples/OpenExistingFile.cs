using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExamples.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ExcelExamples {
    [TestClass]
    public class OpenExistingFile{
        private string FileName = "ExcelFileExtract\\Sample.xlsx";

        [TestMethod]
        public void LoadExcelFile() {
            // Open the document for editing.
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(FileName, false)) {
                // Code removed here.
                Assert.IsNotNull(spreadsheetDocument);
            }
        }

        [TestMethod]
        public void LoadExcelFileFromStream() {
            Stream stream = File.Open(FileName, FileMode.Open);
            OpenAndAddToSpreadsheetStream(stream);
            stream.Close();
        }

        private void OpenAndAddToSpreadsheetStream(Stream stream) {
            // Open a SpreadsheetDocument based on a stream.
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, false);

            // Add a new worksheet.
            WorksheetPart worksheetPart = spreadsheetDocument.WorkbookPart.WorksheetParts.FirstOrDefault();
            Assert.IsNotNull(worksheetPart);
            Worksheet sheet = worksheetPart.Worksheet;

            var cells = sheet.Descendants<Cell>();
            var rows = sheet.Descendants<Row>();

            Console.WriteLine("Row count = {0}", rows.LongCount());
            Console.WriteLine("Cell count = {0}", cells.LongCount());

            // One way: go through each cell in the sheet
            foreach (Cell cell in cells) {
                if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString)) {
                    int ssid = int.Parse(cell.CellValue.Text);
                    string str = sst.ChildElements[ssid].InnerText;
                    Console.WriteLine("Shared string {0}: {1}", ssid, str);
                } else if (cell.CellValue != null) {
                    Console.WriteLine("Cell contents: {0}", cell.CellValue.Text);
                }
            }

            // Or... via each row
            foreach (Row row in rows) {
                foreach (Cell c in row.Elements<Cell>()) {
                    if ((c.DataType != null) && (c.DataType == CellValues.SharedString)) {
                        int ssid = int.Parse(c.CellValue.Text);
                        string str = sst.ChildElements[ssid].InnerText;
                        Console.WriteLine("Shared string {0}: {1}", ssid, str);
                    } else if (c.CellValue != null) {
                        Console.WriteLine("Cell contents: {0}", c.CellValue.Text);
                    }
                }
            }
        }

    }
}
