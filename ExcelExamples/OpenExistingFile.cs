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
            SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(stream, true);

            // Add a new worksheet.
            WorksheetPart newWorksheetPart = spreadsheetDocument.WorkbookPart.AddNewPart<WorksheetPart>();
            newWorksheetPart.Worksheet = new Worksheet(new SheetData());
            newWorksheetPart.Worksheet.Save();

            Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>();
            string relationshipId = spreadsheetDocument.WorkbookPart.GetIdOfPart(newWorksheetPart);

            // Get a unique ID for the new worksheet.
            uint sheetId = 1;
            if (sheets.Elements<Sheet>().Count() > 0) {
                sheetId = sheets.Elements<Sheet>().Select(s => s.SheetId.Value).Max() + 1;
            }

            // Give the new worksheet a name.
            string sheetName = "Sheet" + sheetId;

            // Append the new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
            sheets.Append(sheet);
            spreadsheetDocument.WorkbookPart.Workbook.Save();

            // Close the document handle.
            spreadsheetDocument.Close();

            // Caller must close the stream.
        }

    }
}
