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

        [TestMethod]
        public void GetCellValueFromCellReference() {
            var spreadSheetDocument = ExcelHelper.LoadSpreadSheetDocument("ExcelFileExtract\\test.xlsx", true);
            string value1 = ExcelHelper.GetCellValue(spreadSheetDocument, "testing", "A1");
            Assert.AreEqual(value1, "sssd");
            string value2 = ExcelHelper.GetCellValue(spreadSheetDocument, "testing", "B2");
            Assert.AreEqual(value2, "fdfdf");
        }

        [TestMethod]
        public void CellIsHighlightedTest() {
            var spreadSheetDocument = ExcelHelper.LoadSpreadSheetDocument("ExcelFileExtract\\test.xlsx", true);
            var value1 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "A1", 1);
            Assert.IsTrue(value1);
            var value2 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "B1", 2);
            Assert.IsTrue(value2);
        }

        [TestMethod]
        public void HighlightCell() {
            var spreadSheetDocument = ExcelHelper.LoadSpreadSheetDocument("ExcelFileExtract\\test.xlsx", true);
            var value1 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "A1", 1);
            Assert.IsTrue(value1);
            ExcelHelper.HighlightCell(spreadSheetDocument, "testing", "B1", 2);
            var value2 = ExcelHelper.CellIsHighlighted(spreadSheetDocument, "testing", "B1", 2);
            Assert.IsTrue(value2);
        }


        [TestMethod]
        public void CreateStyleAndAssociateIt() {
            var doc = ExcelHelper.LoadSpreadSheetDocument("ExcelFileExtract\\test.xlsx", true);
            var stylesSheet = doc.WorkbookPart.WorkbookStylesPart.Stylesheet;
            // add font
            var f1 = ExcelHelper.AddFontStyle(doc, 14, "Arial", "FFFFFFF0");
            var f2 = ExcelHelper.AddFontStyle(doc, 12, "Arial", "F0F0F0F0");
            Assert.AreNotEqual(f1, f2);
            // add fill
            var fill1 = ExcelHelper.AddFillStyle(doc, "F0F00000", "F0F00000");
            var fill2 = ExcelHelper.AddFillStyle(doc, "1010AAAA", "FFF000AA");
            Assert.AreNotEqual(fill1, fill2);
            // add border
            var b1 = ExcelHelper.AddBorderStyle(doc);
            var b2 = ExcelHelper.AddBorderStyle(doc);
            Assert.AreNotEqual(b1, b2);
            // add cellformat
            var c1 = ExcelHelper.AddCellFormatStyle(doc, 0, 0, 0);
            var c2 = ExcelHelper.AddCellFormatStyle(doc, 1, 1, 1);
            Assert.AreNotEqual(c1, c2);
            stylesSheet.Save();
        }
    }
}
