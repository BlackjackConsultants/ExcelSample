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
            int sc = ExcelHelper.GetStyleCount(doc);
            Assert.AreEqual(sc, 1);
            ExcelHelper.AddStyle(doc, 14, "Arial", "FFFFFFFF", "00000000", "22222222", 0, 0, 0);
            int sc2 = ExcelHelper.GetStyleCount(doc);
            Assert.AreEqual(sc2, 2);
        }
    }
}
