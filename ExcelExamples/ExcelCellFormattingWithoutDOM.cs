using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelExamples.Helpers;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using X14 = DocumentFormat.OpenXml.Office2010.Excel;

namespace ExcelExamples {
    [TestClass]
    public class ExcelCellFormattingWithoutDOM {
        [TestMethod]
        public void ChangeCellColor() {
            string zipFileName = @"ExcelFileExtract\sample.xlsx";
            using (Package package = Package.Open(zipFileName, FileMode.Open)) {

            }
            Assert.IsTrue(File.Exists(zipFileName));
        }
    }
}
