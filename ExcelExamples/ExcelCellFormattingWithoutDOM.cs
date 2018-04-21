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

            string zipFileName = @"test.zip";

            using (Package package = ZipPackage.Open(zipFileName, FileMode.Create)) {
                string startFolder = @"ExcelFileExtract\Sample";

                foreach (string currentFile in Directory.GetFiles(startFolder, "*.*", SearchOption.AllDirectories)) {
                    System.Diagnostics.Debug.WriteLine("------------------------------------------------------------------------------------------------------------");
                    System.Diagnostics.Debug.WriteLine("Packing " + currentFile);
                    Uri relUri = PackageHelper.GetRelativeUri(currentFile);

                    PackagePart packagePart = package.CreatePart(relUri, System.Net.Mime.MediaTypeNames.Application.Octet, CompressionOption.Maximum);
                    using (FileStream fileStream = new FileStream(currentFile, FileMode.Open, FileAccess.Read)) {
                        if (packagePart != null)
                            PackageHelper.CopyStream(fileStream, packagePart.GetStream());
                    }
                    System.Diagnostics.Debug.WriteLine("PackagePart Uri: " + packagePart.Uri);
                }
            }

            Assert.IsTrue(File.Exists(zipFileName));
        }
    }
}
