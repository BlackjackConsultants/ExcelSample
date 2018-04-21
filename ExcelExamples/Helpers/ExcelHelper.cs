using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.Helpers {
    public class ExcelHelper{
        private SpreadsheetDocument doc;
        private WorkbookPart wbp;
        private WorksheetPart wsp;

        public ExcelHelper(string filename){
            doc = SpreadsheetDocument.Create(filename, SpreadsheetDocumentType.Workbook);
            wbp = doc.AddWorkbookPart();
            wsp = wbp.AddNewPart<WorksheetPart>();
            CreateStyles();
        }

        private void CreateStyles() {
            // add styles to sheet
            WorkbookStylesPart wbsp = wbp.AddNewPart<WorkbookStylesPart>();
            wbsp.Stylesheet = CreateStylesheet();
            wbsp.Stylesheet.Save();
        }

        public void Open(Stream stream){
            doc?.Close();
            doc = SpreadsheetDocument.Open(stream, true);
        }

        /// <summary>
        /// changes the color of a cell.
        /// </summary>
        /// <param name="color"></param>
        /// <param name="cellRange"></param>
        public void ColorCell(Color color, string cellRange, SheetData sheetData){
            foreach (Row r in sheetData.Elements<Row>()) {
                foreach (Cell c in r.Elements<Cell>()){
                    c.StyleIndex = 1;
                    //text = c.CellValue.Text;
                    //Console.Write(text + " 111111111");
                }
            }
            doc.WorkbookPart.Workbook.Save();
        }

        ~ExcelHelper() {
            this.Dispose();
        }

        public void Dispose(){
           doc.Close();
        }

        private Stylesheet CreateStylesheet() {
            Stylesheet stylesheet1 = new Stylesheet() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "x14ac" } };
            stylesheet1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            stylesheet1.AddNamespaceDeclaration("x14ac", "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac");

            // fills
            Fills fills = new Fills() { Count = (UInt32Value)5U };
            Fill fill = new Fill();
            PatternFill patternFill3 = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFF0000" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };
            patternFill3.Append(foregroundColor1);
            patternFill3.Append(backgroundColor1);
            fill.Append(patternFill3);

            fills.Append(fill);

            stylesheet1.Append(fills);
            return stylesheet1;
        }
    }
}
