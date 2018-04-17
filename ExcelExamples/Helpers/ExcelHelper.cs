using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.Helpers {
    public static class ExcelHelper {
        public static void SetCellBorder(WorksheetPart workSheetPart, WorkbookPart workbookPart, Cell cell) {
            ////Cell cell = GetCell(workSheetPart, "B2");

            CellFormat cellFormat = cell.StyleIndex != null ? GetCellFormat(workbookPart, cell.StyleIndex).CloneNode(true) as CellFormat : new CellFormat();
            cellFormat.FillId = InsertFill(workbookPart, GenerateFill());
            cellFormat.BorderId = InsertBorder(workbookPart, GenerateBorder());

            cell.StyleIndex = InsertCellFormat(workbookPart, cellFormat);
        }

        private static Border GenerateBorder() {
            Border border2 = new Border();

            LeftBorder leftBorder2 = new LeftBorder() { Style = BorderStyleValues.Thin };
            Color color1 = new Color() { Indexed = (UInt32Value)64U };

            leftBorder2.Append(color1);

            RightBorder rightBorder2 = new RightBorder() { Style = BorderStyleValues.Thin };
            Color color2 = new Color() { Indexed = (UInt32Value)64U };

            rightBorder2.Append(color2);

            TopBorder topBorder2 = new TopBorder() { Style = BorderStyleValues.Thin };
            Color color3 = new Color() { Indexed = (UInt32Value)64U };

            topBorder2.Append(color3);

            BottomBorder bottomBorder2 = new BottomBorder() { Style = BorderStyleValues.Thin };
            Color color4 = new Color() { Indexed = (UInt32Value)64U };

            bottomBorder2.Append(color4);
            DiagonalBorder diagonalBorder2 = new DiagonalBorder();

            border2.Append(leftBorder2);
            border2.Append(rightBorder2);
            border2.Append(topBorder2);
            border2.Append(bottomBorder2);
            border2.Append(diagonalBorder2);

            return border2;
        }

        private static Fill GenerateFill() {
            Fill fill = new Fill();

            PatternFill patternFill = new PatternFill() { PatternType = PatternValues.Solid };
            ForegroundColor foregroundColor1 = new ForegroundColor() { Rgb = "FFFFFF00" };
            BackgroundColor backgroundColor1 = new BackgroundColor() { Indexed = (UInt32Value)64U };

            patternFill.Append(foregroundColor1);
            patternFill.Append(backgroundColor1);

            fill.Append(patternFill);

            return fill;
        }

        private static uint InsertBorder(WorkbookPart workbookPart, Border border) {
            Borders borders = workbookPart.WorkbookStylesPart.Stylesheet.Elements<Borders>().First();
            borders.Append(border);
            return (uint)borders.Count++;
        }

        private static uint InsertFill(WorkbookPart workbookPart, Fill fill) {
            Fills fills = workbookPart.WorkbookStylesPart.Stylesheet.Elements<Fills>().First();
            fills.Append(fill);
            return (uint)fills.Count++;
        }

        private static Cell GetCell(WorksheetPart workSheetPart, string cellAddress) {
            return workSheetPart.Worksheet.Descendants<Cell>()
                                        .SingleOrDefault(c => cellAddress.Equals(c.CellReference));
        }

        private static CellFormat GetCellFormat(WorkbookPart workbookPart, uint styleIndex) {
            return workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().First().Elements<CellFormat>().ElementAt((int)styleIndex);
        }

        private static uint InsertCellFormat(WorkbookPart workbookPart, CellFormat cellFormat) {
            CellFormats cellFormats = workbookPart.WorkbookStylesPart.Stylesheet.Elements<CellFormats>().First();
            cellFormats.Append(cellFormat);
            return (uint)cellFormats.Count++;
        }
    }
}
