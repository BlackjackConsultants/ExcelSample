using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelExamples.Helpers {
    /// <summary>
    /// Helper class for reading Excel 2007 file format.  This class does not use OpenXml SDK
    /// </summary>
    public class ExcelXmlHelper {
        /// <summary>
        /// Retrieves an excel table from a given sheet in an xlsx file
        /// </summary>
        /// <param name="fileName">xlsx file name</param>
        /// <param name="sheetName">Sheet name (e.g. "Sheet1"</param>
        /// <returns>table[rows][cols] containing all the table information for the given sheet in the given file
        /// Cells with no information are returned as null. The table starts at the first element containing data in the sheet.</returns>
        public static List<List<string>> XLGetTable(string fileName, string sheetName) {
            //  Return the value of the specified cell.
            const string documentRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
            const string worksheetSchema = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
            const string sharedStringSchema = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            List<List<string>> table = new List<List<string>>();

            //  Retrieve the stream containing the requested
            //  worksheet's info:
            using (Package xlPackage = Package.Open(fileName, FileMode.Open, FileAccess.Read)) {
                PackagePart documentPart = null;
                Uri documentUri = null;

                //  Get the main document part (workbook.xml).
                foreach (System.IO.Packaging.PackageRelationship relationship in xlPackage.GetRelationshipsByType(documentRelationshipType)) {
                    //  There should only be one document part in the package. 
                    documentUri = PackUriHelper.ResolvePartUri(new Uri("/", UriKind.Relative), relationship.TargetUri);
                    documentPart = xlPackage.GetPart(documentUri);
                    //  There should only be one instance, but get out no matter what.
                    break;
                }

                if (documentPart == null) {
                    return table;
                }

                // Load the contents of the workbook.
                XmlDocument doc = new XmlDocument();
                doc.Load(documentPart.GetStream());

                //  Create a namespace manager, so you can search.
                //  Add a prefix (d) for the default namespace.
                NameTable nt = new NameTable();
                XmlNamespaceManager nsManager = new XmlNamespaceManager(nt);
                nsManager.AddNamespace("d", worksheetSchema);
                nsManager.AddNamespace("s", sharedStringSchema);

                string searchString = string.Format("//d:sheet[@name='{0}']", sheetName);
                XmlNode sheetNode = doc.SelectSingleNode(searchString, nsManager);
                if (sheetNode == null) {
                    // sheet not found
                    return table;
                }

                //  Get the relId attribute:
                XmlAttribute relationAttribute = sheetNode.Attributes["r:id"];
                if (relationAttribute == null) {
                    return table;
                }

                string relId = relationAttribute.Value;

                //  First, get the relation between the document and the sheet.
                PackageRelationship sheetRelation = documentPart.GetRelationship(relId);
                Uri sheetUri = PackUriHelper.ResolvePartUri(documentUri, sheetRelation.TargetUri);
                PackagePart sheetPart = xlPackage.GetPart(sheetUri);

                //  Load the contents of the workbook.
                XmlDocument sheetDoc = new XmlDocument(nt);
                sheetDoc.Load(sheetPart.GetStream());

                XmlNode dimensionNode = sheetDoc.SelectSingleNode("//d:dimension", nsManager);

                if (dimensionNode == null) {
                    throw new InvalidDataException("Dimension information not found in excel spreadsheet file");
                }

                XmlAttribute dimensionRefAttr = dimensionNode.Attributes["ref"];
                string dimensions;
                if (dimensionRefAttr == null) {
                    throw new InvalidDataException("Dimension ref information not found in excel spreadsheet file");
                }
                dimensions = dimensionRefAttr.Value;

                int startRow, startCol, endRow, endCol, rows, cols;
                ParseDimensions(dimensions, out startRow, out startCol, out endRow, out endCol);
                rows = endRow - startRow + 1;
                cols = endCol - startCol + 1;

                if ((rows == 0) || (cols == 0)) {
                    return table;
                }

                // uncomment the lines below to output the sheet information in a separate xml file
                //                XmlTextWriter writer = new XmlTextWriter(@"output.xml", System.Text.Encoding.UTF8);
                //                sheetDoc.WriteTo(writer);
                //                writer.Close();

                XmlNode sheetDataNode = sheetDoc.SelectSingleNode("//d:sheetData", nsManager);
                if (sheetDataNode == null) {
                    throw new InvalidDataException("Sheet data information not found in excel spreadsheet file");
                }

                table.Capacity = rows;
                for (int i = 0; i < rows; i++) {
                    List<string> colList = new List<string>(cols);
                    for (int j = 0; j < cols; j++) {
                        colList.Add(null);
                    }
                    table.Add(colList);
                }

                foreach (XmlNode row in sheetDataNode.ChildNodes) {
                    if (row.LocalName != "row") {
                        continue;
                    }

                    XmlAttribute rowRAttr = row.Attributes["r"];
                    if (rowRAttr == null) {
                        throw new InvalidDataException("Dimension ref information not found in excel spreadsheet file");
                    }

                    int rowIndex = int.Parse(rowRAttr.Value) - startRow;
                    // rowIndex is valid index inside List<List<string>> table

                    List<string> colList = table[rowIndex]; // columns for this table
                    foreach (XmlNode col in row.ChildNodes) {
                        if (col.LocalName != "c") {
                            continue;
                        }

                        XmlAttribute rAttr = col.Attributes["r"];
                        if (rAttr == null) {
                            throw new InvalidDataException("No r attribute found for col in xlxs data");
                        }

                        int colIndex, tempI;
                        GetColNumber(rAttr.Value, out colIndex, out tempI);
                        colIndex -= startCol;
                        colList[colIndex] = GetValue(col, nsManager, documentPart, documentUri, xlPackage, nt);
                    }
                }

                // uncomment the code below to find the value of a specific cell directly (e.g. A1)
                //XmlNode cellNode = sheetDoc.SelectSingleNode(string.Format("//d:sheetData/d:row/d:c[@r='{0}']", "A1"), nsManager);
                //if (cellNode != null) {
                //    string cellValue = GetValue(cellNode, nsManager, documentPart, documentUri, xlPackage, nt);
                //}
            }

            return table;
        }

        /// <summary>
        /// Retrieves the value as string of a cell node in an xlsx file
        /// </summary>
        private static string GetValue(XmlNode cellNode, XmlNamespaceManager nsManager, PackagePart documentPart, Uri documentUri, Package xlPackage, NameTable nt) {
            const string sharedStringsRelationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
            const string sharedStringSchema = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

            string cellValue = null;

            //  Retrieve the value. The value may be stored within 
            //  this element. If the "t" attribute contains "s", then
            //  the cell contains a shared string, and you must look 
            //  up the value individually.
            XmlAttribute typeAttr = cellNode.Attributes["t"];
            string cellType = string.Empty;
            if (typeAttr != null) {
                cellType = typeAttr.Value;
            }

            XmlNode valueNode = cellNode.SelectSingleNode("d:v", nsManager);
            if (valueNode != null) {
                cellValue = valueNode.InnerText;
            }

            //  Check the cell type. At this point, this code only checks
            //  for booleans and strings individually.
            if (cellType == "b") {
                if (cellValue == "1") {
                    cellValue = "TRUE";
                } else {
                    cellValue = "FALSE";
                }
            } else if (cellType == "s") {
                //  Go retrieve the actual string from the associated string file.
                foreach (System.IO.Packaging.PackageRelationship stringRelationship in documentPart.GetRelationshipsByType(sharedStringsRelationshipType)) {
                    //  There should only be one shared string reference, so you'll exit this loop immediately.
                    Uri sharedStringsUri = PackUriHelper.ResolvePartUri(documentUri, stringRelationship.TargetUri);
                    PackagePart stringPart = xlPackage.GetPart(sharedStringsUri);
                    if (stringPart != null) {
                        //  Load the contents of the shared strings.
                        XmlDocument stringDoc = new XmlDocument(nt);
                        stringDoc.Load(stringPart.GetStream());

                        //  Add the string schema to the namespace manager:
                        nsManager.AddNamespace("s", sharedStringSchema);

                        int requestedString = Convert.ToInt32(cellValue);
                        string strSearch = string.Format("//s:sst/s:si[{0}]", requestedString + 1);
                        XmlNode stringNode = stringDoc.SelectSingleNode(strSearch, nsManager);
                        if (stringNode != null) {
                            cellValue = stringNode.InnerText;
                        }
                    }
                }
            }

            return cellValue;
        }

        /// <summary>
        /// Returns column index given an excel column as string (e.g. "AB" -> 28)
        /// </summary>
        private static int ColNumFromStr(string str) {
            str = str.ToUpper();
            int result = 0;
            int strLen = str.Length;
            int mul = 1;

            for (int i = strLen - 1; i >= 0; i--) {
                result += (mul * (str[i] - 'A' + 1));
                mul *= 26;
            }

            return result;
        }

        /// <summary>
        /// Retrieves the column number given a string from Excel (e.g. A5 or C4)
        /// </summary>
        private static void GetColNumber(string rowColStr, out int col, out int indexAfterColInStr) {
            int len = rowColStr.Length;
            for (indexAfterColInStr = 0; indexAfterColInStr < len; indexAfterColInStr++) {
                if (!char.IsLetter(rowColStr[indexAfterColInStr])) {
                    break;
                }
            }

            if (indexAfterColInStr == 0) {
                throw new InvalidDataException("xlsx GetColNumber(): invalid rowColStr");
            }

            col = ColNumFromStr(rowColStr.Substring(0, indexAfterColInStr));
        }

        /// <summary>
        /// Converts row and col string in Excel (something like C1 or A4) to 1-based indexes
        /// </summary>
        private static void GetRowColNumbers(string rowColStr, out int row, out int col) {
            int i = 0;
            while (char.IsLetter(rowColStr[i])) {
                i++;
            }

            int indexAfterColInStr;
            GetColNumber(rowColStr, out col, out indexAfterColInStr);
            row = int.Parse(rowColStr.Substring(indexAfterColInStr));
        }

        /// <summary>
        /// Converts dimensions in Excel (e.g. C3:D4) to 1-based indexes
        /// </summary>
        private static void ParseDimensions(string dimensions, out int startRow, out int startCol,
            out int endRow, out int endCol) {

            int colonIndex = dimensions.IndexOf(':');
            if (colonIndex < 0) {
                throw new InvalidDataException("Invalid xlss dimension information: colon not found");
            }

            GetRowColNumbers(dimensions.Substring(0, colonIndex), out startRow, out startCol);
            GetRowColNumbers(dimensions.Substring(colonIndex + 1), out endRow, out endCol);
        }
    }
}
