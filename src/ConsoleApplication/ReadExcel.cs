using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Collections.Generic;

namespace ConsoleApplication
{
    class ReadExcel
    {
        public static void ReadExcelFile(string fileName)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                // open the document read-only

                SharedStringTable sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                string cellValue = null;

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.WorksheetParts.Where(x => x.Uri.OriginalString.Contains("sheet1")).Single();

                foreach (SheetData sheetData in worksheetPart.Worksheet.Elements<SheetData>())
                    {
                        if (sheetData.HasChildren)
                        {
                            foreach (Row row in sheetData.Elements<Row>())
                            {
                                foreach (Cell cell in row.Elements<Cell>())
                                {
                                    cellValue = cell.InnerText;

                                    if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                                    {
                                        Console.WriteLine("cell val: " + sharedStringTable.ElementAt(Int32.Parse(cellValue)).InnerText);

                                }
                                    else
                                    {
                                        Console.WriteLine("else: " + cellValue);
                                    }
                                }
                            }
                        }
                    }
                document.Close();
            }
        }
    }
}
