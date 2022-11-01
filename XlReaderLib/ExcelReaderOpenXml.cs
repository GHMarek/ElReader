using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;


namespace XlReaderLib
{
    // Basic OpenXml example with xlsx files used here https://stackoverflow.com/questions/3321082/from-excel-to-datatable-in-c-sharp-with-open-xml/47600574#47600574
    public static class ExcelReaderOpenXml
    {
        /// <summary>
        /// Reads xlsx file path to DataTable.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        public static DataTable ReadXlsx(string path)
        {
            var dt = new DataTable();

            using (var ssDoc = SpreadsheetDocument.Open(path, false))
            {
                var sheets = ssDoc.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                var relationshipId = sheets.First().Id.Value;
                var worksheetPart = (WorksheetPart)ssDoc.WorkbookPart.GetPartById(relationshipId);
                var workSheet = worksheetPart.Worksheet;
                var sheetData = workSheet.GetFirstChild<SheetData>();
                //var rows = sheetData.Descendants<Row>().ToList();
                var rows = sheetData.Elements<Row>();

                Regex MyRegex = new Regex("[^a-z]", RegexOptions.IgnoreCase);
                Dictionary<string, int> cellRefIndexDict = new Dictionary<string, int>();
                int maxCellSum = 0;

                // Loop over all rows, including header.
                foreach (var row in rows) 
                {

                    var tempRow = dt.NewRow();
                    // Header.
                    if (row.RowIndex.Value == 1)
                    {

                        foreach (var cell in row.Descendants<Cell>())
                        {
                            var index = GetIndex(cell.CellReference);
                            // Atm header cells must be filled, can't have empty cells between nonempty.

                            // Store CellReference and index in dict, to save original structure.
                            cellRefIndexDict.Add(cell.CellReference, index);

                            // Add Columns
                            for (var i = dt.Columns.Count; i <= index; i++)
                                dt.Columns.Add(GetCellValue(ssDoc, cell));

                        }
                        string maxCol = cellRefIndexDict.Where(x => x.Value == cellRefIndexDict.Max(y => y.Value)).Select(z => z.Key).FirstOrDefault();
                        
                        foreach (var i in Regex.Replace(maxCol, @"[\d-]", string.Empty))
                        {
                            maxCellSum += Convert.ToInt32(Convert.ToChar(i));
                        }
                        continue;
                    }

                    // Next rows
                    int localIdx = 0;
                    

                    foreach (var cell in row.Descendants<Cell>())
                    {

                        // If CellReference is out of dict max, it means used cells are out of columns range.

                        int cellSum = 0;

                        foreach (var i in Regex.Replace(cell.CellReference.ToString(), @"[\d-]", string.Empty))
                        {
                            cellSum += Convert.ToInt32(Convert.ToChar(i));
                        }

                        if (cellSum > maxCellSum)
                        {
                            continue;
                        }


                        // Find real index of cell, using dictionary filled during columns iteration.
                        int cellIdx = cellRefIndexDict
                                    .Where(x => MyRegex.Replace(x.Key, @"") == MyRegex.Replace(cell.CellReference, @""))
                                    .Select(x => x.Value).FirstOrDefault();

                        // If real cell index equals var localIdx, then it is right column.
                        if (localIdx == cellIdx)
                        {
                            tempRow[localIdx] = GetCellValue(ssDoc, cell);
                            localIdx++;
                            continue;
                        }
                        // Else..
                        // It isn't right column but this cell won't happen again in loop.

                        // How many indices are missing.

                        int missingIdxCnt = cellIdx - localIdx;
                        var rangeIdx = Enumerable.Range(1, missingIdxCnt);

                        // Add null to row index. If localIdx == 5, and cellIdx == 8, adds 5+1...5+3    
                        rangeIdx.ToList().ForEach(x => tempRow[localIdx + x] = null);
                        tempRow[cellIdx] = GetCellValue(ssDoc, cell);

                        // Now localIdx is equal with real CellIdx.
                        localIdx = cellIdx;

                    }

                    dt.Rows.Add(tempRow);

                }
            }

            return dt;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            var stringTablePart = document.WorkbookPart.SharedStringTablePart;

            // Some used cells doesn't have value, yet reads in as non empty.
            if (cell.CellValue is null)
            {
                return null;
            }

            var value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
            }

            return value;
        }

        private static int GetIndex(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return -1;
            }
                
            int index = 0;

            foreach (var ch in name)
            {
                if (!char.IsLetter(ch))
                {
                    break;
                }

                int value = ch - 'A' + 1;
                index = value + index * 26;

            }

            return index - 1;
        }

    }
}
