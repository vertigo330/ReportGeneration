using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace RemedyRoom.FceAutomation.Services.DocumentWriter
{
    public class Word2007DocumentWriter : IWordDocumentWriter, IDisposable
    {
        private readonly WordprocessingDocument _document;

        public Word2007DocumentWriter(string documentTemplatePath, string documentOutputPath)
        {
            var settings = new OpenSettings
                            {
                                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
                            };

            File.Copy(documentTemplatePath, documentOutputPath, true);

            _document = WordprocessingDocument.Open(documentOutputPath, true, settings);
        }

        public void AppendRowsToTable(string contentControlTag, string[,] tabularData)
        {
            var contentControlContainingTable = GetContentControlByTagName(contentControlTag);
            AppendRowsToTableInContentControl(tabularData, contentControlContainingTable);
        }

        public void AppendColumnsToChartData(string contentControlTag, string[,] tabularData)
        {
            var contentControlContainingTable = GetContentControlByTagName(contentControlTag);
            AppendColumnsToChartDataInContentControl(tabularData, contentControlContainingTable);
        }

        public void InsertText(string contentControlTag, string text)
        {
            throw new System.NotImplementedException();
        }

        public void Dispose()
        {
            _document.Close();
            _document.Dispose();
        }

        private SdtBlock GetContentControlByTagName(string contentControlTag)
        {
            var contentControls = _document.MainDocumentPart.Document.Body.Descendants<SdtBlock>();
            var contentControl = contentControls.FirstOrDefault(cc => cc.SdtProperties.GetFirstChild<Tag>().Val == contentControlTag);

            if (contentControl == null)
                throw new InvalidOperationException("The content control specified does not exist");

            return contentControl;
        }

        private static void AppendRowsToTableInContentControl(string[,] tabularData, OpenXmlElement contentControlContainingTable)
        {
            if (!contentControlContainingTable.Descendants<Table>().Any())
                throw new InvalidOperationException("The content control specified does not contain a table to append to");

            var targetTable = contentControlContainingTable.Descendants<Table>().Single();
            for (var rowIndex = 0; rowIndex < tabularData.GetLength(0); rowIndex++)
            {
                var tableRow = targetTable.Elements<TableRow>().Last().CloneNode(true);
                for (var colIndex = 0; colIndex < tabularData.GetLength(1); colIndex++)
                {
                    var tableCell = tableRow.Descendants<TableCell>().ElementAt(colIndex);
                    var tableParagraph = tableCell.Elements<Paragraph>().First();
                    var tableText = new Text(tabularData[rowIndex, colIndex]);
                    var tableRun = new Run(tableText);

                    tableParagraph.RemoveAllChildren<Run>();
                    tableParagraph.AppendChild(tableRun);
                }
                targetTable.AppendChild(tableRow);
            }
        }

        private void AppendColumnsToChartDataInContentControl(string[,] tabularData, OpenXmlElement contentControlContainingChart)
        {
            if (!contentControlContainingChart.Descendants<ChartReference>().Any())
                throw new InvalidOperationException("The content control specified does not contain a chart to append to");

            var targetChartReference = contentControlContainingChart.Descendants<ChartReference>().Single();

            if (targetChartReference.Id == null) return;

            var relationshipId = targetChartReference.Id;
            var part = _document.MainDocumentPart.Parts.FirstOrDefault(p => p.RelationshipId == relationshipId);
            if (part == null) return;

            var chartPart = part.OpenXmlPart as ChartPart;

            if (chartPart != null)
            {
                var externalData = chartPart.ChartSpace.Elements<ExternalData>().FirstOrDefault();

                if (externalData != null)
                {
                    var extDataId = externalData.Id;
                    var embeddedPackagePart = chartPart.Parts.FirstOrDefault(p => p.RelationshipId == extDataId);

                    if (embeddedPackagePart != null)
                    {
                        var embeddedPackage = embeddedPackagePart.OpenXmlPart as EmbeddedPackagePart;
                        WriteDataToChartSource(embeddedPackage);
                    }
                }
            }
        }

        private void WriteDataToChartSource(OpenXmlPart embeddedPackage)
        {
            using (var packageStream = embeddedPackage.GetStream())
            using (var memStream = new MemoryStream())
            {
                CopyStream(packageStream, memStream);
                using (var spreadsheetDocument = SpreadsheetDocument.Open(memStream, true))
                {
                    var sheet = (Sheet)spreadsheetDocument.WorkbookPart.Workbook.Sheets.FirstOrDefault();
                    if (sheet != null)
                    {
                        var sheetId = sheet.Id;
                        var part = spreadsheetDocument.WorkbookPart.Parts.FirstOrDefault(prt => prt.RelationshipId == sheetId);

                        if (part != null)
                        {
                            var worksheetPart = (WorksheetPart)part.OpenXmlPart;


                            SharedStringTablePart shareStringPart;
                            if (spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().Count() > 0)
                            {
                                shareStringPart = spreadsheetDocument.WorkbookPart.GetPartsOfType<SharedStringTablePart>().First();
                            }
                            else
                            {
                                shareStringPart = spreadsheetDocument.WorkbookPart.AddNewPart<SharedStringTablePart>();
                            }

                            // Insert the text into the SharedStringTablePart.
                            int index = InsertSharedStringItem("Jonnie Test 1", shareStringPart);

                            var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                            if (sheetData != null)
                            {
                                var rows = sheetData.Elements<Row>();
                                var row = rows.Skip(1).First();

                                var cell = row.Elements<Cell>().FirstOrDefault(c => c.CellReference == "B2");

                                if(cell == null)
                                {
                                    //Create the cell
                                    cell = InsertCellInWorksheet("B", 1, worksheetPart);
                                    cell.CellValue = new CellValue(Convert.ToString(index));
                                    cell.DataType = new EnumValue<CellValues>(CellValues.SharedString);

                                    var cell2 = InsertCellInWorksheet("B", 2, worksheetPart);
                                    cell2.CellValue = new CellValue("10.56");
                                    cell2.DataType = new EnumValue<CellValues>(CellValues.String);

                                    var cell3 = InsertCellInWorksheet("B", 3, worksheetPart);
                                    cell3.CellValue = new CellValue("5.58");
                                    cell3.DataType = new EnumValue<CellValues>(CellValues.String);
                                    //cell.StyleIndex = 2;

                                    
                                }
                                else
                                {
                                    //Ammend the cell
                                }
                                
                                
                            
                                //var refCell = row.Elements<Cell>().First(c => c.CellReference == "A2");

                                //row.InsertBefore(cell, refCell);
                            }
                            worksheetPart.Worksheet.Save();
                        }
                    }

                    using (var s = embeddedPackage.GetStream())
                        memStream.WriteTo(s);
                }
            }
        }

        public static void CopyStream(Stream input, Stream output)
        {
            byte[] buffer = new byte[32768];
            while (true)
            {
                int read = input.Read(buffer, 0, buffer.Length);
                if (read <= 0)
                    return;
                output.Write(buffer, 0, read);
            }
        }

        #region experimental code

        private static int InsertSharedStringItem(string text, SharedStringTablePart shareStringPart)
        {
            // If the part does not contain a SharedStringTable, create one.
            if (shareStringPart.SharedStringTable == null)
            {
                shareStringPart.SharedStringTable = new SharedStringTable();
            }

            int i = 0;

            // Iterate through all the items in the SharedStringTable. If the text already exists, return its index.
            foreach (SharedStringItem item in shareStringPart.SharedStringTable.Elements<SharedStringItem>())
            {
                if (item.InnerText == text)
                {
                    return i;
                }

                i++;
            }

            // The text does not exist in the part. Create the SharedStringItem and return its index.
            shareStringPart.SharedStringTable.AppendChild(new SharedStringItem(new DocumentFormat.OpenXml.Spreadsheet.Text(text)));
            shareStringPart.SharedStringTable.Save();

            return i;
        }

        private static Cell InsertCellInWorksheet(string columnName, uint rowIndex, WorksheetPart worksheetPart)
        {
            var worksheet = worksheetPart.Worksheet;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            var cellReference = columnName + rowIndex;

            // If the worksheet does not contain a row with the specified row index, insert one.
            Row row;
            if (sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).Count() != 0)
            {
                row = sheetData.Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
            }
            else
            {
                row = new Row() { RowIndex = rowIndex };
                sheetData.Append(row);
            }

            // If there is not a cell with the specified column name, insert one.  
            if (row.Elements<Cell>().Where(c => c.CellReference.Value == columnName + rowIndex).Count() > 0)
            {
                return row.Elements<Cell>().Where(c => c.CellReference.Value == cellReference).First();
            }
            
            // Cells must be in sequential order according to CellReference. Determine where to insert the new cell.
            Cell refCell = null;
            foreach (Cell cell in row.Elements<Cell>())
            {
                if (string.Compare(cell.CellReference.Value, cellReference, true) > 0)
                {
                    refCell = cell;
                    break;
                }
            }

            Cell newCell = new Cell() { CellReference = cellReference };
            row.InsertBefore(newCell, refCell);

            worksheet.Save();
            return newCell;
        }

        private static void WriteDataToChartSourceExperimental(OpenXmlPart embeddedPackage)
        {
            var sheetName = "Sheet1";
            var addressName = "A1";
            string value = null;

            using (var packageStream = embeddedPackage.GetStream())
            {
                using (var memStream = new MemoryStream())
                {
                    CopyStream(packageStream, memStream);
                    using (var document = SpreadsheetDocument.Open(memStream, true))
                    {
                        WorkbookPart wbPart = document.WorkbookPart;

                        // Find the sheet with the supplied name, and then use that Sheet
                        // object to retrieve a reference to the appropriate worksheet.
                        Sheet theSheet =
                            wbPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName).FirstOrDefault();

                        if (theSheet == null)
                        {
                            throw new ArgumentException("sheetName");
                        }

                        // Retrieve a reference to the worksheet part, and then use its 
                        // Worksheet property to get a reference to the cell whose 
                        // address matches the address you supplied:
                        WorksheetPart wsPart = (WorksheetPart)(wbPart.GetPartById(theSheet.Id));
                        Cell theCell =
                            wsPart.Worksheet.Descendants<Cell>().Where(c => c.CellReference == addressName).
                                FirstOrDefault();

                        // If the cell does not exist, return an empty string:
                        if (theCell != null)
                        {
                            value = theCell.InnerText;

                            // If the cell represents a numeric value, you are done. 
                            // For dates, this code returns the serialized value that 
                            // represents the date. The code handles strings and Booleans
                            // individually. For shared strings, the code looks up the 
                            // corresponding value in the shared string table. For Booleans, 
                            // the code converts the value into the words TRUE or FALSE.
                            if (theCell.DataType != null)
                            {
                                switch (theCell.DataType.Value)
                                {
                                    case CellValues.SharedString:
                                        // For shared strings, look up the value in the shared 
                                        // strings table.
                                        var stringTable =
                                            wbPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                                        // If the shared string table is missing, something is 
                                        // wrong. Return the index that you found in the cell.
                                        // Otherwise, look up the correct text in the table.
                                        if (stringTable != null)
                                        {
                                            value = stringTable.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                                        }
                                        break;

                                    case CellValues.Boolean:
                                        switch (value)
                                        {
                                            case "0":
                                                value = "FALSE";
                                                break;
                                            default:
                                                value = "TRUE";
                                                break;
                                        }
                                        break;
                                }
                            }
                        }
                    }
                }
            }
        }

        #endregion

    }
}