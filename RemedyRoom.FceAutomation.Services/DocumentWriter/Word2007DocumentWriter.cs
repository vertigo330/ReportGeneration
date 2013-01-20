using System;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

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
            AppendRowsToTable(tabularData, contentControlContainingTable);
        }
        
        public void AppendColumnsToChartData(string contentControlTag, string[,] tabularData)
        {
            var contentControlContainingTable = GetContentControlByTagName(contentControlTag);
            AppendColumnsToChartData(tabularData, contentControlContainingTable);
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
            var contentControlContainingTable = contentControls.FirstOrDefault(cc => cc.SdtProperties.GetFirstChild<Tag>().Val == contentControlTag);

            if (contentControlContainingTable == null)
                throw new InvalidOperationException("The content control specified does not exist");

            return contentControlContainingTable;
        }

        private static void AppendRowsToTable(string[,] tabularData, OpenXmlElement contentControlContainingTable)
        {
            if(!contentControlContainingTable.Descendants<Table>().Any())
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

        private static void AppendColumnsToChartData(string[,] tabularData, SdtBlock contentControlContainingTable)
        {
            
        }
    }
}