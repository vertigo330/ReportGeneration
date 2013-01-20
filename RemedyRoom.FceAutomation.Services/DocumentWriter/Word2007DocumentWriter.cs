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

        public void WriteTextToBookmark()
        {
            throw new System.NotImplementedException();
        }

        public void AppendTable(string contentControlTag, string[,] tabularData)
        {
            var contentControls = _document.MainDocumentPart.Document.Body.Descendants<SdtBlock>();
            var contentControlContainingTable = contentControls.Single(cc => cc.SdtProperties.GetFirstChild<Tag>().Val == contentControlTag);
            
            if (contentControlContainingTable == null) return;
            AppendRowsToTable(tabularData, contentControlContainingTable);
        }
        
        public void WriteChartToBookmark()
        {
            throw new System.NotImplementedException();
        }
        
        public void Dispose()
        {
            _document.Close();
            _document.Dispose();
        }

        private static void AppendRowsToTable(string[,] tabularData, OpenXmlElement contentControlContainingTable)
        {
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
    }
}