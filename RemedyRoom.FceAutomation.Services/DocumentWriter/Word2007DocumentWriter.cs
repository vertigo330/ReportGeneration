using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

namespace RemedyRoom.FceAutomation.Services.DocumentWriter
{
    public class Word2007DocumentWriter : IWordDocumentWriter, IDisposable
    {
        private readonly WordprocessingDocument _document;

        public Word2007DocumentWriter(string path)
        {
            var settings = new OpenSettings
                            {
                                MarkupCompatibilityProcessSettings = new MarkupCompatibilityProcessSettings(MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
                            };
            _document = WordprocessingDocument.Open(path, true, settings);
        }

        public void WriteTextToBookmark()
        {
            throw new System.NotImplementedException();
        }

        public void UpdateTableWithData()
        {
            throw new System.NotImplementedException();
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
    }
}