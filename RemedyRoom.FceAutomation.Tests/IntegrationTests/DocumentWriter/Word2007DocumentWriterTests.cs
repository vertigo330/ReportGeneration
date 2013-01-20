using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RemedyRoom.FceAutomation.Services.DocumentWriter;

namespace RemedyRoom.FceAutomation.Tests.IntegrationTests.DocumentWriter
{
    [TestClass]
    public class Word2007DocumentWriterTests
    {
        private const string DocTemplatePath = @"C:\Projects\RemedyRoom.FceAutomation\fce_partial_template.docx";
        private const string DocOutputPath = @"C:\Projects\RemedyRoom.FceAutomation\fce_partial_report.docx";

        #region AppendRowsToTable

        [TestMethod]
        // ReSharper disable InconsistentNaming
        public void AppendRowsToTable_WhenContentControlFoundAndContentControlContainsTable_WriteTabularDataToTable()
        // ReSharper restore InconsistentNaming
        {
            //Arrange
            const string contentControlTagName = "ClericalProductivityTable";
            var tabularTestData = GetTestTabularData();

            //Act
            using (var writer = new Word2007DocumentWriter(DocTemplatePath, DocOutputPath))
            {
                writer.AppendRowsToTable(contentControlTagName, tabularTestData);
            }

            //Assert
            using (var documentUnderTest = GetTestWordProcessingDocument(DocOutputPath))
            {
                var rowIndex = 0;
                var colIndex = 0;
                var contentControls = documentUnderTest.MainDocumentPart.Document.Body.Descendants<SdtBlock>();
                var contentControlContainingTable =
                    contentControls.Single(cc => cc.SdtProperties.GetFirstChild<Tag>().Val == contentControlTagName);
                var targetTable = contentControlContainingTable.Descendants<Table>().Single();

                //Has the correct number of rows
                Assert.IsTrue(targetTable.Elements<TableRow>().Skip(1).Count() == tabularTestData.GetLength(0));

                foreach (var tableRow in targetTable.Elements<TableRow>().Skip(1)) //Skip the header
                {
                    foreach (var tableCell in tableRow.Descendants<TableCell>())
                    {
                        var paragraph = tableCell.Elements<Paragraph>().First();
                        var run = paragraph.Elements<Run>().First();
                        var text = run.Elements<Text>().First();

                        //Contents of the table cell are the same as the contents of the tabular data
                        Assert.AreEqual(text.Text, tabularTestData[rowIndex, colIndex]);
                        colIndex++;
                    }
                    colIndex = 0;
                    rowIndex++;
                }
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        // ReSharper disable InconsistentNaming
        public void AppendRowsToTable_WhenContentControlFoundAndContentControlDoenstContainTable_ThrowsException()
        // ReSharper restore InconsistentNaming
        {
            //Arrange
            const string contentControlTagName = "EmptyContentControl";
            var tabularTestData = GetTestTabularData();

            //Act
            try
            {
                using (var writer = new Word2007DocumentWriter(DocTemplatePath, DocOutputPath))
                {
                    writer.AppendRowsToTable(contentControlTagName, tabularTestData);
                }
            }
            catch (Exception ex)
            {
                //Assert
                Assert.AreEqual(ex.Message, "The content control specified does not contain a table to append to");
                throw;
            }
        }

        [TestMethod]
        [ExpectedException(typeof(InvalidOperationException))]
        // ReSharper disable InconsistentNaming
        public void AppendRowsToTable_WhenContentControlNotFound_ThrowsException()
        // ReSharper restore InconsistentNaming
        {
            //Arrange
            const string contentControlTagName = "NonExistantContentControl";
            var tabularTestData = GetTestTabularData();

            //Act
            try
            {
                using (var writer = new Word2007DocumentWriter(DocTemplatePath, DocOutputPath))
                {
                    writer.AppendRowsToTable(contentControlTagName, tabularTestData);
                }
            }
            catch (Exception ex)
            {
                //Assert
                Assert.AreEqual(ex.Message, "The content control specified does not exist");
                throw;
            }
        }

        #endregion

        #region helper methods

        private string[,] GetTestTabularData()
        {
            return new[,]
                {
                    {"a1", "a2", "a3", "a4", "a5"}, 
                    {"b1", "b2", "b3", "b4", "b5"},
                    {"c1", "c2", "c3", "c4", "c5"}
                };
        }

        private WordprocessingDocument GetTestWordProcessingDocument(string documenPath)
        {
            var settings = new OpenSettings
                               {
                                   MarkupCompatibilityProcessSettings =
                                       new MarkupCompatibilityProcessSettings(
                                       MarkupCompatibilityProcessMode.ProcessAllParts, FileFormatVersions.Office2007)
                               };
            return WordprocessingDocument.Open(DocOutputPath, true, settings);
        }

        #endregion

    }
}
