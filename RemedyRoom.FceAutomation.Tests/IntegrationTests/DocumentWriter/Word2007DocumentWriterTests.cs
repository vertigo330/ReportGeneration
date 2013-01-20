using Microsoft.VisualStudio.TestTools.UnitTesting;
using RemedyRoom.FceAutomation.Services.DocumentWriter;

namespace RemedyRoom.FceAutomation.Tests.IntegrationTests.DocumentWriter
{
    [TestClass]
    public class Word2007DocumentWriterTests
    {
        [TestMethod]

        // ReSharper disable InconsistentNaming
        public void AppendTable_WhenContentControlFoundAndContentControlContainsTable_WriteTabularDataToTable()
        // ReSharper restore InconsistentNaming
        {
            //Arrange
            var writer = new Word2007DocumentWriter
            (
                @"C:\Projects\RemedyRoom.FceAutomation\fce_partial_template.docx",
                @"C:\Projects\RemedyRoom.FceAutomation\fce_partial_report.docx"
            );
            string[,] tabularTestData = { { "a1", "a2", "a3", "a4", "a5" }, { "b1", "b2", "b3", "b4", "b5" }, { "c1", "c2", "c3", "c4", "c5" } };

            //Act
            using (writer)
            {
                writer.AppendTable("ClericalProductivityTable", tabularTestData);    
            }
            
            //Assert
            //???
        }
    }
}
