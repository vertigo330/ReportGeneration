namespace RemedyRoom.FceAutomation.Services.DocumentWriter
{
    public class Word2010DocumentWriter : IWordDocumentWriter
    {
        public void InsertText(string contentControlTag, string text)
        {
            throw new System.NotImplementedException();
        }

        public void AppendRowsToTable(string contentControlTag, string[,] tabularData)
        {
            throw new System.NotImplementedException();
        }

        public void AppendColumnsToChartData(string contentControlTag, string[,] tabularData)
        {
            throw new System.NotImplementedException();
        }
    }
}