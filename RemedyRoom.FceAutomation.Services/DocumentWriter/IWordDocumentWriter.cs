namespace RemedyRoom.FceAutomation.Services.DocumentWriter
{
    public interface IWordDocumentWriter
    {
        void InsertText(string contentControlTag, string text);
        void AppendRowsToTable(string contentControlTag, string[,] tabularData);
        void AppendColumnsToChartData(string contentControlTag, string[,] tabularData);
    }
}