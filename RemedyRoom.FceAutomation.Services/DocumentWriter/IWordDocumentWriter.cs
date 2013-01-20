namespace RemedyRoom.FceAutomation.Services.DocumentWriter
{
    public interface IWordDocumentWriter
    {
        void WriteTextToBookmark();
        void AppendTable(string contentControlTag, string[,] tabularData);
        void AppendChartData(string contentControlTag, string[,] tabularData);
    }
}