namespace RemedyRoom.FceAutomation.Services.DocumentWriter
{
    public interface IWordDocumentWriter
    {
        void WriteTextToBookmark();
        void UpdateTableWithData();
        void WriteChartToBookmark();
    }
}