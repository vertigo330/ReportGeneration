using System.Collections.Generic;
using RemedyRoom.FceAutomation.Services.DomainObjects.WorkSamples;

namespace RemedyRoom.FceAutomation.Services.DocumentWriter
{
    public interface IWordDocumentWriter
    {
        void WriteTextToBookmark();
        void AppendTable(string contentControlTag, string[,] tabularData);
        void WriteChartToBookmark();
    }
}