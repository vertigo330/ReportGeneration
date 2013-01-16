using System.Collections.Generic;
using System.Xml;
using RemedyRoom.FceAutomation.Services.DomainObjects.Reports;

namespace RemedyRoom.FceAutomation.Services.ReportGenerator
{
    public class WordDocumentReportGenerator : IReportGenerator
    {
        public Report GenerateReport(XmlDocument source, IDictionary<string, string> options)
        {
            //EXTRACT
            //Parses source data using the appropriate parser. Pass in the parser?

            //TRANSFORM
            //Apply transformation rules

            //LOAD
            //Use the word document writer to write the data

            return new Report();
        }
    }
}