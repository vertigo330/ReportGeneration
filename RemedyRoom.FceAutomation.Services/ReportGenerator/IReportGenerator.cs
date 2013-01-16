using System.Collections.Generic;
using System.Xml;
using RemedyRoom.FceAutomation.Services.DomainObjects;
using RemedyRoom.FceAutomation.Services.DomainObjects.Reports;

namespace RemedyRoom.FceAutomation.Services.ReportGenerator
{
    public interface IReportGenerator
    {
        Report GenerateReport(XmlDocument source, IDictionary<string, string> options);
    }
}