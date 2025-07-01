using BatchAndReport.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BatchAndReport.Services
{
    public interface IWordWFService
    {
        byte[] GenAnnualWorkProcesses();
        byte[] ConvertWordToPdf(byte[] wordBytes);
        byte[] GenWorkSystem();

    }
}
