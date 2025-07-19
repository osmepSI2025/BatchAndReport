using BatchAndReport.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BatchAndReport.Services
{
    public interface IWordWFService
    {
        byte[] GenAnnualWorkProcesses(WFProcessDetailModels model);
        byte[] ConvertWordToPdf(byte[] wordBytes);
        byte[] GenWorkSystem(WorkSystemModels model);
        byte[] GenInternalControlSystem(List<WFInternalControlProcessModels> model);
        Task<byte[]> GenWorkProcessPoint(WFSubProcessDetailModels model);
        byte[] GenWorkProcessPointPreview();
        byte[] GenWFProcessDetail(WFProcessDetailModels model);
    }
}
