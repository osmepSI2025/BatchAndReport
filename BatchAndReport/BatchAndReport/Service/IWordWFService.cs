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
        byte[] GenCreateWFStatus(List<WFCreateProcessStatusModels> model);
        byte[] GenInternalControlSystem(List<WFInternalControlProcessModels> model);
        Task<byte[]> GenInternalControlSystemWord(List<WFInternalControlProcessModels> model, WFSubProcessDetailModels detail2);
        Task<byte[]> GenWorkProcessPoint(WFSubProcessDetailModels model);
        Task<byte[]> GenWorkProcessPointHtmlToPdf(WFSubProcessDetailModels model);
        //  byte[] GenWorkProcessPointPreview();
        byte[] GenWFProcessDetail(WFProcessDetailModels model);
    }
}
