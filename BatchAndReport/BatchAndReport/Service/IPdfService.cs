using BatchAndReport.Models;

namespace BatchAndReport.Services
{
    public interface IPdfService
    {
        Task<SMEProjectDetailModels?> GetProjectDetailAsync(string projectCode);
        byte[] GeneratePdf(SMEProjectDetailModels model);

    }
}
