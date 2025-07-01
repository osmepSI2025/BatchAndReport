using BatchAndReport.Models;

namespace BatchAndReport.Services
{
    public interface IWordService
    {
        byte[] GenerateWord(SMEProjectDetailModels model);
        byte[] ConvertWordToPdf(byte[] wordBytes);
        public byte[] GenerateSummaryWord(List<SMESummaryProjectModels> projects, List<SMEStrategyDetailModels> strategies, string year);

    }
}
