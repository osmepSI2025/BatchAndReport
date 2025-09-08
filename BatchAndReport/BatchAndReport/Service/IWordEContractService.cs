using BatchAndReport.Models;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BatchAndReport.Services
{
    public interface IWordEContractService
    {
        byte[] GenJointContractAgreement(ConJointContractModels model);
        byte[] ConvertWordToPdf(byte[] wordBytes);
        byte[] GenImportContract(IEnumerable<ImportContractModels> model);

    }
}
