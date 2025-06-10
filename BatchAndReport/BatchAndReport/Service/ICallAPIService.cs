using BatchAndReport.Models;

namespace BatchAndReport.Services
{
    public interface ICallAPIService
    {
        Task<string> GetDataApiAsync(MapiInformationModels apiModels, object xdata);
        Task<string> GetDataByParamApiAsync(MapiInformationModels apiModels, string typeValue);
        Task<string> GetDataEmpMovementApiAsync(MapiInformationModels apiModels, string empId);

    }
}
