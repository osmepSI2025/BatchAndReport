using BatchAndReport.Models;
using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace BatchAndReport.Services
{
    public interface IWordEContract_AllowanceService
    {

        byte[] ConvertWordToPdf(byte[] wordBytes);
        byte[] GenerateWordContactAllowance();

    }

}

