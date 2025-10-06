using BatchAndReport.Entities;
using BatchAndReport.Models;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Threading.Tasks;

namespace BatchAndReport.DAO
{
    public class Econtract_Report_PMLDAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_PMLDAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_PMLModels?> GetPMLAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_PML_R314_60", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@PML_R314_60_ID_Input", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_PMLModels
                {
                    PML_R314_60_ID = reader["PML_R314_60_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["PML_R314_60_ID"]),
                    Contract_Number = reader["Contract_Number"] as string,
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("ContractSignDate")),
                    Contract_Sign_Address = reader["Contract_Sign_Address"] as string,
                    Contract_Organization = reader["Contract_Organization"] as string,
                    SignatoryName = reader["SignatoryName"] as string,
                    SignatoryPosition = reader["SignatoryPosition"] as string,
                    ContractorType = reader["ContractorType"] as string,
                    ContractorName = reader["ContractorName"] as string,
                    ContractorCompany = reader["ContractorCompany"] as string,
                    ContractorAddressNo = reader["ContractorAddressNo"] as string,
                    ContractorStreet = reader["ContractorStreet"] as string,
                    ContractorSubDistrict = reader["ContractorSubDistrict"] as string,
                    ContractorDistrict = reader["ContractorDistrict"] as string,
                    ContractorProvince = reader["ContractorProvince"] as string,
                    ContractorZipcode = reader["ContractorZipcode"] as string,
                    ContractorSignatoryName = reader["ContractorSignatoryName"] as string,
                    ContractorSignatoryPosition = reader["ContractorSignatoryPosition"] as string,
                    ContractorAuthorize = reader["ContractorAuthorize"] as string,
                    ContractorAuthDate = reader["ContractorAuthDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("ContractorAuthDate")),
                    ContractorAuthNumber = reader["ContractorAuthNumber"] as string,
                    RentalCopierBrand = reader["RentalCopierBrand"] as string,
                    RentalCopierModel = reader["RentalCopierModel"] as string,
                    RentalCopierNumber = reader["RentalCopierNumber"] as string,
                    RentalCopierAmount = reader["RentalCopierAmount"] == DBNull.Value ? null : Convert.ToInt32(reader["RentalCopierAmount"]),
                    RentalYears = reader["RentalYears"] == DBNull.Value ? null : Convert.ToInt32(reader["RentalYears"]),
                    RentalMonths = reader["RentalMonths"] == DBNull.Value ? null : Convert.ToInt32(reader["RentalMonths"]),
                    RentalStartDate = reader["RentalStartDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("RentalStartDate")),
                    RentalEndDate = reader["RentalEndDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("RentalEndDate")),
                    RatePerUnit = reader["RatePerUnit"] == DBNull.Value ? null : Convert.ToDecimal(reader["RatePerUnit"]),
                    RateTotal = reader["RateTotal"] == DBNull.Value ? null : Convert.ToDecimal(reader["RateTotal"]),
                    EstCopiesPerMonth = reader["EstCopiesPerMonth"] == DBNull.Value ? null : Convert.ToInt32(reader["EstCopiesPerMonth"]),
                    IfNotCopiesAmount = reader["IfNotCopiesAmount"] == DBNull.Value ? null : Convert.ToInt32(reader["IfNotCopiesAmount"]),
                    CopiesRate = reader["CopiesRate"] == DBNull.Value ? null : Convert.ToDecimal(reader["CopiesRate"]),
                    SaleBankName = reader["SaleBankName"] as string,
                    SaleBankBranch = reader["SaleBankBranch"] as string,
                    SaleBankAccountName = reader["SaleBankAccountName"] as string,
                    SaleBankAccountNumber = reader["SaleBankAccountNumber"] as string,
                    DeliveryLocation = reader["DeliveryLocation"] as string,
                    DeliveryType = reader["DeliveryType"] as string,
                    TotalDay = reader["TotalDay"] == DBNull.Value ? null : Convert.ToInt32(reader["TotalDay"]),
                    DeliveryDate = reader["DeliveryDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("DeliveryDate")),
                    NotiSendLocation = reader["NotiSendLocation"] as string,
                    NotiDaysBeforeDelivery = reader["NotiDaysBeforeDelivery"] == DBNull.Value ? null : Convert.ToInt32(reader["NotiDaysBeforeDelivery"]),
                    RespReplaceDays = reader["RespReplaceDays"] == DBNull.Value ? null : Convert.ToInt32(reader["RespReplaceDays"]),
                    MaintenancePermonth = reader["MaintenancePermonth"] == DBNull.Value ? null : Convert.ToInt32(reader["MaintenancePermonth"]),
                    MaintenanceInterval = reader["MaintenanceInterval"] == DBNull.Value ? null : Convert.ToInt32(reader["MaintenanceInterval"]),
                    CopierFixDays = reader["CopierFixDays"] == DBNull.Value ? null : Convert.ToInt32(reader["CopierFixDays"]),
                    ReplaceFixDays = reader["ReplaceFixDays"] == DBNull.Value ? null : Convert.ToInt32(reader["ReplaceFixDays"]),
                    FinePerDays = reader["FinePerDays"] == DBNull.Value ? null : Convert.ToDecimal(reader["FinePerDays"]),
                    GuaranteeType = reader["GuaranteeType"] as string,
                    GuaranteeAmount = reader["GuaranteeAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteeAmount"]),
                    GuaranteePercent = reader["GuaranteePercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteePercent"]),
                    NewGuaranteeDays = reader["NewGuaranteeDays"] == DBNull.Value ? null : Convert.ToInt32(reader["NewGuaranteeDays"]),
                    TeminationReplaceDays = reader["TeminationReplaceDays"] == DBNull.Value ? null : Convert.ToInt32(reader["TeminationReplaceDays"]),
                    LateFinePerDays = reader["LateFinePerDays"] == DBNull.Value ? null : Convert.ToDecimal(reader["LateFinePerDays"]),
                    EnforcementOfFineDays = reader["EnforcementOfFineDays"] == DBNull.Value ? null : Convert.ToInt32(reader["EnforcementOfFineDays"]),
                    OutstandingPeriodDays = reader["OutstandingPeriodDays"] == DBNull.Value ? null : Convert.ToInt32(reader["OutstandingPeriodDays"]),
                    OSMEP_Signer = reader["OSMEP_Signer"] as string,
                    OSMEP_Witness = reader["OSMEP_Witness"] as string,
                    Contract_Signer = reader["Contract_Signer"] as string,
                    Contract_Witness = reader["Contract_Witness"] as string,
                    CreatedDate = reader["CreatedDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("CreatedDate")),
                    CreateBy = reader["CreateBy"] as string,
                    UpdateDate = reader["UpdateDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("UpdateDate")),
                    UpdateBy = reader["UpdateBy"] as string,
                    Flag_Delete = reader["Flag_Delete"] as string,
                    AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? null : (bool?)Convert.ToBoolean(reader["AttorneyFlag"]),
                    AttorneyLetterDate = reader["AttorneyLetterDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("AttorneyLetterDate")),
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] as string,
                    CitizenId = reader["CitizenId"] as string,
                    CitizenCardRegisDate = reader["CitizenCardRegisDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("CitizenCardRegisDate")),
                    CitizenCardExpireDate = reader["CitizenCardExpireDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("CitizenCardExpireDate")),
                    CopierSendBackDays = reader["CopierSendBackDays"] == DBNull.Value ? null : Convert.ToInt32(reader["CopierSendBackDays"]),
                    Request_ID = reader["Request_ID"] as string,
                    Contract_Status = reader["Contract_Status"] as string
                };

                return detail;
            }
            catch (Exception ex)
            {
                return null; // Consider logging the exception
            }
        }

    }
}