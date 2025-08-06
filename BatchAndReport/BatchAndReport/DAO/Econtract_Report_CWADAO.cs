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
    public class Econtract_Report_CWADAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_CWADAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_CWAModels?> GetCWAAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_CWA", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@CWA_ID_Input", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_CWAModels
                {
                    CWA_ID = reader["CWA_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["CWA_ID"]),
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
                    WorkName = reader["WorkName"] as string,
                    GuaranteeType = reader["GuaranteeType"] as string,
                    GuaranteeAmount = reader["GuaranteeAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteeAmount"]),
                    GuaranteePercent = reader["GuaranteePercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteePercent"]),
                    GuaranteePaymentPeriod = reader["GuaranteePaymentPeriod"] == DBNull.Value ? null : Convert.ToInt32(reader["GuaranteePaymentPeriod"]),
                    PaymentMethod = reader["PaymentMethod"] as string,
                    PrepaidAmount = reader["PrepaidAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["PrepaidAmount"]),
                    PrepaidPercents = reader["PrepaidPercents"] == DBNull.Value ? null : Convert.ToDecimal(reader["PrepaidPercents"]),
                    PrepaidGuaranteeType = reader["PrepaidGuaranteeType"] as string,
                    PrepaidDeductPercent = reader["PrepaidDeductPercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["PrepaidDeductPercent"]),
                    WorkStartDate = reader["WorkStartDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("WorkStartDate")),
                    WorkEndDate = reader["WorkEndDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("WorkEndDate")),
                    WarrantyPeriodYears = reader["WarrantyPeriodYears"] == DBNull.Value ? null : Convert.ToInt32(reader["WarrantyPeriodYears"]),
                    WarrantyPeriodMonths = reader["WarrantyPeriodMonths"] == DBNull.Value ? null : Convert.ToInt32(reader["WarrantyPeriodMonths"]),
                    SubcontractPenaltyPercent = reader["SubcontractPenaltyPercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["SubcontractPenaltyPercent"]),
                    FinePerDays = reader["FinePerDays"] == DBNull.Value ? null : Convert.ToDecimal(reader["FinePerDays"]),
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
                    DaysToRepairIn = reader["DaysToRepairIn"] == DBNull.Value ? null : Convert.ToInt32(reader["DaysToRepairIn"]),
                    Request_ID = reader["Request_ID"] == DBNull.Value ? null : Convert.ToInt32(reader["Request_ID"]),
                    Contract_Status = reader["Contract_Status"] as string,
                    Install_PayAMT = reader["Install_PayAMT"] == DBNull.Value ? null : Convert.ToDecimal(reader["Install_PayAMT"]),
                    Install_PayVat = reader["Install_PayVat"] == DBNull.Value ? null : Convert.ToDecimal(reader["Install_PayVat"]),
                    Install_Num = reader["Install_Num"] == DBNull.Value ? null : Convert.ToInt32(reader["Install_Num"])
                    ,BankAccountName = reader["BankAccountName"] as string,
                    BankAccountNumber = reader["BankAccountNum"] as string,
                    BankName = reader["BankName"] as string,
                    BankBranch = reader["BankBranch"] as string,
                };

                return detail;
            }
            catch (Exception ex)
            {
                return null; // Consider logging the exception
            }
        }
        public async Task<List<CWA_Installment>> GetCWAInstallmentAsync(string? id = "0")
        {
            try
            {
                var result = new List<CWA_Installment>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT 
                CWA_Inst_ID,
                CWA_ID,
                PayRound,
                TotalAmount,
                WorkName,
                DeliverDate,
                Flag_Inst_Final
            FROM CWA_Installment
            WHERE CWA_ID = @CWA_ID", connection);

                command.Parameters.AddWithValue("@CWA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new CWA_Installment
                    {
                        CWA_Inst_ID = reader["CWA_Inst_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["CWA_Inst_ID"]),
                        CWA_ID = reader["CWA_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["CWA_ID"]),
                        PayRound = reader["PayRound"] == DBNull.Value ? 0 : Convert.ToInt32(reader["PayRound"]),
                        TotalAmount = reader["TotalAmount"] == DBNull.Value ? 0 : Convert.ToDecimal(reader["TotalAmount"]),
                        WorkName = reader["WorkName"] as string,
                        DeliverDate = reader["DeliverDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("DeliverDate")),
                        Flag_Inst_Final = reader["Flag_Inst_Final"] == DBNull.Value ? false : Convert.ToBoolean(reader["Flag_Inst_Final"])
                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}