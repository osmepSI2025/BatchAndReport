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
    public class Econtract_Report_SMCDAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_SMCDAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_SMCModels?> GetSMCAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_SMC_R310_60", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@SMC_R310_60_ID_Input", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_SMCModels
                {
                    SMC_R310_60_ID = reader["SMC_R310_60_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["SMC_R310_60_ID"]),
                    Contract_Number = reader["Contract_Number"] as string,
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["ContractSignDate"]),
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
                    CompSetLocation = reader["CompSetLocation"] as string,
                    RentalBrandName = reader["RentalBrandName"] as string,
                    ServiceStartDate = reader["ServiceStartDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["ServiceStartDate"]),
                    ServiceEndDate = reader["ServiceEndDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["ServiceEndDate"]),
                    ServiceTotalYears = reader["ServiceTotalYears"] == DBNull.Value ? null : Convert.ToInt32(reader["ServiceTotalYears"]),
                    ServiceTotalMonths = reader["ServiceTotalMonths"] == DBNull.Value ? null : Convert.ToInt32(reader["ServiceTotalMonths"]),
                    ServiceTotalDays = reader["ServiceTotalDays"] == DBNull.Value ? null : Convert.ToInt32(reader["ServiceTotalDays"]),
                    ServiceFee = reader["ServiceFee"] == DBNull.Value ? null : Convert.ToDecimal(reader["ServiceFee"]),
                    VatAmount = reader["VatAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["VatAmount"]),
                    PaymentInstallment = reader["PaymentInstallment"] == DBNull.Value ? null : Convert.ToInt32(reader["PaymentInstallment"]),
                    MaximumDownTimeHours = reader["MaximumDownTimeHours"] == DBNull.Value ? null : Convert.ToInt32(reader["MaximumDownTimeHours"]),
                    MaximumDownPercents = reader["MaximumDownPercents"] == DBNull.Value ? null : Convert.ToDecimal(reader["MaximumDownPercents"]),
                    PenaltyPerHours = reader["PenaltyPerHours"] == DBNull.Value ? null : Convert.ToDecimal(reader["PenaltyPerHours"]),
                    ServiceFixPerMonths = reader["ServiceFixPerMonths"] == DBNull.Value ? null : Convert.ToInt32(reader["ServiceFixPerMonths"]),
                    ServiceFixStartIn = reader["ServiceFixStartIn"] == DBNull.Value ? null : Convert.ToInt32(reader["ServiceFixStartIn"]),
                    ServiceFixStartUnit = reader["ServiceFixStartUnit"] as string,
                    ServiceTimeIn = reader["ServiceTimeIn"] == DBNull.Value ? null : Convert.ToInt32(reader["ServiceTimeIn"]),
                    ServiceTimeUnit = reader["ServiceTimeUnit"] as string,
                    ServicePenaltyPercent = reader["ServicePenaltyPercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["ServicePenaltyPercent"]),
                    ContPenaltyPercent = reader["ContPenaltyPercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["ContPenaltyPercent"]),
                    ContPenaltyPerDays = reader["ContPenaltyPerDays"] == DBNull.Value ? null : Convert.ToInt32(reader["ContPenaltyPerDays"]),
                    GuaranteeType = reader["GuaranteeType"] as string,
                    GuaranteeAmount = reader["GuaranteeAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteeAmount"]),
                    GuaranteePercent = reader["GuaranteePercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteePercent"]),
                    NewGuaranteeDays = reader["NewGuaranteeDays"] == DBNull.Value ? null : Convert.ToInt32(reader["NewGuaranteeDays"]),
                    SubcontractPenaltyPercent = reader["SubcontractPenaltyPercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["SubcontractPenaltyPercent"]),
                    EnforcementOfFineDays = reader["EnforcementOfFineDays"] == DBNull.Value ? null : Convert.ToInt32(reader["EnforcementOfFineDays"]),
                    OutstandingPeriodDays = reader["OutstandingPeriodDays"] == DBNull.Value ? null : Convert.ToInt32(reader["OutstandingPeriodDays"]),
                    OSMEP_Signer = reader["OSMEP_Signer"] as string,
                    OSMEP_Witness = reader["OSMEP_Witness"] as string,
                    Contract_Signer = reader["Contract_Signer"] as string,
                    Contract_Witness = reader["Contract_Witness"] as string,
                    CreatedDate = reader["CreatedDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["CreatedDate"]),
                    CreateBy = reader["CreateBy"] as string,
                    UpdateDate = reader["UpdateDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["UpdateDate"]),
                    UpdateBy = reader["UpdateBy"] as string,
                    Flag_Delete = reader["Flag_Delete"] as string,
                    LegalEntityRegisNumber = reader["LegalEntityRegisNumber"] as string,
                    BusinessRegistrationCertDate = reader["BusinessRegistrationCertDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["BusinessRegistrationCertDate"]),
                    AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? (bool?)null : reader["AttorneyFlag"].ToString() == "1" || reader["AttorneyFlag"].ToString().ToLower() == "true",
                    AttorneyLetterDate = reader["AttorneyLetterDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["AttorneyLetterDate"]),
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] as string,
                    CitizenId = reader["CitizenId"] as string,
                    CitizenCardRegisDate = reader["CitizenCardRegisDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["CitizenCardRegisDate"]),
                    CitizenCardExpireDate = reader["CitizenCardExpireDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["CitizenCardExpireDate"]),
                    Request_ID = reader["Request_ID"] == DBNull.Value ? null : Convert.ToInt32(reader["Request_ID"]),
                    Contract_Status = reader["Contract_Status"] as string
                };

                return detail;
            }
            catch (Exception ex)
            {
                return null; // Consider logging the exception
            }
        }
        public async Task<List<E_ConReport_SMCInstallmentModels>> GetSMCInstallmentAsync(string? id = "0")
        {
            try
            {
                var result = new List<E_ConReport_SMCInstallmentModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT SMC_Inst_ID, SMC_R310_60_ID, PayRound, TotalAmount, RepairMonth, Flag_Delete
            FROM SMC_R310_60_Installment
            WHERE SMC_R310_60_ID = @SMC_R310_60_ID", connection);

                command.Parameters.AddWithValue("@SMC_R310_60_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_SMCInstallmentModels
                    {
                        SMC_Inst_ID = reader["SMC_Inst_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["SMC_Inst_ID"]),
                        SMC_R310_60_ID = reader["SMC_R310_60_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["SMC_R310_60_ID"]),
                        PayRound = reader["PayRound"] == DBNull.Value ? 0 : Convert.ToInt32(reader["PayRound"]),
                        TotalAmount = reader["TotalAmount"] == DBNull.Value ? 0 : Convert.ToDecimal(reader["TotalAmount"]),
                        RepairMonth = reader["RepairMonth"] == DBNull.Value ? 0 : Convert.ToInt32(reader["RepairMonth"]),
                        Flag_Delete = reader["Flag_Delete"] as string
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