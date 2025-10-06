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
    public class Econtract_Report_CTRDAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_CTRDAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_CTRModels?> GetCTRAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_CTR_R317_60", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@CTR_R317_60_ID_Input", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_CTRModels
                {
                    CTR_R317_60_ID = reader["CTR_R317_60_ID"] as int? ?? Convert.ToInt32(reader["CTR_R317_60_ID"]),
                    Contract_Number = reader["Contract_Number"] as string,
                    ContractSignDate = reader["ContractSignDate"] as DateTime? ?? (reader["ContractSignDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["ContractSignDate"])),
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
                    ContractorAuthDate = reader["ContractorAuthDate"] as DateTime? ?? (reader["ContractorAuthDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["ContractorAuthDate"])),
                    ContractorAuthNumber = reader["ContractorAuthNumber"] as string,
                    ProjectName = reader["ProjectName"] as string,
                    ProjectDesc = reader["ProjectDesc"] as string,
                    ConsultExpertise = reader["ConsultExpertise"] as string,
                    ProjectReference = reader["ProjectReference"] as string,
                    ProjectStartDate = reader["ProjectStartDate"] as DateTime? ?? (reader["ProjectStartDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["ProjectStartDate"])),
                    ProjectEndDate = reader["ProjectEndDate"] as DateTime? ?? (reader["ProjectEndDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["ProjectEndDate"])),
                    ContractTotalAmount = reader["ContractTotalAmount"] as decimal? ?? (reader["ContractTotalAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["ContractTotalAmount"])),
                    VatAmount = reader["VatAmount"] as decimal? ?? (reader["VatAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["VatAmount"])),
                    ContractInstallment = reader["ContractInstallment"] as int? ?? (reader["ContractInstallment"] == DBNull.Value ? null : Convert.ToInt32(reader["ContractInstallment"])),
                    ContractRef = reader["ContractRef"] as string,
                    ContractBankName = reader["ContractBankName"] as string,
                    ContractBankBranch = reader["ContractBankBranch"] as string,
                    ContractBankAccountName = reader["ContractBankAccountName"] as string,
                    ContractBankAccountNumber = reader["ContractBankAccountNumber"] as string,
                    PrepaidAmount = reader["PrepaidAmount"] as decimal? ?? (reader["PrepaidAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["PrepaidAmount"])),
                    PrepaidPercents = reader["PrepaidPercents"] as decimal? ?? (reader["PrepaidPercents"] == DBNull.Value ? null : Convert.ToDecimal(reader["PrepaidPercents"])),
                    PrepaidGuaranteeType = reader["PrepaidGuaranteeType"] as string,
                    PrepaidBankName = reader["PrepaidBankName"] as string,
                    PrepaidDeductPercent = reader["PrepaidDeductPercent"] as decimal? ?? (reader["PrepaidDeductPercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["PrepaidDeductPercent"])),
                    SendWorkMethod = reader["SendWorkMethod"] as string,
                    WorkAmount = reader["WorkAmount"] as decimal? != null
    ? (int?)(decimal?)reader["WorkAmount"]
    : (reader["WorkAmount"] == DBNull.Value ? null : (int?)Convert.ToDecimal(reader["WorkAmount"])),
                    RelateExpertise = reader["RelateExpertise"] as string,
                    FixDaysAfterNoti = reader["FixDaysAfterNoti"] as int? ?? (reader["FixDaysAfterNoti"] == DBNull.Value ? null : Convert.ToInt32(reader["FixDaysAfterNoti"])),
                    NotiDaysAfterTerminate = reader["NotiDaysAfterTerminate"] as int? ?? (reader["NotiDaysAfterTerminate"] == DBNull.Value ? null : Convert.ToInt32(reader["NotiDaysAfterTerminate"])),
                    TerminationReferene = reader["TerminationReferene"] as string,
                    ConsultScopeRef = reader["ConsultScopeRef"] as string,
                    FinePerDays = reader["FinePerDays"] as decimal? ?? (reader["FinePerDays"] == DBNull.Value ? null : Convert.ToDecimal(reader["FinePerDays"])),
                    EnforcementOfFineDays = reader["EnforcementOfFineDays"] as int? ?? (reader["EnforcementOfFineDays"] == DBNull.Value ? null : Convert.ToInt32(reader["EnforcementOfFineDays"])),
                    OutstandingPeriodDays = reader["OutstandingPeriodDays"] as int? ?? (reader["OutstandingPeriodDays"] == DBNull.Value ? null : Convert.ToInt32(reader["OutstandingPeriodDays"])),
                    RetentionRatePercent = reader["RetentionRatePercent"] as decimal? ?? (reader["RetentionRatePercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["RetentionRatePercent"])),
                    GuaranteeType = reader["GuaranteeType"] as string,
                    GuaranteeAmount = reader["GuaranteeAmount"] as decimal? ?? (reader["GuaranteeAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteeAmount"])),
                    GuaranteePercent = reader["GuaranteePercent"] as decimal? ?? (reader["GuaranteePercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteePercent"])),
                    NewGuaranteeDays = reader["NewGuaranteeDays"] as int? ?? (reader["NewGuaranteeDays"] == DBNull.Value ? null : Convert.ToInt32(reader["NewGuaranteeDays"])),
                    SubcontractPenaltyPercent = reader["SubcontractPenaltyPercent"] as decimal? ?? (reader["SubcontractPenaltyPercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["SubcontractPenaltyPercent"])),
                    OSMEP_Signer = reader["OSMEP_Signer"] as string,
                    OSMEP_Witness = reader["OSMEP_Witness"] as string,
                    Contract_Signer = reader["Contract_Signer"] as string,
                    Contract_Witness = reader["Contract_Witness"] as string,
                    CreatedDate = reader["CreatedDate"] as DateTime? ?? (reader["CreatedDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["CreatedDate"])),
                    CreateBy = reader["CreateBy"] as string,
                    UpdateDate = reader["UpdateDate"] as DateTime? ?? (reader["UpdateDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["UpdateDate"])),
                    UpdateBy = reader["UpdateBy"] as string,
                    Flag_Delete = reader["Flag_Delete"] as string,
                    AttorneyFlag = reader["AttorneyFlag"] as bool? ?? (reader["AttorneyFlag"] == DBNull.Value ? null : Convert.ToBoolean(reader["AttorneyFlag"])),
                    AttorneyLetterDate = reader["AttorneyLetterDate"] as DateTime? ?? (reader["AttorneyLetterDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["AttorneyLetterDate"])),
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] as string,
                    CitizenId = reader["CitizenId"] as string,
                    CitizenCardRegisDate = reader["CitizenCardRegisDate"] as DateTime? ?? (reader["CitizenCardRegisDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["CitizenCardRegisDate"])),
                    CitizenCardExpireDate = reader["CitizenCardExpireDate"] as DateTime? ?? (reader["CitizenCardExpireDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["CitizenCardExpireDate"])),
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