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
    public class Econtract_Report_CLADAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_CLADAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_CLAModels?> GetCLAAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_CLA_R309_60", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@CLA_R309_60_ID", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_CLAModels
                {
                    CLA_R309_60_ID = reader["CLA_R309_60_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["CLA_R309_60_ID"]),
                    CLAContractNumber = reader["CLAContractNumber"] as string,
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
                    RentalSysName = reader["RentalSysName"] as string,
                    RentalBrandName = reader["RentalBrandName"] as string,
                    RentalPeriodYear = reader["RentalPeriodYear"] == DBNull.Value ? null : Convert.ToInt32(reader["RentalPeriodYear"]),
                    RentalPeriodMonth = reader["RentalPeriodMonth"] == DBNull.Value ? null : Convert.ToInt32(reader["RentalPeriodMonth"]),
                    SaleBankName = reader["SaleBankName"] as string,
                    SaleBankBranch = reader["SaleBankBranch"] as string,
                    SaleBankAccountName = reader["SaleBankAccountName"] as string,
                    SaleBankAccountNumber = reader["SaleBankAccountNumber"] as string,
                    DeliveryLocation = reader["DeliveryLocation"] as string,
                    DeliveryDateIn = reader["DeliveryDateIn"] == DBNull.Value ? null : Convert.ToInt32(reader["DeliveryDateIn"]),

                    NotiLocation = reader["NotiLocation"] as string,
                    NotiDaysBeforeDelivery = reader["NotiDaysBeforeDelivery"] == DBNull.Value ? null : Convert.ToInt32(reader["NotiDaysBeforeDelivery"]),
                    LocationDesignDays = reader["LocationDesignDays"] == DBNull.Value ? null : Convert.ToInt32(reader["LocationDesignDays"]),
                    MaintenancePermonth = reader["MaintenancePermonth"] == DBNull.Value ? null : Convert.ToInt32(reader["MaintenancePermonth"]),
                    MaintenanceInterval = reader["MaintenanceInterval"] == DBNull.Value ? null : Convert.ToInt32(reader["MaintenanceInterval"]),
                    MaximumDownTimeHours = reader["MaximumDownTimeHours"] == DBNull.Value ? null : Convert.ToInt32(reader["MaximumDownTimeHours"]),
                    MaximumDownPercents = reader["MaximumDownPercents"] == DBNull.Value ? null : Convert.ToDecimal(reader["MaximumDownPercents"]),
                    PenaltyPerHours = reader["PenaltyPerHours"] == DBNull.Value ? null : Convert.ToDecimal(reader["PenaltyPerHours"]),
                  //  NormalTimeFixDays = reader["NormalTimeFixDays"] == DBNull.Value ? null : Convert.ToInt32(reader["NormalTimeFixDays"]),
                //    OffTimeFixDays = reader["OffTimeFixDays"] == DBNull.Value ? null : Convert.ToInt32(reader["OffTimeFixDays"]),
                    FixPenaltyPerHours = reader["FixPenaltyPerHours"] == DBNull.Value ? null : Convert.ToDecimal(reader["FixPenaltyPerHours"]),
                    FixReplaceCompDays = reader["FixReplaceCompDays"] == DBNull.Value ? null : Convert.ToInt32(reader["FixReplaceCompDays"]),
                 //   FixReplacePenaltyPerHours = reader["FixReplacePenaltyPerHours"] == DBNull.Value ? null : Convert.ToDecimal(reader["FixReplacePenaltyPerHours"]),
                    TrainingPeriodDays = reader["TrainingPeriodDays"] == DBNull.Value ? null : Convert.ToInt32(reader["TrainingPeriodDays"]),
                    ComputerManualsCount = reader["ComputerManualsCount"] == DBNull.Value ? null : Convert.ToInt32(reader["ComputerManualsCount"]),
                    GuaranteeType = reader["GuaranteeType"] as string,
                    GuaranteeAmount = reader["GuaranteeAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteeAmount"]),
                    GuaranteePercent = reader["GuaranteePercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteePercent"]),
                    NewGuaranteeDays = reader["NewGuaranteeDays"] == DBNull.Value ? null : Convert.ToInt32(reader["NewGuaranteeDays"]),
                    RespReplaceDays = reader["RespReplaceDays"] == DBNull.Value ? null : Convert.ToInt32(reader["RespReplaceDays"]),
                    TeminationNewMonths = reader["TeminationNewMonths"] as string,
                    FinePerDaysPercent = reader["FinePerDaysPercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["FinePerDaysPercent"]),
                    ComputerSendBackDays = reader["ComputerSendBackDays"] == DBNull.Value ? null : Convert.ToInt32(reader["ComputerSendBackDays"]),
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
                    LegalEntityRegisNumber = reader["LegalEntityRegisNumber"] as string,
                    BusinessRegistrationCertDate = reader["BusinessRegistrationCertDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("BusinessRegistrationCertDate")),
                    AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? null : (bool?)Convert.ToBoolean(reader["AttorneyFlag"]),
                    AttorneyLetterDate = reader["AttorneyLetterDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("AttorneyLetterDate")),
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] as string,
                    CitizenId = reader["CitizenId"] as string,
                    RespReplaceYears = reader["RespReplaceYears"] == DBNull.Value ? null : Convert.ToInt32(reader["RespReplaceYears"]),
                    RespReplaceMonth = reader["RespReplaceMonth"] == DBNull.Value ? null : Convert.ToInt32(reader["RespReplaceMonth"]),
                    Request_ID = reader["Request_ID"] as string,
                    Contract_Status = reader["Contract_Status"] as string,
                    GuaranteeTypeOther = reader["GuaranteeTypeOther"] as string,
                    NormalTimeFixHours= reader["NormalTimeFixHours"] ==  DBNull.Value ? null : Convert.ToDecimal(reader["NormalTimeFixHours"]),
                    OffTimeFixHours = reader["OffTimeFixHours"] == DBNull.Value ? null : Convert.ToDecimal(reader["OffTimeFixHours"]),
                    FixReplacePenaltyPerDays = reader["FixReplacePenaltyPerDays"] == DBNull.Value ? null : Convert.ToDecimal(reader["FixReplacePenaltyPerDays"]),
               
                     NeedAttachCuS = reader["NeedAttachCuS"] == DBNull.Value ? null : (bool?)Convert.ToBoolean(reader["NeedAttachCuS"]),
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