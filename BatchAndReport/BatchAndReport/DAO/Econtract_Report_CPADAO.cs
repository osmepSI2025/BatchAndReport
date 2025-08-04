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
    public class Econtract_Report_CPADAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_CPADAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_CPAModels?> GetCPAAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_CPA", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@CPA_ID_INPUT", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_CPAModels
                {
                    CPA_ID = reader["CPA_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["CPA_ID"]),
                    Contract_Sign_Address = reader["Contract_Sign_Address"] == DBNull.Value ? null : reader["Contract_Sign_Address"].ToString(),
                    Contract_Organization = reader["Contract_Organization"] == DBNull.Value ? null : reader["Contract_Organization"].ToString(),
                    SignatoryName = reader["SignatoryName"] == DBNull.Value ? null : reader["SignatoryName"].ToString(),
                    SignatoryPosition = reader["SignatoryPosition"] == DBNull.Value ? null : reader["SignatoryPosition"].ToString(),
                    ContractorType = reader["ContractorType"] == DBNull.Value ? null : reader["ContractorType"].ToString(),
                    ContractorName = reader["ContractorName"] == DBNull.Value ? null : reader["ContractorName"].ToString(),
                    ContractorAddressNo = reader["ContractorAddressNo"] == DBNull.Value ? null : reader["ContractorAddressNo"].ToString(),
                    ContractorStreet = reader["ContractorStreet"] == DBNull.Value ? null : reader["ContractorStreet"].ToString(),
                    ContractorSubDistrict = reader["ContractorSubDistrict"] == DBNull.Value ? null : reader["ContractorSubDistrict"].ToString(),
                    ContractorDistrict = reader["ContractorDistrict"] == DBNull.Value ? null : reader["ContractorDistrict"].ToString(),
                    ContractorProvince = reader["ContractorProvince"] == DBNull.Value ? null : reader["ContractorProvince"].ToString(),
                    ContractorZipcode = reader["ContractorZipcode"] == DBNull.Value ? null : reader["ContractorZipcode"].ToString(),
                    ContractorSignatoryName = reader["ContractorSignatoryName"] == DBNull.Value ? null : reader["ContractorSignatoryName"].ToString(),
                    ContractorSignatoryPosition = reader["ContractorSignatoryPosition"] == DBNull.Value ? null : reader["ContractorSignatoryPosition"].ToString(),
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(reader["ContractSignDate"]),
                    ContractorAuthorize = reader["ContractorAuthorize"] == DBNull.Value ? null : reader["ContractorAuthorize"].ToString(),
                    Computer_Model = reader["Computer_Model"] == DBNull.Value ? null : reader["Computer_Model"].ToString(),
                    TotalAmount = reader["TotalAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["TotalAmount"]),
                    VatAmount = reader["VatAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["VatAmount"]),
                    DeliveryLocation = reader["DeliveryLocation"] == DBNull.Value ? null : reader["DeliveryLocation"].ToString(),
                    DeliveryDateIn = reader["DeliveryDateIn"] == DBNull.Value ? null : Convert.ToInt32(reader["DeliveryDateIn"]),
                    NotiDaysBeforeDelivery = reader["NotiDaysBeforeDelivery"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["NotiDaysBeforeDelivery"]),
                    LocationPrepareDays = reader["LocationPrepareDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["LocationPrepareDays"]),
                    PaymentMethod = reader["PaymentMethod"] == DBNull.Value ? null : reader["PaymentMethod"].ToString(),
                    AdvancePayment = reader["AdvancePayment"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["AdvancePayment"]),
                    PaymentDueDays = reader["PaymentDueDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["PaymentDueDays"]),
                    PaymentGuaranteeType = reader["PaymentGuaranteeType"] == DBNull.Value ? null : reader["PaymentGuaranteeType"].ToString(),
                    RemainingPaymentAmount = reader["RemainingPaymentAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["RemainingPaymentAmount"]),
                    SaleBankName = reader["SaleBankName"] == DBNull.Value ? null : reader["SaleBankName"].ToString(),
                    SaleBankBranch = reader["SaleBankBranch"] == DBNull.Value ? null : reader["SaleBankBranch"].ToString(),
                    SaleBankAccountName = reader["SaleBankAccountName"] == DBNull.Value ? null : reader["SaleBankAccountName"].ToString(),
                    SaleBankAccountNumber = reader["SaleBankAccountNumber"] == DBNull.Value ? null : reader["SaleBankAccountNumber"].ToString(),
                    WarrantyPeriodYears = reader["WarrantyPeriodYears"] == DBNull.Value ? null : reader["WarrantyPeriodYears"].ToString(),
                    WarrantyPeriodMonths = reader["WarrantyPeriodMonths"] == DBNull.Value ? null : reader["WarrantyPeriodMonths"].ToString(),
                    DaysToRepairAfterNoti = reader["DaysToRepairAfterNoti"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["DaysToRepairAfterNoti"]),
                    MaximumDownTimeHours = reader["MaximumDownTimeHours"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["MaximumDownTimeHours"]),
                    MaximumDownTimePercent = reader["MaximumDownTimePercent"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["MaximumDownTimePercent"]),
                    PenaltyPerHourPercent = reader["PenaltyPerHourPercent"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["PenaltyPerHourPercent"]),
                    PenaltyPerHour = reader["PenaltyPerHour"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["PenaltyPerHour"]),
                    PenaltyDueDaysIn = reader["PenaltyDueDaysIn"] == DBNull.Value ? null : Convert.ToInt32(reader["PenaltyDueDaysIn"]),
                    PerformanceGuarantee = reader["PerformanceGuarantee"] == DBNull.Value ? null : reader["PerformanceGuarantee"].ToString(),
                    GuaranteeAmount = reader["GuaranteeAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["GuaranteeAmount"]),
                    GuaranteePercent = reader["GuaranteePercent"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["GuaranteePercent"]),
                    NewGuaranteeDays = reader["NewGuaranteeDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["NewGuaranteeDays"]),
                    TrainingPeriodDays = reader["TrainingPeriodDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TrainingPeriodDays"]),
                    ComputerManualsCount = reader["ComputerManualsCount"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["ComputerManualsCount"]),
                    TeminationNewMonths = reader["TeminationNewMonths"] == DBNull.Value ? null : reader["TeminationNewMonths"].ToString(),
                    ReturnDaysIn = reader["ReturnDaysIn"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["ReturnDaysIn"]),
                    FinePerDaysPercent = reader["FinePerDaysPercent"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["FinePerDaysPercent"]),
                    EnforcementOfFineDays = reader["EnforcementOfFineDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["EnforcementOfFineDays"]),
                    OutstandingPeriodDays = reader["OutstandingPeriodDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["OutstandingPeriodDays"]),
                    OSMEP_Signer = reader["OSMEP_Signer"] == DBNull.Value ? null : reader["OSMEP_Signer"].ToString(),
                    OSMEP_Witness = reader["OSMEP_Witness"] == DBNull.Value ? null : reader["OSMEP_Witness"].ToString(),
                    Contract_Signer = reader["Contract_Signer"] == DBNull.Value ? null : reader["Contract_Signer"].ToString(),
                    Contract_Witness = reader["Contract_Witness"] == DBNull.Value ? null : reader["Contract_Witness"].ToString(),
                    CreatedDate = reader["CreatedDate"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(reader["CreatedDate"]),
                    CreateBy = reader["CreateBy"] == DBNull.Value ? null : reader["CreateBy"].ToString(),
                    UpdateDate = reader["UpdateDate"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(reader["UpdateDate"]),
                    UpdateBy = reader["UpdateBy"] == DBNull.Value ? null : reader["UpdateBy"].ToString(),
                    Flag_Delete = reader["Flag_Delete"] == DBNull.Value ? null : reader["Flag_Delete"].ToString(),
                    CPAContractNumber = reader["CPAContractNumber"] == DBNull.Value ? null : reader["CPAContractNumber"].ToString(),
                    LegalEntityRegisNumber = reader["LegalEntityRegisNumber"] == DBNull.Value ? null : reader["LegalEntityRegisNumber"].ToString(),
                    CompanyOrganizer = reader["CompanyOrganizer"] == DBNull.Value ? null : reader["CompanyOrganizer"].ToString(),
                    AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? null : (bool?)reader["AttorneyFlag"],
                    AttorneyLetterDate = reader["AttorneyLetterDate"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(reader["AttorneyLetterDate"]),
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] == DBNull.Value ? null : reader["AttorneyLetterNumber"].ToString(),
                    CitizenId = reader["CitizenId"] == DBNull.Value ? null : reader["CitizenId"].ToString(),
                    CitizenCardRegisDate = reader["CitizenCardRegisDate"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(reader["CitizenCardRegisDate"]),
                    CitizenCardExpireDate = reader["CitizenCardExpireDate"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(reader["CitizenCardExpireDate"]),
                    BusinessRegistrationCertDate = reader["BusinessRegistrationCertDate"] == DBNull.Value ? (DateTime?)null : Convert.ToDateTime(reader["BusinessRegistrationCertDate"]),
                    DeliveryNotifyLocation = reader["DeliveryNotifyLocation"] == DBNull.Value ? null : reader["DeliveryNotifyLocation"].ToString(),
                    PaymentSumAMT = reader["PaymentSumAMT"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["PaymentSumAMT"]),
                    Request_ID = reader["Request_ID"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["Request_ID"]),
                    Contract_Status = reader["Contract_Status"] == DBNull.Value ? null : reader["Contract_Status"].ToString(),
                    PaymentGuaranteeTypeOther = reader["PaymentGuaranteeTypeOther"] == DBNull.Value ? null : reader["PaymentGuaranteeTypeOther"].ToString()
                };

                int conId = detail.CPA_ID;

                await reader.CloseAsync();

                // 🔹 Load Signatory list from SP_Preview_Signatory_List_Report
                await using var signatoryCmd = new SqlCommand("SP_Preview_Signatory_List_Report", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };
                signatoryCmd.Parameters.AddWithValue("@con_id", conId);
                signatoryCmd.Parameters.AddWithValue("@con_type", "CPA"); // ใช้ค่าตามที่ระบบระบุ

                using var signatoryReader = await signatoryCmd.ExecuteReaderAsync();
                while (await signatoryReader.ReadAsync())
                {
                    detail.Signatories.Add(new E_ConReport_SignatoryModels
                    {
                        Signatory_Name = signatoryReader["Signatory_Name"] as string,
                        Position = signatoryReader["Position"] as string,
                        BU_UNIT = signatoryReader["BU_UNIT"] as string,
                        DS_FILE = signatoryReader["DS_FILE"] as string
                    });
                }

                return detail;
            }
            catch (Exception ex)
            {
                return null; // Consider logging the exception
            }
        }



    }
}