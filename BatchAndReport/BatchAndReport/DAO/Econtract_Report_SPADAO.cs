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
    public class Econtract_Report_SPADAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_SPADAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_SPAModels?> GetSPAAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_SPA_R305_60", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@SPA_R305_60_ID_INPUT", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

             



                var detail = new E_ConReport_SPAModels
                {
                    SPA_R305_60_Id = reader["SPA_R305_60_Id"] == DBNull.Value ? null : reader["SPA_R305_60_Id"].ToString(),
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
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("ContractSignDate")),
                    ContractorAuthorize = reader["ContractorAuthorize"] == DBNull.Value ? null : reader["ContractorAuthorize"].ToString(),
                    ProductDescription = reader["ProductDescription"] == DBNull.Value ? null : reader["ProductDescription"].ToString(),
                    Quantity = reader["Quantity"] == DBNull.Value ? null : Convert.ToInt32(reader["Quantity"]),
                    Unit = reader["Unit"] == DBNull.Value ? null : reader["Unit"].ToString(),
                    UnitPrice = reader["UnitPrice"] == DBNull.Value ? null : Convert.ToDecimal(reader["UnitPrice"]),
                    TotalAmount = reader["TotalAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["TotalAmount"]),
                    VatAmount = reader["VatAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["VatAmount"]),
                    DeliveryLocation = reader["DeliveryLocation"] == DBNull.Value ? null : reader["DeliveryLocation"].ToString(),
                    DeliveryDate = reader["DeliveryDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("DeliveryDate")),
                    DeliveryNotifyLocation = reader["DeliveryNotifyLocation"] == DBNull.Value ? null : reader["DeliveryNotifyLocation"].ToString(),
                    DeliveryNotifyDays = reader["DeliveryNotifyDays"] == DBNull.Value ? null : (int?)reader.GetInt32(reader.GetOrdinal("DeliveryNotifyDays")),
                    PaymentMethod = reader["PaymentMethod"] == DBNull.Value ? null : reader["PaymentMethod"].ToString(),
                    AdvancePayment = reader["AdvancePayment"] == DBNull.Value ? null : Convert.ToDecimal(reader["AdvancePayment"]),
                    PaymentDueDays = reader["PaymentDueDays"] == DBNull.Value ? null : (int?)reader.GetInt32(reader.GetOrdinal("PaymentDueDays")),
                    PaymentGuaranteeType = reader["PaymentGuaranteeType"] == DBNull.Value ? null : reader["PaymentGuaranteeType"].ToString(),
                    RemainingPaymentAmount = reader["RemainingPaymentAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["RemainingPaymentAmount"]),
                    SaleBankName = reader["SaleBankName"] == DBNull.Value ? null : reader["SaleBankName"].ToString(),
                    SaleBankBranch = reader["SaleBankBranch"] == DBNull.Value ? null : reader["SaleBankBranch"].ToString(),
                    SaleBankAccountName = reader["SaleBankAccountName"] == DBNull.Value ? null : reader["SaleBankAccountName"].ToString(),
                    SaleBankAccountNumber = reader["SaleBankAccountNumber"] == DBNull.Value ? null : reader["SaleBankAccountNumber"].ToString(),
                    WarrantyPeriodMonths = reader["WarrantyPeriodMonths"] == DBNull.Value ? null : reader["WarrantyPeriodMonths"].ToString(),
                    WarrantyPeriodDays = reader["WarrantyPeriodDays"] == DBNull.Value ? null : reader["WarrantyPeriodDays"].ToString(),
                    DaysToRepairAfterNoti = reader["DaysToRepairAfterNoti"] == DBNull.Value ? null : (int?)reader.GetInt32(reader.GetOrdinal("DaysToRepairAfterNoti")),
                    GuaranteeType = reader["GuaranteeType"] == DBNull.Value ? null : reader["GuaranteeType"].ToString(),
                    GuaranteeAmount = reader["GuaranteeAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteeAmount"]),


                    GuaranteePercent = reader["GuaranteePercent"] == DBNull.Value ? null : Convert.ToDecimal(reader["GuaranteePercent"]),
                    NewGuaranteeDays = reader["NewGuaranteeDays"] == DBNull.Value ? null : (int?)reader.GetInt32(reader.GetOrdinal("NewGuaranteeDays")),
                    TerminationNewMonths = reader["TerminationNewMonths"] == DBNull.Value ? null : reader["TerminationNewMonths"].ToString(),
                    FineRatePerDay = reader["FineRatePerDay"] == DBNull.Value ? null : Convert.ToDecimal(reader["FineRatePerDay"]),
                    FinePeriodAfterNoti = reader["FinePeriodAfterNoti"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("FinePeriodAfterNoti")),
                    DefectFinePeriod = reader["DefectFinePeriod"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("DefectFinePeriod")),
                    OSMEP_Signer = reader["OSMEP_Signer"] == DBNull.Value ? null : reader["OSMEP_Signer"].ToString(),
                    OSMEP_Witness = reader["OSMEP_Witness"] == DBNull.Value ? null : reader["OSMEP_Witness"].ToString(),
                    Contract_Signer = reader["Contract_Signer"] == DBNull.Value ? null : reader["Contract_Signer"].ToString(),
                    Contract_Witness = reader["Contract_Witness"] == DBNull.Value ? null : reader["Contract_Witness"].ToString(),
                    CreatedDate = reader["CreatedDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("CreatedDate")),
                    CreateBy = reader["CreateBy"] == DBNull.Value ? null : reader["CreateBy"].ToString(),
                    UpdateDate = reader["UpdateDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("UpdateDate")),
                    UpdateBy = reader["UpdateBy"] == DBNull.Value ? null : reader["UpdateBy"].ToString(),


                    Flag_Delete = reader["Flag_Delete"] == DBNull.Value ? null : (bool?)reader.GetBoolean(reader.GetOrdinal("Flag_Delete")),
                    SPAContractNumber = reader["SPAContractNumber"] == DBNull.Value ? null : reader["SPAContractNumber"].ToString(),
                    LegalEntityRegisNumber = reader["LegalEntityRegisNumber"] == DBNull.Value ? null : reader["LegalEntityRegisNumber"].ToString(),
                    CompanyOrganizer = reader["CompanyOrganizer"] == DBNull.Value ? null : reader["CompanyOrganizer"].ToString(),
                    AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? null : (bool?)reader["AttorneyFlag"],
                    AttorneyLetterDate = reader["AttorneyLetterDate"] == DBNull.Value ? (DateTime?)null : (DateTime?)reader["AttorneyLetterDate"],
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] == DBNull.Value ? null : reader["AttorneyLetterNumber"].ToString(),
                    CitizenId = reader["CitizenId"] == DBNull.Value ? null : reader["CitizenId"].ToString(),
                    CitizenCardRegisDate = reader["CitizenCardRegisDate"] == DBNull.Value ? (DateTime?)null : (DateTime?)reader["CitizenCardRegisDate"],
                    CitizenCardExpireDate = reader["CitizenCardExpireDate"] == DBNull.Value ? (DateTime?)null : (DateTime?)reader["CitizenCardExpireDate"],
                    DeliverType = reader["DeliverType"] == DBNull.Value ? null : reader["DeliverType"].ToString(),
                    DeliverDays = reader["DeliverDays"] == DBNull.Value ? null : (int?)reader["DeliverDays"],
                    PaymentSumAMT = reader["PaymentSumAMT"] == DBNull.Value ? null : Convert.ToDecimal(reader["PaymentSumAMT"]),

                    WarrantyPeriodYears = reader["WarrantyPeriodYears"] == DBNull.Value ? null : reader["WarrantyPeriodYears"].ToString(),
                    FinePeriodAfterNotiDays = reader["FinePeriodAfterNotiDays"] == DBNull.Value ? null : (int?)reader["FinePeriodAfterNotiDays"],
                    DefectFinePeriodDays = reader["DefectFinePeriodDays"] == DBNull.Value ? null : (int?)reader["DefectFinePeriodDays"],
                    BusinessRegistrationCertDate = reader["BusinessRegistrationCertDate"] == DBNull.Value ? (DateTime?)null : (DateTime?)reader["BusinessRegistrationCertDate"],
                    Request_ID = reader["Request_ID"] == DBNull.Value ? null : reader["Request_ID"].ToString(),
                    Contract_Status = reader["Contract_Status"] == DBNull.Value ? null : reader["Contract_Status"].ToString(),
                    GuaranteeTypeOther = reader["GuaranteeTypeOther"] == DBNull.Value ? null : reader["GuaranteeTypeOther"].ToString()
                };

                return detail;
            }
            catch(Exception ex)
            {
            return null; // Consider logging the exception
            }
       
        }

    }
}