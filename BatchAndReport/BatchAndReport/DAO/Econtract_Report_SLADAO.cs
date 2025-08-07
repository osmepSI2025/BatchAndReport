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
    public class Econtract_Report_SLADAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_SLADAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_SLAModels?> GetSLAAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_SLA_R308_60", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@SLA_R308_60_ID", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_SLAModels
                {
                    SLA_R308_60_ID = reader["SLA_R308_60_ID"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SLA_R308_60_ID"]),
                    Contract_Number = reader["Contract_Number"] as string,
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["ContractSignDate"]),
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
                    TotalAmount = reader["TotalAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["TotalAmount"]),
                    VatAmount = reader["VatAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["VatAmount"]),
                    SWRight_1 = reader["SWRight_1"] as string,
                    SWRight_1_Detail = reader["SWRight_1_Detail"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SWRight_1_Detail"]),
                    SWRight_2 = reader["SWRight_2"] as string,
                    SWRight_2_Detail_1 = reader["SWRight_2_Detail_1"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SWRight_2_Detail_1"]),
                    SWRight_2_Detail_2 = reader["SWRight_2_Detail_2"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SWRight_2_Detail_2"]),
                    SWRight_3 = reader["SWRight_3"] as string,
                    SWRight_3_Detail = reader["SWRight_3_Detail"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SWRight_3_Detail"]),
                    SWRight_4 = reader["SWRight_4"] as string,
                    SWRight_5 = reader["SWRight_5"] as string,
                    SWRight_5_Detail = reader["SWRight_5_Detail"] as string,
                    TotalYears = reader["TotalYears"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TotalYears"]),
                    TotalMonths = reader["TotalMonths"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TotalMonths"]),
                    TotalDays = reader["TotalDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TotalDays"]),
                    DeliveryLocation = reader["DeliveryLocation"] as string,
                    DeliveryDateIn = reader["DeliveryDateIn"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["DeliveryDateIn"]),
                    NotiLocation = reader["NotiLocation"] as string,
                    NotiDaysBeforeDelivery = reader["NotiDaysBeforeDelivery"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["NotiDaysBeforeDelivery"]),
                    PaymentMethod = reader["PaymentMethod"] as string,
                    PaymentInstallment = reader["PaymentInstallment"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["PaymentInstallment"]),
                    LastPayRound = reader["LastPayRound"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["LastPayRound"]),
                    SaleBankName = reader["SaleBankName"] as string,
                    SaleBankBranch = reader["SaleBankBranch"] as string,
                    SaleBankAccountName = reader["SaleBankAccountName"] as string,
                    SaleBankAccountNumber = reader["SaleBankAccountNumber"] as string,
                    BackupQty = reader["BackupQty"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["BackupQty"]),
                    ComputerModel = reader["ComputerModel"] as string,
                    ComputerBrand = reader["ComputerBrand"] as string,
                    BuyerAddressNo = reader["BuyerAddressNo"] as string,
                    BuyerStreet = reader["BuyerStreet"] as string,
                    BuyerSubDistrict = reader["BuyerSubDistrict"] as string,
                    BuyerDistrict = reader["BuyerDistrict"] as string,
                    BuyerProvince = reader["BuyerProvince"] as string,
                    BuyerZipcode = reader["BuyerZipcode"] as string,
                    WarrantyPeriodYears = reader["WarrantyPeriodYears"] as string,
                    WarrantyPeriodMonths = reader["WarrantyPeriodMonths"] as string,
                    DaysToRepairIn = reader["DaysToRepairIn"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["DaysToRepairIn"]),
                    GuaranteeType = reader["GuaranteeType"] as string,
                    GuaranteeAmount = reader["GuaranteeAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["GuaranteeAmount"]),
                    GuaranteePercent = reader["GuaranteePercent"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["GuaranteePercent"]),
                    NewGuaranteeDays = reader["NewGuaranteeDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["NewGuaranteeDays"]),
                    TrainingPeriodDays = reader["TrainingPeriodDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TrainingPeriodDays"]),
                    ComputerManualsCount = reader["ComputerManualsCount"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["ComputerManualsCount"]),
                    TeminationNewMonths = reader["TeminationNewMonths"] as string,
                    ReturnDaysIn = reader["ReturnDaysIn"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["ReturnDaysIn"]),
                    EnforcementOfFineDays = reader["EnforcementOfFineDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["EnforcementOfFineDays"]),
                    OSMEP_Signer = reader["OSMEP_Signer"] as string,
                    OSMEP_Witness = reader["OSMEP_Witness"] as string,
                    Contract_Signer = reader["Contract_Signer"] as string,
                    Contract_Witness = reader["Contract_Witness"] as string,
                    CreateBy = reader["CreateBy"] as string,
                    UpdateBy = reader["UpdateBy"] as string,
                    Flag_Delete = reader["Flag_Delete"] as string,
                    LegalEntityRegisNumber = reader["LegalEntityRegisNumber"] as string,
                    AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? null : (bool?)reader["AttorneyFlag"],
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] as string,
                    CitizenId = reader["CitizenId"] as string,
                    Request_ID = reader["Request_ID"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["Request_ID"]),
                    Contract_Status = reader["Contract_Status"] as string,
                    GuaranteeTypeOther = reader["GuaranteeTypeOther"] as string
                };

                return detail;
            }
            catch (Exception ex)
            {
                return null; // Consider logging the exception
            }
        }
        public async Task<List<E_ConReport_SLAInstallmentModels>> GetSLAInstallmentAsync(string? id = "0")
        {
            try
            {
                var result = new List<E_ConReport_SLAInstallmentModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT SLA_Inst_ID, SLA_R308_60_ID, PayRound, TotalAmount, UseMonth, Flag_Delete
            FROM SLA_R308_60_Installment
            WHERE SLA_R308_60_ID = @SLA_R308_60_ID", connection);

                command.Parameters.AddWithValue("@SLA_R308_60_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_SLAInstallmentModels
                    {
                        SLA_Inst_ID = reader["SLA_Inst_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["SLA_Inst_ID"]),
                        SLA_R308_60_ID = reader["SLA_R308_60_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["SLA_R308_60_ID"]),
                        PayRound = reader["PayRound"] == DBNull.Value ? 0 : Convert.ToInt32(reader["PayRound"]),
                        TotalAmount = reader["TotalAmount"] == DBNull.Value ? 0 : Convert.ToDecimal(reader["TotalAmount"]),
                        UseMonth = reader["UseMonth"] == DBNull.Value ? 0 : Convert.ToInt32(reader["UseMonth"]),
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

        public async Task<E_ConReport_SLAModels?> GetSLAInfoAsync(string id)
        {
            try
            {
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT 
                SLA_R308_60_ID,
                Contract_Number,
                ContractSignDate,
                Contract_Sign_Address,
                Contract_Organization,
                SignatoryName,
                SignatoryPosition,
                ContractorType,
                ContractorName,
                ContractorCompany,
                ContractorAddressNo,
                ContractorStreet,
                ContractorSubDistrict,
                ContractorDistrict,
                ContractorProvince,
                ContractorZipcode,
                ContractorSignatoryName,
                ContractorSignatoryPosition,
                ContractorAuthorize,
                TotalAmount,
                VatAmount,
                SWRight_1,
                SWRight_1_Detail,
                SWRight_2,
                SWRight_2_Detail_1,
                SWRight_2_Detail_2,
                SWRight_3,
                SWRight_3_Detail,
                SWRight_4,
                SWRight_5,
                SWRight_5_Detail,
                SWExpiry_Date,
                TotalYears,
                TotalMonths,
                TotalDays,
                DeliveryLocation,
                DeliveryDateIn,
                NotiLocation,
                NotiDaysBeforeDelivery,
                PaymentMethod,
                PaymentInstallment,
                LastPayRound,
                SaleBankName,
                SaleBankBranch,
                SaleBankAccountName,
                SaleBankAccountNumber,
                BackupQty,
                ComputerModel,
                ComputerBrand,
                BuyerAddressNo,
                BuyerStreet,
                BuyerSubDistrict,
                BuyerDistrict,
                BuyerProvince,
                BuyerZipcode,
                WarrantyPeriodYears,
                WarrantyPeriodMonths,
                DaysToRepairIn,
                GuaranteeType,
                GuaranteeAmount,
                GuaranteePercent,
                NewGuaranteeDays,
                TrainingPeriodDays,
                ComputerManualsCount,
                TeminationNewMonths,
                ReturnDaysIn,
                FinePerDaysPercent,
                EnforcementOfFineDays,
                OutstandingPeriodDays,
                ComputerSendBackDays,
                OSMEP_Signer,
                OSMEP_Witness,
                Contract_Signer,
                Contract_Witness,
                CreatedDate,
                CreateBy,
                UpdateDate,
                UpdateBy,
                Flag_Delete,
                LegalEntityRegisNumber,
                BusinessRegistrationCertDate,
                AttorneyFlag,
                AttorneyLetterDate,
                AttorneyLetterNumber,
                CitizenId,
                CitizenCardRegisDate,
                CitizenCardExpireDate,
                Request_ID,
                Contract_Status,
                GuaranteeTypeOther
            FROM SLA_R308_60
            WHERE SLA_R308_60_ID = @SLA_R308_60_ID", connection);

                command.Parameters.AddWithValue("@SLA_R308_60_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var result = new E_ConReport_SLAModels
                {
                    SLA_R308_60_ID = reader["SLA_R308_60_ID"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SLA_R308_60_ID"]),
                    Contract_Number = reader["Contract_Number"] as string,
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["ContractSignDate"]),
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
                    TotalAmount = reader["TotalAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["TotalAmount"]),
                    VatAmount = reader["VatAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["VatAmount"]),
                    SWRight_1 = reader["SWRight_1"] as string,
                    SWRight_1_Detail = reader["SWRight_1_Detail"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SWRight_1_Detail"]),
                    SWRight_2 = reader["SWRight_2"] as string,
                    SWRight_2_Detail_1 = reader["SWRight_2_Detail_1"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SWRight_2_Detail_1"]),
                    SWRight_2_Detail_2 = reader["SWRight_2_Detail_2"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SWRight_2_Detail_2"]),
                    SWRight_3 = reader["SWRight_3"] as string,
                    SWRight_3_Detail = reader["SWRight_3_Detail"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["SWRight_3_Detail"]),
                    SWRight_4 = reader["SWRight_4"] as string,
                    SWRight_5 = reader["SWRight_5"] as string,
                    SWRight_5_Detail = reader["SWRight_5_Detail"] as string,
                    SWExpiry_Date = reader["SWExpiry_Date"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["SWExpiry_Date"]),
                    TotalYears = reader["TotalYears"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TotalYears"]),
                    TotalMonths = reader["TotalMonths"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TotalMonths"]),
                    TotalDays = reader["TotalDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TotalDays"]),
                    DeliveryLocation = reader["DeliveryLocation"] as string,
                    DeliveryDateIn = reader["DeliveryDateIn"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["DeliveryDateIn"]),
                    NotiLocation = reader["NotiLocation"] as string,
                    NotiDaysBeforeDelivery = reader["NotiDaysBeforeDelivery"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["NotiDaysBeforeDelivery"]),
                    PaymentMethod = reader["PaymentMethod"] as string,
                    PaymentInstallment = reader["PaymentInstallment"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["PaymentInstallment"]),
                    LastPayRound = reader["LastPayRound"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["LastPayRound"]),
                    SaleBankName = reader["SaleBankName"] as string,
                    SaleBankBranch = reader["SaleBankBranch"] as string,
                    SaleBankAccountName = reader["SaleBankAccountName"] as string,
                    SaleBankAccountNumber = reader["SaleBankAccountNumber"] as string,
                    BackupQty = reader["BackupQty"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["BackupQty"]),
                    ComputerModel = reader["ComputerModel"] as string,
                    ComputerBrand = reader["ComputerBrand"] as string,
                    BuyerAddressNo = reader["BuyerAddressNo"] as string,
                    BuyerStreet = reader["BuyerStreet"] as string,
                    BuyerSubDistrict = reader["BuyerSubDistrict"] as string,
                    BuyerDistrict = reader["BuyerDistrict"] as string,
                    BuyerProvince = reader["BuyerProvince"] as string,
                    BuyerZipcode = reader["BuyerZipcode"] as string,
                    WarrantyPeriodYears = reader["WarrantyPeriodYears"] as string,
                    WarrantyPeriodMonths = reader["WarrantyPeriodMonths"] as string,
                    DaysToRepairIn = reader["DaysToRepairIn"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["DaysToRepairIn"]),
                    GuaranteeType = reader["GuaranteeType"] as string,
                    GuaranteeAmount = reader["GuaranteeAmount"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["GuaranteeAmount"]),
                    GuaranteePercent = reader["GuaranteePercent"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["GuaranteePercent"]),
                    NewGuaranteeDays = reader["NewGuaranteeDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["NewGuaranteeDays"]),
                    TrainingPeriodDays = reader["TrainingPeriodDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["TrainingPeriodDays"]),
                    ComputerManualsCount = reader["ComputerManualsCount"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["ComputerManualsCount"]),
                    TeminationNewMonths = reader["TeminationNewMonths"] as string,
                    ReturnDaysIn = reader["ReturnDaysIn"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["ReturnDaysIn"]),
                    FinePerDaysPercent = reader["FinePerDaysPercent"] == DBNull.Value ? null : (decimal?)Convert.ToDecimal(reader["FinePerDaysPercent"]),
                    EnforcementOfFineDays = reader["EnforcementOfFineDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["EnforcementOfFineDays"]),
                    OutstandingPeriodDays = reader["OutstandingPeriodDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["OutstandingPeriodDays"]),
                    ComputerSendBackDays = reader["ComputerSendBackDays"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["ComputerSendBackDays"]),
                    OSMEP_Signer = reader["OSMEP_Signer"] as string,
                    OSMEP_Witness = reader["OSMEP_Witness"] as string,
                    Contract_Signer = reader["Contract_Signer"] as string,
                    Contract_Witness = reader["Contract_Witness"] as string,
                    CreatedDate = reader["CreatedDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["CreatedDate"]),
                    CreateBy = reader["CreateBy"] as string,
                    UpdateDate = reader["UpdateDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["UpdateDate"]),
                    UpdateBy = reader["UpdateBy"] as string,
                    Flag_Delete = reader["Flag_Delete"] as string,
                    LegalEntityRegisNumber = reader["LegalEntityRegisNumber"] as string,
                    BusinessRegistrationCertDate = reader["BusinessRegistrationCertDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["BusinessRegistrationCertDate"]),
                    AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? null : (bool?)reader["AttorneyFlag"],
                    AttorneyLetterDate = reader["AttorneyLetterDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["AttorneyLetterDate"]),
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] as string,
                    CitizenId = reader["CitizenId"] as string,
                    CitizenCardRegisDate = reader["CitizenCardRegisDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["CitizenCardRegisDate"]),
                    CitizenCardExpireDate = reader["CitizenCardExpireDate"] == DBNull.Value ? null : (DateTime?)Convert.ToDateTime(reader["CitizenCardExpireDate"]),
                    Request_ID = reader["Request_ID"] == DBNull.Value ? null : (int?)Convert.ToInt32(reader["Request_ID"]),
                    Contract_Status = reader["Contract_Status"] as string,
                    GuaranteeTypeOther = reader["GuaranteeTypeOther"] as string
                };

                return result;
            }
            catch (Exception ex)
            {
                return null;
            }
        }
    }
}