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
    public class Econtract_Report_MIWDAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_MIWDAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_MIWModels?> GetMIWAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_MIW", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@MIW_ID_Input", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_MIWModels
                {
                    MIW_ID = reader["MIW_ID"] == DBNull.Value ? null : reader["MIW_ID"].ToString(),
                    Contract_Number = reader["Contract_Number"] == DBNull.Value ? null : reader["Contract_Number"].ToString(),
                    ServiceName = reader["ServiceName"] == DBNull.Value ? null : reader["ServiceName"].ToString(),
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("ContractSignDate")),
                    SignatoryName = reader["SignatoryName"] == DBNull.Value ? null : reader["SignatoryName"].ToString(),
                    SignatoryPosition = reader["SignatoryPosition"] == DBNull.Value ? null : reader["SignatoryPosition"].ToString(),
                    ContractPartyName = reader["ContractPartyName"] == DBNull.Value ? null : reader["ContractPartyName"].ToString(),
                    IdenID = reader["IdenID"] == DBNull.Value ? null : reader["IdenID"].ToString(),
                    IdenIssue_Date = reader["IdenIssue_Date"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("IdenIssue_Date")),
                    IdenExpiry_Date = reader["IdenExpiry_Date"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("IdenExpiry_Date")),
                    AddressNo = reader["AddressNo"] == DBNull.Value ? null : reader["AddressNo"].ToString(),
                    AddressStreet = reader["AddressStreet"] == DBNull.Value ? null : reader["AddressStreet"].ToString(),
                    AddressSubDistrict = reader["AddressSubDistrict"] == DBNull.Value ? null : reader["AddressSubDistrict"].ToString(),
                    AddressDistrict = reader["AddressDistrict"] == DBNull.Value ? null : reader["AddressDistrict"].ToString(),
                    AddressProvince = reader["AddressProvince"] == DBNull.Value ? null : reader["AddressProvince"].ToString(),
                    AddressZipCode = reader["AddressZipCode"] == DBNull.Value ? null : reader["AddressZipCode"].ToString(),
                    HiringAgreement = reader["HiringAgreement"] == DBNull.Value ? null : reader["HiringAgreement"].ToString(),
                    HiringAmount = reader["HiringAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["HiringAmount"]),
                    VatAmount = reader["VatAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["VatAmount"]),
                    Amount = reader["Amount"] == DBNull.Value ? null : Convert.ToDecimal(reader["Amount"]),
                    Salary = reader["Salary"] == DBNull.Value ? null : Convert.ToDecimal(reader["Salary"]),
                    Delivery_Due_Date = reader["Delivery_Due_Date"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("Delivery_Due_Date")),
                    DailyRate = reader["DailyRate"] == DBNull.Value ? null : Convert.ToDecimal(reader["DailyRate"]),
                    ContractBankName = reader["ContractBankName"] == DBNull.Value ? null : reader["ContractBankName"].ToString(),
                    ContractBankBranch = reader["ContractBankBranch"] == DBNull.Value ? null : reader["ContractBankBranch"].ToString(),
                    ContractBankAccountName = reader["ContractBankAccountName"] == DBNull.Value ? null : reader["ContractBankAccountName"].ToString(),
                    ContractBankAccountNumber = reader["ContractBankAccountNumber"] == DBNull.Value ? null : reader["ContractBankAccountNumber"].ToString(),
                    WorkStartDate = reader["WorkStartDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("WorkStartDate")),
                    WorkEndDate = reader["WorkEndDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("WorkEndDate")),
                    DailyFineRate = reader["DailyFineRate"] == DBNull.Value ? null : Convert.ToDecimal(reader["DailyFineRate"]),
                    OSMEP_Signer = reader["OSMEP_Signer"] == DBNull.Value ? null : reader["OSMEP_Signer"].ToString(),
                    OSMEP_Witness = reader["OSMEP_Witness"] == DBNull.Value ? null : reader["OSMEP_Witness"].ToString(),
                    Contract_Signer = reader["Contract_Signer"] == DBNull.Value ? null : reader["Contract_Signer"].ToString(),
                    Contract_Witness = reader["Contract_Witness"] == DBNull.Value ? null : reader["Contract_Witness"].ToString(),
                    CreatedDate = reader["CreatedDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("CreatedDate")),
                    CreateBy = reader["CreateBy"] == DBNull.Value ? null : reader["CreateBy"].ToString(),
                    UpdateDate = reader["UpdateDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("UpdateDate")),
                    UpdateBy = reader["UpdateBy"] == DBNull.Value ? null : reader["UpdateBy"].ToString(),
                    Flag_Delete = reader["Flag_Delete"] == DBNull.Value ? null : reader["Flag_Delete"].ToString(),
                    Request_ID = reader["Request_ID"] == DBNull.Value ? null : reader["Request_ID"].ToString(),
                    Contract_Status = reader["Contract_Status"] == DBNull.Value ? null : reader["Contract_Status"].ToString(),
                    NeedAttachCuS = reader["NeedAttachCuS"] == DBNull.Value ? null : (bool?)reader["NeedAttachCuS"],
                    Organization_Logo = reader["Organization_Logo"] == DBNull.Value ? null : reader["Organization_Logo"].ToString(),
                    OSMEP_NAME = reader["OSMEP_NAME"] == DBNull.Value ? null : reader["OSMEP_NAME"].ToString(),
                    OSMEP_POSITION = reader["OSMEP_POSITION"] == DBNull.Value ? null : reader["OSMEP_POSITION"].ToString(),
                    CP_S_NAME = reader["CP_S_NAME"] == DBNull.Value ? null : reader["CP_S_NAME"].ToString(),
                    CP_S_POSITION = reader["CP_S_POSITION"] == DBNull.Value ? null : reader["CP_S_POSITION"].ToString(),
                    CP_S_AttorneyFlag = reader["CP_S_AttorneyFlag"] == DBNull.Value ? null : (bool?)reader["CP_S_AttorneyFlag"],
                    CP_S_AttorneyLetterNumbers = reader["CP_S_AttorneyLetterNumbers"] == DBNull.Value ? null : reader["CP_S_AttorneyLetterNumbers"].ToString(),
                    AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? null : (bool?)reader["AttorneyFlag"],
                    AttorneyLetterNumber = reader["AttorneyLetterNumber"] == DBNull.Value ? null : reader["AttorneyLetterNumber"].ToString(),
                    Grant_Date = reader["Grant_Date"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("Grant_Date"))
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