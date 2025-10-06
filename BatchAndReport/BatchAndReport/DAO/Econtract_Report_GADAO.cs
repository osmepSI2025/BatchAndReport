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
    public class Econtract_Report_GADAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_GADAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_GAModels?> GetGAAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_GA", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@GA_ID_Input", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_GAModels
                {
                    GA_ID = reader["GA_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["GA_ID"]),
                    Contract_Number = reader["Contract_Number"] as string,
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["ContractSignDate"]),
                    SignAddress = reader["SignAddress"] as string,
                    OrganizationName = reader["OrganizationName"] as string,
                    SignatoryName = reader["SignatoryName"] as string,
                    SignatoryPosition = reader["SignatoryPosition"] as string,
                    TaxID = reader["TaxID"] as string,
                    ContractPartyName = reader["ContractPartyName"] as string,
                    RegType = reader["RegType"] as string,
                    RegOrganization = reader["RegOrganization"] as string,
                    HQLocationAddressNo = reader["HQLocationAddressNo"] as string,
                    HQLocationStreet = reader["HQLocationStreet"] as string,
                    HQLocationSubDistrict = reader["HQLocationSubDistrict"] as string,
                    HQLocationDistrict = reader["HQLocationDistrict"] as string,
                    HQLocationProvince = reader["HQLocationProvince"] as string,
                    HQLocationZipCode = reader["HQLocationZipCode"] as string,
                    RegEmail = reader["RegEmail"] as string,
                    RegPersonalName = reader["RegPersonalName"] as string,
                    RegIdenID = reader["RegIdenID"] as string,
                    GrantAmount = reader["GrantAmount"] == DBNull.Value ? null : Convert.ToDecimal(reader["GrantAmount"]),
                    GrantStartDate = reader["GrantStartDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["GrantStartDate"]),
                    GrantEndDate = reader["GrantEndDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["GrantEndDate"]),
                    SpendingPurpose = reader["SpendingPurpose"] as string,
                    OSMEP_Signer = reader["OSMEP_Signer"] as string,
                    OSMEP_Witness = reader["OSMEP_Witness"] as string,
                    Contract_Signer = reader["Contract_Signer"] as string,
                    Contract_Witness = reader["Contract_Witness"] as string,
                    CreatedDate = reader["CreatedDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["CreatedDate"]),
                    CreateBy = reader["CreateBy"] as string,
                    UpdateDate = reader["UpdateDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["UpdateDate"]),
                    UpdateBy = reader["UpdateBy"] as string,
                    Flag_Delete = reader["Flag_Delete"] as string,

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