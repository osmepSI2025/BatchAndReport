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
    public class Econtract_Report_ECDAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_ECDAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
 
        public async Task<E_ConReport_ECModels?> GetECAsync(string? id = "0")
        {
            try
            {
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT 
                EC_ID,
                Contract_Number,
                ContractSignDate,
                SignatoryName,
                EmploymentName,
                IdenID,
                EmpAddress,
                WorkDetail,
                WorkPosition,
                HiringStartDate,
                HiringEndDate,
                Salary,
                OSMEP_Signer,
                OSMEP_Witness,
                Contract_Signer,
                Contract_Witness,
                CreatedDate,
                CreateBy,
                UpdateDate,
                UpdateBy,
                Flag_Delete,
                Request_ID,
                Contract_Status,
AttorneyFlag,
AttorneyLetterDate,AttorneyLetterNumber
            FROM EC
            WHERE EC_ID = @EC_ID", connection);

                command.Parameters.AddWithValue("@EC_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    return new E_ConReport_ECModels
                    {
                        EC_ID = reader["EC_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["EC_ID"]),
                        Contract_Number = reader["Contract_Number"] as string,
                        ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("ContractSignDate")),
                        SignatoryName = reader["SignatoryName"] as string,
                        EmploymentName = reader["EmploymentName"] as string,
                        IdenID = reader["IdenID"] as string,
                        EmpAddress = reader["EmpAddress"] as string,
                        WorkDetail = reader["WorkDetail"] as string,
                        WorkPosition = reader["WorkPosition"] as string,
                        HiringStartDate = reader["HiringStartDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("HiringStartDate")),
                        HiringEndDate = reader["HiringEndDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("HiringEndDate")),
                        Salary = reader["Salary"] == DBNull.Value ? null : Convert.ToDecimal(reader["Salary"]),
                        OSMEP_Signer = reader["OSMEP_Signer"] as string,
                        OSMEP_Witness = reader["OSMEP_Witness"] as string,
                        Contract_Signer = reader["Contract_Signer"] as string,
                        Contract_Witness = reader["Contract_Witness"] as string,
                        CreatedDate = reader["CreatedDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("CreatedDate")),
                        CreateBy = reader["CreateBy"] as string,
                        UpdateDate = reader["UpdateDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("UpdateDate")),
                        UpdateBy = reader["UpdateBy"] as string,
                        Flag_Delete = reader["Flag_Delete"] as string,
                        Request_ID = reader["Request_ID"] as string,
                        Contract_Status = reader["Contract_Status"] as string,
                         AttorneyFlag = reader["AttorneyFlag"] == DBNull.Value ? null : Convert.ToBoolean(reader["AttorneyFlag"]),
                         AttorneyLetterDate = reader["AttorneyLetterDate"] == DBNull.Value ? null : reader.GetDateTime(reader.GetOrdinal("AttorneyLetterDate")),
                         AttorneyLetterNumber = reader["AttorneyLetterNumber"] as string
                    };
                }
                return null;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}