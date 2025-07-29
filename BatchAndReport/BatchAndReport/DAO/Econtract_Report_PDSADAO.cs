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
    public class Econtract_Report_PDSADAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_PDSADAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_PDSAModels> GetPDSAAsync(string? id = "0")
        {
            try
            {
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT 
                PDSA_ID,
                Contract_Number,
                Project_Name,
                Contract_Organization,
                Master_Contract_Number,
                Master_Contract_Sign_Date,
                ContractPartyName,
                ContractPartyCommonName,
                ContractPartyType,
                ContractPartyType_Other,
                Contract_Category,
                Contract_Storage,
                RetentionPeriodDays,
                IncidentNotifyPeriod,
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
                Contract_Status
            FROM PDSA
            WHERE PDSA_ID = @PDSA_ID", connection);

                command.Parameters.AddWithValue("@PDSA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (await reader.ReadAsync())
                {
                    return new E_ConReport_PDSAModels
                    {
                        PDSA_ID = reader["PDSA_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["PDSA_ID"]),
                        Contract_Number = reader["Contract_Number"] as string,
                        Project_Name = reader["Project_Name"] as string,
                        Contract_Organization = reader["Contract_Organization"] as string,
                        Master_Contract_Number = reader["Master_Contract_Number"] as string,
                        Master_Contract_Sign_Date = reader["Master_Contract_Sign_Date"] == DBNull.Value ? null : Convert.ToDateTime(reader["Master_Contract_Sign_Date"]),
                        ContractPartyName = reader["ContractPartyName"] as string,
                        ContractPartyCommonName = reader["ContractPartyCommonName"] as string,
                        ContractPartyType = reader["ContractPartyType"] as string,
                        ContractPartyType_Other = reader["ContractPartyType_Other"] as string,
                        Contract_Category = reader["Contract_Category"] as string,
                        Contract_Storage = reader["Contract_Storage"] as string,
                        RetentionPeriodDays = reader["RetentionPeriodDays"] == DBNull.Value ? null : Convert.ToInt32(reader["RetentionPeriodDays"]),
                        IncidentNotifyPeriod = reader["IncidentNotifyPeriod"] == DBNull.Value ? null : Convert.ToInt32(reader["IncidentNotifyPeriod"]),
                        OSMEP_Signer = reader["OSMEP_Signer"] as string,
                        OSMEP_Witness = reader["OSMEP_Witness"] as string,
                        Contract_Signer = reader["Contract_Signer"] as string,
                        Contract_Witness = reader["Contract_Witness"] as string,
                        CreatedDate = reader["CreatedDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["CreatedDate"]),
                        CreateBy = reader["CreateBy"] as string,
                        UpdateDate = reader["UpdateDate"] == DBNull.Value ? null : Convert.ToDateTime(reader["UpdateDate"]),
                        UpdateBy = reader["UpdateBy"] as string,
                        Flag_Delete = reader["Flag_Delete"] as string,
                        Request_ID = reader["Request_ID"] == DBNull.Value ? null : Convert.ToInt32(reader["Request_ID"]),
                        Contract_Status = reader["Contract_Status"] as string
                    };
                }
                return null;
            }
            catch (Exception ex)
            {
                return null;
            }
        }

        public async Task<List<PDSA_Shared_Data>> GetPDSA_Shared_DataAsync(string? id = "0")
        {
            try
            {
                var result = new List<PDSA_Shared_Data>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT 
                SharePD_ID,
                PDSA_ID,
                Detail,
                Objective,
                Owner,
                Flag_Delete
            FROM PDSA_Shared_Data
            WHERE PDSA_ID = @PDSA_ID", connection);

                command.Parameters.AddWithValue("@PDSA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new PDSA_Shared_Data
                    {
                        SharePD_ID = reader["SharePD_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["SharePD_ID"]),
                        PDSA_ID = reader["PDSA_ID"] == DBNull.Value ? null : Convert.ToInt32(reader["PDSA_ID"]),
                        Detail = reader["Detail"] as string,
                        Objective = reader["Objective"] as string,
                        Owner = reader["Owner"] as string,
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
        public async Task<List<PDSA_LegalBasisSharing>> GetPDSA_LegalBasisSharingAsync(string? id = "0")
        {
            try
            {
                var result = new List<PDSA_LegalBasisSharing>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            SELECT 
                LglShare_ID,
                PDSA_ID,
                Detail,
                Owner,
                Flag_Delete
            FROM PDSA_LegalBasisSharing 
            WHERE PDSA_ID = @PDSA_ID", connection);

                command.Parameters.AddWithValue("@PDSA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new PDSA_LegalBasisSharing
                    {
                        LglShare_ID = reader["LglShare_ID"] == DBNull.Value ? 0 : Convert.ToInt32(reader["LglShare_ID"]),
                        PDSA_ID = reader["PDSA_ID"] == DBNull.Value ? null : Convert.ToInt32(reader["PDSA_ID"]),
                        Detail = reader["Detail"] as string,
                        Owner = reader["Owner"] as string,
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