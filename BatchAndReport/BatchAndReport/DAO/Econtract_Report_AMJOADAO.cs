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
    public class Econtract_Report_AMJOADAO // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public Econtract_Report_AMJOADAO(SqlConnectionDAO connectionDAO
            ,K2DBContext_EContract context
            
            ) // Fixed spelling error: Changed "Econtract_SPADAO" to "Econtract_SPADAO"  
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = context;

        }
        public async Task<E_ConReport_AMJOAModels?> GetAMJOAAsync(string id)
        {
            try
            {
                var conn = _k2context_EContract.Database.GetDbConnection();
                await using var connection = new SqlConnection(conn.ConnectionString);
                await using var command = new SqlCommand("sp_Preview_AMJOA", connection)
                {
                    CommandType = CommandType.StoredProcedure
                };

                command.Parameters.AddWithValue("@AMJOA_ID_INPUT", id);
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                if (!await reader.ReadAsync()) return null;

                var detail = new E_ConReport_AMJOAModels
                {
                    AMJOA_ID = reader["AMJOA_ID"] == DBNull.Value ? null : reader["AMJOA_ID"].ToString(),
                    Contract_Number = reader["Contract_Number"] == DBNull.Value ? null : reader["Contract_Number"].ToString(),
                    RefContract_Number = reader["RefContract_Number"] == DBNull.Value ? null : reader["RefContract_Number"].ToString(),
                    Contract_Name = reader["Contract_Name"] == DBNull.Value ? null : reader["Contract_Name"].ToString(),
                    ContractSignDate = reader["ContractSignDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("ContractSignDate")),
                    Start_Unit = reader["Start_Unit"] == DBNull.Value ? null : reader["Start_Unit"].ToString(),
                    Contract_Partner = reader["Contract_Partner"] == DBNull.Value ? null : reader["Contract_Partner"].ToString(),
                    Contract_Description = reader["Contract_Description"] == DBNull.Value ? null : reader["Contract_Description"].ToString(),
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
                    CP_S_AttorneyLetterDate = reader["CP_S_AttorneyLetterDate"] == DBNull.Value ? (DateTime?)null : reader.GetDateTime(reader.GetOrdinal("CP_S_AttorneyLetterDate")),
                    CP_S_AttorneyLetterNumbers = reader["CP_S_AttorneyLetterNumbers"] == DBNull.Value ? null : reader["CP_S_AttorneyLetterNumbers"].ToString()
                };

                return detail;
            }
            catch (Exception ex)
            {
                return null; // Consider logging the exception
            }
        }
        public async Task<List<E_ConReport_AMJOAObjectModels>> GetJDCA_JointPurpAsync(string? id = "0")
        {
            try
            {

                var result = new List<E_ConReport_AMJOAObjectModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT JP_ID, JDCA_ID, Detail
        FROM JDCA_JointPurp
        WHERE JDCA_ID = @JDCA_ID", connection);

                command.Parameters.AddWithValue("@JDCA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_AMJOAObjectModels
                    {
                        PD_Objectives_ID = reader["PD_Objectives_ID"] as int?,
                        PDPA_ID = reader["PDPA_ID"] as int?,
                        Objective_Description = reader["Objective_Description"] as string
                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public async Task<List<E_ConReport_PDPAAgreementListModels>> GetJDCA_PurpMeansAsync(string? id = "0")
        {
            try
            {

                var result = new List<E_ConReport_PDPAAgreementListModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT PD_List_ID, PDPA_ID, PD_Detail
        FROM PDPA_Agreement_List
        WHERE JDCA_ID = @JDCA_ID", connection);

                command.Parameters.AddWithValue("@JDCA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_PDPAAgreementListModels
                    {
                        PD_List_ID = reader["PD_List_ID"] as int?,
                        PDPA_ID = reader["PDPA_ID"] as string,
                        PD_Detail = reader["PD_Detail"] as string
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