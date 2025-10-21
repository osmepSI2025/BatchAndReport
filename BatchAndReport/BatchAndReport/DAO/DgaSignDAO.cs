using BatchAndReport.Entities;
using BatchAndReport.Models;
using BatchAndReport.Models.BatchAndReport.Models;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf.Canvas.Wmf;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
//using Org.BouncyCastle.Asn1.X509;
using QuestPDF.Infrastructure;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Threading.Tasks;

namespace BatchAndReport.DAO
{
    public class DgaSignDAO
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public DgaSignDAO(SqlConnectionDAO connectionDAO, K2DBContext_EContract k2context_EContract)
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = k2context_EContract;
        }



        public async Task<List<DgaEsignUrlModels>> GetDgaEsignUrlAsync()
        {
            try
            {

                var result = new List<DgaEsignUrlModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT Id,ServiceCode, ServiceName, Method,UrlProd,UrlDev,Example,CreateDate
        FROM DgaEsignUrl
       ", connection);


                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new DgaEsignUrlModels
                    {
                        CreateDate = reader.GetDateTime(reader.GetOrdinal("CreateDate")),
                        ServiceCode = reader.GetString(reader.GetOrdinal("ServiceCode")) ?? "",
                        ServiceName = reader.GetString(reader.GetOrdinal("ServiceName")) ?? "",
                        Method = reader.GetString(reader.GetOrdinal("Method")) ?? "",
                        UrlProd = reader.GetString(reader.GetOrdinal("UrlProd")) ?? "",
                        UrlDev = reader.GetString(reader.GetOrdinal("UrlDev")) ?? "",
                        //      Example = reader.GetString(reader.GetOrdinal("Example"))??"",
                        ID = reader.GetInt32(reader.GetOrdinal("ID"))

                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public async Task<List<DgaEsingConfigModels>> GetDgaEsignConfigAsync()
        {
            try
            {

                var result = new List<DgaEsingConfigModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT ID,ConsumerKey, ConsumerSecret,Email
        FROM DgaEsignConfig
       ", connection);


                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new DgaEsingConfigModels
                    {

                        ConsumerKey = reader.GetString(reader.GetOrdinal("ConsumerKey")) ?? "",
                        ConsumerSecret = reader.GetString(reader.GetOrdinal("ConsumerSecret")) ?? "",
                        ID = reader.GetInt32(reader.GetOrdinal("ID")),
                        Email = reader.GetString(reader.GetOrdinal("Email")) ?? "",

                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public async Task<List<DgaEsignTemplateModels>> GetDgaEsignTemplateAsync()
        {
            try
            {

                var result = new List<DgaEsignTemplateModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT ID,ContractType, DocumentName,TemplateID,ConsumerKey,FlagActive
        FROM DgaEsignTemplate
       ", connection);


                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new DgaEsignTemplateModels
                    {

                        ConsumerKey = reader.GetString(reader.GetOrdinal("ConsumerKey")) ?? "",
                        ContractType = reader.GetString(reader.GetOrdinal("ContractType")) ?? "",
                        DocumentName = reader.GetString(reader.GetOrdinal("DocumentName")) ?? "",
                        TemplateID = reader.GetString(reader.GetOrdinal("TemplateID")) ?? "",
                        ID = reader.GetInt32(reader.GetOrdinal("ID")),
                        FlagActive = reader.GetString(reader.GetOrdinal("FlagActive")) ?? ""

                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public async Task<int> InsertDgaEsignDocumentAsync(DgaEsignDocumentModels model)
        {
            try
            {
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        INSERT INTO DgaEsign (WFTypeCode, ContractID,DGA_TemplateID
, DGA_DocumentID, DGA_SignatureID, DGA_DocumentDataFile, DGA_DocumentPathFile, SignBy, CreateBy, CreateDate)
        VALUES (@WFTypeCode, @ContractID,@DGA_TemplateID, @DGA_DocumentID, @DGA_SignatureID, @DGA_DocumentDataFile, @DGA_DocumentPathFile, @SignBy, @CreateBy, @CreateDate);
        SELECT SCOPE_IDENTITY();
         ", connection);

                command.Parameters.AddWithValue("@WFTypeCode", model.WFTypeCode ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@ContractID", model.ContractID);
                command.Parameters.AddWithValue("@DGA_TemplateID", model.DGA_TemplateID ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@DGA_DocumentID", model.DGA_DocumentID ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@DGA_SignatureID", model.DGA_SignatureID ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@DGA_DocumentDataFile", model.DGA_DocumentDataFile ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@DGA_DocumentPathFile", model.DGA_DocumentPathFile ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@SignBy", model.SignBy ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@CreateBy", model.CreateBy ?? (object)DBNull.Value);
                command.Parameters.AddWithValue("@CreateDate", model.CreateDate);
                 

                await connection.OpenAsync();
                var result = await command.ExecuteScalarAsync();
                return Convert.ToInt32(result);
            }
            catch (Exception ex)
            {
                return 0;
            }
        }

        public async Task<int> UpdateDgaEsignDocumentAsync(string docId, byte[] datafile)
        {
            try
            {
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
            UPDATE DgaEsign
            SET DGA_DocumentDataFile = @DGA_DocumentDataFile,
                UpdateDate = @UpdateDate
            WHERE DGA_DocumentID = @DGA_DocumentID;
        ", connection);

                command.Parameters.AddWithValue("@DGA_DocumentDataFile", datafile ?? (object)DBNull.Value);       
                command.Parameters.AddWithValue("@UpdateDate", DateTime.Now);
                command.Parameters.AddWithValue("@DGA_DocumentID", docId);

                await connection.OpenAsync();
                var rowsAffected = await command.ExecuteNonQueryAsync();
                return rowsAffected;
            }
            catch (Exception)
            {
                return 0;
            }
        }
        public async Task<List<DgaEsignModels>> GetDgaEsignAsync(string? wfTypeCode = null, int? contractId = null)
        {
            try
            {
                var result = new List<DgaEsignModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();

                var sql = new StringBuilder();
                sql.AppendLine("SELECT");
                sql.AppendLine("ID,");
                sql.AppendLine("WFTypeCode,");
                sql.AppendLine("ContractID,");
                sql.AppendLine("DGA_TemplateID,");
                sql.AppendLine("DGA_DocumentID,");
                sql.AppendLine("DGA_SignatureID,");
                sql.AppendLine("DGA_DocumentDataFile,");
                sql.AppendLine("DGA_DocumentPathFile,");
                sql.AppendLine("SignBy,");
                sql.AppendLine("CreateBy,");
                sql.AppendLine("CreateDate,");
                sql.AppendLine("UpdateDate");
                sql.AppendLine("FROM DgaEsign");

                var whereAdded = false;
                if (!string.IsNullOrWhiteSpace(wfTypeCode))
                {
                    sql.AppendLine(whereAdded ? "AND WFTypeCode = @WFTypeCode" : "WHERE WFTypeCode = @WFTypeCode");
                    whereAdded = true;
                }

                if (contractId.HasValue)
                {
                    sql.AppendLine(whereAdded ? "AND ContractID = @ContractID" : "WHERE ContractID = @ContractID");
                    whereAdded = true;
                }

                sql.AppendLine("ORDER BY ID DESC;");

                await using var command = new SqlCommand(sql.ToString(), connection);

                if (!string.IsNullOrWhiteSpace(wfTypeCode))
                    command.Parameters.AddWithValue("@WFTypeCode", wfTypeCode);
                if (contractId.HasValue)
                    command.Parameters.AddWithValue("@ContractID", contractId.Value);

                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new DgaEsignModels
                    {
                        ID = reader.IsDBNull(reader.GetOrdinal("ID")) ? 0 : reader.GetInt32(reader.GetOrdinal("ID")),
                        WFTypeCode = reader.IsDBNull(reader.GetOrdinal("WFTypeCode")) ? "" : reader.GetString(reader.GetOrdinal("WFTypeCode")),
                        ContractID = reader.IsDBNull(reader.GetOrdinal("ContractID")) ? 0 : reader.GetInt32(reader.GetOrdinal("ContractID")),
                        DGA_DocumentID = reader.IsDBNull(reader.GetOrdinal("DGA_DocumentID")) ? "" : reader.GetString(reader.GetOrdinal("DGA_DocumentID")),
                        DGA_SignatureID = reader.IsDBNull(reader.GetOrdinal("DGA_SignatureID")) ? "" : reader.GetString(reader.GetOrdinal("DGA_SignatureID")),
                        DGA_DocumentDataFile = reader.IsDBNull(reader.GetOrdinal("DGA_DocumentDataFile")) ? Array.Empty<byte>() : (byte[])reader["DGA_DocumentDataFile"],
                        DGA_DocumentPathFile = reader.IsDBNull(reader.GetOrdinal("DGA_DocumentPathFile")) ? "" : reader.GetString(reader.GetOrdinal("DGA_DocumentPathFile")),
                        SignBy = reader.IsDBNull(reader.GetOrdinal("SignBy")) ? "" : reader.GetString(reader.GetOrdinal("SignBy")),
                        CreateBy = reader.IsDBNull(reader.GetOrdinal("CreateBy")) ? "" : reader.GetString(reader.GetOrdinal("CreateBy")),
                        CreateDate = reader.IsDBNull(reader.GetOrdinal("CreateDate")) ? DateTime.MinValue : reader.GetDateTime(reader.GetOrdinal("CreateDate")),
                        UpdateDate = reader.IsDBNull(reader.GetOrdinal("UpdateDate")) ? null : reader.GetDateTime(reader.GetOrdinal("UpdateDate"))
                    });
                }

                return result;
            }
            catch (Exception)
            {
                return null;
            }
        }
    }
}