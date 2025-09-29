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
    }
}