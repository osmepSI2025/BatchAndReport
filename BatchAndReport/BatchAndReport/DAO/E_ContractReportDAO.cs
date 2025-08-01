﻿using BatchAndReport.Entities;
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
using System.Threading.Tasks;

namespace BatchAndReport.DAO
{
    public class E_ContractReportDAO
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_EContract _k2context_EContract;

        public E_ContractReportDAO(SqlConnectionDAO connectionDAO, K2DBContext_EContract k2context_EContract)
        {
            _connectionDAO = connectionDAO;
            _k2context_EContract = k2context_EContract;
        }


        public async Task<E_ConReport_JOADetailModels?> GetJOAAsync(string Joa_id)
        {
            var conn = _k2context_EContract.Database.GetDbConnection();
            await using var connection = new SqlConnection(conn.ConnectionString);
            await using var command = new SqlCommand("sp_Preview_JOA", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            command.Parameters.AddWithValue("@JOA_ID_INPUT", Joa_id);
            await connection.OpenAsync();

            using var reader = await command.ExecuteReaderAsync();
            if (!await reader.ReadAsync()) return null;

            var detail = new E_ConReport_JOADetailModels
            {
                JOA_ID = reader["JOA_ID"] as string,
                Contract_Number = reader["Contract_Number"] as string,
                Project_Name = reader["Project_Name"] as string,
                Organization = reader["Organization"] as string,
                Contract_SignDate = reader["Contract_SignDate"] as DateTime?,
                IssueOwner = reader["IssueOwner"] as string,
                IssueOwnerPosition = reader["IssueOwnerPosition"] as string,
                JointOfficer = reader["JointOfficer"] as string,
                JointOfficerPosition = reader["JointOfficerPosition"] as string,
                Contract_Type = reader["Contract_Type"] as string,
                Contract_Type_Other = reader["Contract_Type_Other"] as string,
                Grant_Date = reader["Grant_Date"] as DateTime?,
                OfficeLoc = reader["OfficeLoc"] as string,
                Contract_Start_Date = reader["Contract_Start_Date"] as DateTime?,
                Contract_End_Date = reader["Contract_End_Date"] as DateTime?,
                Contract_Value = reader["Contract_Value"] as decimal?,
                Contract_Category = reader["Contract_Category"] as string,
                Contract_Storage = reader["Contract_Storage"] as string,
                OSMEP_Signer = reader["OSMEP_Signer"] as string,
                Contract_Signer = reader["Contract_Signer"] as string,
                CreatedDate = reader["CreatedDate"] as DateTime?,
                CreateBy = reader["CreateBy"] as string,
                UpdateDate = reader["UpdateDate"] as DateTime?,
                UpdateBy = reader["UpdateBy"] as string,
                Flag_Delete = reader["Flag_Delete"] as bool?,
                Request_ID = reader["Request_ID"] as string,
                Contract_Status = reader["Contract_Status"] as string
            };

            return detail;
        }

        public async Task<List<E_ConReport_JOAPoposeModels>> GetJOAPoposeAsync(string? Joa_id = "0")
        {
            try
            {

                var result = new List<E_ConReport_JOAPoposeModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT JOAP_ID, JOA_ID, Detail
        FROM JOA_Purpose
        WHERE JOA_ID = @JOA_ID", connection);

                command.Parameters.AddWithValue("@JOA_ID", Joa_id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_JOAPoposeModels
                    {
                        JOAP_ID = reader["JOAP_ID"] as int?,
                        JOA_ID = reader["JOA_ID"] as int?,
                        Detail = reader["Detail"] as string
                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public async Task<E_ConReport_MOUModels?> GetMOUAsync(string Joa_id)
        {
            var conn = _k2context_EContract.Database.GetDbConnection();
            await using var connection = new SqlConnection(conn.ConnectionString);
            await using var command = new SqlCommand("sp_Preview_MOU", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            command.Parameters.AddWithValue("@MOU_ID_INPUT", Joa_id);
            await connection.OpenAsync();

            using var reader = await command.ExecuteReaderAsync();
            if (!await reader.ReadAsync()) return null;

            var detail = new E_ConReport_MOUModels
            {
                MOU_ID = reader["MOU_ID"] as string,
                MOU_Number = reader["MOU_Number"] as string,
                ProjectTitle = reader["ProjectTitle"] as string,
                OrgName = reader["OrgName"] as string,
                OrgCommonName = reader["OrgCommonName"] as string,
                Sign_Date = reader["Sign_Date"] as DateTime?,
                Requestor = reader["Requestor"] as string,
                RequestorPosition = reader["RequestorPosition"] as string,
                Org_Requestor = reader["Org_Requestor"] as string,
                Org_RequestorPosition = reader["Org_RequestorPosition"] as string,
                Contract_Type = reader["Contract_Type"] as string,
                Contract_Type_Other = reader["Contract_Type_Other"] as string,
                Effective_Date = reader["Effective_Date"] as DateTime?,
                Office_Loc = reader["Office_Loc"] as string,
                Start_Date = reader["Start_Date"] as DateTime?,
                End_Date = reader["End_Date"] as DateTime?,
                Contract_Value = reader["Contract_Value"] as decimal?,
                Contract_Category = reader["Contract_Category"] as string,
                Contract_Storage = reader["Contract_Storage"] as string,
                OSMEP_Signer = reader["OSMEP_Signer"] as string,
                OSMEP_Witness = reader["OSMEP_Witness"] as string,
                Contract_Signer = reader["Contract_Signer"] as string,
                Contract_Witness = reader["Contract_Witness"] as string,
                CreatedDate = reader["CreatedDate"] as DateTime?,
                CreateBy = reader["CreateBy"] as string,
                UpdateDate = reader["UpdateDate"] as DateTime?,
                UpdateBy = reader["UpdateBy"] as string,
                Flag_Delete = reader["Flag_Delete"] as bool?,
                Request_ID = reader["Request_ID"] as string,
                Contract_Status = reader["Contract_Status"] as string
            };

            return detail;
        }

        public async Task<List<E_ConReport_MOUPoposeModels>> GetMOUPoposeAsync(string? Mou_id = "0")
        {
            try
            {

                var result = new List<E_ConReport_MOUPoposeModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT MOUP_ID, MOU_ID, Detail
        FROM MOU_Purpose
        WHERE MOU_ID = @MOU_ID", connection);

                command.Parameters.AddWithValue("@MOU_ID", Mou_id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_MOUPoposeModels
                    {
                        MOUP_ID = reader["MOUP_ID"] as int?,
                        MOU_ID = reader["MOU_ID"] as int?,
                        Detail = reader["Detail"] as string
                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }



        #region PDPA    
        public async Task<E_ConReport_PDPAModels?> GetPDPAAsync(string Joa_id)
        {
            var conn = _k2context_EContract.Database.GetDbConnection();
            await using var connection = new SqlConnection(conn.ConnectionString);
            await using var command = new SqlCommand("sp_Preview_PDPA", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            command.Parameters.AddWithValue("@PDPA_ID_INPUT", Joa_id);
            await connection.OpenAsync();

            using var reader = await command.ExecuteReaderAsync();
            if (!await reader.ReadAsync()) return null;

            var detail = new E_ConReport_PDPAModels
            {
                PDPA_ID = reader["PDPA_ID"] as int? ?? 0,
                Contract_Number = reader["Contract_Number"] as string ?? "",
                Project_Name = reader["Project_Name"] as string ?? "",
                Contract_Organization = reader["Contract_Organization"] as string ?? "",
                Master_Contract_Number = reader["Master_Contract_Number"] as string ?? "",
                Master_Contract_Sign_Date = reader["Master_Contract_Sign_Date"] as DateTime?,
                ContractPartyName = reader["ContractPartyName"] as string ?? "",
                ContractPartyCommonName = reader["ContractPartyCommonName"] as string ?? "",
                ContractPartyType = reader["ContractPartyType"] as string ?? "",
                ContractPartyType_Other = reader["ContractPartyType_Other"] as string ?? "",
                OSMEP_ScopeRightsDuties = reader["OSMEP_ScopeRightsDuties"] as string ?? "",
                Contract_Ref_Name = reader["Contract_Ref_Name"] as string ?? "",
                Start_Date = reader["Start_Date"] as DateTime?,
                End_Date = reader["End_Date"] as DateTime?,
                Contract_Category = reader["Contract_Category"] as string ?? "",
                Contract_Storage = reader["Contract_Storage"] as string ?? "",
                Objectives = reader["Objectives"] as string ?? "",
                Objectives_Other = reader["Objectives_Other"] as string ?? "",
                RecordFreq =  reader["RecordFreq"] is int val ? val : 0,
                RecordFreqUnit = reader["RecordFreqUnit"] as string ?? "",
                RetentionPeriodDays = reader["RetentionPeriodDays"] as int?,
                IncidentNotifyPeriod = reader["IncidentNotifyPeriod"] as int?,
                OSMEP_Signer = reader["OSMEP_Signer"] as string ?? "",
                OSMEP_Witness = reader["OSMEP_Witness"] as string ?? "",
                Contract_Signer = reader["Contract_Signer"] as string ?? "",
                Contract_Witness = reader["Contract_Witness"] as string ?? "",
                CreatedDate = reader["CreatedDate"] as DateTime?,
                CreateBy = reader["CreateBy"] as string ?? "",
                UpdateDate = reader["UpdateDate"] as DateTime?,
                UpdateBy = reader["UpdateBy"] as string ?? "",
                Flag_Delete = reader["Flag_Delete"] as string ?? "",
                Request_ID = reader["Request_ID"] as string ?? "",
                Contract_Status = reader["Contract_Status"] as string ?? ""
            };

            return detail;
        }

        public async Task<List<E_ConReport_PDPAObjectModels>> GetPDPA_ObjectivesAsync(string? id = "0")
        {
            try
            {

                var result = new List<E_ConReport_PDPAObjectModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT PD_Objectives_ID, PDPA_ID, Objective_Description
        FROM PDPA_Objectives
        WHERE PDPA_ID = @PDPA_ID", connection);

                command.Parameters.AddWithValue("@PDPA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_PDPAObjectModels
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

        public async Task<List<E_ConReport_PDPAAgreementListModels>> GetPDPA_AgreementListAsync(string? id = "0")
        {
            try
            {

                var result = new List<E_ConReport_PDPAAgreementListModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT PD_List_ID, PDPA_ID, PD_Detail
        FROM PDPA_Agreement_List
        WHERE PDPA_ID = @PDPA_ID", connection);

                command.Parameters.AddWithValue("@PDPA_ID", id ?? "0");
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
        #endregion PDPA

        #region 4.1.1.2.5.บันทึกข้อตกลงการเป็นผู้ควบคุมข้อมูลส่วนบุคคลร่วมตัวอย่างหน้าจอ
        public async Task<E_ConReport_JDCAModels?> GetJDCAAsync(string Joa_id)
        {
            var conn = _k2context_EContract.Database.GetDbConnection();
            await using var connection = new SqlConnection(conn.ConnectionString);
            await using var command = new SqlCommand("sp_Preview_JDCA", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            command.Parameters.AddWithValue("@JDCA_ID_INPUT", Joa_id);
            await connection.OpenAsync();

            using var reader = await command.ExecuteReaderAsync();
            if (!await reader.ReadAsync()) return null;

            var detail = new E_ConReport_JDCAModels
            {
                JDCA_ID = reader["JDCA_ID"] is int id ? id : 0,
                Contract_Number = reader["Contract_Number"] as string,
                Project_Name = reader["Project_Name"] as string,
                Contract_Party_Name = reader["Contract_Party_Name"] as string,
                Contract_Party_Abb_Name = reader["Contract_Party_Abb_Name"] as string,
                Contract_Party_Type = reader["Contract_Party_Type"] as string,
                Contract_Party_Type_Other = reader["Contract_Party_Type_Other"] as string,
                MOU_Name = reader["MOU_Name"] as string,
                Master_Contract_Number = reader["Master_Contract_Number"] as string,
                Master_Contract_Sign_Date = reader["Master_Contract_Sign_Date"] as DateTime?,
                Contract_Category = reader["Contract_Category"] as string,
                Contract_Storage = reader["Contract_Storage"] as string,
                OSMEP_ContRep = reader["OSMEP_ContRep"] as string,
                OSMEP_ContRep_Contact = reader["OSMEP_ContRep_Contact"] as string,
                OSMEP_DPO = reader["OSMEP_DPO"] as string,
                OSMEP_DPO_Contact = reader["OSMEP_DPO_Contact"] as string,
                CP_ContRep = reader["CP_ContRep"] as string,
                CP_ContRep_Contact = reader["CP_ContRep_Contact"] as string,
                CP_DPO = reader["CP_DPO"] as string,
                CP_DPO_Contact = reader["CP_DPO_Contact"] as string,
                OSMEP_Signer = reader["OSMEP_Signer"] as string,
                OSMEP_Witness = reader["OSMEP_Witness"] as string,
                Contract_Signer = reader["Contract_Signer"] as string,
                Contract_Witness = reader["Contract_Witness"] as string,
                CreatedDate = reader["CreatedDate"] as DateTime?,
                CreateBy = reader["CreateBy"] as string,
                UpdateDate = reader["UpdateDate"] as DateTime?,
                UpdateBy = reader["UpdateBy"] as string,
                Flag_Delete = reader["Flag_Delete"] is bool flag ? flag : false,
                Request_ID = reader["Request_ID"] as string,
                Contract_Status = reader["Contract_Status"] as string
            };

            return detail;
        }

        public async Task<List<E_ConReport_PDPAObjectModels>> GetJDCA_JointPurpAsync(string? id = "0")
        {
            try
            {

                var result = new List<E_ConReport_PDPAObjectModels>();
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
                    result.Add(new E_ConReport_PDPAObjectModels
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

        public async Task<List<E_ConReport_JDCA_SubProcessActivitiesModels>> GetJDCA_SubProcessActivitiesAsync(string? id = "0")
        {
            try
            {

                var result = new List<E_ConReport_JDCA_SubProcessActivitiesModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT SubPA_ID, JDCA_ID, Activity,LegalBasis,PersonalData,Owner
        FROM JDCA_SubProcessActivities
        WHERE JDCA_ID = @JDCA_ID", connection);

                command.Parameters.AddWithValue("@JDCA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_JDCA_SubProcessActivitiesModels
                    {
                        SubPA_ID = reader["SubPA_ID"] as int?,
                        JDCA_ID = reader["JDCA_ID"] as int?,
                        Activity = reader["Activity"] as string,
                        LegalBasis = reader["LegalBasis"] as string,
                        PersonalData = reader["PersonalData"] as string,
                        Owner = reader["Owner"] as string

                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        #endregion


        #region 4.1.1.2.7.สัญญาการรักษาข้อมูลที่เป็นความลับ
        public async Task<E_ConReport_NDAModels?> GetNDAAsync(string Joa_id)
        {
            var conn = _k2context_EContract.Database.GetDbConnection();
            await using var connection = new SqlConnection(conn.ConnectionString);
            await using var command = new SqlCommand("sp_Preview_NDA", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            command.Parameters.AddWithValue("@NDA_ID_INPUT", Joa_id);
            await connection.OpenAsync();

            using var reader = await command.ExecuteReaderAsync();
            if (!await reader.ReadAsync()) return null;

            var detail = new E_ConReport_NDAModels
            {
                NDA_ID = reader["NDA_ID"] is int id ? id : 0,
                Contract_Number = reader["Contract_Number"] as string,
                Contract_Party_Name = reader["Contract_Party_Name"] as string,
                Sign_Date = reader["Sign_Date"] as DateTime?,
                OSMEP_Signatory = reader["OSMEP_Signatory"] as string,
                OSMEP_Position = reader["OSMEP_Position"] as string,
                CP_Signatory = reader["CP_Signatory"] as string,
                CP_Position = reader["CP_Position"] as string,
                Contract_Type = reader["Contract_Type"] as string,
                Contract_Type_Other = reader["Contract_Type_Other"] as string,
                OfficeLoc = reader["OfficeLoc"] as string,
                Contract_Category = reader["Contract_Category"] as string,
                Contract_Storage = reader["Contract_Storage"] as string,
                Ref_Name = reader["Ref_Name"] as string,
                EnforcePeriods = reader["EnforcePeriods"] as string,
                OSMEP_Signer = reader["OSMEP_Signer"] as string,
                OSMEP_Witness = reader["OSMEP_Witness"] as string,
                Contract_Signer = reader["Contract_Signer"] as string,
                Contract_Witness = reader["Contract_Witness"] as string,
                CreatedDate = reader["CreatedDate"] as DateTime?,
                CreateBy = reader["CreateBy"] as string,
                UpdateDate = reader["UpdateDate"] as DateTime?,
                UpdateBy = reader["UpdateBy"] as string,
                Flag_Delete = reader["Flag_Delete"] is bool flag ? flag : false,
                Request_ID = reader["Request_ID"] as string,
                Contract_Status = reader["Contract_Status"] as string
            };

            return detail;
        }
        public async Task<List<E_ConReport_NDAConfidentialTypeModels>> GetNDA_ConfidentialTypeAsync(string? id = "0")
        {
            try
            {

                var result = new List<E_ConReport_NDAConfidentialTypeModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT Conf_ID, NDA_ID, Detail
        FROM NDA_ConfidentialType
        WHERE NDA_ID = @NDA_ID", connection);

                command.Parameters.AddWithValue("@NDA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new E_ConReport_NDAConfidentialTypeModels
                    {
                        Conf_ID = reader["Conf_ID"] as int?,
                        NDA_ID = reader["NDA_ID"] as int?,
                        Detail = reader["Detail"] as string
                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        public async Task<List<E_ConReport_NDA_RequestPurposeModels>> GetNDA_RequestPurposeAsync(string? id = "0")
        {
            try
            {

                var result = new List<E_ConReport_NDA_RequestPurposeModels>();
                await using var connection = _connectionDAO.GetConnectionK2Econctract();
                await using var command = new SqlCommand(@"
        SELECT Detail, RP_ID, NDA_ID
        FROM PDPA_Agreement_List
        WHERE NDA_ID = @NDA_ID", connection);

                command.Parameters.AddWithValue("@NDA_ID", id ?? "0");
                await connection.OpenAsync();

                using var reader = await command.ExecuteReaderAsync();
                while (await reader.ReadAsync())
                {
                    result.Add(new  E_ConReport_NDA_RequestPurposeModels
                    {
                      Detail = reader["Detail"] as string,
                        RP_ID = reader["RP_ID"] as int?,
                        NDA_ID = reader["NDA_ID"] as int?
                    });
                }
                return result;
            }
            catch (Exception ex)
            {
                return null;
            }

        }

        #endregion
    }
}