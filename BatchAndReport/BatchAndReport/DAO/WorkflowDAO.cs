using BatchAndReport.Entities;
using BatchAndReport.Models;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf.Canvas.Wmf;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
//using Org.BouncyCastle.Asn1.X509;
using QuestPDF.Infrastructure;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Threading.Tasks;

namespace BatchAndReport.DAO
{
    public class WorkflowDAO
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_Workflow _k2context_workflow;
        private readonly K2DBContext _dbContext; // Add missing field declaration  

        public WorkflowDAO(SqlConnectionDAO connectionDAO, K2DBContext_Workflow k2context_workflow, K2DBContext dbContext)
        {
            _connectionDAO = connectionDAO;
            _k2context_workflow = k2context_workflow;
            _dbContext = dbContext; // Initialize the missing field  
        }

        public async Task<WFProcessDetailModels> GetProcessDetailAsync(int fiscalYear)
        {
            var review = await _k2context_workflow.AnnualProcessReviews
                .Include(r => r.ProjectFiscalYear)
                .FirstOrDefaultAsync(r => r.ProjectFiscalYear.FiscalYearDesc == fiscalYear.ToString());

            if (review == null) return null;

            var details = await _k2context_workflow.AnnualProcessReviewDetails
                .Where(d => d.AnnualProcessReviewId == review.AnnualProcessReviewId)
                .ToListAsync();

            var history = await _k2context_workflow.AnnualProcessReviewHistories
                .Where(h => h.AnnualProcessReviewId == review.AnnualProcessReviewId)
                .ToListAsync();

            var approverIds = history
                .Where(h => h.StatusCode == "APRH01" || h.StatusCode == "APRH03")
                .Select(h => h.EmployeeId)
                .Distinct()
                .ToList();

            var approverInfo = await _dbContext.Employees
                .Where(e => approverIds.Contains(e.EmployeeId))
                .Include(e => e.Position) // ensure navigation property exists
                .Select(e => new
                {
                    e.EmployeeId,
                    Name = e.NameTh,
                    Position = e.Position.NameTh
                })
                .ToDictionaryAsync(e => e.EmployeeId);

            var model = new WFProcessDetailModels
            {
                FiscalYear = fiscalYear, // Fix: Change from fiscalYear.ToString() to fiscalYear (int type expected)
                FiscalYearPrevious = (fiscalYear - 1).ToString(),
                ReviewDetails = review.ProcessReviewDetail?.Split('\n') ?? [],
                PrevProcesses = details.Select(d => d.PrevProcessName).Where(name => !string.IsNullOrWhiteSpace(name)).ToList(),
                CurrentProcesses = details.Where(d => d.ProcessReviewTypeId != null).Select(d => d.ProcessName).ToList(),
                ControlActivities = details.Where(d => d.IsCgdControlProcess == true).Select(d => d.ProcessName).ToList(),
                WorkflowProcesses = details.Where(d => d.IsWorkflow == true).Select(d => d.ProcessName).ToList(),
                ApproveRemarks = review.ApproveRemark?.Split('\n') ?? [],

                Approver1Name = approverInfo.GetValueOrDefault(history.FirstOrDefault(h => h.StatusCode == "APRH01")?.EmployeeId)?.Name,
                Approver1Position = approverInfo.GetValueOrDefault(history.FirstOrDefault(h => h.StatusCode == "APRH01")?.EmployeeId)?.Position,
                Approver2Name = approverInfo.GetValueOrDefault(history.FirstOrDefault(h => h.StatusCode == "APRH03")?.EmployeeId)?.Name,
                Approver2Position = approverInfo.GetValueOrDefault(history.FirstOrDefault(h => h.StatusCode == "APRH03")?.EmployeeId)?.Position,
                Approve1Date = history.FirstOrDefault(h => h.StatusCode == "APRH01")?.CreatedDateTime?.ToString("d MMM yyyy", new CultureInfo("th-TH")),
                Approve2Date = history.FirstOrDefault(h => h.StatusCode == "APRH03")?.CreatedDateTime?.ToString("d MMM yyyy", new CultureInfo("th-TH")),
            };

            return model;
        }
        public async Task<List<WFInternalControlProcessModels>> GetInternalControlProcessesAsync(int processID)
        {
            var query = from detail in _k2context_workflow.AnnualProcessReviewDetails
                        join process_review in _k2context_workflow.AnnualProcessReviews
                            on detail.AnnualProcessReviewId equals process_review.AnnualProcessReviewId
                        join plan_cat_detail in _k2context_workflow.PlanCategoriesDetails
                            on process_review.OwnerBusinessUnitId equals plan_cat_detail.BusinessUnitId
                        join plan_cat_detail1 in _k2context_workflow.PlanCategoriesDetails
                            on plan_cat_detail.PlanCategoriesId equals plan_cat_detail1.PlanCategoriesId
                        join plan_cat in _k2context_workflow.PlanCategories
                            on plan_cat_detail.PlanCategoriesId equals plan_cat.PlanCategoriesId
                        where plan_cat_detail.IsActive == true
                              && plan_cat_detail.IsDeleted == false
                              && detail.IsCgdControlProcess == true
                              && detail.AnnualProcessReviewDetailId == processID
                        select new WFInternalControlProcessModels
                        {
                            PlanCategoryName = plan_cat.PlanCategoriesName,
                            BusinessUnitId = plan_cat_detail1.BusinessUnitId,
                            Objective = plan_cat_detail1.Objective,
                            ProcessCode = detail.ProcessCode,
                            ProcessName = detail.ProcessName
                        };

            return await query
                .GroupBy(x => new { x.PlanCategoryName, x.BusinessUnitId, x.Objective, x.ProcessCode, x.ProcessName })
                .Select(g => g.First())
                .ToListAsync();
        }
        public async Task<WorkSystemModels> GetWorkSystemDataAsync(
        int? fiscalYearId,
        string? businessUnitId = null,
        string? processTypeCode = null,
        string? processGroupCode = null,
        string? processCode = null,
        int? processCategory = null)
            {
                var result = new WorkSystemModels();
                var processDetails = new List<ProcessDetailDto>();

                var query = @"
        SELECT 
            d.PROCESS_CODE,
            d.PROCESS_NAME,
            d.PREV_PROCESS_CODE,
            bu.NameTh AS Department,
            CASE WHEN d.IS_WORKFLOW = 1 THEN N'✓' ELSE N'-' END AS Workflow,
            CASE WHEN d.PREV_IS_WORKFLOW = 1 THEN N'✓' ELSE N'-' END AS PrevWorkflow,
            CASE WHEN d.IS_WI = 1 THEN N'✓' ELSE N'-' END AS WI,
            t.PROCESS_REVIEW_TYPE_NAME AS ReviewType
        FROM ANNUAL_PROCESS_REVIEW_DETAIL d
        INNER JOIN ANNUAL_PROCESS_REVIEW r
            ON d.ANNUAL_PROCESS_REVIEW_ID = r.ANNUAL_PROCESS_REVIEW_ID
        INNER JOIN PROCESS_MASTER_DETAIL pm
            ON d.PROCESS_GROUP_CODE = pm.PROCESS_GROUP_CODE
            AND r.FISCAL_YEAR_ID = pm.FISCAL_YEAR_ID
        INNER JOIN PROCESS_REVIEW_TYPE t
            ON d.PROCESS_REVIEW_TYPE_ID = t.PROCESS_REVIEW_TYPE_ID
        INNER JOIN HR.dbo.BusinessUnits bu
            ON LTRIM(RTRIM(LOWER(r.OWNER_BusinessUnitId))) = LTRIM(RTRIM(LOWER(bu.BusinessUnitId)))
        WHERE 
            d.IS_DELETED != 1
            AND (@FISCAL_YEAR_ID IS NULL OR r.FISCAL_YEAR_ID = @FISCAL_YEAR_ID)
            AND (@OWNER_BusinessUnitId IS NULL OR r.OWNER_BusinessUnitId = @OWNER_BusinessUnitId)
            AND (@PROCESS_TYPE_CODE IS NULL OR pm.PROCESS_TYPE_CODE = @PROCESS_TYPE_CODE)
            AND (@PROCESS_GROUP_CODE IS NULL OR d.PROCESS_GROUP_CODE = @PROCESS_GROUP_CODE)
            AND (@PROCESS_CODE IS NULL OR d.PROCESS_CODE = @PROCESS_CODE)
            AND (@PROCESS_CATEGORY IS NULL OR d.PROCESS_REVIEW_TYPE_ID = @PROCESS_CATEGORY);";

                var connStr = _k2context_workflow.Database.GetDbConnection().ConnectionString;

                using (var conn = new SqlConnection(connStr))
                {
                    await conn.OpenAsync();

                    using (var cmd = new SqlCommand(query, conn))
                    {
                        cmd.Parameters.AddWithValue("@FISCAL_YEAR_ID", (object?)fiscalYearId ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@OWNER_BusinessUnitId", (object?)businessUnitId ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@PROCESS_TYPE_CODE", (object?)processTypeCode ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@PROCESS_GROUP_CODE", (object?)processGroupCode ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@PROCESS_CODE", (object?)processCode ?? DBNull.Value);
                        cmd.Parameters.AddWithValue("@PROCESS_CATEGORY", (object?)processCategory ?? DBNull.Value);

                        using var reader = await cmd.ExecuteReaderAsync();
                        while (await reader.ReadAsync())
                        {
                            var dto = new ProcessDetailDto
                            {
                                ProcessCode = reader["PROCESS_CODE"]?.ToString(),
                                ProcessName = reader["PROCESS_NAME"]?.ToString(),
                                PrevProcessCode = reader["PREV_PROCESS_CODE"]?.ToString(),
                                Department = reader["Department"]?.ToString(),
                                Workflow = reader["Workflow"]?.ToString(),
                                PrevWorkflow = reader["PrevWorkflow"]?.ToString(),
                                WI = reader["WI"]?.ToString(),
                                ReviewType = reader["ReviewType"]?.ToString()
                            };

                            processDetails.Add(dto);
                        }
                    }
                }

                result.ProcessDetails = processDetails;

                Console.WriteLine($"ProcessDetails count: {result.ProcessDetails.Count}");

                return result;
            }

        public async Task<WFSubProcessDetailModels?> GetSubProcessDetailAsync(int subProcessId)
        {
            var model = new WFSubProcessDetailModels();

            var header = await _k2context_workflow.SubProcessMasters
                .AsNoTracking()
                .FirstOrDefaultAsync(x => x.SubProcessMasterId == subProcessId);

            if (header == null)
                return null;

            model.Header = header;

            var ownerBU = await _k2context_workflow.AnnualProcessReviewDetails
                .Include(d => d.AnnualProcessReview)
                .Where(d =>
                    d.ProcessCode == header.ProcessCode &&
                    d.ProcessGroupCode == header.ProcessGroupCode &&
                    d.AnnualProcessReview.FiscalYearId == header.FiscalYearId)
                .Select(d => d.AnnualProcessReview.OwnerBusinessUnitId)
                .FirstOrDefaultAsync();

            if (!string.IsNullOrEmpty(ownerBU))
            {
                model.OwnerBusinessUnitName = await _dbContext.BusinessUnits
                    .Where(b => b.BusinessUnitId == ownerBU)
                    .Select(b => b.NameTh)
                    .FirstOrDefaultAsync();
            }

            model.Approvals = await _k2context_workflow.SubProcessReviewApprovals
                .Where(x => x.SubProcessMasterId == subProcessId)
                .ToListAsync();

            var revisionRaw = await _k2context_workflow.Set<SubProcessMasterHistory>()
                .FromSqlRaw(@"
                    SELECT 
                        SUB_PROCESS_MASTER_HISTORY_ID, 
                        EDIT_DETAIL, 
                        TRY_CONVERT(datetime, DATETIME) AS DATETIME,  -- 👈 ต้องใช้ alias เป็น 'DATETIME'
                        PROCESS_MASTER_HISTORY_TYPE,
                        SUB_PROCESS_MASTER_ID,
                        STATUS_CODE,
                        EMPLOYEE_ID,
                        CREATED_DATETIME,
                        UPDATED_DATETIME,
                        CREATED_BY,
                        UPDATED_BY,
                        IS_DELETED
                    FROM SUB_PROCESS_MASTER_HISTORY
                    WHERE SUB_PROCESS_MASTER_ID = {0} 
                    AND PROCESS_MASTER_HISTORY_TYPE = 'PMH01'", subProcessId)
                .AsNoTracking()
                .ToListAsync();


            model.Revisions = revisionRaw
                .OrderBy(x => x.DateTime ?? DateTime.MinValue)
                .ToList();

            model.Evaluations = await _k2context_workflow.Evaluations
                .Where(x => x.SubProcessMasterId == subProcessId)
                .ToListAsync();

            model.ControlPoints = await _k2context_workflow.SubProcessControlDetails
                .Where(x => x.SubProcessMasterId == subProcessId)
                .ToListAsync();
            model.DiagramAttachFile = header.DiagramAttachFile;

            return model;
        }
        public async Task<WFProcessDetailModels?> GetWFProcessDetailAsync(int idParam)
        {
            var all = await _k2context_workflow.ProcessMasterDetails
                .Where(p => p.ProcessMasterId == idParam)
                .ToListAsync();

            var detail = new WFProcessDetailModels
            {
                FiscalYear = 2568,
                CoreProcesses = all
                    .Where(p => p.ProcessTypeCode == "CORE")
                    .OrderBy(p => p.ProcessGroupCode)
                    .Select(p => new ProcessGroupItem
                    {
                        ProcessGroupCode = p.ProcessGroupCode,
                        ProcessGroupName = p.ProcessGroupName
                    }).ToList(),

                SupportProcesses = all
                    .Where(p => p.ProcessTypeCode == "SUPPORT")
                    .OrderBy(p => p.ProcessGroupCode)
                    .Select(p => new ProcessGroupItem
                    {
                        ProcessGroupCode = p.ProcessGroupCode,
                        ProcessGroupName = p.ProcessGroupName
                    }).ToList()
            };

            return detail;
        }

    }
}