using BatchAndReport.Entities;
using BatchAndReport.Models;
using DocumentFormat.OpenXml.InkML;
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
using System.Linq;
using System.Text.RegularExpressions;
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
        public async Task<WFProcessDetailModels?> GetProcessDetailAsync(int annualProcessReviewId)
        {
            // 1. JOIN AnnualProcessReview กับ ProjectFiscalYear (ใช้ _k2context_workflow เท่านั้น)
            var reviewJoin = await (
                from a in _k2context_workflow.AnnualProcessReviews
                join f in _k2context_workflow.ProjectFiscalYears on a.FiscalYearId equals f.FiscalYearId into fiscalGroup
                from f in fiscalGroup.DefaultIfEmpty()
                where a.AnnualProcessReviewId == annualProcessReviewId
                select new
                {
                    Review = a,
                    ProjectFiscalYear = f
                }
            ).FirstOrDefaultAsync();

            if (reviewJoin == null) return null;

            var review = reviewJoin.Review;
            var fiscalYearDesc = reviewJoin.ProjectFiscalYear?.FiscalYearDesc;
            int fiscalYear = int.TryParse(fiscalYearDesc, out var parsedYear) ? parsedYear : 0;

            // 🔹 Query BusinessUnitOwner (ใช้ _dbContext แยกต่างหาก)
            string? businessUnitOwner = await _dbContext.BusinessUnits
                .Where(b => b.BusinessUnitId == review.OwnerBusinessUnitId)
                .Select(b => b.NameTh)
                .FirstOrDefaultAsync();

            // 🔹 รายการ Review Detail
            var details = await _k2context_workflow.AnnualProcessReviewDetails
                .Where(d => d.AnnualProcessReviewId == annualProcessReviewId)
                .ToListAsync();

            // 🔹 ชื่อกระบวนการเก่า
            var prevDetailIds = details
                .Where(d => d.PrevAnnualProcessReviewDetailId != null)
                .Select(d => d.PrevAnnualProcessReviewDetailId!.Value)
                .Distinct()
                .ToList();

            var prevProcessNames = await _k2context_workflow.AnnualProcessReviewDetails
                .Where(p => prevDetailIds.Contains(p.AnnualProcessReviewDetailId))
                .Select(p => p.ProcessName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct()
                .ToListAsync();

            // 🔹 ประวัติอนุมัติ
            var history = await _k2context_workflow.AnnualProcessReviewHistories
                .Where(h => h.AnnualProcessReviewId == annualProcessReviewId)
                .ToListAsync();

            var approverIds = history
                .Where(h => h.StatusCode == "APRH01" || h.StatusCode == "APRH03")
                .Select(h => h.EmployeeId)
                .Distinct()
                .ToList();

            var approverInfo = await _dbContext.Employees
                .Where(e => approverIds.Contains(e.EmployeeId))
                .Include(e => e.Position)
                .Select(e => new
                {
                    e.EmployeeId,
                    Name = e.NameTh,
                    Position = e.Position.NameTh
                })
                .ToDictionaryAsync(e => e.EmployeeId);

            var approver1Id = history.FirstOrDefault(h => h.StatusCode == "APRH01")?.EmployeeId;
            var approver2Id = history.FirstOrDefault(h => h.StatusCode == "APRH03")?.EmployeeId;

            return new WFProcessDetailModels
            {
                FiscalYear = fiscalYear,
                FiscalYearPrevious = (fiscalYear - 1).ToString(),
                BusinessUnitOwner = businessUnitOwner,
                ReviewDetails = review.ProcessReviewDetail?.Split('\n') ?? [],
                ApproveRemarks = review.ApproveRemark?.Split('\n') ?? [],
                PrevProcesses = prevProcessNames,
                CurrentProcesses = details
                .Where(d => d.ProcessReviewTypeId != null)
                .Select(d => d.ProcessReviewTypeId switch
                {
                    1 => d.ProcessName + "*",
                    2 => d.ProcessName + "**",
                    3 => d.ProcessName + "***",
                    _ => d.ProcessName
                })
                .ToList(),

                ControlActivities = details
                .Where(d => d.IsCgdControlProcess == true)
                .Select(d => d.ProcessReviewTypeId switch
                {
                    1 => d.ProcessName + "*",
                    2 => d.ProcessName + "**",
                    3 => d.ProcessName + "***",
                    null => d.ProcessName,
                    _ => d.ProcessName
                })
                .ToList(),

                WorkflowProcesses = details.Where(d => d.IsWorkflow == true).Select(d => d.ProcessName).ToList(),

                Approver1Name = approver1Id != null && approverInfo.ContainsKey(approver1Id) ? approverInfo[approver1Id].Name : null,
                Approver1Position = approver1Id != null && approverInfo.ContainsKey(approver1Id) ? approverInfo[approver1Id].Position : null,
                Approver2Name = approver2Id != null && approverInfo.ContainsKey(approver2Id) ? approverInfo[approver2Id].Name : null,
                Approver2Position = approver2Id != null && approverInfo.ContainsKey(approver2Id) ? approverInfo[approver2Id].Position : null,
                Approve1Date = history.FirstOrDefault(h => h.StatusCode == "APRH01")?.CreatedDateTime?.ToString("d MMM yyyy", new CultureInfo("th-TH")),
                Approve2Date = history.FirstOrDefault(h => h.StatusCode == "APRH03")?.CreatedDateTime?.ToString("d MMM yyyy", new CultureInfo("th-TH")),
            };
        }



        public async Task<List<WFInternalControlProcessModels>> GetInternalControlProcessesAsync(
        int? fiscalYearId,
        string? businessUnitId = null,
        string? processTypeCode = null,
        string? processGroupCode = null,
        string? processCode = null,
        int? processCategory = null)
        {
            var workflowData = (from detail in _k2context_workflow.AnnualProcessReviewDetails
                                join review in _k2context_workflow.AnnualProcessReviews
                                    on detail.AnnualProcessReviewId equals review.AnnualProcessReviewId
                                join plan_cat_detail in _k2context_workflow.PlanCategoriesDetails
                                    on review.OwnerBusinessUnitId equals plan_cat_detail.BusinessUnitId
                                join plan_cat in _k2context_workflow.PlanCategories
                                    on plan_cat_detail.PlanCategoriesId equals plan_cat.PlanCategoriesId
                                join pm in _k2context_workflow.ProcessMasterDetails
                                    on new { detail.ProcessGroupCode, FiscalYearId = review.FiscalYearId }
                                    equals new { pm.ProcessGroupCode, pm.FiscalYearId }
                                join fcy in _k2context_workflow.ProjectFiscalYears
                                    on review.FiscalYearId equals fcy.FiscalYearId
                                where plan_cat_detail.IsActive == true
                                      && plan_cat_detail.IsDeleted == false
                                      && detail.IsCgdControlProcess == true
                                      && (fiscalYearId == null || review.FiscalYearId == fiscalYearId)
                                      && (businessUnitId == null || review.OwnerBusinessUnitId == businessUnitId)
                                      && (processTypeCode == null || pm.ProcessTypeCode == processTypeCode)
                                      && (processGroupCode == null || detail.ProcessGroupCode == processGroupCode)
                                      && (processCode == null || detail.ProcessCode == processCode)
                                      && (processCategory == null || detail.ProcessReviewTypeId == processCategory)
                                select new WFInternalControlProcessModels
                                {
                                    PlanCategoryName = plan_cat.PlanCategoriesName ?? string.Empty,
                                    BusinessUnitId = plan_cat_detail.BusinessUnitId ?? string.Empty,
                                    Objective = plan_cat_detail.Objective ?? string.Empty,
                                    ProcessCode = detail.ProcessCode ?? string.Empty,
                                    ProcessName = detail.ProcessName ?? string.Empty
                                }).GroupBy(x => new { x.PlanCategoryName, x.BusinessUnitId, x.Objective, x.ProcessCode, x.ProcessName })
                                .Select(g => g.First())
                                .ToList();


            var query = from wf in workflowData
                        join bu in _dbContext.BusinessUnits.ToList()
                            on wf.BusinessUnitId equals bu.BusinessUnitId
                        where wf.BusinessUnitId.Contains(bu.BusinessUnitId)
                        && bu.BusinessUnitLevel == 3

                        select new WFInternalControlProcessModels
                        {
                            PlanCategoryName = wf.PlanCategoryName,
                            BusinessUnitId = bu.NameTh ?? string.Empty,
                            Objective = wf.Objective,
                            ProcessCode = wf.ProcessCode,
                            ProcessName = wf.ProcessName
                        };

            var ordered = query
            .AsEnumerable()
            .OrderBy(x => Regex.Match(x.ProcessCode ?? "", @"^\D+").Value) // prefix
            .ThenBy(x =>
            {
                var match = Regex.Match(x.ProcessCode ?? "", @"\d+");
                return match.Success ? int.Parse(match.Value) : 0;
            });

            return await Task.FromResult(
                ordered
                    .ToList()
            );
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
            t.PROCESS_REVIEW_TYPE_NAME AS ReviewType,
            CASE WHEN d.IS_DIGITAL = 1 THEN N'ใช่' ELSE N'ไม่ใช่' END AS isDigital
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
                            isDigital = reader["isDigital"]?.ToString()
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
            // Fetch related ProcessMasterDetails for idParam
            var all = await _k2context_workflow.ProcessMasterDetails
                .Where(p => p.ProcessMasterId == idParam)
                .ToListAsync();

            if (all == null || !all.Any())
                return null;

            var fiscalYearId = all.First().FiscalYearId;

            var fiscalYearDesc = await _k2context_workflow.ProjectFiscalYears
                .Where(f => f.FiscalYearId == fiscalYearId)
                .Select(f => f.FiscalYearDesc)
                .FirstOrDefaultAsync();

            if (!int.TryParse(fiscalYearDesc, out var fiscalYear))
                return null;

            // ใช้ Regex แยก Prefix + Number เพื่อจัดเรียงเหมือน SQL
            static (string prefix, int number) SplitCode(string code)
            {
                var match = Regex.Match(code ?? "", @"^([A-Z]+)(\d+)$", RegexOptions.IgnoreCase);
                return match.Success
                    ? (match.Groups[1].Value, int.TryParse(match.Groups[2].Value, out var num) ? num : 0)
                    : (code ?? "", 0);
            }

            var coreProcesses = all
                .Where(p => p.ProcessTypeCode == "CORE")
                .OrderBy(p => SplitCode(p.ProcessGroupCode).prefix)
                .ThenBy(p => SplitCode(p.ProcessGroupCode).number)
                .Select(p => new ProcessGroupItem
                {
                    ProcessGroupCode = p.ProcessGroupCode,
                    ProcessGroupName = p.ProcessGroupName
                })
                .ToList();

            var supportProcesses = all
                .Where(p => p.ProcessTypeCode == "SUPPORT")
                .OrderBy(p => SplitCode(p.ProcessGroupCode).prefix)
                .ThenBy(p => SplitCode(p.ProcessGroupCode).number)
                .Select(p => new ProcessGroupItem
                {
                    ProcessGroupCode = p.ProcessGroupCode,
                    ProcessGroupName = p.ProcessGroupName
                })
                .ToList();

            return new WFProcessDetailModels
            {
                FiscalYear = fiscalYear,
                CoreProcesses = coreProcesses,
                SupportProcesses = supportProcesses
            };
        }
        public async Task<List<WFInternalControlProcessModels>> GetInternalControlProcessesByProcessID(int processId)
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
                              && detail.AnnualProcessReviewDetailId == processId
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
        public async Task<List<WFCreateProcessStatusModels>> GetCreateProcessStatusAsync(
            int? fiscalYearId,
            string? businessUnitId,
            string? processTypeCode,
            string? processGroupCode,
            string? processCode,
            bool? isST01,
            bool? isST0101,
            bool? isST0102,
            bool? isST0103,
            bool? isST0104,
            bool? isST0105
            )
        {
            var result = new WFCreateProcessStatusModels();
            var filter = new List<WFCreateProcessStatusModels>();

            var query = @"
        SELECT 
	        sub_process.SUB_PROCESS_MASTER_ID,
	        fiscal.FISCAL_YEAR_DESC,
	        bu.NameTh, 
	        lookup.LookupValue as PROCESS_TYPE_NAME,
	        sub_process.PROCESS_GROUP_CODE, 
	        pm.PROCESS_GROUP_NAME, 
	        sub_process.PROCESS_CODE,
	        sub_process.PROCESS_NAME,
	        process_review.FISCAL_YEAR_ID,
	        pm.PROCESS_TYPE_CODE,
	        lookup1.LookupValue as STATUS
	
	
        FROM [Workflow].[dbo].[SUB_PROCESS_MASTER] sub_process
        left join (SELECT  a.[ANNUAL_PROCESS_REVIEW_DETAIL_ID]
            ,a.[ANNUAL_PROCESS_REVIEW_ID]
            ,a.[PROCESS_GROUP_CODE]
            ,a.[PROCESS_CODE]
            ,b.OWNER_BusinessUnitId
	        ,b.FISCAL_YEAR_ID
        FROM [Workflow].[dbo].[ANNUAL_PROCESS_REVIEW_DETAIL] a
        left join  [Workflow].[dbo].[ANNUAL_PROCESS_REVIEW] b on a.ANNUAL_PROCESS_REVIEW_ID = b.ANNUAL_PROCESS_REVIEW_ID) process_review 
        on sub_process.PROCESS_CODE = process_review.PROCESS_CODE and sub_process.PROCESS_GROUP_CODE = process_review.PROCESS_GROUP_CODE 
        and sub_process.FISCAL_YEAR_ID = process_review.FISCAL_YEAR_ID and sub_process.[OWNER_BusinessUnitId] = process_review.[OWNER_BusinessUnitId]
        left join [Workflow].[dbo].[PROJECT_FISCAL_YEAR] fiscal on sub_process.FISCAL_YEAR_ID = fiscal.FISCAL_YEAR_ID
        left join hr.dbo.BusinessUnits bu on process_review.OWNER_BusinessUnitId = bu.BusinessUnitId
        left join [Workflow].[dbo].[PROCESS_MASTER_DETAIL] pm on pm.PROCESS_GROUP_CODE = sub_process.PROCESS_GROUP_CODE 
        and process_review.FISCAL_YEAR_ID = pm.FISCAL_YEAR_ID
        left join [Workflow].[dbo].[WF_Lookup] lookup on pm.PROCESS_TYPE_CODE = lookup.LookupCode and lookup.LookupType='ProcessType'  
        left join [Workflow].[dbo].[WF_WFTaskList] task on task.Request_ID = sub_process.SUB_PROCESS_MASTER_ID and task.WF_TYPE = '01'
        left join [Workflow].[dbo].[WF_Lookup] lookup1 on task.STATUS = lookup1.LookupCode and lookup1.LookupType = 'WORKFLOW_STATUS' and lookup1.FlagDelete = 'N'
        WHERE (@pFISCAL_YEAR_ID IS NULL OR process_review.FISCAL_YEAR_ID = @pFISCAL_YEAR_ID) 
        and (@pBusinessUnitId IS NULL OR sub_process.OWNER_BusinessUnitId = @pBusinessUnitId) 
        and (@pProcessTypeCode IS NULL OR pm.PROCESS_TYPE_CODE = @pProcessTypeCode) 
        and (@pProcessGroupCode IS NULL OR sub_process.[PROCESS_GROUP_CODE] = @pProcessGroupCode) 
        and (@pProcessCode IS NULL OR sub_process.[PROCESS_CODE] = @pProcessCode) 
        and sub_process.IS_DELETED !=1 and task.STATUS != 'ST0106'
	    and  (
		    (ISNULL(@is_ST01, 0) = 0 AND ISNULL(@is_ST0101, 0) = 0 AND ISNULL(@is_ST0102, 0) = 0 
			    AND ISNULL(@is_ST0103, 0) = 0 AND ISNULL(@is_ST0104, 0) = 0 AND ISNULL(@is_ST0105, 0) = 0) -- ไม่กรองเลย
            OR
            (task.STATUS = 'ST01' AND @is_ST01 = 1)
            OR
            (task.STATUS = 'ST0101' AND @is_ST0101 = 1)
		        OR
            (task.STATUS = 'ST0102' AND @is_ST0102 = 1)
		        OR
            (task.STATUS = 'ST0103' AND @is_ST0103 = 1)
		        OR
            (task.STATUS = 'ST0104' AND @is_ST0104 = 1)
		        OR
            (task.STATUS = 'ST0105' AND @is_ST0105 = 1)
        )
        ORDER BY sub_process.SUB_PROCESS_MASTER_ID asc;";

            var connStr = _k2context_workflow.Database.GetDbConnection().ConnectionString;

            using (var conn = new SqlConnection(connStr))
            {
                await conn.OpenAsync();

                using (var cmd = new SqlCommand(query, conn))
                {
                    cmd.Parameters.AddWithValue("@pFISCAL_YEAR_ID", (object?)fiscalYearId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@pBusinessUnitId", (object?)businessUnitId ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@pProcessTypeCode", (object?)processTypeCode ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@pProcessGroupCode", (object?)processGroupCode ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@pProcessCode", (object?)processCode ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@is_ST01", (object?)isST01 ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@is_ST0101", (object?)isST0101 ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@is_ST0102", (object?)isST0102 ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@is_ST0103", (object?)isST0103 ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@is_ST0104", (object?)isST0104 ?? DBNull.Value);
                    cmd.Parameters.AddWithValue("@is_ST0105", (object?)isST0105 ?? DBNull.Value);

                    using var reader = await cmd.ExecuteReaderAsync();
                    while (await reader.ReadAsync())
                    {
                        var dto = new WFCreateProcessStatusModels
                        {
                            subProcessMasterId = (int)reader["SUB_PROCESS_MASTER_ID"],
                            FiscalYearDesc = fiscalYearId == null ? null : reader["FISCAL_YEAR_DESC"]?.ToString(),
                            BUNameTh = reader["NameTh"]?.ToString(),
                            ProcessType = reader["PROCESS_TYPE_NAME"]?.ToString(),
                            ProcessGroupCode = reader["PROCESS_GROUP_CODE"]?.ToString(),
                            ProcessGroupName = reader["PROCESS_GROUP_NAME"]?.ToString(),
                            ProcessCode = reader["PROCESS_CODE"]?.ToString(),
                            ProcessName = reader["PROCESS_NAME"]?.ToString(),
                            FiscalYearId = (int?)reader["FISCAL_YEAR_ID"],
                            ProcessTypeCode = reader["PROCESS_TYPE_CODE"]?.ToString(),
                            Status = reader["STATUS"]?.ToString()
                        };

                        filter.Add(dto);
                    }
                }
            }

            result.CreateProcessStatusModels = filter;

            Console.WriteLine($"ProcessDetails count: {result.CreateProcessStatusModels.Count}");

            return result.CreateProcessStatusModels;
        }

    }

}