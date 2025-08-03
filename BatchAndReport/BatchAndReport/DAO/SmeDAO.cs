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
using System.Threading.Tasks;

namespace BatchAndReport.DAO
{
    public class SmeDAO
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_SME _k2context_sme;

        public SmeDAO(SqlConnectionDAO connectionDAO, K2DBContext_SME k2context_sme)
        {
            _connectionDAO = connectionDAO;
            _k2context_sme = k2context_sme;
        }

        // CREATE OR UPDATE
        public async Task InsertOrUpdateProjectMasterAsync(List<MProjectMasterModels> projects)
        {
            try
            {
                foreach (var pro in projects)
                {
                    if (pro.KeyId == null)
                        continue;

                    var existingPro = await _k2context_sme.ProjectMasters
                        .FirstOrDefaultAsync(e => e.KeyId == pro.KeyId.Value);

                    if (existingPro != null)
                    {
                        existingPro.ProjectName = pro.ProjectName ?? "";
                        existingPro.BudgetAmount = pro.BudgetAmount ?? 0;
                        existingPro.Issue = pro.Issue ?? "";
                        existingPro.Strategy = pro.Strategy ?? "";
                        existingPro.FiscalYear = pro.FiscalYear ?? "";

                        _k2context_sme.ProjectMasters.Update(existingPro);
                    }
                    else
                    {
                        var newPro = new ProjectMaster
                        {
                            KeyId = pro.KeyId.Value,
                            ProjectName = pro.ProjectName ?? "",
                            BudgetAmount = pro.BudgetAmount ?? 0,
                            Issue = pro.Issue ?? "",
                            Strategy = pro.Strategy ?? "",
                            FiscalYear = pro.FiscalYear ?? "",
                        };

                        await _k2context_sme.ProjectMasters.AddAsync(newPro);
                    }
                }

                await _k2context_sme.SaveChangesAsync();
            }
            catch (Exception ex)
            {
                // Log to console or use ILogger
                Console.WriteLine($"InsertOrUpdateProjectMasterAsync ERROR: {ex.Message}\n{ex.StackTrace}");
                throw; // Let the controller catch and return the error
            }
        }

        public async Task<List<string>> GetDistinctFiscalYearsAsync()
        {
            return await _k2context_sme.ProjectYears
                .Select(f => f.FISCAL_YEAR_DESC)
                .Distinct()
                .ToListAsync();
        }

        public async Task<SMEProjectDetailModels?> GetProjectDetailAsync(string projectCode)
        {
            var conn = _k2context_sme.Database.GetDbConnection();
            await using var connection = new SqlConnection(conn.ConnectionString);
            await using var command = new SqlCommand("SP_SME_PROJECT_DETAIL", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            command.Parameters.AddWithValue("@ProjectCode", projectCode);
            await connection.OpenAsync();

            using var reader = await command.ExecuteReaderAsync();
            if (!await reader.ReadAsync()) return null;

            var detail = new SMEProjectDetailModels
            {
                ProjectCode = reader["PROJECT_CODE"]?.ToString(),
                ProjectName = reader["PROJECT_NAME"]?.ToString(),
                MinistryName = reader["MINISTRY_NAME"]?.ToString(),
                DepartmentCode = reader["DEPARTMENT_CODE"]?.ToString(),
                DepartmentName = reader["NameTh"]?.ToString(),
                FiscalYear = Convert.ToInt32(reader["FISCAL_YEAR"]),
                BudgetAmount = Convert.ToDecimal(reader["BUDGET_AMOUNT"]),
                BudgetAmountApprove = reader["BUDGET_AMOUNT_APPROVE"] as decimal?,
                StrategyDesc = reader["STRATEGY_DESC"]?.ToString(),
                Score = reader["SCORE"] as int?,
                ProjectStatusName = reader["PROJECT_STATUS_NAME"]?.ToString(),
                ProjectRationale = reader["PROJECT_RATIONALE"]?.ToString(),
                ProjectObjective = reader["PROJECT_OBJECTIVE"]?.ToString(),
                TargetGroup = reader["TARGET_GROUP_CODE"]?.ToString(),
                StartDate = Convert.ToDateTime(reader["START_DATE"]),
                EndDate = Convert.ToDateTime(reader["END_DATE"]),
                ProjectFocus = reader["PROJECT_HIGHLIGHT"]?.ToString(),

                OperationArea = new List<string>(),
                IndustrySector = new List<string>(),
                OutputIndicators = new List<Indicator>() // Initialize OutputIndicators as a list of Indicator objects
         
            , FiscalYearDesc = reader["FISCAL_YEAR_DESC"]?.ToString(),
            };

            var requestId = reader["REQUEST_ID"]?.ToString();
            var budYear = reader["FISCAL_YEAR_DESC"]?.ToString();

            await reader.CloseAsync();

            if (!string.IsNullOrEmpty(requestId))
            {
                await using var areaCmd = new SqlCommand(
                    "SELECT OPERATION_AREA_NAME FROM SME_PROJECT_REQUEST_OPERATION_AREA WHERE REQUEST_ID = @RequestId",
                    connection
                );
                areaCmd.Parameters.AddWithValue("@RequestId", requestId);

                using var areaReader = await areaCmd.ExecuteReaderAsync();
                while (await areaReader.ReadAsync())
                {
                    var areaName = areaReader["OPERATION_AREA_NAME"]?.ToString();
                    if (!string.IsNullOrEmpty(areaName))
                        detail.OperationArea.Add(areaName);
                }
                await areaReader.CloseAsync();
            }

            if (!string.IsNullOrEmpty(requestId))
            {
                await using var targetCmd = new SqlCommand(
                    "SELECT TARGET_SECTOR_DESC FROM SME_TARGET_SECTOR WHERE REQUEST_ID = @RequestId",
                    connection
                );
                targetCmd.Parameters.AddWithValue("@RequestId", requestId);

                using var targetReader = await targetCmd.ExecuteReaderAsync();
                while (await targetReader.ReadAsync())
                {
                    var targetName = targetReader["TARGET_SECTOR_DESC"]?.ToString();
                    if (!string.IsNullOrEmpty(targetName))
                        detail.IndustrySector.Add(targetName);
                }
                await targetReader.CloseAsync();
            }

            if (!string.IsNullOrEmpty(requestId))
            {
                await using var indCmd = new SqlCommand(
                    "SELECT INDICATOR_NAME , TARGET_VALUE , UNIT_NAME , MEASUREMENT_METHOD FROM SME_PERFORMANCE_INDICATORS WHERE REQUEST_ID = @RequestId and INDICATOR_TYPE = '01'",
                    connection
                );
                indCmd.Parameters.AddWithValue("@RequestId", requestId);

                using var indReader = await indCmd.ExecuteReaderAsync();
                while (await indReader.ReadAsync())
                {
                    var indicator = new Indicator
                    {
                        Name = indReader["INDICATOR_NAME"]?.ToString(),
                        Target = indReader["TARGET_VALUE"]?.ToString(),
                        Unit = indReader["UNIT_NAME"]?.ToString(),
                        Method = indReader["MEASUREMENT_METHOD"]?.ToString()
                    };

                    if (!string.IsNullOrEmpty(indicator.Name))
                        detail.OutputIndicators.Add(indicator); // Add Indicator object to the list
                }
                await indReader.CloseAsync();

                await using var indCmd1 = new SqlCommand(
                    "SELECT INDICATOR_NAME , TARGET_VALUE , UNIT_NAME , MEASUREMENT_METHOD FROM SME_PERFORMANCE_INDICATORS WHERE REQUEST_ID = @RequestId and INDICATOR_TYPE = '02'",
                    connection
                );
                indCmd1.Parameters.AddWithValue("@RequestId", requestId);

                using var indReader1 = await indCmd1.ExecuteReaderAsync();
                while (await indReader1.ReadAsync())
                {
                    var indicator = new Indicator
                    {
                        Name = indReader1["INDICATOR_NAME"]?.ToString(),
                        Target = indReader1["TARGET_VALUE"]?.ToString(),
                        Unit = indReader1["UNIT_NAME"]?.ToString(),
                        Method = indReader1["MEASUREMENT_METHOD"]?.ToString()
                    };

                    if (!string.IsNullOrEmpty(indicator.Name))
                        detail.OutputIndicators.Add(indicator); // Add Indicator object to the list
                }
                await indReader1.CloseAsync();

                await using var stdCmd = new SqlCommand(
                    "SELECT ST.STRATEGY_ID,ST.TOPIC,STD.STRATEGY_DESC FROM SME_SME_PROJECT_STRATEGY ST " +
                    "LEFT JOIN SME_PROJECT_STRATEGY_DETAIL STD ON ST.STRATEGY_ID = STD.STRATEGY_ID " +
                    "LEFT JOIN SME_PROJECT_FISCAL_YEAR FY ON ST.FISCAL_YEAR_ID = FY.FISCAL_YEAR_ID WHERE FY.FISCAL_YEAR_DESC = @BudYear",
                    connection
                );
                stdCmd.Parameters.AddWithValue("@BudYear", budYear);

                using var stdReader = await stdCmd.ExecuteReaderAsync();
                while (await stdReader.ReadAsync())
                {
                    var strategy = new Strategy
                    {
                        Topic = stdReader["TOPIC"]?.ToString(),
                        StrategyDesc = stdReader["STRATEGY_DESC"]?.ToString(),
                    };

                    if (!string.IsNullOrEmpty(strategy.Topic))
                        detail.Strategies.Add(strategy); // Add Indicator object to the list
                }
                await stdReader.CloseAsync();
            }

            return detail;
        }


        public async Task<List<SMESummaryProjectModels>> GetSummaryProjectAsync(string budYear)
        {
            var result = new List<SMESummaryProjectModels>();
            var conn = _k2context_sme.Database.GetDbConnection();
            await using var connection = new SqlConnection(conn.ConnectionString); // เปิด connection ใหม่
            await using var command = new SqlCommand("SP_SME_PROJECT_LIST", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            command.Parameters.AddWithValue("@BudYear", budYear);
            await connection.OpenAsync();

            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                result.Add(new SMESummaryProjectModels
                {
                    IssueName = reader["TOPIC"].ToString(),
                    IssueId = reader["STRATEGY_ID"] != DBNull.Value ? Convert.ToInt32(reader["STRATEGY_ID"]) : 0,
                    Budget = reader["SumBudget"] != DBNull.Value ? Convert.ToDecimal(reader["SumBudget"]) : 0,
                    ProjectCount = reader["ProjectCount"] != DBNull.Value ? Convert.ToInt32(reader["ProjectCount"]) : 0
                });
            }

            return result;
        }

        public async Task<List<SMEStrategyDetailModels>> GetProjectStrategyAsync(string budYear)
        {
            var results = new List<SMEStrategyDetailModels>();

            var conn = _k2context_sme.Database.GetDbConnection();
            await using var connection = new SqlConnection(conn.ConnectionString);
            await using var command = new SqlCommand("SP_SME_PROJECT_STRATEGY", connection)
            {
                CommandType = CommandType.StoredProcedure
            };

            command.Parameters.AddWithValue("@BudYear", budYear);
            await connection.OpenAsync();

            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var item = new SMEStrategyDetailModels
                {
                    StrategyId = reader["STRATEGY_ID"] != DBNull.Value ? Convert.ToInt32(reader["STRATEGY_ID"]) : (int?)null, // Fix for CS0029
                    Topic = reader["TOPIC"]?.ToString(),
                    StrategyDesc = reader["STRATEGY_DESC"]?.ToString(),
                    DepartmentCode = reader["DEPARTMENT_CODE"]?.ToString(),
                    Department = reader["NameTh"]?.ToString(),
                    ProjectName = reader["PROJECT_NAME"]?.ToString(),
                    BudgetAmount = reader["BUDGET_AMOUNT"] as decimal?,
                    ProjectStatus = reader["ProjectStatus"]?.ToString()
                    , Ministry_Id = reader["MINISTRY_ID"]?.ToString(),
                    Ministry_Name = reader["MINISTRY_NAME"]?.ToString()
                };

                results.Add(item);
            }

            return results;
        }

        public async Task InsertOrUpdateFiscalYearsAsync(List<int> fiscalYearsChristianEra)
        {
            // แปลง ค.ศ. → พ.ศ.
            var fiscalYears = fiscalYearsChristianEra.Select(year => year + 543).ToList();

            // ดึงปี พ.ศ. ที่มีอยู่แล้วจากฐานข้อมูล
            List<int> existingYears; // Declare the variable outside the try block
            try
            {
                existingYears = await _k2context_sme.FiscalYears
                    .Select(y => int.Parse(y.FISCAL_YEAR_DESC)) // Convert string to int
                    .Distinct()
                    .ToListAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("🔥 ERROR in FiscalYears query:");
                Console.WriteLine(ex.Message);
                Console.WriteLine(ex.InnerException?.Message);
                Console.WriteLine(ex.StackTrace);
                throw; // rethrow เพื่อดูผลลัพธ์ 500 หรือให้ API handler จัดการ
            }

            foreach (var year in fiscalYears)
            {
                if (!existingYears.Contains(year))
                {
                    _k2context_sme.FiscalYears.Add(new FiscalYear
                    {
                        FISCAL_YEAR_DESC = year.ToString(), // Convert int back to string
                        // CREATE_DATE = DateTime.Now (ถ้ามี)
                    });
                }
            }

            await _k2context_sme.SaveChangesAsync();
        }




    }
}