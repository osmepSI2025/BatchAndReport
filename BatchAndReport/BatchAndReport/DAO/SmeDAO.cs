using BatchAndReport.Entities;
using BatchAndReport.Models;
using Microsoft.Data.SqlClient;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.EntityFrameworkCore;

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
    }
}