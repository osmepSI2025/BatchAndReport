using BatchAndReport.Entities;
using BatchAndReport.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

public class WorkflowProcessModel : PageModel
{
    private readonly K2DBContext_Workflow _k2context_workflow;

    public WorkflowProcessModel(K2DBContext_Workflow context)
    {
        _k2context_workflow = context;
    }

    public WFProcessDetailModels Detail { get; set; }

    public async Task<IActionResult> OnGetAsync(int id_param)
    {
        // Fetch related ProcessMasterDetails for idParam
        var all = await _k2context_workflow.ProcessMasterDetails
            .Where(p => p.ProcessMasterId == id_param)
            .ToListAsync();

        if (all == null || !all.Any())
            return null;

        // Fetch the first FiscalYearId from the retrieved list
        var fiscalYearId = all.First().FiscalYearId;

        // JOIN to fetch the FiscalYearDesc
        var fiscalYearDesc = await _k2context_workflow.ProjectFiscalYears
            .Where(f => f.FiscalYearId == fiscalYearId)
            .Select(f => f.FiscalYearDesc)
            .FirstOrDefaultAsync();

        // Fix: Parse the FiscalYearDesc string to an integer
        if (!int.TryParse(fiscalYearDesc, out var fiscalYear))
        {
            // Handle the case where parsing fails (e.g., log an error or return null)
            return null;
        }

        var detail = new WFProcessDetailModels
        {
            FiscalYear = fiscalYear, // Assign the parsed integer value
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

        return Page();
    }
}