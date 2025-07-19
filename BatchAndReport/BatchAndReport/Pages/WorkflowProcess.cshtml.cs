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
        var all = await _k2context_workflow.TempProcessMasterDetails
            .Where(p => p.ProcessMasterId == id_param)
            .ToListAsync();

        Detail = new WFProcessDetailModels
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

        return Page();
    }
}