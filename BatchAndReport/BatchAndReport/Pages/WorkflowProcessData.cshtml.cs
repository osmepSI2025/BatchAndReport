using BatchAndReport.Entities;
using BatchAndReport.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.Data;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

public class WorkflowProcessDataModel : PageModel
{
    private readonly K2DBContext_Workflow _k2context_workflow;

    public WorkflowProcessDataModel(K2DBContext_Workflow context)
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
            
            FiscalYear = fiscalYear,

            CoreProcesses = all
        .Where(p => p.ProcessTypeCode == "CORE")
        .OrderBy(p => Regex.Match(p.ProcessGroupCode, @"^[A-Za-z]+").Value) // ตัวอักษรนำหน้า
        .ThenBy(p =>
        {
            var match = Regex.Match(p.ProcessGroupCode, @"\d+");
            return match.Success ? int.Parse(match.Value) : int.MaxValue;
        }) // ตัวเลขต่อท้าย
        .Select(p => new ProcessGroupItem
        {
            ProcessGroupCode = p.ProcessGroupCode,
            ProcessGroupName = p.ProcessGroupName
        })
        .ToList(),

            SupportProcesses = all
        .Where(p => p.ProcessTypeCode == "SUPPORT")
        .OrderBy(p => Regex.Match(p.ProcessGroupCode, @"^[A-Za-z]+").Value)
        .ThenBy(p =>
        {
            var match = Regex.Match(p.ProcessGroupCode, @"\d+");
            return match.Success ? int.Parse(match.Value) : int.MaxValue;
        })
        .Select(p => new ProcessGroupItem
        {
            ProcessGroupCode = p.ProcessGroupCode,
            ProcessGroupName = p.ProcessGroupName
        })
        .ToList()
        };

        Detail = detail; // ✅ ต้องมีบรรทัดนี้
        return Page();
    }
}