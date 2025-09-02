using BatchAndReport.Entities;
using BatchAndReport.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.RazorPages;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Data;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

public class MultiUploadFile : PageModel
{
    private readonly K2DBContext_Workflow _k2context_workflow;

    public MultiUploadFile(K2DBContext_Workflow context)
    {
        _k2context_workflow = context;
    }

}