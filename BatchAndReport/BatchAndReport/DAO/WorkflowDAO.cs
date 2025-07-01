using BatchAndReport.Entities;
using BatchAndReport.Models;
using DocumentFormat.OpenXml.Wordprocessing;
using iText.Kernel.Pdf.Canvas.Wmf;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using Org.BouncyCastle.Asn1.X509;
using QuestPDF.Infrastructure;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace BatchAndReport.DAO
{
    public class WorkflowDAO
    {
        private readonly SqlConnectionDAO _connectionDAO;
        private readonly K2DBContext_Workflow _k2context_workflow;

        public WorkflowDAO(SqlConnectionDAO connectionDAO, K2DBContext_Workflow k2context_workflow)
        {
            _connectionDAO = connectionDAO;
            _k2context_workflow = k2context_workflow;
        }

        // CREATE OR UPDATE
        

    }
}