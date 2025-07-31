using BatchAndReport.Entities;
using BatchAndReport.Models;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class WFSubProcessDetailModels
    {
        public SubProcessMaster? Header { get; set; }
        public List<SubProcessReviewApproval> Approvals { get; set; } = new();
        public List<Evaluation> Evaluations { get; set; } = new();
        public List<SubProcessMasterHistory> Revisions { get; set; } = new();
        public List<SubProcessControlDetail> ControlPoints { get; set; } = new();
        public string? OwnerBusinessUnitName { get; set; }
        public string? DiagramAttachFile { get; set; }

        public List<relate_LawsModels> Listrelate_Laws { get; set; } = new();

        public List<SubProcessReviewApprovalModels> ApprovalsDetail { get; set; } = new();
    }

}
