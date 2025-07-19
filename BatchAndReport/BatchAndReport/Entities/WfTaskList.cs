using System;

namespace BatchAndReport.Entities
{
    public class WfTaskList
    {
        public int? WfTaskListId { get; set; }

        public int? WfId { get; set; }

        public string? Status { get; set; }

        public int? RequestId { get; set; }

        public string? WfType { get; set; }

        public string? CreateBy { get; set; }

        public string? UpdateBy { get; set; }

        public DateTime? CreateDate { get; set; }

        public DateTime? LastUpdate { get; set; }

        public DateTime? CompleteOn { get; set; }

        public string? Owner { get; set; }
    }
}
