using System.ComponentModel.DataAnnotations;

namespace BatchAndReport.Entities
{
    public class DiagramFileEntity
    {
        public int DiagramId { get; set; }

        public string? DiagramAttachFile { get; set; }

        public bool? IsDeleted { get; set; }

        public DateTime? CreatedDatetime { get; set; }

        public DateTime? UpdatedDatetime { get; set; }

     
        public string? CreatedBy { get; set; }

     
        public string? UpdatedBy { get; set; }

        public int? SubProcessMasterId { get; set; }
    }
}
