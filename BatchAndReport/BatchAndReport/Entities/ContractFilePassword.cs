using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class ContractFilePassword
    {
        public Guid GuId { get; set; }
        public string? Password { get; set; }
        public string? EmpId { get; set; }
        public string? Email { get; set; }
        public DateTime? CreateDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public string FlagActive { get; set; } = "1";
    }


}
