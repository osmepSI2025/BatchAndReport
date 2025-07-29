using System;

namespace BatchAndReport.Models
{
    public class E_ConReport_ECModels
    {
        public int EC_ID { get; set; }
        public string? Contract_Number { get; set; }
        public DateTime? ContractSignDate { get; set; }
        public string? SignatoryName { get; set; }
        public string? EmploymentName { get; set; }
        public string? IdenID { get; set; }
        public string? EmpAddress { get; set; }
        public string? WorkDetail { get; set; }
        public string? WorkPosition { get; set; }
        public DateTime? HiringStartDate { get; set; }
        public DateTime? HiringEndDate { get; set; }
        public decimal? Salary { get; set; }
        public string? OSMEP_Signer { get; set; }
        public string? OSMEP_Witness { get; set; }
        public string? Contract_Signer { get; set; }
        public string? Contract_Witness { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public string? Flag_Delete { get; set; }
        public int? Request_ID { get; set; }
        public string? Contract_Status { get; set; }
    }
}