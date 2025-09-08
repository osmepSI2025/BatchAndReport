using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class ImportContract
    {
        public int IcId { get; set; }      // PK Identity
        public string? ImportRunningNum { get; set; }
        public string? ContractNumber { get; set; }
        public string? ContractType { get; set; }
        public string? ProjectName { get; set; }
        public string? Owner { get; set; }
        public string? ContractParty { get; set; }
        public string? Domicile { get; set; }
        public DateTime? StartDate { get; set; }
        public DateTime? EndDate { get; set; }
        public string? Status { get; set; }
        public decimal? Amount { get; set; }
        public int? Installment { get; set; }
        public string? ContractStorage { get; set; }
        public int? InstallmentNo { get; set; }
        public DateTime? PaymentDate { get; set; }
        public decimal? InstallmentAmount { get; set; }
        public DateTime? CreateDate { get; set; }
        public string? CreateBy { get; set; }
        public DateTime? UpdateDate { get; set; }
        public string? UpdateBy { get; set; }
        public bool? FlagDelete { get; set; }
    }



}
