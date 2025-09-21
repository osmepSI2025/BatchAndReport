namespace BatchAndReport.Models
{
    public class E_ConReport_PDPAModels
    {
        public int PDPA_ID { get; set; } // int
        public string Contract_Number { get; set; } // nvarchar(255)
        public string Project_Name { get; set; } // nvarchar(255)
        public string Contract_Organization { get; set; } // nvarchar(100)
        public string Master_Contract_Number { get; set; } // nvarchar(100)
        public DateTime? Master_Contract_Sign_Date { get; set; } // date
        public string ContractPartyName { get; set; } // nvarchar(100)
        public string ContractPartyCommonName { get; set; } // nvarchar(100)
        public string ContractPartyType { get; set; } // nvarchar(100)
        public string ContractPartyType_Other { get; set; } // nvarchar(100)
        public string OSMEP_ScopeRightsDuties { get; set; } // nvarchar(255)
        public string Contract_Ref_Name { get; set; } // nvarchar(100)
        public DateTime? Start_Date { get; set; } // date
        public DateTime? End_Date { get; set; } // date
        public string Contract_Category { get; set; } // nvarchar(100)
        public string Contract_Storage { get; set; } // nvarchar(100)
        public string Objectives { get; set; } // nvarchar(100)
        public string Objectives_Other { get; set; } // nvarchar(100)
        public int? RecordFreq { get; set; } // int
        public string RecordFreqUnit { get; set; } // nvarchar(10)
        public int? RetentionPeriodDays { get; set; } // int
        public int? IncidentNotifyPeriod { get; set; } // int
        public string OSMEP_Signer { get; set; } // varchar(255)
        public string OSMEP_Witness { get; set; } // varchar(255)
        public string Contract_Signer { get; set; } // varchar(255)
        public string Contract_Witness { get; set; } // varchar(255)
        public DateTime? CreatedDate { get; set; } // datetime
        public string CreateBy { get; set; } // nvarchar(100)
        public DateTime? UpdateDate { get; set; } // datetime
        public string UpdateBy { get; set; } // nvarchar(100)
        public string Flag_Delete { get; set; } // nvarchar(1)
        public string Request_ID { get; set; } // nvarchar(50)
        public string Contract_Status { get; set; } // nvarchar(10)
        public string Organization_Logo { get; set; }
        
    }


    public class E_ConReport_PDPAObjectModels {
        public int? PDPA_ID { get; set; }
        public int? PD_Objectives_ID { get; set; }
        public string? Objective_Description { get; set; }
    }
    public class E_ConReport_PDPAAgreementListModels
    {
        public int? PD_List_ID { get; set; }
        public string? PDPA_ID { get; set; }
        public string? PD_Detail { get; set; }
    }
}
