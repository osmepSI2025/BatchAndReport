namespace BatchAndReport.Models
{
    public class ContractAllowanceModels
    {
        public string? AgreementNumber { get; set; }
        public string? Name { get; set; }
        public string? Age { get; set; }
        public string? nationality { get; set; }
        public string? HouseNo { get; set; }
        public string? BuildingName { get; set; }

        public string? Mootee { get; set; }
        public string? SubRoad { get; set; }
        public string? Road { get; set; }
        public string? Subdistrict { get; set; }
        public string? District { get; set; }
        public string? Province { get; set; }
        public string? IDCardNo { get; set; }
        public string? DataIdCardNo { get; set; }

        #region ข้อ1
        public decimal? DepositAmount { get; set; }

        public decimal? StrDepositAmount { get; set; }
        public int? NoOfMonth { get; set; }
        #endregion ข้อ1
        #region ข้อ2
        public string? BankBranchName { get; set; }

        public string? NoBookBank { get; set; }
        public string? NameBookBank { get; set; }

        public decimal? Amount2 { get; set; }
        public string? StrAmount2 { get; set; }

        #endregion ข้อ2

        #region ข้อ5
        public decimal? Amount5 { get; set; }

        public string? strAmount5 { get; set; }
        public string? MonthName5 { get; set; }

        public string? Date5 { get; set; }
        public string? Year5 { get; set; }
        public string? ToYear5 { get; set; }

        #endregion ข้อ5


        #region หน้า3
        public string? nameSupportForm { get; set; }
        public string? nameSupportForm_Sign { get; set; }

        public string? nameSupportTo { get; set; }
        public string? nameSupportTo_Sign { get; set; }

        public string? nameSpouse { get; set; }
        public string? nameSpouse_Sign { get; set; }
        public string? nameWitness1 { get; set; }
        public string? nameWitness1_Sign { get; set; }
        public string? nameWitness2 { get; set; }
        public string? nameWitness2_Sign { get; set; }

        #endregion หน้า3

        #region หน้า4
        public string? nameMarried { get; set; }
        public string? statusMarried { get; set; }

        public string? nameGuarantee { get; set; }
        public string? nameGuarantee_Sign { get; set; }

        public string? nameMarriedWitness1 { get; set; }
        public string? nameMarriedWitness1_Sign { get; set; }
        public string? nameMarriedWitness2 { get; set; }
        public string? nameMarriedWitness2_Sign { get; set; }

        #endregion หน้า4

    }
}
