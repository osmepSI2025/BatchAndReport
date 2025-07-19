using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class ContractParty
    {
        public int Id { get; set; }
        public string? ContractPartyName { get; set; }
        public string? RegType { get; set; }
        public string? RegIden { get; set; }
        public string? RegDetail { get; set; }
        public string? AddressNo { get; set; }
        public string? SubDistrict { get; set; }
        public string? District { get; set; }
        public string? Province { get; set; }
        public string? PostalCode { get; set; }
        public string? FlagActive { get; set; }
    }



}
