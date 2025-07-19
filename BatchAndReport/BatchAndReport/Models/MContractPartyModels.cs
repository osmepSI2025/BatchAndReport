using System;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class MContractPartyModels
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
    public class OrganizationJuristicRoot
    {
        [JsonPropertyName("data")]
        public OrganizationJuristicData? Data { get; set; }
    }

    public class OrganizationJuristicData
    {
        [JsonPropertyName("cd:OrganizationJuristicPerson")]
        public OrganizationJuristicPerson? Person { get; set; }
    }

    public class OrganizationJuristicListRoot
    {
        [JsonPropertyName("data")]
        public List<OrganizationJuristicPerson> Data { get; set; }
    }

 

    public class OrganizationJuristicPerson
    {
        [JsonPropertyName("cd:OrganizationJuristicID")]
        public string? RegIden { get; set; }

        [JsonPropertyName("cd:OrganizationJuristicNameTH")]
        public string? ContractPartyName { get; set; }

        [JsonPropertyName("cd:OrganizationJuristicNameEN")]
        public string? RegDetail { get; set; }

        [JsonPropertyName("cd:OrganizationJuristicType")]
        public string? RegType { get; set; }

        [JsonPropertyName("cd:OrganizationJuristicStatus")]
        public string? FlagActive { get; set; }

        [JsonPropertyName("cd:OrganizationJuristicAddress")]
        public JuristicAddress? Address { get; set; }
    }

    public class JuristicAddress
    {
        [JsonPropertyName("cr:AddressType")]
        public AddressType? AddressType { get; set; }
    }

    public class AddressType
    {
        [JsonPropertyName("cd:AddressNo")]
        public string? AddressNo { get; set; }

        [JsonPropertyName("cd:CitySubDivision")]
        public CitySubDivision? CitySubDivision { get; set; }

        [JsonPropertyName("cd:City")]
        public City? City { get; set; }

        [JsonPropertyName("cd:CountrySubDivision")]
        public CountrySubDivision? CountrySubDivision { get; set; }
    }

    public class CitySubDivision
    {
        [JsonPropertyName("cr:CitySubDivisionTextTH")]
        public string? SubDistrict { get; set; }
    }

    public class City
    {
        [JsonPropertyName("cr:CityTextTH")]
        public string? District { get; set; }
    }

    public class CountrySubDivision
    {
        [JsonPropertyName("cr:CountrySubDivisionTextTH")]
        public string? Province { get; set; }
    }
}
