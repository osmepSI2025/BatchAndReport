using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class MEmployeeProfileModels
    {
        public int Id { get; set; }
        public string EmployeeId { get; set; } = null!;
        public string? InternalPhone { get; set; }
        public string? MilitaryStatus { get; set; }
        public string? MailingAddrTh { get; set; }
        public string? MailingAddrEn { get; set; }
        public string? MailingSubdistrict { get; set; }
        public string? MailingDistrict { get; set; }
        public string? MailingProvince { get; set; }
        public string? MailingCountry { get; set; }
        public string? MailingPostCode { get; set; }
        public string? MailingPhoneNo { get; set; }
        public string? RegisAddrTh { get; set; }
        public string? RegisAddrEn { get; set; }
        public string? RegisSubdistrict { get; set; }
        public string? RegisDistrict { get; set; }
        public string? RegisProvince { get; set; }
        public string? RegisCountry { get; set; }
        public string? RegisPostCode { get; set; }
        public string? RegisPhoneNo { get; set; }
        public string? BloodGroup { get; set; }
        public string? Religion { get; set; }
        public string? Race { get; set; }
        public string? Nationality { get; set; }
        public string? JobDetails { get; set; }
        public string? NickName { get; set; }
    }

    public class EmployeeProfileResult
    {
        [JsonPropertyName("id")]
        public int Id { get; set; }
        [JsonPropertyName("employeeId")]
        public string EmployeeId { get; set; } = null!;
        [JsonPropertyName("internalPhone")]
        public string? InternalPhone { get; set; }
        [JsonPropertyName("militaryStatus")]
        public string? MilitaryStatus { get; set; }
        [JsonPropertyName("mailingAddrTh")]
        public string? MailingAddrTh { get; set; }
        [JsonPropertyName("mailingAddrEn")]
        public string? MailingAddrEn { get; set; }
        [JsonPropertyName("mailingSubdistrict")]
        public string? MailingSubdistrict { get; set; }
        [JsonPropertyName("mailingDistrict")]
        public string? MailingDistrict { get; set; }
        [JsonPropertyName("mailingProvince")]
        public string? MailingProvince { get; set; }
        [JsonPropertyName("mailingCountry")]
        public string? MailingCountry { get; set; }
        [JsonPropertyName("mailingPostCode")]
        public string? MailingPostCode { get; set; }
        [JsonPropertyName("mailingPhoneNo")]
        public string? MailingPhoneNo { get; set; }
        [JsonPropertyName("regisAddrTh")]
        public string? RegisAddrTh { get; set; }
        [JsonPropertyName("regisAddrEn")]
        public string? RegisAddrEn { get; set; }
        [JsonPropertyName("regisSubdistrict")]
        public string? RegisSubdistrict { get; set; }
        [JsonPropertyName("regisDistrict")]
        public string? RegisDistrict { get; set; }
        [JsonPropertyName("regisProvince")]
        public string? RegisProvince { get; set; }
        [JsonPropertyName("regisCountry")]
        public string? RegisCountry { get; set; }
        [JsonPropertyName("regisPostCode")]
        public string? RegisPostCode { get; set; }
        [JsonPropertyName("regisPhoneNo")]
        public string? RegisPhoneNo { get; set; }
        [JsonPropertyName("bloodGroup")]
        public string? BloodGroup { get; set; }
        [JsonPropertyName("religion")]
        public string? Religion { get; set; }
        [JsonPropertyName("race")]
        public string? Race { get; set; }
        [JsonPropertyName("nationality")]
        public string? Nationality { get; set; }
        [JsonPropertyName("jobDetails")]
        public string? JobDetails { get; set; }
        [JsonPropertyName("nickName")]
        public string? NickName { get; set; }
    }

    public class ApiEmployeeProfileResponse
    {
        [JsonPropertyName("results")]
        public EmployeeProfileResult Results { get; set; }
    }

    public class ApiListEmployeeProfileResponse
    {
        [JsonPropertyName("pagination")]
        public Pagination? Pagination { get; set; }
        [JsonPropertyName("results")]
        public List<EmployeeProfileResult?> Results { get; set; } = new();
    }

    public class SearchEmployeeProfileModels
    {
        public int page { get; set; }
        public int perPage { get; set; }
    }
}