using BatchAndReport.Models;
using System.Text.Json.Serialization;

namespace BatchAndReport.Models
{
    public class ConJointContractModels
    {
        public string? ProjectName { get; set; }
        public string? AgencyName { get; set; }
        public string? SMEOfficialName { get; set; }
        public string? SMEOfficialPosition { get; set; }
        public string? AgencyRepresentative { get; set; }
        public string? AgencyPosition { get; set; }
        public string? SignDay { get; set; }
        public string? SignMonth { get; set; }
        public string? SignYear { get; set; }
        public string? SignDateText => $"{SignDay} {SignMonth} พ.ศ. {SignYear}";

        public List<ObjectiveItem> Objectives { get; set; } = new();
        public List<string> SMEDuties { get; set; } = new();
        public List<string> AgencyDuties { get; set; } = new();
        public List<string> OtherTerms { get; set; } = new();
    }

    public class ObjectiveItem
    {
        public string? Number { get; set; }
        public string? Description { get; set; }
    }

}
