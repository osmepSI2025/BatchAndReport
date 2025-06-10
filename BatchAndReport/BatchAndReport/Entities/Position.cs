using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class Position
    {
        public int Id { get; set; }

        public string? ProjectCode { get; set; }

        public string? PositionId { get; set; }

        public string? TypeCode { get; set; }

        public string? Module { get; set; }

        public string? NameTh { get; set; }

        public string? NameEn { get; set; }

        public string? DescriptionTh { get; set; }

        public string? DescriptionEn { get; set; }
    }


}
