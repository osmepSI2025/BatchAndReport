using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities
{
    public class ProjectYear
    {
        public string FISCAL_YEAR_DESC { get; set; } = null!;

        public DateTime START_DATE { get; set; }

        public DateTime END_DATE { get; set; }

        public DateTime CREATE_DATE { get; set; }

        public string CREATE_BY { get; set; } = null!;

        public string ACTIVE_FLAG { get; set; } = null!;

        public DateTime? UPDATE_DATE { get; set; }

        public string? UPDATE_BY { get; set; }
    }


}
