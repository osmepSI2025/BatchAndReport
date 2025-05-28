using System;
using System.Collections.Generic;

namespace BatchAndReport.Entities;

public partial class MscheduledJob
{
    public int Id { get; set; }

    public string? JobName { get; set; }

    public short? RunHour { get; set; }

    public short? RunMinute { get; set; }

    public bool? IsActive { get; set; }
}
