using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class MasterCostCenter
{
    public long No { get; set; }

    public string OldShopCode { get; set; } = null!;

    public string CostCenter { get; set; } = null!;

    public string GeneralName { get; set; } = null!;

    public string Description { get; set; } = null!;

    public string Department { get; set; } = null!;

    public string Name1 { get; set; } = null!;

    public DateTime Createdatetime { get; set; }

    public DateTime Updatedatetime { get; set; }

    public string FileName { get; set; } = null!;
}
