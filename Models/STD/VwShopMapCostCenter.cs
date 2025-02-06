using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class VwShopMapCostCenter
{
    public string ShopCode { get; set; } = null!;

    public int ColumnNo { get; set; }

    public string? OldShopCode { get; set; }

    public string? CostCenter { get; set; }

    public string? GeneralName { get; set; }

    public string? Description { get; set; }

    public string? Department { get; set; }

    public string? Name1 { get; set; }
}
