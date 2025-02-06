using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class ProcessOhcostSap
{
    public string Model { get; set; } = null!;

    public string Plant { get; set; } = null!;

    public string FiscalYear { get; set; } = null!;

    public string CostCenter { get; set; } = null!;

    public string? TsQuantity { get; set; }

    public string? UnitQuantity { get; set; }

    public string? PricePerUnit { get; set; }

    public string? PriceQtyUnit { get; set; }

    public string? ProcCostRate { get; set; }

    public string? OhCostRate { get; set; }

    public string? TotalProcessCost { get; set; }

    public string? TotalOh { get; set; }

    public string? TotalValue { get; set; }
}
