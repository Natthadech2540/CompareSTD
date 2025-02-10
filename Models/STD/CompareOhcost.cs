using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace TemplateSTD.Models.STD;

public partial class CompareOhcost
{
    public string SapModel { get; set; } = null!;

    public string SapPlant { get; set; } = null!;

    public string SapFiscalYear { get; set; } = null!;

    public string SapCostCenter { get; set; } = null!;

    public decimal SapTsQuantity { get; set; }

    public string SapUnitQuantity { get; set; } = null!;

    public decimal SapPricePerUnit { get; set; }

    public string SapPriceQtyUnit { get; set; } = null!;

    public decimal SapCostRate { get; set; }

    public decimal SapOhCostRate { get; set; }

    public decimal SapTotalProcessCost { get; set; }

    public decimal SapTotalOh { get; set; }

    public decimal SapTotalValue { get; set; }

    public string As400Model { get; set; } = null!;

    public string As400Plant { get; set; } = null!;

    public string As400FiscalYear { get; set; } = null!;

    public string As400CostCenter { get; set; } = null!;

    public decimal As400TsQuantity { get; set; }

    public string As400UnitQuantity { get; set; } = null!;

    public decimal As400PricePerUnit { get; set; }

    public string As400PriceQtyUnit { get; set; } = null!;

    public decimal As400CostRate { get; set; }

    public decimal As400OhCostRate { get; set; }

    public decimal As400TotalProcessCost { get; set; }

    public decimal As400TotalOh { get; set; }

    public decimal As400TotalValue { get; set; }

    public decimal? DiffTsQuantity { get; set; }

    public decimal? DiffProcessCostRate { get; set; }

    public decimal? DiffOhCostRate { get; set; }

    public decimal? DiffTotalProcessCost { get; set; }

    public decimal? DiffTotalOh { get; set; }

    public decimal? DiffTotalValue { get; set; }

    public string PercentDiffTsQuantity { get; set; } = null!;

    public string PercentDiffProcessCostRate { get; set; } = null!;

    public string PercentDiffOhCostRate { get; set; } = null!;

    public string PercentDiffTotalProcessCost { get; set; } = null!;

    public string PercentDiffTotalOh { get; set; } = null!;

    public string PercentDiffTotalValue { get; set; } = null!;
    [NotMapped]
    public string Reason { get; set; } = null!;
}
