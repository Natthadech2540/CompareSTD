using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class TotalCostSap
{
    public string FiscalYear { get; set; } = null!;

    public string Model { get; set; } = null!;

    public decimal MaterialCost { get; set; }

    public decimal ProcessCost { get; set; }

    public decimal Ohcost { get; set; }

    public string? Srvpackingpercent { get; set; }

    public string? SrvpackingCost { get; set; }

    public decimal? TotalStdcost { get; set; }

    public decimal? TotalStdcostRound { get; set; }

    public string? Unit { get; set; }

    public string? Priceperunit { get; set; }

    public decimal? TotalTs { get; set; }

    public string? Tsunit { get; set; }
}
