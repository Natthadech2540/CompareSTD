using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class CompareTotalStdcost
{
    public string SapFiscalYear { get; set; } = null!;

    public string SapModel { get; set; } = null!;

    public decimal SapMaterialCost { get; set; }

    public decimal SapProcessCost { get; set; }

    public decimal SapOhcost { get; set; }

    public string SapSrvpackingpercent { get; set; } = null!;

    public string SapSrvpackingCost { get; set; } = null!;

    public decimal SapTotalStdcost { get; set; }

    public decimal SapTotalStdcostRound { get; set; }

    public string SapUnit { get; set; } = null!;

    public string SapPriceperunit { get; set; } = null!;

    public decimal SapTotalTs { get; set; }

    public string SapTsunit { get; set; } = null!;

    public string As400FiscalYear { get; set; } = null!;

    public string As400Model { get; set; } = null!;

    public decimal As400MaterialCost { get; set; }

    public decimal As400ProcessCost { get; set; }

    public decimal As400Ohcost { get; set; }

    public string As400Srvpackingpercent { get; set; } = null!;

    public string As400SrvpackingCost { get; set; } = null!;

    public decimal As400TotalStdcost { get; set; }

    public decimal As400TotalStdcostRound { get; set; }

    public string As400Unit { get; set; } = null!;

    public string As400Priceperunit { get; set; } = null!;

    public decimal As400TotalTs { get; set; }

    public string As400Tsunit { get; set; } = null!;

    public decimal? DiffMaterialCost { get; set; }

    public decimal? DiffProcessCost { get; set; }

    public decimal? DiffOhcost { get; set; }

    public decimal? DiffSrvpackingCost { get; set; }

    public decimal? DiffTotalStdcost { get; set; }

    public decimal? DiffTotalStdcostRound { get; set; }

    public decimal? DiffTotalTs { get; set; }

    public string? PercentDiffMaterialCost { get; set; }

    public string? PercentDiffProcessCost { get; set; }

    public string? PercentDiffOhcost { get; set; }

    public string? PercentDiffSrvpackingCost { get; set; }

    public string? PercentDiffTotalStdcost { get; set; }

    public string? PercentDiffTotalStdcostRound { get; set; }

    public string? PercentDiffTotalTs { get; set; }
}
