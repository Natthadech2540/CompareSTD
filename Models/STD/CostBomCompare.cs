using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class CostBomCompare
{
    public string? KeyAs400 { get; set; }

    public string As400Model { get; set; } = null!;

    public string As400Plant { get; set; } = null!;

    public string As400ParentMat { get; set; } = null!;

    public string As400Component { get; set; } = null!;

    public string? As400QuantityUnit { get; set; }

    public string? As400NoCost { get; set; }

    public string? As400PhantomItem { get; set; }

    public string? As400StdPrice { get; set; }

    public string? As400PriceQtyUnit { get; set; }

    public decimal? As400TotalScrap { get; set; }

    public decimal? As400TotalQuantity { get; set; }

    public decimal? As400SumValue { get; set; }

    public decimal? As400SumTotalValue { get; set; }

    public string? KeySap { get; set; }

    public string SapModel { get; set; } = null!;

    public string SapPlant { get; set; } = null!;

    public string SapParentMat { get; set; } = null!;

    public string SapComponent { get; set; } = null!;

    public string? SapQuantityUnit { get; set; }

    public string? SapNoCost { get; set; }

    public string? SapPhantomItem { get; set; }

    public string? SapStdPrice { get; set; }

    public string? SapPriceQtyUnit { get; set; }

    public decimal? SapTotalScrap { get; set; }

    public decimal? SapTotalQuantity { get; set; }

    public decimal? SapSumValue { get; set; }

    public decimal? SapSumTotalValue { get; set; }
}
