using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations.Schema;

namespace TemplateSTD.Models.STD;

public partial class CompareCostBom
{
    public string Sapkey { get; set; } = null!;

    public string SapModel { get; set; } = null!;

    public string SapPlant { get; set; } = null!;

    public string SapParentMat { get; set; } = null!;

    public string SapComponent { get; set; } = null!;

    public int SapLv { get; set; }

    public string SapQuantityUnit { get; set; } = null!;

    public string SapNoCost { get; set; } = null!;

    public string SapPhantomItem { get; set; } = null!;

    public string SapStdPrice { get; set; } = null!;

    public decimal SapTotalScrap { get; set; }

    public string SapPriceQtyUnit { get; set; } = null!;

    public decimal SapTotalQuantity { get; set; }

    public decimal SapSumValue { get; set; }

    public decimal SapSumTotalValue { get; set; }

    public string As400key { get; set; } = null!;

    public string As400Model { get; set; } = null!;

    public string As400Plant { get; set; } = null!;

    public string As400ParentMat { get; set; } = null!;

    public string As400Component { get; set; } = null!;

    public int As400Lv { get; set; }

    public string As400QuantityUnit { get; set; } = null!;

    public string As400NoCost { get; set; } = null!;

    public string As400PhantomItem { get; set; } = null!;

    public string As400StdPrice { get; set; } = null!;

    public decimal As400TotalScrap { get; set; }

    public string As400PriceQtyUnit { get; set; } = null!;

    public decimal As400TotalQuantity { get; set; }

    public decimal As400SumValue { get; set; }

    public decimal As400SumTotalValue { get; set; }

    public string AlternativeUnit { get; set; } = null!;

    public string BaseUnit { get; set; } = null!;

    public string StatusCheckUnit { get; set; } = null!;

    public decimal Numerator { get; set; }

    public decimal Denominator { get; set; }

    public decimal? StdpricePerConvertToSap { get; set; }

    public decimal? SumDeductedScrapInsapbaseunit { get; set; }

    public decimal? SumTotalQuantityInsapbaseunit { get; set; }

    public decimal? SumValueDeductedScrapInsapbaseunit { get; set; }

    public decimal? SumTotalValueInsapbaseunit { get; set; }

    public decimal? DiffStdPrice { get; set; }

    public decimal? DiffDeductedScrapInbaseunit { get; set; }

    public decimal? DiffTotalQuantityInbaseunit { get; set; }

    public decimal? DiffSumValueDeductedScrapInbaseunit { get; set; }

    public decimal? DiffSumTotalValueInsapbaseunit { get; set; }

    public string PercentDiffStdPrice { get; set; } = null!;

    public string PercentDiffDeductedScrapInbaseunit { get; set; } = null!;

    public string PercentDiffTotalQuantityInbaseunit { get; set; } = null!;

    public string PercentDiffSumValueDeductedScrapInbaseunit { get; set; } = null!;

    public string PercentDiffSumTotalValueInsapbaseunit { get; set; } = null!;
    [NotMapped]
    public string Reasondedected { get; set; } = null!;
    [NotMapped]
    public string Reasonincluded { get; set; } = null!;
}
