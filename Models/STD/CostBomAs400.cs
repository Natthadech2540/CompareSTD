using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class CostBomAs400
{
    public string Model { get; set; } = null!;

    public string Plant { get; set; } = null!;

    public string? CostingRun { get; set; }

    public string? CostingRundt { get; set; }

    public int? RunningNo { get; set; }

    public int? Lv { get; set; }

    public string ParentMat { get; set; } = null!;

    public string? ParentMatDesc { get; set; }

    public string? ParentProcTypeMm { get; set; }

    public string? ParentSpTypeMm { get; set; }

    public string Component { get; set; } = null!;

    public string? ComponentDesc { get; set; }

    public string? ItemType { get; set; }

    public string? CompProcTypeMm { get; set; }

    public string? CompSpTypeMm { get; set; }

    public string? CompSpTypeBom { get; set; }

    public string? BulkMat { get; set; }

    public string? CostRelevancyBom { get; set; }

    public string? PhantomItem { get; set; }

    public string? DeletionIndicator { get; set; }

    public string? MatProvisionIndicator { get; set; }

    public string? CompEffectDtFrom { get; set; }

    public string? CompEffectDtTo { get; set; }

    public string? Unit { get; set; }

    public string? MSize { get; set; }

    public string? QuantityModel { get; set; }

    public string? Scrap { get; set; }

    public decimal? TotalScrap { get; set; }

    public decimal? TotalQuantity { get; set; }

    public string QuantityUnit { get; set; } = null!;

    public string? Import { get; set; }

    public string? TaxExp { get; set; }

    public string? Local { get; set; }

    public string? StdPrice { get; set; }

    public string? PriceQtyUnit { get; set; }

    public string? NoPrice { get; set; }

    public string? NoCost { get; set; }

    public decimal? SumValue { get; set; }

    public decimal? SumTotalValue { get; set; }

    public string? PurPriceCur { get; set; }

    public string? ItemClass { get; set; }
}
