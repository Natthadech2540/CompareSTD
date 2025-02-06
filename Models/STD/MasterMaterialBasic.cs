using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class MasterMaterialBasic
{
    public string As400Material { get; set; } = null!;

    public string SapMaterial { get; set; } = null!;

    public string? Description { get; set; }

    public string? BaseUnit { get; set; }

    public string? MaterialType { get; set; }

    public string? MaterialGroup { get; set; }

    public decimal? NetWeight { get; set; }

    public string? WeightUnit { get; set; }
}
