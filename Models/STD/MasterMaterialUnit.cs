using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class MasterMaterialUnit
{
    public string Material { get; set; } = null!;

    public string? Description { get; set; }

    public string? AlternativeUnit { get; set; }

    public decimal? Numerator { get; set; }

    public decimal? Denominator { get; set; }

    public string? BaseUnit { get; set; }
}
