using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class ConversionUnit
{
    public long No { get; set; }

    public string Material { get; set; } = null!;

    public string? MaterialDescription { get; set; }

    public string? OldMaterialNumber { get; set; }

    public string MaterialType { get; set; } = null!;

    public string BaseUnitOfMeasure { get; set; } = null!;

    public string? OrderUnit { get; set; }

    public string AlternativeUnitOfMeasure { get; set; } = null!;

    public decimal YNumerator { get; set; }

    public decimal XDenominator { get; set; }

    public string? BulkMaterial { get; set; }

    public string? IssueUnit { get; set; }

    public DateTime? Createdatetime { get; set; }

    public DateTime? Updatedatetime { get; set; }
}
