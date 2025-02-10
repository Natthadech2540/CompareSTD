using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class MappingSapMaterial
{
    public long No { get; set; }

    public string As400ItemNumber { get; set; } = null!;

    public string Decription { get; set; } = null!;

    public string SapMaterialCode { get; set; } = null!;

    public DateTime Createdatetime { get; set; }

    public DateTime Updatedatetime { get; set; }
}
