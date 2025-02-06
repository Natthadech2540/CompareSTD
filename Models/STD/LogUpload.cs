using System;
using System.Collections.Generic;

namespace TemplateSTD.Models.STD;

public partial class LogUpload
{
    public int No { get; set; }

    public string? FileName { get; set; }

    public string? Category { get; set; }

    public DateOnly? OrderDate { get; set; }

    public int? Model { get; set; }

    public int? TotalRecord { get; set; }

    public DateTime? DateCreated { get; set; }
}
