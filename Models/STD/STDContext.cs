using System;
using System.Collections.Generic;
using Microsoft.EntityFrameworkCore;

namespace TemplateSTD.Models.STD;

public partial class STDContext : DbContext
{
    public STDContext()
    {
    }

    public STDContext(DbContextOptions<STDContext> options)
        : base(options)
    {
    }

    public virtual DbSet<CompareCostBom> CompareCostBoms { get; set; }

    public virtual DbSet<CompareOhcost> CompareOhcosts { get; set; }

    public virtual DbSet<CompareTotalStdcost> CompareTotalStdcosts { get; set; }

    public virtual DbSet<ConversionUnit> ConversionUnits { get; set; }

    public virtual DbSet<CostBomAs400> CostBomAs400s { get; set; }

    public virtual DbSet<CostBomSap> CostBomSaps { get; set; }

    public virtual DbSet<LogUpload> LogUploads { get; set; }

    public virtual DbSet<MappingSapMaterial> MappingSapMaterials { get; set; }

    public virtual DbSet<MasterCostCenter> MasterCostCenters { get; set; }

    public virtual DbSet<MasterMaterialBasic> MasterMaterialBasics { get; set; }

    public virtual DbSet<MasterMaterialUnit> MasterMaterialUnits { get; set; }

    public virtual DbSet<ProcessOhcostAs400> ProcessOhcostAs400s { get; set; }

    public virtual DbSet<ProcessOhcostSap> ProcessOhcostSaps { get; set; }

    public virtual DbSet<TotalCostAs400> TotalCostAs400s { get; set; }

    public virtual DbSet<TotalCostSap> TotalCostSaps { get; set; }

    public virtual DbSet<VwShopMapCostCenter> VwShopMapCostCenters { get; set; }

    protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
#warning To protect potentially sensitive information in your connection string, you should move it out of source code. You can avoid scaffolding the connection string by using the Name= syntax to read it from configuration - see https://go.microsoft.com/fwlink/?linkid=2131148. For more guidance on storing connection strings, see https://go.microsoft.com/fwlink/?LinkId=723263.
        => optionsBuilder.UseSqlServer("Server=10.236.35.235;Database=DataMigration;Persist Security Info=False;User ID=admin_query;Password=adqu@123456789;MultipleActiveResultSets=False;Encrypt=False;TrustServerCertificate=True;Connection Timeout=30;");

    protected override void OnModelCreating(ModelBuilder modelBuilder)
    {
        modelBuilder.Entity<CompareCostBom>(entity =>
        {
            entity
                .HasNoKey()
                .ToView("CompareCostBom", "STD");

            entity.Property(e => e.AlternativeUnit)
                .HasMaxLength(10)
                .HasColumnName("Alternative_Unit");
            entity.Property(e => e.As400Component)
                .HasMaxLength(100)
                .HasColumnName("AS400_Component");
            entity.Property(e => e.As400Lv).HasColumnName("AS400_Lv");
            entity.Property(e => e.As400Model)
                .HasMaxLength(100)
                .HasColumnName("AS400_Model");
            entity.Property(e => e.As400NoCost)
                .HasMaxLength(100)
                .HasColumnName("AS400_No_Cost");
            entity.Property(e => e.As400ParentMat)
                .HasMaxLength(100)
                .HasColumnName("AS400_Parent_Mat");
            entity.Property(e => e.As400PhantomItem)
                .HasMaxLength(100)
                .HasColumnName("AS400_Phantom_Item");
            entity.Property(e => e.As400Plant)
                .HasMaxLength(100)
                .HasColumnName("AS400_Plant");
            entity.Property(e => e.As400PriceQtyUnit)
                .HasMaxLength(100)
                .HasColumnName("AS400_Price_Qty_Unit");
            entity.Property(e => e.As400QuantityUnit)
                .HasMaxLength(10)
                .HasColumnName("AS400_Quantity_Unit");
            entity.Property(e => e.As400StdPrice)
                .HasMaxLength(100)
                .HasColumnName("AS400_STD_Price");
            entity.Property(e => e.As400SumTotalValue)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("AS400_SumTotalValue");
            entity.Property(e => e.As400SumValue)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("AS400_SumValue");
            entity.Property(e => e.As400TotalQuantity)
                .HasColumnType("decimal(18, 5)")
                .HasColumnName("AS400_Total_Quantity");
            entity.Property(e => e.As400TotalScrap)
                .HasColumnType("decimal(18, 5)")
                .HasColumnName("AS400_Total_Scrap");
            entity.Property(e => e.As400key)
                .HasMaxLength(410)
                .HasColumnName("AS400KEY");
            entity.Property(e => e.BaseUnit)
                .HasMaxLength(10)
                .HasColumnName("Base_Unit");
            entity.Property(e => e.Denominator).HasColumnType("numeric(11, 4)");
            entity.Property(e => e.DiffDeductedScrapInbaseunit)
                .HasColumnType("decimal(16, 5)")
                .HasColumnName("Diff_deducted_scrap_inbaseunit");
            entity.Property(e => e.DiffStdPrice)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("Diff_Std_Price");
            entity.Property(e => e.DiffSumTotalValueInsapbaseunit)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("Diff_Sum_Total_value_insapbaseunit");
            entity.Property(e => e.DiffSumValueDeductedScrapInbaseunit)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("Diff_Sum_value_deducted_scrap_inbaseunit");
            entity.Property(e => e.DiffTotalQuantityInbaseunit)
                .HasColumnType("decimal(16, 5)")
                .HasColumnName("Diff_Total_Quantity_inbaseunit");
            entity.Property(e => e.Numerator).HasColumnType("numeric(11, 4)");
            entity.Property(e => e.PercentDiffDeductedScrapInbaseunit)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("percent_Diff_deducted_scrap_inbaseunit");
            entity.Property(e => e.PercentDiffStdPrice)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("percent_Diff_Std_Price");
            entity.Property(e => e.PercentDiffSumTotalValueInsapbaseunit)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("percent_Diff_Sum_Total_value_insapbaseunit");
            entity.Property(e => e.PercentDiffSumValueDeductedScrapInbaseunit)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("percent_Diff_Sum_value_deducted_scrap_inbaseunit");
            entity.Property(e => e.PercentDiffTotalQuantityInbaseunit)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("percent_Diff_Total_Quantity_inbaseunit");
            entity.Property(e => e.SapComponent)
                .HasMaxLength(100)
                .HasColumnName("SAP_Component");
            entity.Property(e => e.SapLv).HasColumnName("SAP_Lv");
            entity.Property(e => e.SapModel)
                .HasMaxLength(100)
                .HasColumnName("SAP_Model");
            entity.Property(e => e.SapNoCost)
                .HasMaxLength(100)
                .HasColumnName("SAP_No_Cost");
            entity.Property(e => e.SapParentMat)
                .HasMaxLength(100)
                .HasColumnName("SAP_Parent_Mat");
            entity.Property(e => e.SapPhantomItem)
                .HasMaxLength(100)
                .HasColumnName("SAP_Phantom_Item");
            entity.Property(e => e.SapPlant)
                .HasMaxLength(100)
                .HasColumnName("SAP_Plant");
            entity.Property(e => e.SapPriceQtyUnit)
                .HasMaxLength(100)
                .HasColumnName("SAP_Price_Qty_Unit");
            entity.Property(e => e.SapQuantityUnit)
                .HasMaxLength(10)
                .HasColumnName("SAP_Quantity_Unit");
            entity.Property(e => e.SapStdPrice)
                .HasMaxLength(100)
                .HasColumnName("SAP_STD_Price");
            entity.Property(e => e.SapSumTotalValue)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_SumTotalValue");
            entity.Property(e => e.SapSumValue)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_SumValue");
            entity.Property(e => e.SapTotalQuantity)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_Total_Quantity");
            entity.Property(e => e.SapTotalScrap)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_Total_Scrap");
            entity.Property(e => e.Sapkey)
                .HasMaxLength(410)
                .HasColumnName("SAPKEY");
            entity.Property(e => e.StatusCheckUnit)
                .HasMaxLength(9)
                .IsUnicode(false);
            entity.Property(e => e.StdpricePerConvertToSap)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("STDPrice_per_convert_to_SAP");
            entity.Property(e => e.SumDeductedScrapInsapbaseunit)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("Sum_deducted_scrap_insapbaseunit");
            entity.Property(e => e.SumTotalQuantityInsapbaseunit)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("Sum_total_quantity_insapbaseunit");
            entity.Property(e => e.SumTotalValueInsapbaseunit)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("Sum_Total_value_insapbaseunit");
            entity.Property(e => e.SumValueDeductedScrapInsapbaseunit)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("Sum_value_deducted_scrap_insapbaseunit");
        });

        modelBuilder.Entity<CompareOhcost>(entity =>
        {
            entity
                .HasNoKey()
                .ToView("CompareOHCost", "STD");

            entity.Property(e => e.As400CostCenter)
                .HasMaxLength(100)
                .HasColumnName("AS400_Cost_Center");
            entity.Property(e => e.As400CostRate)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("AS400_Cost_Rate");
            entity.Property(e => e.As400FiscalYear)
                .HasMaxLength(100)
                .HasColumnName("AS400_Fiscal_Year");
            entity.Property(e => e.As400Model)
                .HasMaxLength(100)
                .HasColumnName("AS400_Model");
            entity.Property(e => e.As400OhCostRate)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("AS400_OH_Cost_Rate");
            entity.Property(e => e.As400Plant)
                .HasMaxLength(100)
                .HasColumnName("AS400_Plant");
            entity.Property(e => e.As400PricePerUnit)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("AS400_Price_Per_Unit");
            entity.Property(e => e.As400PriceQtyUnit)
                .HasMaxLength(100)
                .HasColumnName("AS400_Price_Qty_Unit");
            entity.Property(e => e.As400TotalOh)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("AS400_Total_OH");
            entity.Property(e => e.As400TotalProcessCost)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("AS400_Total_Process_Cost");
            entity.Property(e => e.As400TotalValue)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("AS400_Total_Value");
            entity.Property(e => e.As400TsQuantity)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("AS400_TS_Quantity");
            entity.Property(e => e.As400UnitQuantity)
                .HasMaxLength(100)
                .HasColumnName("AS400_Unit_Quantity");
            entity.Property(e => e.DiffOhCostRate)
                .HasColumnType("decimal(17, 4)")
                .HasColumnName("Diff_OH_Cost_Rate");
            entity.Property(e => e.DiffProcessCostRate)
                .HasColumnType("decimal(17, 4)")
                .HasColumnName("Diff_ProcessCost_Rate");
            entity.Property(e => e.DiffTotalOh)
                .HasColumnType("decimal(17, 4)")
                .HasColumnName("Diff_Total_OH");
            entity.Property(e => e.DiffTotalProcessCost)
                .HasColumnType("decimal(17, 4)")
                .HasColumnName("Diff_Total_Process_Cost");
            entity.Property(e => e.DiffTotalValue)
                .HasColumnType("decimal(17, 4)")
                .HasColumnName("Diff_Total_Value");
            entity.Property(e => e.DiffTsQuantity)
                .HasColumnType("decimal(17, 4)")
                .HasColumnName("Diff_TS_Quantity");
            entity.Property(e => e.PercentDiffOhCostRate)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("Percent_Diff_OH_Cost_Rate");
            entity.Property(e => e.PercentDiffProcessCostRate)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("Percent_Diff_ProcessCost_Rate");
            entity.Property(e => e.PercentDiffTotalOh)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("Percent_Diff_Total_OH");
            entity.Property(e => e.PercentDiffTotalProcessCost)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("Percent_Diff_Total_Process_Cost");
            entity.Property(e => e.PercentDiffTotalValue)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("Percent_Diff_Total_Value");
            entity.Property(e => e.PercentDiffTsQuantity)
                .HasMaxLength(42)
                .IsUnicode(false)
                .HasColumnName("Percent_Diff_TS_Quantity");
            entity.Property(e => e.SapCostCenter)
                .HasMaxLength(100)
                .HasColumnName("SAP_Cost_Center");
            entity.Property(e => e.SapCostRate)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("SAP_Cost_Rate");
            entity.Property(e => e.SapFiscalYear)
                .HasMaxLength(100)
                .HasColumnName("SAP_Fiscal_Year");
            entity.Property(e => e.SapModel)
                .HasMaxLength(100)
                .HasColumnName("SAP_Model");
            entity.Property(e => e.SapOhCostRate)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("SAP_OH_Cost_Rate");
            entity.Property(e => e.SapPlant)
                .HasMaxLength(100)
                .HasColumnName("SAP_Plant");
            entity.Property(e => e.SapPricePerUnit)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("SAP_Price_Per_Unit");
            entity.Property(e => e.SapPriceQtyUnit)
                .HasMaxLength(100)
                .HasColumnName("SAP_Price_Qty_Unit");
            entity.Property(e => e.SapTotalOh)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("SAP_Total_OH");
            entity.Property(e => e.SapTotalProcessCost)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("SAP_Total_Process_Cost");
            entity.Property(e => e.SapTotalValue)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("SAP_Total_Value");
            entity.Property(e => e.SapTsQuantity)
                .HasColumnType("decimal(16, 4)")
                .HasColumnName("SAP_TS_Quantity");
            entity.Property(e => e.SapUnitQuantity)
                .HasMaxLength(100)
                .HasColumnName("SAP_Unit_Quantity");
        });

        modelBuilder.Entity<CompareTotalStdcost>(entity =>
        {
            entity
                .HasNoKey()
                .ToView("CompareTotalSTDCost", "STD");

            entity.Property(e => e.As400FiscalYear)
                .HasMaxLength(100)
                .HasColumnName("AS400_FiscalYear");
            entity.Property(e => e.As400MaterialCost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("AS400_MaterialCost");
            entity.Property(e => e.As400Model)
                .HasMaxLength(100)
                .HasColumnName("AS400_Model");
            entity.Property(e => e.As400Ohcost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("AS400_OHCost");
            entity.Property(e => e.As400Priceperunit)
                .HasMaxLength(100)
                .HasColumnName("AS400_Priceperunit");
            entity.Property(e => e.As400ProcessCost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("AS400_ProcessCost");
            entity.Property(e => e.As400SrvpackingCost)
                .HasMaxLength(100)
                .HasColumnName("AS400_SRVPackingCost");
            entity.Property(e => e.As400Srvpackingpercent)
                .HasMaxLength(100)
                .HasColumnName("AS400_SRVPackingpercent");
            entity.Property(e => e.As400TotalStdcost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("AS400_TotalSTDCost");
            entity.Property(e => e.As400TotalStdcostRound)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("AS400_TotalSTDCostRound");
            entity.Property(e => e.As400TotalTs)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("AS400_TotalTS");
            entity.Property(e => e.As400Tsunit)
                .HasMaxLength(100)
                .HasColumnName("AS400_TSUnit");
            entity.Property(e => e.As400Unit)
                .HasMaxLength(100)
                .HasColumnName("AS400_Unit");
            entity.Property(e => e.DiffMaterialCost)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("Diff_MaterialCost");
            entity.Property(e => e.DiffOhcost)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("Diff_OHCost");
            entity.Property(e => e.DiffProcessCost)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("Diff_ProcessCost");
            entity.Property(e => e.DiffSrvpackingCost)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("Diff_SRVPackingCost");
            entity.Property(e => e.DiffTotalStdcost)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("Diff_TotalSTDCost");
            entity.Property(e => e.DiffTotalStdcostRound)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("Diff_TotalSTDCostRound");
            entity.Property(e => e.DiffTotalTs)
                .HasColumnType("decimal(19, 4)")
                .HasColumnName("Diff_TotalTS");
            entity.Property(e => e.PercentDiffMaterialCost)
                .HasMaxLength(4000)
                .HasColumnName("Percent_Diff_MaterialCost");
            entity.Property(e => e.PercentDiffOhcost)
                .HasMaxLength(4000)
                .HasColumnName("Percent_Diff_OHCost");
            entity.Property(e => e.PercentDiffProcessCost)
                .HasMaxLength(4000)
                .HasColumnName("Percent_Diff_ProcessCost");
            entity.Property(e => e.PercentDiffSrvpackingCost)
                .HasMaxLength(4000)
                .HasColumnName("Percent_Diff_SRVPackingCost");
            entity.Property(e => e.PercentDiffTotalStdcost)
                .HasMaxLength(4000)
                .HasColumnName("Percent_Diff_TotalSTDCost");
            entity.Property(e => e.PercentDiffTotalStdcostRound)
                .HasMaxLength(4000)
                .HasColumnName("Percent_Diff_TotalSTDCostRound");
            entity.Property(e => e.PercentDiffTotalTs)
                .HasMaxLength(4000)
                .HasColumnName("Percent_Diff_TotalTS");
            entity.Property(e => e.SapFiscalYear)
                .HasMaxLength(100)
                .HasColumnName("SAP_FiscalYear");
            entity.Property(e => e.SapMaterialCost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_MaterialCost");
            entity.Property(e => e.SapModel)
                .HasMaxLength(100)
                .HasColumnName("SAP_Model");
            entity.Property(e => e.SapOhcost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_OHCost");
            entity.Property(e => e.SapPriceperunit)
                .HasMaxLength(100)
                .HasColumnName("SAP_Priceperunit");
            entity.Property(e => e.SapProcessCost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_ProcessCost");
            entity.Property(e => e.SapSrvpackingCost)
                .HasMaxLength(100)
                .HasColumnName("SAP_SRVPackingCost");
            entity.Property(e => e.SapSrvpackingpercent)
                .HasMaxLength(100)
                .HasColumnName("SAP_SRVPackingpercent");
            entity.Property(e => e.SapTotalStdcost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_TotalSTDCost");
            entity.Property(e => e.SapTotalStdcostRound)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("SAP_TotalSTDCostRound");
            entity.Property(e => e.SapTotalTs)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("SAP_TotalTS");
            entity.Property(e => e.SapTsunit)
                .HasMaxLength(100)
                .HasColumnName("SAP_TSUnit");
            entity.Property(e => e.SapUnit)
                .HasMaxLength(100)
                .HasColumnName("SAP_Unit");
        });

        modelBuilder.Entity<ConversionUnit>(entity =>
        {
            entity.HasKey(e => e.No).HasName("PK_CONVERSION_UNIT_No");

            entity.ToTable("CONVERSION_UNIT", "BOM");

            entity.Property(e => e.No).HasColumnName("NO");
            entity.Property(e => e.AlternativeUnitOfMeasure)
                .HasMaxLength(10)
                .HasColumnName("Alternative Unit of Measure");
            entity.Property(e => e.BaseUnitOfMeasure)
                .HasMaxLength(10)
                .HasColumnName("Base Unit of Measure");
            entity.Property(e => e.BulkMaterial)
                .HasMaxLength(10)
                .HasColumnName("Bulk material");
            entity.Property(e => e.Createdatetime)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime")
                .HasColumnName("CREATEDATETIME");
            entity.Property(e => e.IssueUnit)
                .HasMaxLength(10)
                .HasColumnName("Issue Unit");
            entity.Property(e => e.Material).HasMaxLength(15);
            entity.Property(e => e.MaterialDescription)
                .HasMaxLength(200)
                .HasColumnName("Material Description");
            entity.Property(e => e.MaterialType)
                .HasMaxLength(10)
                .HasColumnName("Material Type");
            entity.Property(e => e.OldMaterialNumber)
                .HasMaxLength(30)
                .HasColumnName("Old material number");
            entity.Property(e => e.OrderUnit)
                .HasMaxLength(10)
                .HasColumnName("Order Unit");
            entity.Property(e => e.Updatedatetime)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime")
                .HasColumnName("UPDATEDATETIME");
            entity.Property(e => e.XDenominator)
                .HasColumnType("decimal(11, 4)")
                .HasColumnName("X (Denominator)");
            entity.Property(e => e.YNumerator)
                .HasColumnType("decimal(11, 4)")
                .HasColumnName("Y (Numerator)");
        });

        modelBuilder.Entity<CostBomAs400>(entity =>
        {
            entity.HasKey(e => new { e.Model, e.Plant, e.ParentMat, e.Component, e.QuantityUnit });

            entity.ToTable("CostBomAS400", "STD");

            entity.Property(e => e.Model).HasMaxLength(100);
            entity.Property(e => e.Plant).HasMaxLength(100);
            entity.Property(e => e.ParentMat)
                .HasMaxLength(100)
                .HasColumnName("Parent_Mat");
            entity.Property(e => e.Component).HasMaxLength(100);
            entity.Property(e => e.QuantityUnit)
                .HasMaxLength(10)
                .HasColumnName("Quantity_Unit");
            entity.Property(e => e.BulkMat)
                .HasMaxLength(100)
                .HasColumnName("Bulk_Mat");
            entity.Property(e => e.CompEffectDtFrom)
                .HasMaxLength(100)
                .HasColumnName("Comp_Effect_dtFrom");
            entity.Property(e => e.CompEffectDtTo)
                .HasMaxLength(100)
                .HasColumnName("Comp_Effect_dtTo");
            entity.Property(e => e.CompProcTypeMm)
                .HasMaxLength(100)
                .HasColumnName("Comp_Proc_Type_MM");
            entity.Property(e => e.CompSpTypeBom)
                .HasMaxLength(100)
                .HasColumnName("Comp_SP_Type_BOM");
            entity.Property(e => e.CompSpTypeMm)
                .HasMaxLength(100)
                .HasColumnName("Comp_SP_Type_MM");
            entity.Property(e => e.ComponentDesc)
                .HasMaxLength(100)
                .HasColumnName("Component_Desc");
            entity.Property(e => e.CostRelevancyBom)
                .HasMaxLength(100)
                .HasColumnName("Cost_Relevancy_BOM");
            entity.Property(e => e.CostingRun)
                .HasMaxLength(100)
                .HasColumnName("Costing_Run");
            entity.Property(e => e.CostingRundt)
                .HasMaxLength(100)
                .HasColumnName("Costing_Rundt");
            entity.Property(e => e.DeletionIndicator)
                .HasMaxLength(100)
                .HasColumnName("Deletion_Indicator");
            entity.Property(e => e.Import).HasMaxLength(100);
            entity.Property(e => e.ItemClass)
                .HasMaxLength(100)
                .HasColumnName("Item_Class");
            entity.Property(e => e.ItemType)
                .HasMaxLength(100)
                .HasColumnName("Item_Type");
            entity.Property(e => e.Local).HasMaxLength(100);
            entity.Property(e => e.MSize)
                .HasMaxLength(100)
                .HasColumnName("M/Size");
            entity.Property(e => e.MatProvisionIndicator)
                .HasMaxLength(100)
                .HasColumnName("Mat_Provision_Indicator");
            entity.Property(e => e.NoCost)
                .HasMaxLength(100)
                .HasColumnName("No_Cost");
            entity.Property(e => e.NoPrice)
                .HasMaxLength(100)
                .HasColumnName("No_Price");
            entity.Property(e => e.ParentMatDesc)
                .HasMaxLength(100)
                .HasColumnName("Parent_Mat_Desc");
            entity.Property(e => e.ParentProcTypeMm)
                .HasMaxLength(100)
                .HasColumnName("Parent_Proc_Type_MM");
            entity.Property(e => e.ParentSpTypeMm)
                .HasMaxLength(100)
                .HasColumnName("Parent_SP_Type_MM");
            entity.Property(e => e.PhantomItem)
                .HasMaxLength(100)
                .HasColumnName("Phantom_Item");
            entity.Property(e => e.PriceQtyUnit)
                .HasMaxLength(100)
                .HasColumnName("Price_Qty_Unit");
            entity.Property(e => e.PurPriceCur)
                .HasMaxLength(100)
                .HasColumnName("Pur_Price_Cur");
            entity.Property(e => e.QuantityModel)
                .HasMaxLength(100)
                .HasColumnName("Quantity/Model");
            entity.Property(e => e.RunningNo).HasColumnName("Running_No");
            entity.Property(e => e.Scrap)
                .HasMaxLength(100)
                .HasColumnName("Scrap%");
            entity.Property(e => e.StdPrice)
                .HasMaxLength(100)
                .HasColumnName("STD_Price");
            entity.Property(e => e.SumTotalValue).HasColumnType("decimal(18, 4)");
            entity.Property(e => e.SumValue).HasColumnType("decimal(18, 4)");
            entity.Property(e => e.TaxExp)
                .HasMaxLength(100)
                .HasColumnName("Tax&Exp");
            entity.Property(e => e.TotalQuantity)
                .HasColumnType("decimal(18, 5)")
                .HasColumnName("Total_Quantity");
            entity.Property(e => e.TotalScrap)
                .HasColumnType("decimal(18, 5)")
                .HasColumnName("Total_Scrap");
            entity.Property(e => e.Unit)
                .HasMaxLength(100)
                .HasColumnName("/Unit");
        });

        modelBuilder.Entity<CostBomSap>(entity =>
        {
            entity.HasKey(e => new { e.Model, e.Plant, e.ParentMat, e.Component, e.QuantityUnit });

            entity.ToTable("CostBomSAP", "STD");

            entity.Property(e => e.Model).HasMaxLength(100);
            entity.Property(e => e.Plant).HasMaxLength(100);
            entity.Property(e => e.ParentMat)
                .HasMaxLength(100)
                .HasColumnName("Parent_Mat");
            entity.Property(e => e.Component).HasMaxLength(100);
            entity.Property(e => e.QuantityUnit)
                .HasMaxLength(10)
                .HasColumnName("Quantity_Unit");
            entity.Property(e => e.BulkMat)
                .HasMaxLength(100)
                .HasColumnName("Bulk_Mat");
            entity.Property(e => e.CompEffectDtFrom)
                .HasMaxLength(100)
                .HasColumnName("Comp_Effect_dtFrom");
            entity.Property(e => e.CompEffectDtTo)
                .HasMaxLength(100)
                .HasColumnName("Comp_Effect_dtTo");
            entity.Property(e => e.CompProcTypeMm)
                .HasMaxLength(100)
                .HasColumnName("Comp_Proc_Type_MM");
            entity.Property(e => e.CompSpTypeBom)
                .HasMaxLength(100)
                .HasColumnName("Comp_SP_Type_BOM");
            entity.Property(e => e.CompSpTypeMm)
                .HasMaxLength(100)
                .HasColumnName("Comp_SP_Type_MM");
            entity.Property(e => e.ComponentDesc)
                .HasMaxLength(100)
                .HasColumnName("Component_Desc");
            entity.Property(e => e.CostRelevancyBom)
                .HasMaxLength(100)
                .HasColumnName("Cost_Relevancy_BOM");
            entity.Property(e => e.CostingRun)
                .HasMaxLength(100)
                .HasColumnName("Costing_Run");
            entity.Property(e => e.CostingRundt)
                .HasMaxLength(100)
                .HasColumnName("Costing_Rundt");
            entity.Property(e => e.DeletionIndicator)
                .HasMaxLength(100)
                .HasColumnName("Deletion_Indicator");
            entity.Property(e => e.Import).HasMaxLength(100);
            entity.Property(e => e.ItemClass)
                .HasMaxLength(100)
                .HasColumnName("Item_Class");
            entity.Property(e => e.ItemType)
                .HasMaxLength(100)
                .HasColumnName("Item_Type");
            entity.Property(e => e.Local).HasMaxLength(100);
            entity.Property(e => e.MSize)
                .HasMaxLength(100)
                .HasColumnName("M/Size");
            entity.Property(e => e.MatProvisionIndicator)
                .HasMaxLength(100)
                .HasColumnName("Mat_Provision_Indicator");
            entity.Property(e => e.NoCost)
                .HasMaxLength(100)
                .HasColumnName("No_Cost");
            entity.Property(e => e.NoPrice)
                .HasMaxLength(100)
                .HasColumnName("No_Price");
            entity.Property(e => e.ParentMatDesc)
                .HasMaxLength(100)
                .HasColumnName("Parent_Mat_Desc");
            entity.Property(e => e.ParentProcTypeMm)
                .HasMaxLength(100)
                .HasColumnName("Parent_Proc_Type_MM");
            entity.Property(e => e.ParentSpTypeMm)
                .HasMaxLength(100)
                .HasColumnName("Parent_SP_Type_MM");
            entity.Property(e => e.PhantomItem)
                .HasMaxLength(100)
                .HasColumnName("Phantom_Item");
            entity.Property(e => e.PriceQtyUnit)
                .HasMaxLength(100)
                .HasColumnName("Price_Qty_Unit");
            entity.Property(e => e.PurPriceCur)
                .HasMaxLength(100)
                .HasColumnName("Pur_Price_Cur");
            entity.Property(e => e.QuantityModel)
                .HasMaxLength(100)
                .HasColumnName("Quantity/Model");
            entity.Property(e => e.RunningNo).HasColumnName("Running_No");
            entity.Property(e => e.Scrap)
                .HasMaxLength(100)
                .HasColumnName("Scrap%");
            entity.Property(e => e.StdPrice)
                .HasMaxLength(100)
                .HasColumnName("STD_Price");
            entity.Property(e => e.SumTotalValue).HasColumnType("decimal(18, 4)");
            entity.Property(e => e.SumValue).HasColumnType("decimal(18, 4)");
            entity.Property(e => e.TaxExp)
                .HasMaxLength(100)
                .HasColumnName("Tax&Exp");
            entity.Property(e => e.TotalQuantity)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("Total_Quantity");
            entity.Property(e => e.TotalScrap)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("Total_Scrap");
            entity.Property(e => e.Unit)
                .HasMaxLength(100)
                .HasColumnName("/Unit");
        });

        modelBuilder.Entity<LogUpload>(entity =>
        {
            entity.HasKey(e => e.No).HasName("PK__UploadLo__C3905BAFDE5DA48A");

            entity.ToTable("LogUpload", "STD");

            entity.Property(e => e.No).ValueGeneratedNever();
            entity.Property(e => e.Category)
                .HasMaxLength(100)
                .HasColumnName("category");
            entity.Property(e => e.DateCreated).HasColumnType("datetime");
            entity.Property(e => e.FileName).HasMaxLength(200);
        });

        modelBuilder.Entity<MappingSapMaterial>(entity =>
        {
            entity.HasKey(e => e.No).HasName("PK_MAPPING_SAP_MATERIAL_NO");

            entity.ToTable("MAPPING_SAP_MATERIAL", "BOM");

            entity.Property(e => e.No).HasColumnName("NO");
            entity.Property(e => e.As400ItemNumber)
                .HasMaxLength(50)
                .HasColumnName("AS400 Item Number");
            entity.Property(e => e.Createdatetime)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime")
                .HasColumnName("CREATEDATETIME");
            entity.Property(e => e.SapMaterialCode)
                .HasMaxLength(50)
                .HasColumnName("SAP Material Code");
            entity.Property(e => e.Updatedatetime)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime")
                .HasColumnName("UPDATEDATETIME");
        });

        modelBuilder.Entity<MasterCostCenter>(entity =>
        {
            entity.HasKey(e => e.No).HasName("PK_Master_CostCenter_No");

            entity.ToTable("Master_CostCenter", "TS");

            entity.Property(e => e.No).HasColumnName("NO");
            entity.Property(e => e.CostCenter).HasMaxLength(50);
            entity.Property(e => e.Createdatetime)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime")
                .HasColumnName("CREATEDATETIME");
            entity.Property(e => e.Department).HasMaxLength(200);
            entity.Property(e => e.Description).HasMaxLength(200);
            entity.Property(e => e.FileName).HasMaxLength(200);
            entity.Property(e => e.GeneralName).HasMaxLength(200);
            entity.Property(e => e.Name1).HasMaxLength(200);
            entity.Property(e => e.OldShopCode).HasMaxLength(50);
            entity.Property(e => e.Updatedatetime)
                .HasDefaultValueSql("(getdate())")
                .HasColumnType("datetime")
                .HasColumnName("UPDATEDATETIME");
        });

        modelBuilder.Entity<MasterMaterialBasic>(entity =>
        {
            entity.HasKey(e => new { e.As400Material, e.SapMaterial }).HasName("PK__Material__C3905BAFA242FE01");

            entity.ToTable("Master_MaterialBasic", "STD");

            entity.Property(e => e.As400Material)
                .HasMaxLength(100)
                .HasColumnName("AS400_Material");
            entity.Property(e => e.SapMaterial)
                .HasMaxLength(100)
                .HasColumnName("SAP_Material");
            entity.Property(e => e.BaseUnit)
                .HasMaxLength(10)
                .HasColumnName("Base_Unit");
            entity.Property(e => e.Description).HasMaxLength(100);
            entity.Property(e => e.MaterialGroup)
                .HasMaxLength(10)
                .HasColumnName("Material_Group");
            entity.Property(e => e.MaterialType)
                .HasMaxLength(10)
                .HasColumnName("Material_Type");
            entity.Property(e => e.NetWeight)
                .HasColumnType("decimal(18, 3)")
                .HasColumnName("Net_Weight");
            entity.Property(e => e.WeightUnit)
                .HasMaxLength(10)
                .HasColumnName("Weight_Unit");
        });

        modelBuilder.Entity<MasterMaterialUnit>(entity =>
        {
            entity.HasKey(e => e.Material).HasName("PK__Material__C3905BAF237F8C98");

            entity.ToTable("Master_MaterialUnit", "STD");

            entity.Property(e => e.Material).HasMaxLength(100);
            entity.Property(e => e.AlternativeUnit)
                .HasMaxLength(10)
                .HasColumnName("Alternative_Unit");
            entity.Property(e => e.BaseUnit)
                .HasMaxLength(10)
                .HasColumnName("Base_Unit");
            entity.Property(e => e.Denominator).HasColumnType("decimal(10, 0)");
            entity.Property(e => e.Description).HasMaxLength(100);
            entity.Property(e => e.Numerator).HasColumnType("decimal(10, 0)");
        });

        modelBuilder.Entity<ProcessOhcostAs400>(entity =>
        {
            entity.HasKey(e => new { e.Model, e.Plant, e.FiscalYear, e.CostCenter });

            entity.ToTable("ProcessOHCostAS400", "STD");

            entity.Property(e => e.Model).HasMaxLength(100);
            entity.Property(e => e.Plant).HasMaxLength(100);
            entity.Property(e => e.FiscalYear)
                .HasMaxLength(100)
                .HasColumnName("Fiscal_Year");
            entity.Property(e => e.CostCenter)
                .HasMaxLength(100)
                .HasColumnName("Cost_Center");
            entity.Property(e => e.OhCostRate)
                .HasMaxLength(100)
                .HasColumnName("OH_Cost_Rate");
            entity.Property(e => e.PricePerUnit)
                .HasMaxLength(100)
                .HasColumnName("Price_Per_Unit");
            entity.Property(e => e.PriceQtyUnit)
                .HasMaxLength(100)
                .HasColumnName("Price_Qty_Unit");
            entity.Property(e => e.ProcCostRate)
                .HasMaxLength(100)
                .HasColumnName("Proc_Cost_Rate");
            entity.Property(e => e.TotalOh)
                .HasMaxLength(100)
                .HasColumnName("Total_OH");
            entity.Property(e => e.TotalProcessCost)
                .HasMaxLength(100)
                .HasColumnName("Total_Process_Cost");
            entity.Property(e => e.TotalValue)
                .HasMaxLength(100)
                .HasColumnName("Total_Value");
            entity.Property(e => e.TsQuantity)
                .HasMaxLength(100)
                .HasColumnName("TS_Quantity");
            entity.Property(e => e.UnitQuantity)
                .HasMaxLength(100)
                .HasColumnName("Unit_Quantity");
        });

        modelBuilder.Entity<ProcessOhcostSap>(entity =>
        {
            entity.HasKey(e => new { e.Model, e.Plant, e.FiscalYear, e.CostCenter });

            entity.ToTable("ProcessOHCostSAP", "STD");

            entity.Property(e => e.Model).HasMaxLength(100);
            entity.Property(e => e.Plant).HasMaxLength(100);
            entity.Property(e => e.FiscalYear)
                .HasMaxLength(100)
                .HasColumnName("Fiscal_Year");
            entity.Property(e => e.CostCenter)
                .HasMaxLength(100)
                .HasColumnName("Cost_Center");
            entity.Property(e => e.OhCostRate)
                .HasMaxLength(100)
                .HasColumnName("OH_Cost_Rate");
            entity.Property(e => e.PricePerUnit)
                .HasMaxLength(100)
                .HasColumnName("Price_Per_Unit");
            entity.Property(e => e.PriceQtyUnit)
                .HasMaxLength(100)
                .HasColumnName("Price_Qty_Unit");
            entity.Property(e => e.ProcCostRate)
                .HasMaxLength(100)
                .HasColumnName("Proc_Cost_Rate");
            entity.Property(e => e.TotalOh)
                .HasMaxLength(100)
                .HasColumnName("Total_OH");
            entity.Property(e => e.TotalProcessCost)
                .HasMaxLength(100)
                .HasColumnName("Total_Process_Cost");
            entity.Property(e => e.TotalValue)
                .HasMaxLength(100)
                .HasColumnName("Total_Value");
            entity.Property(e => e.TsQuantity)
                .HasMaxLength(100)
                .HasColumnName("TS_Quantity");
            entity.Property(e => e.UnitQuantity)
                .HasMaxLength(100)
                .HasColumnName("Unit_Quantity");
        });

        modelBuilder.Entity<TotalCostAs400>(entity =>
        {
            entity.HasKey(e => new { e.FiscalYear, e.Model }).HasName("PK__TotalCos__8285231ECACC186F");

            entity.ToTable("TotalCostAS400", "STD");

            entity.Property(e => e.FiscalYear).HasMaxLength(100);
            entity.Property(e => e.Model).HasMaxLength(100);
            entity.Property(e => e.MaterialCost).HasColumnType("decimal(18, 4)");
            entity.Property(e => e.Ohcost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("OHCost");
            entity.Property(e => e.Priceperunit).HasMaxLength(100);
            entity.Property(e => e.ProcessCost).HasColumnType("decimal(18, 4)");
            entity.Property(e => e.SrvpackingCost)
                .HasMaxLength(100)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("SRVPackingCost");
            entity.Property(e => e.Srvpackingpercent)
                .HasMaxLength(100)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("SRVPackingpercent");
            entity.Property(e => e.TotalStdcost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("TotalSTDCost");
            entity.Property(e => e.TotalStdcostRound)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("TotalSTDCostRound");
            entity.Property(e => e.TotalTs)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("TotalTS");
            entity.Property(e => e.Tsunit)
                .HasMaxLength(100)
                .HasColumnName("TSUnit");
            entity.Property(e => e.Unit).HasMaxLength(100);
        });

        modelBuilder.Entity<TotalCostSap>(entity =>
        {
            entity.HasKey(e => new { e.FiscalYear, e.Model }).HasName("PK__TotalCos__8285231E6E8E9992");

            entity.ToTable("TotalCostSAP", "STD");

            entity.Property(e => e.FiscalYear).HasMaxLength(100);
            entity.Property(e => e.Model).HasMaxLength(100);
            entity.Property(e => e.MaterialCost).HasColumnType("decimal(18, 4)");
            entity.Property(e => e.Ohcost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("OHCost");
            entity.Property(e => e.Priceperunit).HasMaxLength(100);
            entity.Property(e => e.ProcessCost).HasColumnType("decimal(18, 4)");
            entity.Property(e => e.SrvpackingCost)
                .HasMaxLength(100)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("SRVPackingCost");
            entity.Property(e => e.Srvpackingpercent)
                .HasMaxLength(100)
                .HasDefaultValueSql("(getdate())")
                .HasColumnName("SRVPackingpercent");
            entity.Property(e => e.TotalStdcost)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("TotalSTDCost");
            entity.Property(e => e.TotalStdcostRound)
                .HasColumnType("decimal(18, 2)")
                .HasColumnName("TotalSTDCostRound");
            entity.Property(e => e.TotalTs)
                .HasColumnType("decimal(18, 4)")
                .HasColumnName("TotalTS");
            entity.Property(e => e.Tsunit)
                .HasMaxLength(100)
                .HasColumnName("TSUnit");
            entity.Property(e => e.Unit).HasMaxLength(100);
        });

        modelBuilder.Entity<VwShopMapCostCenter>(entity =>
        {
            entity
                .HasNoKey()
                .ToView("vw_ShopMapCostCenter", "TS");

            entity.Property(e => e.CostCenter).HasMaxLength(50);
            entity.Property(e => e.Department).HasMaxLength(200);
            entity.Property(e => e.Description).HasMaxLength(200);
            entity.Property(e => e.GeneralName).HasMaxLength(200);
            entity.Property(e => e.Name1).HasMaxLength(200);
            entity.Property(e => e.OldShopCode).HasMaxLength(50);
            entity.Property(e => e.ShopCode).HasMaxLength(50);
        });

        OnModelCreatingPartial(modelBuilder);
    }

    partial void OnModelCreatingPartial(ModelBuilder modelBuilder);
}
