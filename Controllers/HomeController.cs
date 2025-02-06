using System.Diagnostics;
using Microsoft.AspNetCore.Mvc;
using TemplateSTD.Models;
using System.IO;
using OfficeOpenXml;  // To handle Excel files (xlsx)
using CsvHelper;     // To handle CSV files
using System.Dynamic;    // Needed for ExpandoObject
using System.Globalization; // Needed for CultureInfo
using Microsoft.AspNetCore.Http;
using System.Collections.Generic;
using CsvHelper.Configuration;
using TemplateSTD.Models.STD; // For CSV configuration
using System.Text.Json;
using Microsoft.AspNetCore.Hosting;
using EFCore.BulkExtensions;
using Microsoft.EntityFrameworkCore;
using OfficeOpenXml.Style;
using System.Net.Mail;
using System.Net;

namespace TemplateSTD.Controllers;

public class HomeController : Controller
{
    private readonly STDContext _context;
    private readonly IWebHostEnvironment _env;
    private readonly ILogger<HomeController> _logger;

    public HomeController(ILogger<HomeController> logger ,STDContext context,IWebHostEnvironment env)
    {
        _logger  = logger;
        _context = context;
        _env = env;
    }

    public IActionResult Index()
    {
        return View();
    }

    public IActionResult Privacy()
    {
        return View();
    }

    public IActionResult UploadDB()
    {
        return View();
    }

    public IActionResult MasterMaterialBasic()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        // var MasterMaterialBasics = _context.MappingSapMaterials.OrderBy(a => a.No).ToList();
        // return View(MasterMaterialBasics);

        // ดึงข้อมูลจากฐานข้อมูล
        var MasterMaterialBasics = _context.MappingSapMaterials.OrderBy(a => a.No).ToList();

        // สร้าง MemoryStream สำหรับเก็บไฟล์ Excel
        using (var package = new ExcelPackage())
        {
            // สร้าง worksheet ใหม่
            var worksheet = package.Workbook.Worksheets.Add("MasterMaterialBasics");

            // เขียนชื่อคอลัมน์ (Header)
            worksheet.Cells[1, 1].Value = "No";  
            worksheet.Cells[1, 2].Value = "As400ItemNumber";  
            worksheet.Cells[1, 3].Value = "Decription";  
            worksheet.Cells[1, 4].Value = "SapMaterialCode";  
            worksheet.Cells[1, 5].Value = "Createdatetime";  

            // เขียนข้อมูลในแถวถัดไป
            int row = 2;  // เริ่มต้นที่แถวที่ 2
            foreach (var item in MasterMaterialBasics)
            {
                worksheet.Cells[row, 1].Value = item.No;  
                worksheet.Cells[row, 2].Value = item.As400ItemNumber;
                worksheet.Cells[row, 3].Value = item.Decription;
                worksheet.Cells[row, 4].Value = item.SapMaterialCode;
                // แปลงวันที่จาก serial date number เป็น DateTime
                if (DateTime.TryParse(item.Createdatetime.ToString(), out DateTime parsedDate))
                {
                    worksheet.Cells[row, 5].Value = parsedDate.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                }
                else
                {
                    worksheet.Cells[row, 5].Value = "Invalid Date";
                }

                row++;
            }

            // สร้างไฟล์ Excel ใน MemoryStream
            var fileBytes = package.GetAsByteArray();

            // ส่งไฟล์ไปให้ผู้ใช้ดาวน์โหลด
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MasterMaterialBasics.xlsx");
        }
    }

    public IActionResult MasterShopmapCostCenter()
{
    //     try
    //     {
    //         // SMTP Server Configuration
    //         SmtpClient smtpClient = new SmtpClient("10.236.36.206", 25)
    //         {
    //             EnableSsl = false, // ถ้า Server ต้องการ SSL ให้เปลี่ยนเป็น true
    //             DeliveryMethod = SmtpDeliveryMethod.Network,
    //             UseDefaultCredentials = false,
    //             Credentials = new NetworkCredential("natthadech.r@mcp.meap.com", "Ntdch@12345678903") // ใส่รหัสผ่านถ้าจำเป็น
    //         };

    //         // Email Body - HTML สวยๆ
    //         string emailBody = @"
    //             <div style='background: linear-gradient(135deg, #667eea, #764ba2); padding: 40px; text-align: center; font-family: Arial, sans-serif; color: #fff;'>
    //                 <div style='max-width: 600px; background: #ffffff; border-radius: 15px; padding: 30px; margin: auto; 
    //                             box-shadow: 0px 8px 20px rgba(0, 0, 0, 0.3); border: 5px solid #667eea; text-align: center;'>
                        
    //                     <h2 style='color: #2c3e50; font-size: 30px; font-weight : bold'>🚀 Email Notification</h2>
                        
    //                     <p style='font-size: 18px; color: #444;'><strong>✅ Test Sent Email Complete!</strong></p>
                        
    //                     <p style='font-size: 16px; color: #777;'>555+ Everything is working fine 🎉</p>
                        
    //                     <div style='margin: 25px 0;'>
    //                         <a href='#' style='background: #667eea; color: white; padding: 12px 24px; 
    //                                 border-radius: 8px; text-decoration: none; font-weight: bold; 
    //                                 display: inline-block; transition: 0.3s; border: 2px solid transparent;'
    //                                 onmouseover='this.style.background=""#5555dd""; this.style.borderColor=""#fff"";'
    //                                 onmouseout='this.style.background=""#667eea""; this.style.borderColor=""transparent"";'>
    //                                 📩 View Details
    //                         </a>
    //                     </div>

    //                     <hr style='border: 1px solid #ddd; margin: 20px 0;'>

    //                     <p style='font-size: 14px; color: #999;'>This is an automated message, please do not reply.</p>
    //                 </div>
    //             </div>";


    //         // Email Message Configuration
    //         MailMessage mailMessage = new MailMessage
    //         {
    //             From = new MailAddress("natthadech.r@mcp.meap.com", "champpion"),
    //             Subject = "🚀 Test Sent Email",
    //             Body = emailBody,
    //             IsBodyHtml = true // ตั้งค่าให้รองรับ HTML
    //         };

    //         // Recipients
    //         mailMessage.To.Add("natthadech.r@mcp.meap.com");
    //         mailMessage.CC.Add("natthadech.r@mcp.meap.com");

    //         // ส่ง Email
    //         smtpClient.Send(mailMessage);
            
    //         var ShopMapCostCenters  = _context.MasterCostCenters.ToList();
    //         return View(ShopMapCostCenters);
    //     }
    //     catch (Exception ex)
    //     {
    //         return Content("❌ Error: " + ex.Message);
    //     }
    // }

        // ดึงข้อมูลจากฐานข้อมูล
        var ShopMapCostCenters  = _context.MasterCostCenters.ToList();
        return View(ShopMapCostCenters);
    }

    public IActionResult MasterMaterialUnit()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        // var MasterMaterialUnits = _context.ConversionUnits.OrderBy(a => a.No).ToList();
        // return View(MasterMaterialUnits);

        // ดึงข้อมูลจากฐานข้อมูล
        var MasterMaterialUnits = _context.ConversionUnits.OrderBy(a => a.No).ToList();

        // สร้าง MemoryStream สำหรับเก็บไฟล์ Excel
        using (var package = new ExcelPackage())
        {
            // สร้าง worksheet ใหม่
            var worksheet = package.Workbook.Worksheets.Add("MasterMaterialBasics");

            // เขียนชื่อคอลัมน์ (Header)
            worksheet.Cells[1, 1].Value = "No";  
            worksheet.Cells[1, 2].Value = "Material";  
            worksheet.Cells[1, 3].Value = "MaterialDescription";  
            worksheet.Cells[1, 4].Value = "OldMaterialNumber";  
            worksheet.Cells[1, 5].Value = "MaterialType";
            worksheet.Cells[1, 6].Value = "BaseUnitOfMeasure";  
            worksheet.Cells[1, 7].Value = "AlternativeUnitOfMeasure";  
            worksheet.Cells[1, 8].Value = "YNumerator";  
            worksheet.Cells[1, 9].Value = "XDenominator";  
            worksheet.Cells[1, 10].Value = "BulkMaterial";
            worksheet.Cells[1, 11].Value = "IssueUnit";
            worksheet.Cells[1, 12].Value = "Createdatetime";    

            // เขียนข้อมูลในแถวถัดไป
            int row = 2;  // เริ่มต้นที่แถวที่ 2
            foreach (var item in MasterMaterialUnits)
            {
                worksheet.Cells[row, 1].Value = item.No;  
                worksheet.Cells[row, 2].Value = item.Material;
                worksheet.Cells[row, 3].Value = item.MaterialDescription;
                worksheet.Cells[row, 4].Value = item.OldMaterialNumber;
                worksheet.Cells[row, 5].Value = item.MaterialType;
                worksheet.Cells[row, 6].Value = item.BaseUnitOfMeasure;
                worksheet.Cells[row, 7].Value = item.AlternativeUnitOfMeasure;
                worksheet.Cells[row, 8].Value = item.YNumerator;
                worksheet.Cells[row, 9].Value = item.XDenominator;
                worksheet.Cells[row, 10].Value = item.BulkMaterial;
                worksheet.Cells[row, 11].Value = item.IssueUnit;
                // แปลงวันที่จาก serial date number เป็น DateTime
                if (DateTime.TryParse(item.Createdatetime.ToString(), out DateTime parsedDate))
                {
                    worksheet.Cells[row, 12].Value = parsedDate.ToString("dd/MM/yyyy HH:mm:ss", CultureInfo.InvariantCulture);
                }
                else
                {
                    worksheet.Cells[row, 12].Value = "Invalid Date";
                }

                row++;
            }

            // สร้างไฟล์ Excel ใน MemoryStream
            var fileBytes = package.GetAsByteArray();

            // ส่งไฟล์ไปให้ผู้ใช้ดาวน์โหลด
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "MasterMaterialBasics.xlsx");
        }
    }

    public IActionResult TotalCostAS400()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        var TotalCostAs400s = _context.TotalCostAs400s.OrderBy(a => a.Model).ToList();
        return View(TotalCostAs400s);
    }

    
    public IActionResult TotalCostSAP()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        var TotalCostSaps = _context.TotalCostSaps.OrderBy(a => a.Model).ToList();
        return View(TotalCostSaps);
    }

    // Action สำหรับแสดงข้อมูล
    public IActionResult CostBomAS400()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        // var costBomsAS400 = _context.CostBomAs400s.OrderBy(a => a.RunningNo).ToList();
        // return View(costBomsAS400);

        // ดึงข้อมูลจากฐานข้อมูล
        var costBomsAS400 = _context.CostBomAs400s.OrderBy(a => a.RunningNo).ToList();

        // สร้าง MemoryStream สำหรับเก็บไฟล์ Excel
        using (var package = new ExcelPackage())
        {
            // สร้าง worksheet ใหม่
            var worksheet = package.Workbook.Worksheets.Add("CompareCostBoms");

            // เขียนชื่อคอลัมน์ (Header)
            worksheet.Cells[1, 1].Value  = "AS400Model";  
            worksheet.Cells[1, 2].Value  = "AS400Plant";  
            worksheet.Cells[1, 3].Value  = "AS400CostingRun";  
            worksheet.Cells[1, 4].Value  = "AS400CostingRundt";  
            worksheet.Cells[1, 5].Value  = "AS400RunningNo";  
            worksheet.Cells[1, 6].Value  = "AS400Lv";  
            worksheet.Cells[1, 7].Value  = "AS400ParentMat";  
            worksheet.Cells[1, 8].Value  = "AS400ParentMatDesc";  
            worksheet.Cells[1, 9].Value  = "AS400ParentProcTypeMm";  
            worksheet.Cells[1, 10].Value = "AS400ParentSpTypeMm";  
            worksheet.Cells[1, 11].Value = "AS400Component";  
            worksheet.Cells[1, 12].Value = "AS400ComponentDesc";  
            worksheet.Cells[1, 13].Value = "AS400ItemType";  
            worksheet.Cells[1, 14].Value = "AS400CompProcTypeMm";  
            worksheet.Cells[1, 15].Value = "AS400CompSpTypeMm";  
            worksheet.Cells[1, 16].Value = "AS400CompSpTypeBom";  
            worksheet.Cells[1, 17].Value = "AS400BulkMat";  
            worksheet.Cells[1, 18].Value = "AS400CostRelevancyBom";  
            worksheet.Cells[1, 19].Value = "AS400PhantomItem";  
            worksheet.Cells[1, 20].Value = "AS400DeletionIndicator";  
            worksheet.Cells[1, 21].Value = "AS400MatProvisionIndicator";  
            worksheet.Cells[1, 22].Value = "AS400CompEffectDtFrom";  
            worksheet.Cells[1, 23].Value = "AS400CompEffectDtTo";  
            worksheet.Cells[1, 24].Value = "AS400Unit";  
            worksheet.Cells[1, 25].Value = "AS400MSize";  
            worksheet.Cells[1, 26].Value = "AS400QuantityModel";  
            worksheet.Cells[1, 27].Value = "AS400Scrap";  
            worksheet.Cells[1, 28].Value = "AS400TotalScrap";  
            worksheet.Cells[1, 29].Value = "AS400TotalQuantity";  
            worksheet.Cells[1, 30].Value = "AS400QuantityUnit";  
            worksheet.Cells[1, 31].Value = "AS400Import";  
            worksheet.Cells[1, 32].Value = "AS400TaxExp";  
            worksheet.Cells[1, 33].Value = "AS400Local";  
            worksheet.Cells[1, 34].Value = "AS400StdPrice";  
            worksheet.Cells[1, 35].Value = "AS400PriceQtyUnit";  
            worksheet.Cells[1, 36].Value = "AS400NoPrice";  
            worksheet.Cells[1, 37].Value = "AS400NoCost";  
            worksheet.Cells[1, 38].Value = "AS400SumValue";  
            worksheet.Cells[1, 39].Value = "AS400SumTotalValue";  
            worksheet.Cells[1, 40].Value = "AS400PurPriceCur";  
            worksheet.Cells[1, 41].Value = "AS400ItemClass";  
          
            // เขียนข้อมูลในแถวถัดไป
            int row = 2;  // เริ่มต้นที่แถวที่ 2
            foreach (var item in costBomsAS400)
            {
                worksheet.Cells[row, 1].Value  = item.Model;  
                worksheet.Cells[row, 2].Value  = item.Plant;
                worksheet.Cells[row, 3].Value  = item.CostingRun;
                worksheet.Cells[row, 4].Value  = item.CostingRundt;
                worksheet.Cells[row, 5].Value  = item.RunningNo;
                worksheet.Cells[row, 6].Value  = item.Lv;
                worksheet.Cells[row, 7].Value  = item.ParentMat;
                worksheet.Cells[row, 8].Value  = item.ParentMatDesc;
                worksheet.Cells[row, 9].Value  = item.ParentProcTypeMm;
                worksheet.Cells[row, 10].Value = item.ParentSpTypeMm;
                worksheet.Cells[row, 11].Value = item.Component;
                worksheet.Cells[row, 12].Value = item.ComponentDesc;
                worksheet.Cells[row, 13].Value = item.ItemType;
                worksheet.Cells[row, 14].Value = item.CompProcTypeMm; 
                worksheet.Cells[row, 15].Value = item.CompSpTypeMm;
                worksheet.Cells[row, 16].Value = item.CompSpTypeBom;
                worksheet.Cells[row, 17].Value = item.BulkMat;
                worksheet.Cells[row, 18].Value = item.CostRelevancyBom;
                worksheet.Cells[row, 19].Value = item.PhantomItem;
                worksheet.Cells[row, 20].Value = item.DeletionIndicator;
                worksheet.Cells[row, 21].Value = item.MatProvisionIndicator;
                worksheet.Cells[row, 22].Value = item.CompEffectDtFrom;
                worksheet.Cells[row, 23].Value = item.CompEffectDtTo;
                worksheet.Cells[row, 24].Value = item.Unit;
                worksheet.Cells[row, 25].Value = item.MSize;
                worksheet.Cells[row, 26].Value = item.QuantityModel;
                worksheet.Cells[row, 27].Value = item.Scrap;
                worksheet.Cells[row, 28].Value = item.TotalScrap;
                worksheet.Cells[row, 29].Value = item.TotalQuantity;
                worksheet.Cells[row, 30].Value = item.QuantityUnit;
                worksheet.Cells[row, 31].Value = item.Import;
                worksheet.Cells[row, 32].Value = item.TaxExp;
                worksheet.Cells[row, 33].Value = item.Local;
                worksheet.Cells[row, 34].Value = item.StdPrice;
                worksheet.Cells[row, 35].Value = item.PriceQtyUnit;
                worksheet.Cells[row, 36].Value = item.NoPrice;
                worksheet.Cells[row, 37].Value = item.NoCost;
                worksheet.Cells[row, 38].Value = item.SumValue;
                worksheet.Cells[row, 39].Value = item.SumTotalValue;
                worksheet.Cells[row, 40].Value = item.PurPriceCur;
                worksheet.Cells[row, 41].Value = item.ItemClass;
    
                row++;
            }

            // สร้างไฟล์ Excel ใน MemoryStream
            var fileBytes = package.GetAsByteArray();

            // ส่งไฟล์ไปให้ผู้ใช้ดาวน์โหลด
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CostBomAS400.xlsx");
        }
    }

    
    public IActionResult CostBomSAP()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        // var costBomsSAP = _context.CostBomSaps.OrderBy(a => a.RunningNo).ToList();
        // return View(costBomsSAP);

        // ดึงข้อมูลจากฐานข้อมูล
        var costBomsSAP = _context.CostBomSaps.OrderBy(a => a.RunningNo).ToList();

        // สร้าง MemoryStream สำหรับเก็บไฟล์ Excel
        using (var package = new ExcelPackage())
        {
            // สร้าง worksheet ใหม่
            var worksheet = package.Workbook.Worksheets.Add("CompareCostBoms");

            // เขียนชื่อคอลัมน์ (Header)
            worksheet.Cells[1, 1].Value  = "SAPModel";  
            worksheet.Cells[1, 2].Value  = "SAPPlant";  
            worksheet.Cells[1, 3].Value  = "SAPCostingRun";  
            worksheet.Cells[1, 4].Value  = "SAPCostingRundt";  
            worksheet.Cells[1, 5].Value  = "SAPRunningNo";  
            worksheet.Cells[1, 6].Value  = "SAPLv";  
            worksheet.Cells[1, 7].Value  = "SAPParentMat";  
            worksheet.Cells[1, 8].Value  = "SAPParentMatDesc";  
            worksheet.Cells[1, 9].Value  = "SAPParentProcTypeMm";  
            worksheet.Cells[1, 10].Value = "SAPParentSpTypeMm";  
            worksheet.Cells[1, 11].Value = "SAPComponent";  
            worksheet.Cells[1, 12].Value = "SAPComponentDesc";  
            worksheet.Cells[1, 13].Value = "SAPItemType";  
            worksheet.Cells[1, 14].Value = "SAPCompProcTypeMm";  
            worksheet.Cells[1, 15].Value = "SAPCompSpTypeMm";  
            worksheet.Cells[1, 16].Value = "SAPCompSpTypeBom";  
            worksheet.Cells[1, 17].Value = "SAPBulkMat";  
            worksheet.Cells[1, 18].Value = "SAPCostRelevancyBom";  
            worksheet.Cells[1, 19].Value = "SAPPhantomItem";  
            worksheet.Cells[1, 20].Value = "SAPDeletionIndicator";  
            worksheet.Cells[1, 21].Value = "SAPMatProvisionIndicator";  
            worksheet.Cells[1, 22].Value = "SAPCompEffectDtFrom";  
            worksheet.Cells[1, 23].Value = "SAPCompEffectDtTo";  
            worksheet.Cells[1, 24].Value = "SAPUnit";  
            worksheet.Cells[1, 25].Value = "SAPMSize";  
            worksheet.Cells[1, 26].Value = "SAPQuantityModel";  
            worksheet.Cells[1, 27].Value = "SAPScrap";  
            worksheet.Cells[1, 28].Value = "SAPTotalScrap";  
            worksheet.Cells[1, 29].Value = "SAPTotalQuantity";  
            worksheet.Cells[1, 30].Value = "SAPQuantityUnit";  
            worksheet.Cells[1, 31].Value = "SAPImport";  
            worksheet.Cells[1, 32].Value = "SAPTaxExp";  
            worksheet.Cells[1, 33].Value = "SAPLocal";  
            worksheet.Cells[1, 34].Value = "SAPStdPrice";  
            worksheet.Cells[1, 35].Value = "SAPPriceQtyUnit";  
            worksheet.Cells[1, 36].Value = "SAPNoPrice";  
            worksheet.Cells[1, 37].Value = "SAPNoCost";  
            worksheet.Cells[1, 38].Value = "SAPSumValue";  
            worksheet.Cells[1, 39].Value = "SAPSumTotalValue";  
            worksheet.Cells[1, 40].Value = "SAPPurPriceCur";  
            worksheet.Cells[1, 41].Value = "SAPItemClass";  
          
            // เขียนข้อมูลในแถวถัดไป
            int row = 2;  // เริ่มต้นที่แถวที่ 2
            foreach (var item in costBomsSAP)
            {
                worksheet.Cells[row, 1].Value  = item.Model;  
                worksheet.Cells[row, 2].Value  = item.Plant;
                worksheet.Cells[row, 3].Value  = item.CostingRun;
                worksheet.Cells[row, 4].Value  = item.CostingRundt;
                worksheet.Cells[row, 5].Value  = item.RunningNo;
                worksheet.Cells[row, 6].Value  = item.Lv;
                worksheet.Cells[row, 7].Value  = item.ParentMat;
                worksheet.Cells[row, 8].Value  = item.ParentMatDesc;
                worksheet.Cells[row, 9].Value  = item.ParentProcTypeMm;
                worksheet.Cells[row, 10].Value = item.ParentSpTypeMm;
                worksheet.Cells[row, 11].Value = item.Component;
                worksheet.Cells[row, 12].Value = item.ComponentDesc;
                worksheet.Cells[row, 13].Value = item.ItemType;
                worksheet.Cells[row, 14].Value = item.CompProcTypeMm; 
                worksheet.Cells[row, 15].Value = item.CompSpTypeMm;
                worksheet.Cells[row, 16].Value = item.CompSpTypeBom;
                worksheet.Cells[row, 17].Value = item.BulkMat;
                worksheet.Cells[row, 18].Value = item.CostRelevancyBom;
                worksheet.Cells[row, 19].Value = item.PhantomItem;
                worksheet.Cells[row, 20].Value = item.DeletionIndicator;
                worksheet.Cells[row, 21].Value = item.MatProvisionIndicator;
                worksheet.Cells[row, 22].Value = item.CompEffectDtFrom;
                worksheet.Cells[row, 23].Value = item.CompEffectDtTo;
                worksheet.Cells[row, 24].Value = item.Unit;
                worksheet.Cells[row, 25].Value = item.MSize;
                worksheet.Cells[row, 26].Value = item.QuantityModel;
                worksheet.Cells[row, 27].Value = item.Scrap;
                worksheet.Cells[row, 28].Value = item.TotalScrap;
                worksheet.Cells[row, 29].Value = item.TotalQuantity;
                worksheet.Cells[row, 30].Value = item.QuantityUnit;
                worksheet.Cells[row, 31].Value = item.Import;
                worksheet.Cells[row, 32].Value = item.TaxExp;
                worksheet.Cells[row, 33].Value = item.Local;
                worksheet.Cells[row, 34].Value = item.StdPrice;
                worksheet.Cells[row, 35].Value = item.PriceQtyUnit;
                worksheet.Cells[row, 36].Value = item.NoPrice;
                worksheet.Cells[row, 37].Value = item.NoCost;
                worksheet.Cells[row, 38].Value = item.SumValue;
                worksheet.Cells[row, 39].Value = item.SumTotalValue;
                worksheet.Cells[row, 40].Value = item.PurPriceCur;
                worksheet.Cells[row, 41].Value = item.ItemClass;
    
                row++;
            }

            // สร้างไฟล์ Excel ใน MemoryStream
            var fileBytes = package.GetAsByteArray();

            // ส่งไฟล์ไปให้ผู้ใช้ดาวน์โหลด
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "CostBomSAP.xlsx");
        }
    }

    public IActionResult TemplateTotalSTDCostBom()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        var CompareTotalStdcosts = _context.CompareTotalStdcosts
                              .OrderBy(a => a.As400Model)
                              .ToList();
        return View(CompareTotalStdcosts);
    }

    public IActionResult TemplateCompareCostBom()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        // var CompareCostBoms = _context.CompareCostBoms
        //                       .OrderBy(a => a.SapModel)
        //                       .ThenBy(a => a.SapLv)
        //                       .ToList();
        // return View(CompareCostBoms);

        // ดึงข้อมูลจากฐานข้อมูล
        var CompareCostBoms = _context.CompareCostBoms
                              .OrderBy(a => a.SapModel)
                              .ToList();

        // สร้าง MemoryStream สำหรับเก็บไฟล์ Excel
        using (var package = new ExcelPackage())
        {
            // สร้าง worksheet ใหม่
            var worksheet = package.Workbook.Worksheets.Add("CompareCostBoms");

            // เขียนชื่อคอลัมน์ (Header)
            worksheet.Cells[1, 1].Value  = "Sapkey";  
            worksheet.Cells[1, 2].Value  = "SapModel";  
            worksheet.Cells[1, 3].Value  = "SapPlant";  
            worksheet.Cells[1, 4].Value  = "SapParentMat";  
            worksheet.Cells[1, 5].Value  = "SapComponent";  
            worksheet.Cells[1, 6].Value  = "SapLv";  
            worksheet.Cells[1, 7].Value  = "SapQuantityUnit";  
            worksheet.Cells[1, 8].Value  = "SapNoCost";  
            worksheet.Cells[1, 9].Value  = "SapPhantomItem";  
            worksheet.Cells[1, 10].Value = "SapStdPrice";  
            worksheet.Cells[1, 11].Value = "SapTotalScrap";  
            worksheet.Cells[1, 12].Value = "SapPriceQtyUnit";  
            worksheet.Cells[1, 13].Value = "SapTotalQuantity";  
            worksheet.Cells[1, 14].Value = "SapSumValue";  
            worksheet.Cells[1, 15].Value = "SapSumTotalValue";
            // กำหนดสีพื้นหลังของหัวข้อคอลัมน์เป็นสีเขียวอ่อน
                using (var range = worksheet.Cells[1, 1, 1, 15]) // ช่วง A1 ถึง M1
                {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(144, 238, 144)); // สีเขียวอ่อน (Light Green)
                    range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                }  
            worksheet.Cells[1, 16].Value = "As400key";  
            worksheet.Cells[1, 17].Value = "As400Model";  
            worksheet.Cells[1, 18].Value = "As400Plant";  
            worksheet.Cells[1, 19].Value = "As400ParentMat";  
            worksheet.Cells[1, 20].Value = "As400Component";  
            worksheet.Cells[1, 21].Value = "As400Lv";  
            worksheet.Cells[1, 22].Value = "As400QuantityUnit";  
            worksheet.Cells[1, 23].Value = "As400NoCost";  
            worksheet.Cells[1, 24].Value = "As400PhantomItem";  
            worksheet.Cells[1, 25].Value = "As400StdPrice";  
            worksheet.Cells[1, 26].Value = "As400TotalScrap";  
            worksheet.Cells[1, 27].Value = "As400PriceQtyUnit";  
            worksheet.Cells[1, 28].Value = "As400TotalQuantity";  
            worksheet.Cells[1, 29].Value = "As400SumValue";  
            worksheet.Cells[1, 30].Value = "As400SumTotalValue";  
            worksheet.Cells[1, 31].Value = "AlternativeUnit";  
            worksheet.Cells[1, 32].Value = "Numerator";  
            worksheet.Cells[1, 33].Value = "Denominator";  
            worksheet.Cells[1, 34].Value = "BaseUnit";  
            worksheet.Cells[1, 35].Value = "StatusCheckUnit";  
            worksheet.Cells[1, 36].Value = "StdpricePerConvertToSap";  
            worksheet.Cells[1, 37].Value = "SumDeductedScrapInsapbaseunit";  
            worksheet.Cells[1, 38].Value = "SumTotalQuantityInsapbaseunit";  
            worksheet.Cells[1, 39].Value = "SumValueDeductedScrapInsapbaseunit";  
            worksheet.Cells[1, 40].Value = "SumTotalValueInsapbaseunit"; 
            using (var range = worksheet.Cells[1, 16, 1, 30]) // ช่วง A1 ถึง M1
            {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 123, 24)); // สีเขียวอ่อน (Light Green)
                    range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }
            using (var range = worksheet.Cells[1, 31, 1, 35]) // ช่วง A1 ถึง M1
            {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 255, 0)); // สีเขียวอ่อน (Light Green)
                    range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            } 
            using (var range = worksheet.Cells[1, 36, 1, 40]) // ช่วง A1 ถึง M1
            {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 123, 24)); // สีเขียวอ่อน (Light Green)
                    range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }  
            worksheet.Cells[1, 41].Value = "DiffStdPrice";  
            worksheet.Cells[1, 42].Value = "DiffDeductedScrapInbaseunit";  
            worksheet.Cells[1, 43].Value = "DiffTotalQuantityInbaseunit";  
            worksheet.Cells[1, 44].Value = "DiffSumValueDeductedScrapInbaseunit";  
            worksheet.Cells[1, 45].Value = "DiffSumTotalValueInsapbaseunit";
            using (var range = worksheet.Cells[1, 41, 1, 45]) // ช่วง A1 ถึง M1
            {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(20, 234, 214)); // สีเขียวอ่อน (Light Green)
                    range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }  
            worksheet.Cells[1, 46].Value = "PercentDiffStdPrice";  
            worksheet.Cells[1, 47].Value = "PercentDiffDeductedScrapInbaseunit";  
            worksheet.Cells[1, 48].Value = "PercentDiffTotalQuantityInbaseunit";  
            worksheet.Cells[1, 49].Value = "PercentDiffSumValueDeductedScrapInbaseunit";  
            worksheet.Cells[1, 50].Value = "PercentDiffSumTotalValueInsapbaseunit"; 
            using (var range = worksheet.Cells[1, 46, 1, 50]) // ช่วง A1 ถึง M1
            {
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(122, 164, 18)); // สีเขียวอ่อน (Light Green)
                    range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                    range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }  
            worksheet.Cells[1, 51].Value = "Reasondedected";
            using (var range = worksheet.Cells[1, 51]) 
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 123, 111)); // สีชมพู (Light Pink)
                range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }  
            worksheet.Cells[1, 52].Value = "Reasonincluded";
            using (var range = worksheet.Cells[1, 52]) 
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(185, 23, 111)); // สีชมพู (Light Pink)
                range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            } 

            using (var range = worksheet.Cells[1, 1, 1, 52]) // เลือกช่วงที่มีข้อมูลทั้งหมด
            {
                // ขีดเส้นขอบให้ทุกเซลล์
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;    // ขอบบน
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;   // ขอบซ้าย
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;  // ขอบขวา
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin; // ขอบล่าง
            }  

            // เขียนข้อมูลในแถวถัดไป
            int row = 2;  // เริ่มต้นที่แถวที่ 2
            foreach (var item in CompareCostBoms)
            {
                worksheet.Cells[row, 1].Value  = item.Sapkey;  
                worksheet.Cells[row, 2].Value  = item.SapModel;
                worksheet.Cells[row, 3].Value  = item.SapPlant;
                worksheet.Cells[row, 4].Value  = item.SapParentMat;
                worksheet.Cells[row, 5].Value  = item.SapComponent;
                worksheet.Cells[row, 6].Value  = item.SapLv;
                worksheet.Cells[row, 7].Value  = item.SapQuantityUnit;
                worksheet.Cells[row, 8].Value  = item.SapNoCost;
                worksheet.Cells[row, 9].Value  = item.SapPhantomItem;
                worksheet.Cells[row, 10].Value = item.SapStdPrice;
                worksheet.Cells[row, 11].Value = item.SapTotalScrap;
                worksheet.Cells[row, 12].Value = item.SapPriceQtyUnit;
                worksheet.Cells[row, 13].Value = item.SapTotalQuantity;
                worksheet.Cells[row, 14].Value = item.SapSumValue; 
                worksheet.Cells[row, 15].Value = item.SapSumTotalValue;
                worksheet.Cells[row, 16].Value = item.As400key;
                worksheet.Cells[row, 17].Value = item.As400Model;
                worksheet.Cells[row, 18].Value = item.As400Plant;
                worksheet.Cells[row, 19].Value = item.As400ParentMat;
                worksheet.Cells[row, 20].Value = item.As400Component;
                worksheet.Cells[row, 21].Value = item.As400Lv;
                worksheet.Cells[row, 22].Value = item.As400QuantityUnit;
                worksheet.Cells[row, 23].Value = item.As400NoCost;
                worksheet.Cells[row, 24].Value = item.As400PhantomItem;
                worksheet.Cells[row, 25].Value = item.As400StdPrice;
                worksheet.Cells[row, 26].Value = item.As400TotalScrap;
                worksheet.Cells[row, 27].Value = item.As400PriceQtyUnit;
                worksheet.Cells[row, 28].Value = item.As400TotalQuantity;
                worksheet.Cells[row, 29].Value = item.As400SumValue;
                worksheet.Cells[row, 30].Value = item.As400SumTotalValue;
                worksheet.Cells[row, 31].Value = item.AlternativeUnit;
                worksheet.Cells[row, 32].Value = item.Numerator;
                worksheet.Cells[row, 33].Value = item.Denominator;
                worksheet.Cells[row, 34].Value = item.BaseUnit;
                worksheet.Cells[row, 35].Value = item.StatusCheckUnit;
                worksheet.Cells[row, 36].Value = item.StdpricePerConvertToSap;
                worksheet.Cells[row, 37].Value = item.SumDeductedScrapInsapbaseunit;
                worksheet.Cells[row, 38].Value = item.SumTotalQuantityInsapbaseunit;
                worksheet.Cells[row, 39].Value = item.SumValueDeductedScrapInsapbaseunit;
                worksheet.Cells[row, 40].Value = item.SumTotalValueInsapbaseunit;
                worksheet.Cells[row, 41].Value = item.DiffStdPrice;
                worksheet.Cells[row, 42].Value = item.DiffDeductedScrapInbaseunit;
                worksheet.Cells[row, 43].Value = item.DiffTotalQuantityInbaseunit;
                worksheet.Cells[row, 44].Value = item.DiffSumValueDeductedScrapInbaseunit;
                worksheet.Cells[row, 45].Value = item.DiffSumTotalValueInsapbaseunit;
                worksheet.Cells[row, 46].Value = item.PercentDiffStdPrice;
                worksheet.Cells[row, 47].Value = item.PercentDiffDeductedScrapInbaseunit;
                worksheet.Cells[row, 48].Value = item.PercentDiffTotalQuantityInbaseunit;
                worksheet.Cells[row, 49].Value = item.PercentDiffSumValueDeductedScrapInbaseunit;
                worksheet.Cells[row, 50].Value = item.PercentDiffSumTotalValueInsapbaseunit;
                // ทำการเพิ่มขอบให้กับทุกเซลล์ที่มีข้อมูล
                using (var range = worksheet.Cells[row, 1, row, 52]) // ขอบเขตของข้อมูลที่เพิ่มในแถวนี้
                {
                    range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                //  Reasondedected
                    if (item.SapModel?.ToString()          == "" && 
                        item.SapPlant?.ToString()          == "" && 
                        item.SapParentMat?.ToString()      == "" && 
                        item.SapComponent?.ToString()      == "" &&
                        item.SapQuantityUnit?.ToString()   == "")
                    {
                        worksheet.Cells[row, 51].Value = "Z2 Key exists in AS400 but not in SAP" ;
                        worksheet.Cells[row, 52].Value = "Z2 Key exists in AS400 but not in SAP" ;
                    } 
                    else if (item.As400Model?.ToString()        == "" && 
                        item.As400Plant?.ToString()        == "" && 
                        item.As400ParentMat?.ToString()    == "" &&
                        item.As400Component?.ToString()    == "" &&
                        item.As400QuantityUnit?.ToString() == "")
                    {
                        worksheet.Cells[row, 51].Value = "Z1 Key exists in SAP but not in AS400" ;
                        worksheet.Cells[row, 52].Value = "Z1 Key exists in SAP but not in AS400" ;
                    }
                    
                    else if (item.DiffSumValueDeductedScrapInbaseunit?.ToString() == "0.0000")
                    {
                        worksheet.Cells[row, 51].Value = "No diff amount" ;
                    }
                    else if (item.DiffSumValueDeductedScrapInbaseunit?.ToString() != "0.0000" && 
                        item.DiffDeductedScrapInbaseunit?.ToString()      == "0.00000")
                    {
                        worksheet.Cells[row, 51].Value = "Diff amount" ;
                    }
                    else if (item.DiffDeductedScrapInbaseunit?.ToString() != "0.00000" &&
                    item.DiffSumValueDeductedScrapInbaseunit?.ToString()  == "0.0000")
                    {
                        worksheet.Cells[row, 51].Value = "Diff quantity" ;
                    }
                    else if (item.DiffSumValueDeductedScrapInbaseunit?.ToString()  != "0.0000" && 
                        item.DiffDeductedScrapInbaseunit?.ToString()       != "0.00000")
                    {
                        worksheet.Cells[row, 51].Value = "Diff amount & diff quantity" ;
                    }

                   // Reasonincluded
                    if (item.DiffTotalQuantityInbaseunit?.ToString() != "0.00000" &&
                        item.DiffSumTotalValueInsapbaseunit?.ToString()     == "0.0000" &&
                        item.Sapkey?.ToString()   != "" &&
                        item.As400key?.ToString() != "")
                    {
                        worksheet.Cells[row, 52].Value = "Diff quantity" ;
                    }
                   if (item.DiffSumTotalValueInsapbaseunit?.ToString() == "0.0000" &&
                        item.Sapkey?.ToString()   != "" &&
                        item.As400key?.ToString() != "")
                    {
                        worksheet.Cells[row, 52].Value = "No diff amount" ;
                    }
                    if (item.DiffSumTotalValueInsapbaseunit?.ToString() != "0.0000" && 
                        item.DiffTotalQuantityInbaseunit?.ToString()  == "0.00000" &&
                        item.Sapkey?.ToString()   != "" &&
                        item.As400key?.ToString() != "")
                    {
                        worksheet.Cells[row, 52].Value = "Diff amount" ;
                    }
                    if (item.DiffSumTotalValueInsapbaseunit?.ToString()     != "0.0000" && 
                        item.DiffTotalQuantityInbaseunit?.ToString()  != "0.00000" &&
                        item.Sapkey?.ToString()   != "" &&
                        item.As400key?.ToString() != "")
                    {
                        worksheet.Cells[row, 52].Value = "Diff amount & diff quantity" ;
                    }
                row++;
            }

            // สร้างไฟล์ Excel ใน MemoryStream
            var fileBytes = package.GetAsByteArray();

            // ส่งไฟล์ไปให้ผู้ใช้ดาวน์โหลด
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TemplateCompareCostBom.xlsx");
        }
    }

    public IActionResult TemplateOHCost()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        // var CompareOhcosts = _context.CompareOhcosts.Where(a => a.As400CostCenter == "XXXXXX").ToList();
        // return View(CompareOhcosts);
        
        // ดึงข้อมูลจากฐานข้อมูล
        var CompareOhcosts = _context.CompareOhcosts.OrderBy(a => a.SapModel).ToList();

        // สร้าง MemoryStream สำหรับเก็บไฟล์ Excel
        using (var package = new ExcelPackage())
        {
            // สร้าง worksheet ใหม่
            var worksheet = package.Workbook.Worksheets.Add("CompareOHCosts");

            // เขียนชื่อคอลัมน์ (Header)
            worksheet.Cells[1, 1].Value  = "SapModel";  
            worksheet.Cells[1, 2].Value  = "SapPlant";  
            worksheet.Cells[1, 3].Value  = "SapFiscalYear";  
            worksheet.Cells[1, 4].Value  = "SapCostCenter";  
            worksheet.Cells[1, 5].Value  = "SapTsQuantity";  
            worksheet.Cells[1, 6].Value  = "SapUnitQuantity";  
            worksheet.Cells[1, 7].Value  = "SapPricePerUnit";  
            worksheet.Cells[1, 8].Value  = "SapPriceQtyUnit";  
            worksheet.Cells[1, 9].Value  = "SapCostRate";  
            worksheet.Cells[1, 10].Value = "SapOhCostRate";  
            worksheet.Cells[1, 11].Value = "SapTotalProcessCost";  
            worksheet.Cells[1, 12].Value = "SapTotalOh";  
            worksheet.Cells[1, 13].Value = "SapTotalValue"; 
            // กำหนดสีพื้นหลังของหัวข้อคอลัมน์เป็นสีเขียวอ่อน
            using (var range = worksheet.Cells[1, 1, 1, 13]) // ช่วง A1 ถึง M1
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(144, 238, 144)); // สีเขียวอ่อน (Light Green)
                range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            worksheet.Cells[1, 14].Value = "As400Model";  
            worksheet.Cells[1, 15].Value = "As400Plant";  
            worksheet.Cells[1, 16].Value = "As400FiscalYear";  
            worksheet.Cells[1, 17].Value = "As400CostCenter";  
            worksheet.Cells[1, 18].Value = "As400TsQuantity";  
            worksheet.Cells[1, 19].Value = "As400UnitQuantity";  
            worksheet.Cells[1, 20].Value = "As400PricePerUnit";  
            worksheet.Cells[1, 21].Value = "As400PriceQtyUnit";  
            worksheet.Cells[1, 22].Value = "As400CostRate";  
            worksheet.Cells[1, 23].Value = "As400OhCostRate";  
            worksheet.Cells[1, 24].Value = "As400TotalProcessCost";  
            worksheet.Cells[1, 25].Value = "As400TotalOh";  
            worksheet.Cells[1, 26].Value = "As400TotalValue";  
             // กำหนดสีพื้นหลังของหัวข้อคอลัมน์เป็นสีเขียวอ่อน
            using (var range = worksheet.Cells[1, 14, 1, 26]) // ช่วง A1 ถึง M1
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 123, 24)); // สีเขียวอ่อน (Light Green)
                range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            worksheet.Cells[1, 27].Value = "DiffTsQuantity";  
            worksheet.Cells[1, 28].Value = "DiffProcessCostRate";  
            worksheet.Cells[1, 29].Value = "DiffOhCostRate";  
            worksheet.Cells[1, 30].Value = "DiffTotalProcessCost";  
            worksheet.Cells[1, 31].Value = "DiffTotalOh";  
            worksheet.Cells[1, 32].Value = "DiffTotalValue";

            using (var range = worksheet.Cells[1, 27, 1, 32]) // ช่วง A1 ถึง M1
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(20, 234, 214)); // สีเขียวอ่อน (Light Green)
                range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }  
            worksheet.Cells[1, 33].Value = "PercentDiffTsQuantity";  
            worksheet.Cells[1, 34].Value = "PercentDiffProcessCostRate";  
            worksheet.Cells[1, 35].Value = "PercentDiffOhCostRate";  
            worksheet.Cells[1, 36].Value = "PercentDiffTotalProcessCost";  
            worksheet.Cells[1, 37].Value = "PercentDiffTotalOh";  
            worksheet.Cells[1, 38].Value = "PercentDiffTotalValue";  
            using (var range = worksheet.Cells[1, 33, 1, 38]) // ช่วง A1 ถึง M1
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(122, 164, 18)); // สีเขียวอ่อน (Light Green)
                range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }  

            worksheet.Cells[1, 39].Value = "Reason";  
            using (var range = worksheet.Cells[1, 39]) // เซลล์ที่ 39 (คอลัมน์ที่ 39)
            {
                range.Style.Fill.PatternType = ExcelFillStyle.Solid;
                range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.FromArgb(255, 123, 111)); // สีชมพู (Light Pink)
                range.Style.Font.Bold = true; // ทำให้ตัวอักษรในหัวข้อหนา
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            }

            using (var range = worksheet.Cells[1, 1, 1, 39]) // เลือกช่วงที่มีข้อมูลทั้งหมด
            {
                // ขีดเส้นขอบให้ทุกเซลล์
                range.Style.Border.Top.Style = ExcelBorderStyle.Thin;    // ขอบบน
                range.Style.Border.Left.Style = ExcelBorderStyle.Thin;   // ขอบซ้าย
                range.Style.Border.Right.Style = ExcelBorderStyle.Thin;  // ขอบขวา
                range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin; // ขอบล่าง
            }

            // เขียนข้อมูลในแถวถัดไป
            int row = 2;  // เริ่มต้นที่แถวที่ 2
            foreach (var item in CompareOhcosts)
            {
                worksheet.Cells[row, 1].Value  = item.SapModel;  
                worksheet.Cells[row, 2].Value  = item.SapPlant;
                worksheet.Cells[row, 3].Value  = item.SapFiscalYear;
                worksheet.Cells[row, 4].Value  = item.SapCostCenter;
                worksheet.Cells[row, 5].Value  = item.SapTsQuantity;
                worksheet.Cells[row, 6].Value  = item.SapUnitQuantity;
                worksheet.Cells[row, 7].Value  = item.SapPricePerUnit;
                worksheet.Cells[row, 8].Value  = item.SapPriceQtyUnit;
                worksheet.Cells[row, 9].Value  = item.SapCostRate;
                worksheet.Cells[row, 10].Value = item.SapOhCostRate;
                worksheet.Cells[row, 11].Value = item.SapTotalProcessCost;
                worksheet.Cells[row, 12].Value = item.SapTotalOh;
                worksheet.Cells[row, 13].Value = item.SapTotalValue;
                worksheet.Cells[row, 14].Value = item.As400Model; 
                worksheet.Cells[row, 15].Value = item.As400Plant;
                worksheet.Cells[row, 16].Value = item.As400FiscalYear;
                worksheet.Cells[row, 17].Value = item.As400CostCenter;
                worksheet.Cells[row, 18].Value = item.As400TsQuantity;
                worksheet.Cells[row, 19].Value = item.As400UnitQuantity;
                worksheet.Cells[row, 20].Value = item.As400PricePerUnit;
                worksheet.Cells[row, 21].Value = item.As400PriceQtyUnit;
                worksheet.Cells[row, 22].Value = item.As400CostRate;
                worksheet.Cells[row, 23].Value = item.As400OhCostRate;
                worksheet.Cells[row, 24].Value = item.As400TotalProcessCost;
                worksheet.Cells[row, 25].Value = item.As400TotalOh;
                worksheet.Cells[row, 26].Value = item.As400TotalValue;
                worksheet.Cells[row, 27].Value = item.DiffTsQuantity;
                worksheet.Cells[row, 28].Value = item.DiffProcessCostRate;
                worksheet.Cells[row, 29].Value = item.DiffOhCostRate;
                worksheet.Cells[row, 30].Value = item.DiffTotalProcessCost;
                worksheet.Cells[row, 31].Value = item.DiffTotalOh;
                worksheet.Cells[row, 32].Value = item.DiffTotalValue;
                worksheet.Cells[row, 33].Value = item.PercentDiffTsQuantity;
                worksheet.Cells[row, 34].Value = item.PercentDiffProcessCostRate;
                worksheet.Cells[row, 35].Value = item.PercentDiffOhCostRate;
                worksheet.Cells[row, 36].Value = item.PercentDiffTotalProcessCost;
                worksheet.Cells[row, 37].Value = item.PercentDiffTotalOh;
                worksheet.Cells[row, 38].Value = item.PercentDiffTotalValue;
                // ทำการเพิ่มขอบให้กับทุกเซลล์ที่มีข้อมูล
                using (var range = worksheet.Cells[row, 1, row, 39]) // ขอบเขตของข้อมูลที่เพิ่มในแถวนี้
                {
                    range.Style.Border.Top.Style    = ExcelBorderStyle.Thin;
                    range.Style.Border.Left.Style   = ExcelBorderStyle.Thin;
                    range.Style.Border.Right.Style  = ExcelBorderStyle.Thin;
                    range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                }
                    if (item.SapModel?.ToString()      == "" && 
                        item.SapPlant?.ToString()      == "" && 
                        item.SapFiscalYear?.ToString() == "" &&
                        item.SapCostCenter?.ToString() == "")
                    {
                        worksheet.Cells[row, 39].Value = "Z2 Key exists in AS400 but not in SAP" ;
                    }
                    else if (item.As400Model?.ToString() == "" && 
                        item.As400Plant?.ToString()      == "" && 
                        item.As400FiscalYear?.ToString() == "" &&
                        item.As400CostCenter?.ToString() == "")
                    {
                        worksheet.Cells[row, 39].Value = "Z1 Key exists in SAP but not in AS400" ;
                    }
                    
                    else if (item.DiffTotalValue?.ToString() == "0.0000" && item.As400Model?.ToString()      != "" &&
                    item.SapModel?.ToString()   != "")
                    {
                        worksheet.Cells[row, 39].Value = "No value dif" ;
                    }
                    else if (item.DiffProcessCostRate?.ToString() != "0.0000" &&
                            item.DiffTsQuantity?.ToString()       == "0.0000" || 
                            item.DiffOhCostRate?.ToString()       != "0.0000" && 
                            item.DiffTsQuantity?.ToString()       == "0.0000" && 
                            item.As400Model?.ToString()           != "" &&
                            item.SapModel?.ToString()             != "")
                    {
                        worksheet.Cells[row, 39].Value = "Diff rate" ;
                    }
                    else if (item.DiffTsQuantity?.ToString()     != "0.0000" && 
                            item.DiffProcessCostRate?.ToString() == "0.0000" && 
                            item.DiffOhCostRate?.ToString()      == "0.0000" && 
                            item.As400Model?.ToString()          != "" &&
                            item.SapModel?.ToString()            != "")
                    {
                        worksheet.Cells[row, 39].Value = "Diff quantity" ;
                    }
                    else if (item.DiffTsQuantity?.ToString()        != "0.0000" && 
                            item.DiffProcessCostRate?.ToString()    != "0.0000" && 
                            item.DiffOhCostRate?.ToString()         != "0.0000" && 
                            item.As400Model?.ToString()             != "" &&
                            item.SapModel?.ToString()               != "")
                    {
                        worksheet.Cells[row, 39].Value = "Diff rate & diff quantity" ;
                    }
                
                row++;
            }

            // สร้างไฟล์ Excel ใน MemoryStream
            var fileBytes = package.GetAsByteArray();

            // ส่งไฟล์ไปให้ผู้ใช้ดาวน์โหลด
            return File(fileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "TemplatePROC&OHCost.xlsx");
        }
    }

    
    public IActionResult ProcessOhcostAs400()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        var ProcessOhcostAs400s = _context.ProcessOhcostAs400s.OrderBy(a => a.CostCenter).ToList();
        return View(ProcessOhcostAs400s);
    }

    public IActionResult ProcessOhcostSap()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        var ProcessOhcostSaps = _context.ProcessOhcostSaps.OrderBy(a => a.CostCenter).ToList();
        return View(ProcessOhcostSaps);
    }

    public IActionResult LogUpload()
    {
        // ดึงข้อมูลจากฐานข้อมูล
        var LogUpload = _context.LogUploads.OrderByDescending(a => a.No).ToList();
        return View(LogUpload);
    }

    // ฟังก์ชันสำหรับตรวจสอบจำนวนคอลัมน์ในไฟล์ Excel
    private int GetExcelColumnCount(string filePath)
    {
        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var columnCount = worksheet.Dimension.Columns;
            
            return columnCount;
        }
    }

[HttpPost]
public async Task<IActionResult> Upload(IFormFile file, string systemType, string category)
{
    try
    {
        // ตรวจสอบว่าไฟล์ถูกอัปโหลดหรือไม่
        if (file == null || file.Length == 0)
        {
            ViewData["Error"] = "Please select a valid file.";
            return View("UploadDB");
        }

        // ตรวจสอบว่าไฟล์เป็นประเภท Excel หรือ CSV
        if (!file.ContentType.Contains("spreadsheetml.sheet") && !file.FileName.EndsWith(".csv"))
        {
            ViewData["Error"] = "Only Excel or CSV files are supported.";
            return View("UploadDB");
        }

        // สร้างโฟลเดอร์อัปโหลดหากยังไม่มี
        //var uploadDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
        var uploadDirectory = Path.Combine(_env.WebRootPath, "uploads");
        if (!Directory.Exists(uploadDirectory))
        {
            Directory.CreateDirectory(uploadDirectory);
        }

        // ลบไฟล์ที่ไม่ใช่ของวันก่อนหน้าและวันนี้
        var today = DateTime.Today;
        var yesterday = today.AddDays(-1);

        var files = Directory.GetFiles(uploadDirectory);
        foreach (var ckfilePath in files)
        {
            var creationTime = System.IO.File.GetCreationTime(ckfilePath); 

            // ลบไฟล์ที่ไม่ได้อยู่ในช่วงวันที่กำหนด
            //if (creationTime.Date != today && creationTime.Date != yesterday)
            if (creationTime.Date != today)
            {
                System.IO.File.Delete(ckfilePath); 
            }
        }

        // บันทึกไฟล์ลงในโฟลเดอร์อัปโหลด
        var filePath = Path.Combine(uploadDirectory, Path.GetFileName(file.FileName));
        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        if (systemType == "AS400" && category == "totalstdCost")
        {
            // อ่านข้อมูลจากไฟล์ตามประเภทและหมวดหมู่
            var records = new List<TotalCostAs400>();

            // ถ้าไฟล์เป็น .xlsx ใช้ฟังก์ชันสำหรับอ่านไฟล์ Excel
            if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
            {
                var columnCount = GetExcelColumnCount(filePath);

                if (columnCount != 13)
                {
                    // ViewData["Error"] = "The Excel file must have exactly 13 columns.";
                    ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
                    return View("UploadDB");
                }
                records = ParseExcelFileTotalSTDCostAS400(filePath);
            }
            // ถ้าไฟล์เป็น .csv ใช้ฟังก์ชันสำหรับอ่านไฟล์ CSV
            else if (file.FileName.EndsWith(".csv"))
            {
                records = ParseCsvFileTotalSTDCostAS400(filePath);
            }

            // บันทึกข้อมูลลงในฐานข้อมูล
            using (var dbContext = new STDContext()) // ใช้ DbContext ที่ตั้งค่าของคุณ
            {
                // ดึงรายชื่อ Plant และ Model ที่ต้องการลบออก
                var FiscalYearAndModels = records
                    .Select(r => new { r.FiscalYear, r.Model })
                    .Distinct();

                // ลบข้อมูลทั้งหมดในฐานข้อมูลที่ตรงกับ Plant และ Model ใน records
                foreach (var pm in FiscalYearAndModels)
                {
                    var itemsToDelete = dbContext.TotalCostAs400s
                        .Where(r => r.FiscalYear == pm.FiscalYear && r.Model == pm.Model);

                    if (itemsToDelete.Any())
                    {
                        dbContext.TotalCostAs400s.RemoveRange(itemsToDelete);
                    }
                }

                // บันทึกการลบทั้งหมดในครั้งเดียว
                dbContext.SaveChanges();

                // เริ่มเพิ่มข้อมูลใหม่
                foreach (var record in records)
                {
                        // ถ้าไม่มีข้อมูลให้ Insert
                        dbContext.TotalCostAs400s.AddRange(record);
                }

                    // ดึงค่า No ล่าสุดจากฐานข้อมูล (ถ้ามี)
                        int lastNo = dbContext.LogUploads
                                            .OrderByDescending(l => l.No) // เรียงจากล่าสุด
                                            .Select(l => l.No)
                                            .FirstOrDefault();

                        // ถ้าไม่มีข้อมูลในตาราง LogUpload จะตั้งค่า lastNo เป็น 0
                        int nextNo = (lastNo == 0) ? 1 : lastNo + 1; // ถ้าไม่มีข้อมูลให้เริ่มต้นที่ 1, ถ้ามีให้เพิ่ม 1

                        // บันทึกข้อมูลใน LogUpload
                        int countModel = records.Select(r => r.Model).Distinct().Count();  
                        var logUpload = new LogUpload
                        {
                            No = nextNo, // กำหนด No ที่คำนวณได้
                            FileName = file.FileName,
                            Category = category,
                            OrderDate = DateOnly.FromDateTime(DateTime.Now), // กำหนดวันที่อัปโหลด
                            Model = countModel,
                            TotalRecord = records.Count,
                            DateCreated = DateTime.Now
                        };

                        dbContext.LogUploads.Add(logUpload);

                // บันทึกการเปลี่ยนแปลงทั้งหมด
                await dbContext.SaveChangesAsync();

                // ส่งข้อความสำเร็จกลับไปยัง View
                ViewData["Success"] = $"{records.Count} records successfully uploaded.";
            }

        }
        else if (systemType == "SAP" && category == "totalstdCost")
        {
            // อ่านข้อมูลจากไฟล์ตามประเภทและหมวดหมู่
            var records = new List<TotalCostSap>();

            // ถ้าไฟล์เป็น .xlsx ใช้ฟังก์ชันสำหรับอ่านไฟล์ Excel
            if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
            {
                var columnCount = GetExcelColumnCount(filePath);

                if (columnCount != 13)
                {
                    // ViewData["Error"] = "The Excel file must have exactly 13 columns.";
                    ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
                    return View("UploadDB");
                }
                records = ParseExcelFileTotalSTDCostSAP(filePath);
            }
            // ถ้าไฟล์เป็น .csv ใช้ฟังก์ชันสำหรับอ่านไฟล์ CSV
            else if (file.FileName.EndsWith(".csv"))
            {
                records = ParseCsvFileTotalSTDCostSAP(filePath);
            }

            // บันทึกข้อมูลลงในฐานข้อมูล
            using (var dbContext = new STDContext()) // ใช้ DbContext ที่ตั้งค่าของคุณ
            {
                // ดึงรายชื่อ Plant และ Model ที่ต้องการลบออก
                var FiscalYearAndModels = records
                    .Select(r => new { r.FiscalYear, r.Model })
                    .Distinct();

                // ลบข้อมูลทั้งหมดในฐานข้อมูลที่ตรงกับ Plant และ Model ใน records
                foreach (var pm in FiscalYearAndModels)
                {
                    var itemsToDelete = dbContext.TotalCostSaps
                        .Where(r => r.FiscalYear == pm.FiscalYear && r.Model == pm.Model);

                    if (itemsToDelete.Any())
                    {
                        dbContext.TotalCostSaps.RemoveRange(itemsToDelete);
                    }
                }

                // บันทึกการลบทั้งหมดในครั้งเดียว
                dbContext.SaveChanges();

                // เริ่มเพิ่มข้อมูลใหม่
                foreach (var record in records)
                {
                        // ถ้าไม่มีข้อมูลให้ Insert
                        dbContext.TotalCostSaps.AddRange(record);
                }

                    // ดึงค่า No ล่าสุดจากฐานข้อมูล (ถ้ามี)
                        int lastNo = dbContext.LogUploads
                                            .OrderByDescending(l => l.No) // เรียงจากล่าสุด
                                            .Select(l => l.No)
                                            .FirstOrDefault();

                        // ถ้าไม่มีข้อมูลในตาราง LogUpload จะตั้งค่า lastNo เป็น 0
                        int nextNo = (lastNo == 0) ? 1 : lastNo + 1; // ถ้าไม่มีข้อมูลให้เริ่มต้นที่ 1, ถ้ามีให้เพิ่ม 1

                        // บันทึกข้อมูลใน LogUpload
                        int countModel = records.Select(r => r.Model).Distinct().Count();  
                        var logUpload = new LogUpload
                        {
                            No = nextNo, // กำหนด No ที่คำนวณได้
                            FileName = file.FileName,
                            Category = category,
                            OrderDate = DateOnly.FromDateTime(DateTime.Now), // กำหนดวันที่อัปโหลด
                            Model = countModel,
                            TotalRecord = records.Count,
                            DateCreated = DateTime.Now
                        };

                        dbContext.LogUploads.Add(logUpload);

                // บันทึกการเปลี่ยนแปลงทั้งหมด
                await dbContext.SaveChangesAsync();

                // ส่งข้อความสำเร็จกลับไปยัง View
                ViewData["Success"] = $"{records.Count} records successfully uploaded.";
            }

        }
        
        if (systemType == "AS400" && category == "MaterialCost")
{
    var records = new List<CostBomAs400>();

    if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
    {
        var columnCount = GetExcelColumnCount(filePath);
        if (columnCount != 41)
        {
            ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
            return View("UploadDB");
        }

        records = ParseExcelFileMaterialCostAS400(filePath);
    }
    else if (file.FileName.EndsWith(".csv"))
    {
        records = ParseCsvFileMaterialCostAS400(filePath);
    }

    using (var dbContext = new STDContext())
    {

                     //ดึงรายชื่อ Plant และ Model ที่ต้องการลบออก
                    var plantsAndModels = records
                        .Select(r => new { r.Plant, r.Model })
                        .Distinct();

                    // ลบข้อมูลทั้งหมดในฐานข้อมูลที่ตรงกับ Plant และ Model ใน records
                    foreach (var pm in plantsAndModels)
                    {
                        var itemsToDelete = dbContext.CostBomAs400s
                            .Where(r => r.Plant == pm.Plant && r.Model == pm.Model);

                        if (itemsToDelete.Any())
                        {
                            dbContext.CostBomAs400s.RemoveRange(itemsToDelete);
                        }
                    }

                    // บันทึกการลบทั้งหมดในครั้งเดียว
                    dbContext.SaveChanges();

        // 🚀 โหลด MappingSapMaterials เพื่อลด Query
       var materialCodes = dbContext.MappingSapMaterials
        .Where(a => records.Select(r => r.Component).Contains(a.As400ItemNumber) ||
                    records.Select(r => r.ParentMat).Contains(a.As400ItemNumber))
        .GroupBy(a => a.As400ItemNumber)  // ✅ Group กัน Key ซ้ำ
        .ToDictionary(g => g.Key, g => g.First().SapMaterialCode); // ✅ ใช้ค่าแรกในกลุ่ม

        // ✅ แปลงค่า Component & ParentMat ก่อนนำไปใช้
        foreach (var record in records)
        {
           record.Component = (!string.IsNullOrEmpty(record.Component) && materialCodes.TryGetValue(record.Component, out var compCode))
            ? (!string.IsNullOrEmpty(compCode) && compCode != "#N/A" ? compCode : record.Component)  
            : record.Component;

            record.ParentMat = (!string.IsNullOrEmpty(record.ParentMat) && materialCodes.TryGetValue(record.ParentMat, out var parentCode))
                ? (!string.IsNullOrEmpty(parentCode) && parentCode != "#N/A" ? parentCode : record.ParentMat)  
                : record.ParentMat;
        }

        // ✅ ใช้ Dictionary ป้องกัน Key ซ้ำ
        var combinedRecords = new Dictionary<string, CostBomAs400>();

        foreach (var record in records)
        {
            var key = $"{record.Model}_{record.Plant}_{record.ParentMat}_{record.Component}_{record.QuantityUnit}";

            if (combinedRecords.ContainsKey(key))
            {
                var existingRecord = combinedRecords[key];
                existingRecord.TotalScrap += record.TotalScrap;
                existingRecord.TotalQuantity += record.TotalQuantity;
                existingRecord.SumValue += record.SumValue;
                existingRecord.SumTotalValue += record.SumTotalValue;
            }
            else
            {
                combinedRecords[key] = record;
            }
        }

        // 🚀 โหลดข้อมูลที่อาจมีอยู่ในฐานข้อมูล
        var existingRecordsInDb = dbContext.CostBomAs400s
            .Where(r => combinedRecords.Values.Select(x => x.Model).Contains(r.Model))
            .ToList();

        foreach (var combinedRecord in combinedRecords.Values)
        {
            var existingRecord = existingRecordsInDb
                .FirstOrDefault(r => r.Model == combinedRecord.Model &&
                                    r.Plant == combinedRecord.Plant &&
                                    r.ParentMat == combinedRecord.ParentMat &&
                                    r.Component == combinedRecord.Component &&
                                    r.QuantityUnit == combinedRecord.QuantityUnit);

            if (existingRecord != null)
            {
                existingRecord.TotalScrap += combinedRecord.TotalScrap;
                existingRecord.TotalQuantity += combinedRecord.TotalQuantity;
                existingRecord.SumValue += combinedRecord.SumValue;
                existingRecord.SumTotalValue += combinedRecord.SumTotalValue;
            }
            else
            {
                dbContext.CostBomAs400s.Add(combinedRecord);
            }
        }

        await dbContext.SaveChangesAsync(); // ✅ บันทึกการเปลี่ยนแปลงครั้งเดียว

        // ✅ เพิ่มข้อมูล LogUpload
        int lastNo = dbContext.LogUploads
            .OrderByDescending(l => l.No)
            .Select(l => l.No)
            .FirstOrDefault();

        int nextNo = (lastNo == 0) ? 1 : lastNo + 1;

        int countModel = records.Select(r => r.Model).Distinct().Count();
        var logUpload = new LogUpload
        {
            No = nextNo,
            FileName = file.FileName,
            Category = category,
            OrderDate = DateOnly.FromDateTime(DateTime.Now),
            Model = countModel,
            TotalRecord = records.Count,
            DateCreated = DateTime.Now
        };

        dbContext.LogUploads.Add(logUpload);
        await dbContext.SaveChangesAsync();

        ViewData["Success"] = $"{records.Count} records successfully uploaded.";
    }
}

        // if (systemType == "AS400" && category == "MaterialCost")
        // {
        //     // อ่านข้อมูลจากไฟล์ตามประเภทและหมวดหมู่
        //     var records = new List<CostBomAs400>();

        //     // ถ้าไฟล์เป็น .xlsx ใช้ฟังก์ชันสำหรับอ่านไฟล์ Excel
        //     if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
        //     {
        //          var columnCount = GetExcelColumnCount(filePath);

        //         if (columnCount != 41)
        //         {
        //             // ViewData["Error"] = "The Excel file must have exactly 41 columns.";
        //             ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
        //             return View("UploadDB");
        //         }

        //         records = ParseExcelFileMaterialCostAS400(filePath);
        //     }
        //     // ถ้าไฟล์เป็น .csv ใช้ฟังก์ชันสำหรับอ่านไฟล์ CSV
        //     else if (file.FileName.EndsWith(".csv"))
        //     {
        //         records = ParseCsvFileMaterialCostAS400(filePath);
        //     }

        //     // บันทึกข้อมูลลงในฐานข้อมูล
        //     using (var dbContext = new STDContext()) // ใช้ DbContext ที่ตั้งค่าของคุณ
        //     {
        //             // ดึงรายชื่อ Plant และ Model ที่ต้องการลบออก
        //             var plantsAndModels = records
        //                 .Select(r => new { r.Plant, r.Model })
        //                 .Distinct();

        //             // ลบข้อมูลทั้งหมดในฐานข้อมูลที่ตรงกับ Plant และ Model ใน records
        //             foreach (var pm in plantsAndModels)
        //             {
        //                 var itemsToDelete = dbContext.CostBomAs400s
        //                     .Where(r => r.Plant == pm.Plant && r.Model == pm.Model);

        //                 if (itemsToDelete.Any())
        //                 {
        //                     dbContext.CostBomAs400s.RemoveRange(itemsToDelete);
        //                 }
        //             }

        //             // บันทึกการลบทั้งหมดในครั้งเดียว
        //             dbContext.SaveChanges();

        //             // ใช้ Dictionary สำหรับจัดเก็บข้อมูลในหน่วยความจำ
        //             var combinedRecords = new Dictionary<string, CostBomAs400>();

        //             foreach (var record in records)
        //             {
        //                 // สร้างคีย์สำหรับระบุเอกลักษณ์ของข้อมูล
        //                 var key = $"{record.Model}_{record.Plant}_{record.ParentMat}_{record.Component}_{record.QuantityUnit}";

        //                 if (combinedRecords.ContainsKey(key))
        //                 {
        //                     // ถ้าพบข้อมูลซ้ำใน Dictionary ให้รวมค่าที่ต้องการ
        //                     var existingRecord = combinedRecords[key];
        //                     existingRecord.TotalScrap += record.TotalScrap;
        //                     existingRecord.TotalQuantity += record.TotalQuantity;
        //                     existingRecord.SumValue += record.SumValue;
        //                     existingRecord.SumTotalValue += record.SumTotalValue;
        //                 }
        //                 else
        //                 {
        //                     // ถ้าไม่พบ ให้เพิ่มข้อมูลใหม่ใน Dictionary
        //                     combinedRecords[key] = record;
        //                 }
        //             }

        //             // ดึงข้อมูลทั้งหมดจากฐานข้อมูลที่อาจเกี่ยวข้อง
        //             var existingRecordsInDb = dbContext.CostBomAs400s
        //                 .Where(r => records.Select(x => x.Model).Contains(r.Model))
        //                 .ToList();

        //             // อัปเดตและเพิ่มข้อมูลในฐานข้อมูล
        //             foreach (var combinedRecord in combinedRecords.Values)
        //             {

        //                 // ค้นหาในฐานข้อมูล
        //                 var existingRecord = existingRecordsInDb
        //                     .FirstOrDefault(r => r.Model == combinedRecord.Model &&
        //                                         r.Plant == combinedRecord.Plant &&
        //                                         r.ParentMat == combinedRecord.ParentMat &&
        //                                         r.Component == combinedRecord.Component &&
        //                                         r.QuantityUnit == combinedRecord.QuantityUnit);

        //                 if (existingRecord != null)
        //                 {
        //                     // อัปเดตค่าถ้าพบข้อมูลในฐานข้อมูล
        //                     existingRecord.TotalScrap += combinedRecord.TotalScrap;
        //                     existingRecord.TotalQuantity += combinedRecord.TotalQuantity;
        //                     existingRecord.SumValue += combinedRecord.SumValue;
        //                     existingRecord.SumTotalValue += combinedRecord.SumTotalValue;
        //                     // ตรวจสอบข้อมูลที่เกี่ยวข้องจาก MappingSapMaterials
        //                     var getcomponent = _context.MappingSapMaterials.FirstOrDefault(a => a.As400ItemNumber == combinedRecord.Component);
        //                     var getparent = _context.MappingSapMaterials.FirstOrDefault(a => a.As400ItemNumber == combinedRecord.ParentMat);

        //                     if (getcomponent != null)
        //                     {
        //                         combinedRecord.Component = getcomponent.SapMaterialCode;
        //                     }

        //                     if (getparent != null)
        //                     {
        //                         combinedRecord.ParentMat = getparent.SapMaterialCode;
        //                     }
        //                 }
        //                 else
        //                 {
        //                     // เพิ่มข้อมูลใหม่ถ้าไม่พบในฐานข้อมูล
        //                     dbContext.CostBomAs400s.Add(combinedRecord);
        //                 }
        //             }

        //             // บันทึกการเปลี่ยนแปลงทั้งหมดครั้งเดียว
        //             await dbContext.SaveChangesAsync();

        //         // ดึงค่า No ล่าสุดจาก LogUpload
        //         int lastNo = dbContext.LogUploads
        //             .OrderByDescending(l => l.No)
        //             .Select(l => l.No)
        //             .FirstOrDefault();

        //         int nextNo = (lastNo == 0) ? 1 : lastNo + 1;

        //         // บันทึกข้อมูลใน LogUpload
        //         int countModel = records.Select(r => r.Model).Distinct().Count();
        //         var logUpload = new LogUpload
        //         {
        //             No = nextNo,
        //             FileName = file.FileName,
        //             Category = category,
        //             OrderDate = DateOnly.FromDateTime(DateTime.Now),
        //             Model = countModel,
        //             TotalRecord = records.Count,
        //             DateCreated = DateTime.Now
        //         };

        //         dbContext.LogUploads.Add(logUpload);

        //         // บันทึกการเปลี่ยนแปลงทั้งหมด
        //         await dbContext.SaveChangesAsync();

        //         // ส่งข้อความสำเร็จกลับไปยัง View
        //         ViewData["Success"] = $"{records.Count} records successfully uploaded.";
        //     }
        // }


        else if (systemType == "SAP" && category == "MaterialCost")
        {
            // อ่านข้อมูลจากไฟล์ตามประเภทและหมวดหมู่
            var records = new List<CostBomSap>();

            // ถ้าไฟล์เป็น .xlsx ใช้ฟังก์ชันสำหรับอ่านไฟล์ Excel
            if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
            {
                 var columnCount = GetExcelColumnCount(filePath);

                if (columnCount != 41)
                {
                    // ViewData["Error"] = "The Excel file must have exactly 41 columns.";
                    ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
                    return View("UploadDB");
                }

                records = ParseExcelFileMaterialCostSAP(filePath);
            }
            // ถ้าไฟล์เป็น .csv ใช้ฟังก์ชันสำหรับอ่านไฟล์ CSV
            else if (file.FileName.EndsWith(".csv"))
            {
                records = ParseCsvFileMaterialCostSAP(filePath);
            }

            // บันทึกข้อมูลลงในฐานข้อมูล
            using (var dbContext = new STDContext()) // ใช้ DbContext ที่ตั้งค่าของคุณ
            {
                    // ดึงรายชื่อ Plant และ Model ที่ต้องการลบออก
                    var plantsAndModels = records
                        .Select(r => new { r.Plant, r.Model })
                        .Distinct();

                    // ลบข้อมูลทั้งหมดในฐานข้อมูลที่ตรงกับ Plant และ Model ใน records
                    foreach (var pm in plantsAndModels)
                    {
                        var itemsToDelete = dbContext.CostBomSaps
                            .Where(r => r.Plant == pm.Plant && r.Model == pm.Model);

                        if (itemsToDelete.Any())
                        {
                            dbContext.CostBomSaps.RemoveRange(itemsToDelete);
                        }
                    }

                    // บันทึกการลบทั้งหมดในครั้งเดียว
                    dbContext.SaveChanges();

                    // ใช้ Dictionary สำหรับจัดเก็บข้อมูลในหน่วยความจำ
                    var combinedRecords = new Dictionary<string, CostBomSap>();

                    foreach (var record in records)
                    {
                        // สร้างคีย์สำหรับระบุเอกลักษณ์ของข้อมูล
                        var key = $"{record.Model}_{record.Plant}_{record.ParentMat}_{record.Component}_{record.QuantityUnit}";

                        if (combinedRecords.ContainsKey(key))
                        {
                            // ถ้าพบข้อมูลซ้ำใน Dictionary ให้รวมค่าที่ต้องการ
                            var existingRecord = combinedRecords[key];
                            existingRecord.TotalScrap += record.TotalScrap;
                            existingRecord.TotalQuantity += record.TotalQuantity;
                            existingRecord.SumValue += record.SumValue;
                            existingRecord.SumTotalValue += record.SumTotalValue;
                        }
                        else
                        {
                            // ถ้าไม่พบ ให้เพิ่มข้อมูลใหม่ใน Dictionary
                            combinedRecords[key] = record;
                        }
                    }

                    // ดึงข้อมูลทั้งหมดจากฐานข้อมูลที่อาจเกี่ยวข้อง
                    var existingRecordsInDb = dbContext.CostBomSaps
                        .Where(r => records.Select(x => x.Model).Contains(r.Model))
                        .ToList();

                    // อัปเดตและเพิ่มข้อมูลในฐานข้อมูล
                    foreach (var combinedRecord in combinedRecords.Values)
                    {
                        var existingRecord = existingRecordsInDb
                            .FirstOrDefault(r => r.Model == combinedRecord.Model &&
                                                r.Plant == combinedRecord.Plant &&
                                                r.ParentMat == combinedRecord.ParentMat &&
                                                r.Component == combinedRecord.Component &&
                                                r.QuantityUnit == combinedRecord.QuantityUnit);

                        if (existingRecord != null)
                        {
                            // อัปเดตค่าถ้าพบข้อมูลในฐานข้อมูล
                            existingRecord.TotalScrap += combinedRecord.TotalScrap;
                            existingRecord.TotalQuantity += combinedRecord.TotalQuantity;
                            existingRecord.SumValue += combinedRecord.SumValue;
                            existingRecord.SumTotalValue += combinedRecord.SumTotalValue;
                        }
                        else
                        {
                            // เพิ่มข้อมูลใหม่ถ้าไม่พบในฐานข้อมูล
                            dbContext.CostBomSaps.Add(combinedRecord);
                        }
                    }

                    // บันทึกการเปลี่ยนแปลงทั้งหมดครั้งเดียว
                    await dbContext.SaveChangesAsync();

                // ดึงค่า No ล่าสุดจาก LogUpload
                int lastNo = dbContext.LogUploads
                    .OrderByDescending(l => l.No)
                    .Select(l => l.No)
                    .FirstOrDefault();

                int nextNo = (lastNo == 0) ? 1 : lastNo + 1;

                // บันทึกข้อมูลใน LogUpload
                int countModel = records.Select(r => r.Model).Distinct().Count();
                var logUpload = new LogUpload
                {
                    No = nextNo,
                    FileName = file.FileName,
                    Category = category,
                    OrderDate = DateOnly.FromDateTime(DateTime.Now),
                    Model = countModel,
                    TotalRecord = records.Count,
                    DateCreated = DateTime.Now
                };

                dbContext.LogUploads.Add(logUpload);

                // บันทึกการเปลี่ยนแปลงทั้งหมด
                await dbContext.SaveChangesAsync();

                // ส่งข้อความสำเร็จกลับไปยัง View
                ViewData["Success"] = $"{records.Count} records successfully uploaded.";
            }

        }
        else if (systemType == "AS400" && category == "ProcessOH")
        {   
            // อ่านข้อมูลจากไฟล์ตามประเภทและหมวดหมู่
            var records = new List<ProcessOhcostAs400>();

            // ถ้าไฟล์เป็น .xlsx ใช้ฟังก์ชันสำหรับอ่านไฟล์ Excel
            if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
            {
                 var columnCount = GetExcelColumnCount(filePath);

                if (columnCount != 13)
                {
                    // ViewData["Error"] = "The Excel file must have exactly 13 columns.";
                    ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
                    return View("UploadDB");
                }
                records = ParseExcelFileProcessOhcostAs400(filePath);
            }
            // ถ้าไฟล์เป็น .csv ใช้ฟังก์ชันสำหรับอ่านไฟล์ CSV
            else if (file.FileName.EndsWith(".csv"))
            {
                records = ParseCsvFileProcessOhcostAs400(filePath);
            }

            // บันทึกข้อมูลลงในฐานข้อมูล
            using (var dbContext = new STDContext()) // ใช้ DbContext ที่ตั้งค่าของคุณ
            {
                // ลบข้อมูลทั้งหมดในตาราง
                dbContext.ProcessOhcostAs400s.RemoveRange(dbContext.ProcessOhcostAs400s);

                // บันทึกการเปลี่ยนแปลงในฐานข้อมูล
                dbContext.SaveChanges();

                foreach (var record in records)
                {
                    var getmapdata = _context.MasterCostCenters
                    .Where(r => r.OldShopCode == record.CostCenter)
                    .FirstOrDefault();
                    if (getmapdata != null){
                        record.CostCenter = getmapdata?.CostCenter?.Trim() ?? "";
                    }
                     // ถ้าไม่มีข้อมูลให้ Insert
                     dbContext.ProcessOhcostAs400s.AddRange(record);
                }

                    // ดึงค่า No ล่าสุดจากฐานข้อมูล (ถ้ามี)
                        int lastNo = dbContext.LogUploads
                                            .OrderByDescending(l => l.No) // เรียงจากล่าสุด
                                            .Select(l => l.No)
                                            .FirstOrDefault();

                        // ถ้าไม่มีข้อมูลในตาราง LogUpload จะตั้งค่า lastNo เป็น 0
                        int nextNo = (lastNo == 0) ? 1 : lastNo + 1; // ถ้าไม่มีข้อมูลให้เริ่มต้นที่ 1, ถ้ามีให้เพิ่ม 1

                        // บันทึกข้อมูลใน LogUpload
                        int countModel = records.Select(r => r.Model).Distinct().Count();  
                        var logUpload = new LogUpload
                        {
                            No = nextNo, // กำหนด No ที่คำนวณได้
                            FileName = file.FileName,
                            Category = category,
                            OrderDate = DateOnly.FromDateTime(DateTime.Now), // กำหนดวันที่อัปโหลด
                            Model = countModel,
                            TotalRecord = records.Count,
                            DateCreated = DateTime.Now
                        };

                        dbContext.LogUploads.Add(logUpload);

                // บันทึกการเปลี่ยนแปลงทั้งหมด
                await dbContext.SaveChangesAsync();

                // ส่งข้อความสำเร็จกลับไปยัง View
                ViewData["Success"] = $"{records.Count} records successfully uploaded.";
            }
        }
        else if (systemType == "SAP" && category == "ProcessOH")
        {   
            // อ่านข้อมูลจากไฟล์ตามประเภทและหมวดหมู่
            var records = new List<ProcessOhcostSap>();

            // ถ้าไฟล์เป็น .xlsx ใช้ฟังก์ชันสำหรับอ่านไฟล์ Excel
            if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
            {
                var columnCount = GetExcelColumnCount(filePath);

                if (columnCount != 13)
                {
                    // ViewData["Error"] = "The Excel file must have exactly 13 columns.";
                    ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
                    return View("UploadDB");
                }
                records = ParseExcelFileProcessOhcostSap(filePath);
            }
            // ถ้าไฟล์เป็น .csv ใช้ฟังก์ชันสำหรับอ่านไฟล์ CSV
            else if (file.FileName.EndsWith(".csv"))
            {
                records = ParseCsvFileProcessOhcostSap(filePath);
            }

            // บันทึกข้อมูลลงในฐานข้อมูล
            using (var dbContext = new STDContext()) // ใช้ DbContext ที่ตั้งค่าของคุณ
            {
                // ลบข้อมูลทั้งหมดในตาราง
                dbContext.ProcessOhcostSaps.RemoveRange(dbContext.ProcessOhcostSaps);

                // บันทึกการเปลี่ยนแปลงในฐานข้อมูล
                dbContext.SaveChanges();
                foreach (var record in records)
                {
                        // ตรวจสอบว่าข้อมูลนี้มีอยู่ในฐานข้อมูลแล้วหรือไม่
                        var existingRecord = dbContext.ProcessOhcostSaps
                            .FirstOrDefault(r => r.Model == record.Model && 
                                                 r.Plant == record.Plant && 
                                                 r.FiscalYear == record.FiscalYear && 
                                                 r.CostCenter == record.CostCenter);

                        if (existingRecord != null)
                        {
                            existingRecord.TsQuantity     = record.TsQuantity;
                            existingRecord.PricePerUnit   = record.PricePerUnit;
                            existingRecord.ProcCostRate   = record.ProcCostRate;
                            existingRecord.OhCostRate     = record.OhCostRate;
                            // Update ฟิลด์
                            existingRecord.Model             = record.Model;
                            existingRecord.Plant             = record.Plant;
                            existingRecord.FiscalYear        = record.FiscalYear;
                            existingRecord.CostCenter        = record.CostCenter;
                            existingRecord.UnitQuantity      = record.UnitQuantity;
                            existingRecord.PriceQtyUnit      = record.PriceQtyUnit;

                            // ถ้ามีข้อมูลให้ Update
                            dbContext.ProcessOhcostSaps.Update(record);
                        }
                        else
                        {
                            // ถ้าไม่มีข้อมูลให้ Insert
                            dbContext.ProcessOhcostSaps.Add(record);
                        }
                }

                    // ดึงค่า No ล่าสุดจากฐานข้อมูล (ถ้ามี)
                        int lastNo = dbContext.LogUploads
                                            .OrderByDescending(l => l.No) // เรียงจากล่าสุด
                                            .Select(l => l.No)
                                            .FirstOrDefault();

                        // ถ้าไม่มีข้อมูลในตาราง LogUpload จะตั้งค่า lastNo เป็น 0
                        int nextNo = (lastNo == 0) ? 1 : lastNo + 1; // ถ้าไม่มีข้อมูลให้เริ่มต้นที่ 1, ถ้ามีให้เพิ่ม 1

                        // บันทึกข้อมูลใน LogUpload
                        int countModel = records.Select(r => r.Model).Distinct().Count();  
                        var logUpload = new LogUpload
                        {
                            No = nextNo, // กำหนด No ที่คำนวณได้
                            FileName = file.FileName,
                            Category = category,
                            OrderDate = DateOnly.FromDateTime(DateTime.Now), // กำหนดวันที่อัปโหลด
                            Model = countModel,
                            TotalRecord = records.Count,
                            DateCreated = DateTime.Now
                        };

                        dbContext.LogUploads.Add(logUpload);

                // บันทึกการเปลี่ยนแปลงทั้งหมด
                await dbContext.SaveChangesAsync();

                // ส่งข้อความสำเร็จกลับไปยัง View
                ViewData["Success"] = $"{records.Count} records successfully uploaded.";
            }
        }

        else
        {
            ViewData["Error"] = "Unsupported system type or category.";
            return View("UploadDB");
        }

        // กลับไปยัง View ที่มีการอัปโหลด
        return View("UploadDB");

    }
    catch (Exception ex)
    {
        // ส่งข้อความแจ้งเตือนข้อผิดพลาดกลับไปยัง View
        ViewData["Error"] = $"Error processing file: {ex.Message}";
        
        // กลับไปยัง View ที่มีการอัปโหลด
        return View("UploadDB");
    }
}

[HttpPost]
public async Task<IActionResult> UploadMasterData(IFormFile file, string masterType)
{
    try
    {
        // ตรวจสอบว่าไฟล์ถูกอัปโหลดหรือไม่
        if (file == null || file.Length == 0)
        {
            ViewData["Error"] = "Please select a valid file.";
            return View("UploadDB");
        }

        // ตรวจสอบว่าไฟล์เป็นประเภท Excel หรือ CSV
        if (!file.ContentType.Contains("spreadsheetml.sheet") && !file.FileName.EndsWith(".csv"))
        {
            ViewData["Error"] = "Only Excel or CSV files are supported.";
            return View("UploadDB");
        }

        // สร้างโฟลเดอร์อัปโหลดหากยังไม่มี
        //var uploadDirectory = Path.Combine(Directory.GetCurrentDirectory(), "wwwroot", "uploads");
        var uploadDirectory = Path.Combine(_env.WebRootPath, "uploads");
        if (!Directory.Exists(uploadDirectory))
        {
            Directory.CreateDirectory(uploadDirectory);
        }

        // ลบไฟล์ที่ไม่ใช่ของวันก่อนหน้าและวันนี้
        var today = DateTime.Today;
        var yesterday = today.AddDays(-1);

        var files = Directory.GetFiles(uploadDirectory);
        foreach (var ckfilePath in files)
        {
            var creationTime = System.IO.File.GetCreationTime(ckfilePath); 

            // ลบไฟล์ที่ไม่ได้อยู่ในช่วงวันที่กำหนด
            // if (creationTime.Date != today && creationTime.Date != yesterday)
            if (creationTime.Date != today)
            {
                System.IO.File.Delete(ckfilePath); 
            }
        }

        // บันทึกไฟล์ลงในโฟลเดอร์อัปโหลด
        var filePath = Path.Combine(uploadDirectory, Path.GetFileName(file.FileName));
        using (var stream = new FileStream(filePath, FileMode.Create))
        {
            await file.CopyToAsync(stream);
        }

        //Master Basic
        if (masterType == "MasterBasic")
        {
            // อ่านข้อมูลจากไฟล์ตามประเภทและหมวดหมู่
            var records = new List<MasterMaterialBasic>();

            // ถ้าไฟล์เป็น .xlsx ใช้ฟังก์ชันสำหรับอ่านไฟล์ Excel
            if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
            {
                var columnCount = GetExcelColumnCount(filePath);

                if (columnCount != 8 && columnCount != 3)
                {
                    // ViewData["Error"] = "The Excel file must have exactly 8 columns.";
                    ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
                    return View("UploadDB");
                }
                records = ParseExcelFileMasterMaterialBasic(filePath);
            }
            // ถ้าไฟล์เป็น .csv ใช้ฟังก์ชันสำหรับอ่านไฟล์ CSV
            else if (file.FileName.EndsWith(".csv"))
            {
                records = ParseCsvFileMasterMaterialBasic(filePath);
            }

            // บันทึกข้อมูลลงในฐานข้อมูล
            using (var dbContext = new STDContext()) // ใช้ DbContext ที่ตั้งค่าของคุณ
            {
                // ดึงข้อมูลทั้งหมดจาก MasterMaterialBasics
                var deletedatamaster = dbContext.MasterMaterialBasics.ToList();

                // ลบข้อมูลทั้งหมดในตาราง
                dbContext.MasterMaterialBasics.RemoveRange(deletedatamaster);

                // บันทึกการเปลี่ยนแปลง
                dbContext.SaveChanges();

                foreach (var record in records)
                {
                        // ตรวจสอบว่าข้อมูลนี้มีอยู่ในฐานข้อมูลแล้วหรือไม่
                        var existingRecord = dbContext.MasterMaterialBasics
                            .FirstOrDefault(r => r.As400Material == record.As400Material && 
                                                r.SapMaterial == record.SapMaterial);

                        if (existingRecord != null)
                        {
                            // Update ฟิลด์อื่น ๆ ถ้าจำเป็น
                            existingRecord.As400Material      = record.As400Material;
                            existingRecord.SapMaterial        = record.SapMaterial;
                            existingRecord.Description        = record.Description;
                            existingRecord.BaseUnit           = record.BaseUnit;
                            existingRecord.MaterialType       = record.MaterialType;
                            existingRecord.MaterialGroup      = record.MaterialGroup;
                            existingRecord.NetWeight          = record.NetWeight;
                            existingRecord.WeightUnit         = record.WeightUnit;
                        }
                        else
                        {
                            // ถ้าไม่มีข้อมูลให้ Insert
                            dbContext.MasterMaterialBasics.Add(record);
                        }
                            // บันทึกการเปลี่ยนแปลงในฐานข้อมูล
                            dbContext.SaveChanges();
                }

                    // ดึงค่า No ล่าสุดจากฐานข้อมูล (ถ้ามี)
                        int lastNo = dbContext.LogUploads
                                        .OrderByDescending(l => l.No) // เรียงจากล่าสุด
                                        .Select(l => l.No)
                                        .FirstOrDefault();

                        // ถ้าไม่มีข้อมูลในตาราง LogUpload จะตั้งค่า lastNo เป็น 0
                        int nextNo = (lastNo == 0) ? 1 : lastNo + 1; // ถ้าไม่มีข้อมูลให้เริ่มต้นที่ 1, ถ้ามีให้เพิ่ม 1

                        // บันทึกข้อมูลใน LogUpload
                        int countModel = records.Select(r => r.SapMaterial).Distinct().Count();  
                        var logUpload = new LogUpload
                        {
                            No = nextNo, // กำหนด No ที่คำนวณได้
                            FileName = file.FileName,
                            Category = masterType,
                            OrderDate = DateOnly.FromDateTime(DateTime.Now), // กำหนดวันที่อัปโหลด
                            Model = countModel,
                            TotalRecord = records.Count,
                            DateCreated = DateTime.Now
                        };

                        dbContext.LogUploads.Add(logUpload);

                // บันทึกการเปลี่ยนแปลงทั้งหมด
                await dbContext.SaveChangesAsync();

                // ส่งข้อความสำเร็จกลับไปยัง View
                ViewData["Success"] = $"{records.Count} records successfully uploaded.";
            }

        }
        //Master Unit
        else if (masterType == "MasterUnit")
        {
            // อ่านข้อมูลจากไฟล์ตามประเภทและหมวดหมู่
            var records = new List<MasterMaterialUnit>();

            // ถ้าไฟล์เป็น .xlsx ใช้ฟังก์ชันสำหรับอ่านไฟล์ Excel
            if (file.FileName.EndsWith(".xlsx") || file.FileName.EndsWith(".xlsb"))
            {
                var columnCount = GetExcelColumnCount(filePath);

                if (columnCount != 6)
                {
                    // ViewData["Error"] = "The Excel file must have exactly 6 columns.";
                    ViewData["Error"] = "The Excel file mistake incorrect. Please Check File Again!";
                    return View("UploadDB");
                }
                records = ParseExcelFileMasterMaterialUnit(filePath);
            }
            // ถ้าไฟล์เป็น .csv ใช้ฟังก์ชันสำหรับอ่านไฟล์ CSV
            else if (file.FileName.EndsWith(".csv"))
            {
                records = ParseCsvFileMasterMaterialUnit(filePath);
            }

            // บันทึกข้อมูลลงในฐานข้อมูล
            using (var dbContext = new STDContext()) // ใช้ DbContext ที่ตั้งค่าของคุณ
            {
                // ลบข้อมูลทั้งหมดในตาราง
                dbContext.MasterMaterialUnits.RemoveRange(dbContext.MasterMaterialUnits);

                // บันทึกการเปลี่ยนแปลงในฐานข้อมูล
                dbContext.SaveChanges();
                    foreach (var record in records)
                    {
                            // ตรวจสอบว่าข้อมูลนี้มีอยู่ในฐานข้อมูลแล้วหรือไม่
                            var existingRecord = dbContext.MasterMaterialUnits
                                .FirstOrDefault(r => r.Material == record.Material);

                            if (existingRecord != null)
                            {
                                existingRecord.Material           = record.Material;
                                existingRecord.Description        = record.Description;
                                existingRecord.AlternativeUnit    = record.AlternativeUnit;
                                existingRecord.Numerator          = record.Numerator;
                                existingRecord.Denominator        = record.Denominator;
                                existingRecord.BaseUnit           = record.BaseUnit;
                            }
                            else
                            {
                                // ถ้าไม่มีข้อมูลให้ Insert
                                dbContext.MasterMaterialUnits.Add(record);
                            }
                                // บันทึกการเปลี่ยนแปลงในฐานข้อมูล
                                dbContext.SaveChanges();
                    }

                    // ดึงค่า No ล่าสุดจากฐานข้อมูล (ถ้ามี)
                        int lastNo = dbContext.LogUploads
                                            .OrderByDescending(l => l.No) // เรียงจากล่าสุด
                                            .Select(l => l.No)
                                            .FirstOrDefault();

                        // ถ้าไม่มีข้อมูลในตาราง LogUpload จะตั้งค่า lastNo เป็น 0
                        int nextNo = (lastNo == 0) ? 1 : lastNo + 1; // ถ้าไม่มีข้อมูลให้เริ่มต้นที่ 1, ถ้ามีให้เพิ่ม 1

                        // บันทึกข้อมูลใน LogUpload
                        int countModel = records.Select(r => r.Material).Distinct().Count();  
                        var logUpload = new LogUpload
                        {
                            No = nextNo, // กำหนด No ที่คำนวณได้
                            FileName = file.FileName,
                            Category = masterType,
                            OrderDate = DateOnly.FromDateTime(DateTime.Now), // กำหนดวันที่อัปโหลด
                            Model = countModel,
                            TotalRecord = records.Count,
                            DateCreated = DateTime.Now
                        };

                        dbContext.LogUploads.Add(logUpload);

                // บันทึกการเปลี่ยนแปลงทั้งหมด
                await dbContext.SaveChangesAsync();

                // ส่งข้อความสำเร็จกลับไปยัง View
                ViewData["Success"] = $"{records.Count} records successfully uploaded.";
            }

        }
        else
        {
            ViewData["Error"] = "Unsupported system type or category.";
            return View("UploadDB");
        }

        // กลับไปยัง View ที่มีการอัปโหลด
        return View("UploadDB");

    }
    catch (Exception ex)
    {
        // ส่งข้อความแจ้งเตือนข้อผิดพลาดกลับไปยัง View
        ViewData["Error"] = $"Error processing file: {ex.Message}";
        
        // กลับไปยัง View ที่มีการอัปโหลด
        return View("UploadDB");
    }
}

    private List<ProcessOhcostAs400> ParseExcelFileProcessOhcostAs400(string filePath)
    {
        var records = new List<ProcessOhcostAs400>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null) throw new Exception("No worksheet found.");

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Skip header row
            {
                var record = new ProcessOhcostAs400
                {
                    Model               = worksheet.Cells[row, 1].Text.Trim(),
                    Plant               = worksheet.Cells[row, 2].Text.Trim(),
                    FiscalYear          = worksheet.Cells[row, 3].Text.Trim(),
                    CostCenter          = worksheet.Cells[row, 4].Text.Trim(),
                    TsQuantity          = worksheet.Cells[row, 5].Text.Trim(),
                    UnitQuantity        = worksheet.Cells[row, 6].Text.Trim(),
                    PricePerUnit        = worksheet.Cells[row, 7].Text.Trim(),
                    PriceQtyUnit        = worksheet.Cells[row, 8].Text.Trim(),
                    ProcCostRate        = worksheet.Cells[row, 9].Text.Trim(),
                    OhCostRate          = worksheet.Cells[row, 10].Text.Trim(),
                    TotalProcessCost    = worksheet.Cells[row, 11].Text.Trim(),
                    TotalOh             = worksheet.Cells[row, 12].Text.Trim(),
                    TotalValue          = worksheet.Cells[row, 13].Text.Trim(),
                };

                records.Add(record);
            }
        }

        return records;
    }

    public List<ProcessOhcostAs400> ParseCsvFileProcessOhcostAs400(string filePath)
    {
        var records = new List<ProcessOhcostAs400>();

        try
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",  // Use comma delimiter
                BadDataFound = null  // Ignore bad data rows
            }))
            {
                 if (!csv.Read() || !csv.ReadHeader())
                throw new Exception("CSV file does not contain a valid header.");

                // Skip the header row
                csv.Read();
                csv.ReadHeader();

                // Read each record
                while (csv.Read())
                {
                     if (csv.Parser.Count != 13)
                    throw new Exception($"Error at line {csv.Parser.Row}: Expected 13 columns but found {csv.Parser.Count} columns.");

                    var record = new ProcessOhcostAs400
                    {
                        Model               = csv.GetField(0) ?? string.Empty,
                        Plant               = csv.GetField(1) ?? string.Empty,
                        FiscalYear          = csv.GetField(2) ?? string.Empty,
                        CostCenter          = csv.GetField(3) ?? string.Empty,
                        TsQuantity          = csv.GetField(4) ?? string.Empty,
                        UnitQuantity        = csv.GetField(5) ?? string.Empty,
                        PricePerUnit        = csv.GetField(6) ?? string.Empty,
                        PriceQtyUnit        = csv.GetField(7) ?? string.Empty,
                        ProcCostRate        = csv.GetField(8) ?? string.Empty,
                        OhCostRate          = csv.GetField(9) ?? string.Empty,
                        TotalProcessCost    = csv.GetField(10) ?? string.Empty,
                        TotalOh             = csv.GetField(11) ?? string.Empty,
                        TotalValue          = csv.GetField(12) ?? string.Empty,
                    };

                    records.Add(record);
                }
            }

        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error processing file: {ex.Message}");
        }

        return records;
    }

    private List<ProcessOhcostSap> ParseExcelFileProcessOhcostSap(string filePath)
    {
        var records = new List<ProcessOhcostSap>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null) throw new Exception("No worksheet found.");

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Skip header row
            {
                var record = new ProcessOhcostSap
                {
                    Model               = worksheet.Cells[row, 1].Text.Trim(),
                    Plant               = worksheet.Cells[row, 2].Text.Trim(),
                    FiscalYear          = worksheet.Cells[row, 3].Text.Trim(),
                    CostCenter          = worksheet.Cells[row, 4].Text.Trim(),
                    TsQuantity          = worksheet.Cells[row, 5].Text.Trim(),
                    UnitQuantity        = worksheet.Cells[row, 6].Text.Trim(),
                    PricePerUnit        = worksheet.Cells[row, 7].Text.Trim(),
                    PriceQtyUnit        = worksheet.Cells[row, 8].Text.Trim(),
                    ProcCostRate        = worksheet.Cells[row, 9].Text.Trim(),
                    OhCostRate          = worksheet.Cells[row, 10].Text.Trim(),
                    TotalProcessCost    = worksheet.Cells[row, 11].Text.Trim(),
                    TotalOh             = worksheet.Cells[row, 12].Text.Trim(),
                    TotalValue          = worksheet.Cells[row, 13].Text.Trim(),
                };

                records.Add(record);
            }
        }

        return records;
    }

    public List<ProcessOhcostSap> ParseCsvFileProcessOhcostSap(string filePath)
    {
        var records = new List<ProcessOhcostSap>();

        try
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",  // Use comma delimiter
                BadDataFound = null  // Ignore bad data rows
            }))
            {
                 if (!csv.Read() || !csv.ReadHeader())
                throw new Exception("CSV file does not contain a valid header.");

                // Skip the header row
                csv.Read();
                csv.ReadHeader();

                // Read each record
                while (csv.Read())
                {
                    if (csv.Parser.Count != 13)
                    throw new Exception($"Error at line {csv.Parser.Row}: Expected 13 columns but found {csv.Parser.Count} columns.");

                    var record = new ProcessOhcostSap
                    {
                        Model               = csv.GetField(0) ?? string.Empty,
                        Plant               = csv.GetField(1) ?? string.Empty,
                        FiscalYear          = csv.GetField(2) ?? string.Empty,
                        CostCenter          = csv.GetField(3) ?? string.Empty,
                        TsQuantity          = csv.GetField(4) ?? string.Empty,
                        UnitQuantity        = csv.GetField(5) ?? string.Empty,
                        PricePerUnit        = csv.GetField(6) ?? string.Empty,
                        PriceQtyUnit        = csv.GetField(7) ?? string.Empty,
                        ProcCostRate        = csv.GetField(8) ?? string.Empty,
                        OhCostRate          = csv.GetField(9) ?? string.Empty,
                        TotalProcessCost    = csv.GetField(10) ?? string.Empty,
                        TotalOh             = csv.GetField(11) ?? string.Empty,
                        TotalValue          = csv.GetField(12) ?? string.Empty,
                    };

                    records.Add(record);
                }
            }

        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error processing file: {ex.Message}");
        }

        return records;
    }

    private List<MasterMaterialBasic> ParseExcelFileMasterMaterialBasic(string filePath)
    {
        var records = new List<MasterMaterialBasic>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null) throw new Exception("No worksheet found.");

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Skip header row
            {
                var record = new MasterMaterialBasic
                {
                    As400Material          = worksheet.Cells[row, 1].Text.Trim(),
                    SapMaterial            = worksheet.Cells[row, 2].Text.Trim(),
                    Description            = worksheet.Cells[row, 3].Text.Trim(),
                    BaseUnit               = worksheet.Cells[row, 4].Text.Trim(),
                    MaterialType           = worksheet.Cells[row, 5].Text.Trim(),
                    MaterialGroup          = worksheet.Cells[row, 6].Text.Trim(),
                    NetWeight              = decimal.TryParse(worksheet.Cells[row, 7].Text.Trim(), out var Denominator) ? Denominator : 0,
                    WeightUnit             = worksheet.Cells[row, 8].Text.Trim(),
                };

                records.Add(record);
            }
        }

        return records;
    }


    public List<MasterMaterialBasic> ParseCsvFileMasterMaterialBasic(string filePath)
    {
        var records = new List<MasterMaterialBasic>();

        try
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",  // Use comma delimiter
                BadDataFound = null  // Ignore bad data rows
            }))
            {
                 if (!csv.Read() || !csv.ReadHeader())
                throw new Exception("CSV file does not contain a valid header.");

                // Skip the header row
                csv.Read();
                csv.ReadHeader();

                // Read each record
                while (csv.Read())
                {
                    if (csv.Parser.Count != 9 && csv.Parser.Count != 3)
                    throw new Exception($"Error at line {csv.Parser.Row}: Expected columns but found {csv.Parser.Count} columns.");

                    var record = new MasterMaterialBasic
                    {
                        As400Material   = csv.GetField(0) ?? string.Empty,
                        SapMaterial     = csv.GetField(1) ?? string.Empty,
                        Description     = csv.GetField(2) ?? string.Empty,
                        BaseUnit        = csv.GetField(3) ?? string.Empty,
                        MaterialType    = csv.GetField(4) ?? string.Empty,
                        MaterialGroup   = csv.GetField(5) ?? string.Empty,
                        NetWeight       = decimal.TryParse(csv.GetField(7), out var Netweight) ? Netweight : 0,
                        WeightUnit      = csv.GetField(8) ?? string.Empty,
                    };

                    records.Add(record);
                }
            }

        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error processing file: {ex.Message}");
        }

        return records;
    }

    private List<MasterMaterialUnit> ParseExcelFileMasterMaterialUnit(string filePath)
    {
        var records = new List<MasterMaterialUnit>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null) throw new Exception("No worksheet found.");

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Skip header row
            {
                var record = new MasterMaterialUnit
                {
                    Material            = worksheet.Cells[row, 1].Text.Trim(),
                    Description         = worksheet.Cells[row, 2].Text.Trim(),
                    AlternativeUnit     = worksheet.Cells[row, 3].Text.Trim(),
                    Numerator           = decimal.TryParse(worksheet.Cells[row, 4].Text.Trim(), out var Numerator) ? Numerator : 0,
                    Denominator         = decimal.TryParse(worksheet.Cells[row, 5].Text.Trim(), out var Denominator) ? Denominator : 0,
                    BaseUnit            = worksheet.Cells[row, 6].Text.Trim(),
                };

                records.Add(record);
            }
        }

        return records;
    }

    public List<MasterMaterialUnit> ParseCsvFileMasterMaterialUnit(string filePath)
{
    var records = new List<MasterMaterialUnit>();

    try
    {
        using (var reader = new StreamReader(filePath))
        using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
        {
            Delimiter = ",",  // ใช้คอมม่าเป็นตัวแบ่ง
            BadDataFound = null  // ข้ามข้อมูลที่ไม่ถูกต้อง
        }))
        {
            // อ่าน Header
            if (!csv.Read() || !csv.ReadHeader())
                throw new Exception("CSV file does not contain a valid header.");

            while (csv.Read())
            {
                // ตรวจสอบจำนวนคอลัมน์
                if (csv.Parser.Count != 6)
                {
                    throw new Exception($"Error at line {csv.Parser.Row}: Expected 6 columns but found {csv.Parser.Count} columns.");
                }

                // อ่านค่าจากไฟล์หลังจากตรวจสอบผ่าน
                var record = new MasterMaterialUnit
                {
                    Material            = csv.GetField(0) ?? string.Empty,
                    Description         = csv.GetField(1) ?? string.Empty,
                    AlternativeUnit     = csv.GetField(2) ?? string.Empty,
                    Numerator           = decimal.TryParse(csv.GetField(3), out var Numerator) ? Numerator : 0,
                    Denominator         = decimal.TryParse(csv.GetField(4), out var Denominator) ? Denominator : 0,
                    BaseUnit            = csv.GetField(5) ?? string.Empty,
                };

                records.Add(record);
            }
        }
    }
    catch (Exception ex)
    {
        throw new InvalidOperationException($"Error processing file: {ex.Message}");
    }

    return records;
}

          

    private List<CostBomAs400> ParseExcelFileMaterialCostAS400(string filePath)
    {
        var records = new List<CostBomAs400>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null) throw new Exception("No worksheet found.");

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Skip header row
            {
                var record = new CostBomAs400
                {
                    Model = worksheet.Cells[row, 1].Text.Trim(),
                    Plant = worksheet.Cells[row, 2].Text.Trim(),
                    CostingRun = worksheet.Cells[row, 3].Text.Trim(),
                    CostingRundt = worksheet.Cells[row, 4].Text.Trim(),
                    RunningNo = int.TryParse(worksheet.Cells[row, 5].Text.Trim(), out var runningNo) ? runningNo : 0,
                    Lv = int.TryParse(worksheet.Cells[row, 6].Text.Trim(), out var lv) ? lv : 0,
                    ParentMat = worksheet.Cells[row, 7].Text.Trim(),
                    ParentMatDesc = worksheet.Cells[row, 8].Text.Trim(),
                    ParentProcTypeMm = worksheet.Cells[row, 9].Text.Trim(),
                    ParentSpTypeMm = worksheet.Cells[row, 10].Text.Trim(),
                    Component = worksheet.Cells[row, 11].Text.Trim(),
                    ComponentDesc = worksheet.Cells[row, 12].Text.Trim(),
                    ItemType = worksheet.Cells[row, 13].Text.Trim(),
                    CompProcTypeMm = worksheet.Cells[row, 14].Text.Trim(),
                    CompSpTypeMm = worksheet.Cells[row, 15].Text.Trim(),
                    CompSpTypeBom = worksheet.Cells[row, 16].Text.Trim(),
                    BulkMat = worksheet.Cells[row, 17].Text.Trim(),
                    CostRelevancyBom = worksheet.Cells[row, 18].Text.Trim(),
                    PhantomItem = worksheet.Cells[row, 19].Text.Trim(),
                    DeletionIndicator = worksheet.Cells[row, 20].Text.Trim(),
                    MatProvisionIndicator = worksheet.Cells[row, 21].Text.Trim(),
                    CompEffectDtFrom = worksheet.Cells[row, 22].Text.Trim(),
                    CompEffectDtTo = worksheet.Cells[row, 23].Text.Trim(),
                    Unit = worksheet.Cells[row, 24].Text.Trim(),
                    MSize = worksheet.Cells[row, 25].Text.Trim(),
                    QuantityModel = worksheet.Cells[row, 26].Text.Trim(),
                    Scrap = worksheet.Cells[row, 27].Text.Trim(),
                    TotalScrap = decimal.TryParse(worksheet.Cells[row, 28].Text.Trim(), out var totalScrapValue) ? totalScrapValue : 0,
                    TotalQuantity = decimal.TryParse(worksheet.Cells[row, 29].Text.Trim(), out var totalQuantityValue) ? totalQuantityValue : 0,
                    QuantityUnit = worksheet.Cells[row, 30].Text.Trim(),
                    Import = worksheet.Cells[row, 31].Text.Trim(),
                    TaxExp = worksheet.Cells[row, 32].Text.Trim(),
                    Local = worksheet.Cells[row, 33].Text.Trim(),
                    StdPrice = worksheet.Cells[row, 34].Text.Trim(),
                    PriceQtyUnit = worksheet.Cells[row, 35].Text.Trim(),
                    NoPrice = worksheet.Cells[row, 36].Text.Trim(),
                    NoCost = worksheet.Cells[row, 37].Text.Trim(),
                    SumValue = decimal.TryParse(worksheet.Cells[row, 38].Text.Trim(), out var value) ? value : 0,
                    SumTotalValue = decimal.TryParse(worksheet.Cells[row, 39].Text.Trim(), out var totalValue) ? totalValue : 0,
                    PurPriceCur = worksheet.Cells[row, 40].Text.Trim(),
                    ItemClass = worksheet.Cells[row, 41].Text.Trim()
                };

                records.Add(record);
            }
        }

        return records;
    }

    public List<CostBomAs400> ParseCsvFileMaterialCostAS400(string filePath)
    {
        var records = new List<CostBomAs400>();

        try
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",  // Use comma delimiter
                BadDataFound = null  // Ignore bad data rows
            }))
            {
                if (!csv.Read() || !csv.ReadHeader())
                throw new Exception("CSV file does not contain a valid header.");
                // Skip the header row
                csv.Read();
                csv.ReadHeader();

                // Read each record
                while (csv.Read())
                {
                      if (csv.Parser.Count != 41)
                    throw new Exception($"Error at line {csv.Parser.Row}: Expected 41 columns but found {csv.Parser.Count} columns.");

                    var record = new CostBomAs400
                    {
                        Model = csv.GetField(0) ?? string.Empty,
                        Plant = csv.GetField(1) ?? string.Empty,
                        CostingRun = csv.GetField(2) ?? string.Empty,
                        CostingRundt = csv.GetField(3) ?? string.Empty,
                        RunningNo = int.TryParse(csv.GetField(4), out var runningNo) ? runningNo : 0,
                        Lv = int.TryParse(csv.GetField(5), out var lv) ? lv : 0,
                        ParentMat = csv.GetField(6) ?? string.Empty,
                        ParentMatDesc = csv.GetField(7) ?? string.Empty,
                        ParentProcTypeMm = csv.GetField(8) ?? string.Empty,
                        ParentSpTypeMm = csv.GetField(9) ?? string.Empty,
                        Component = csv.GetField(10) ?? string.Empty,
                        ComponentDesc = csv.GetField(11) ?? string.Empty,
                        ItemType = csv.GetField(12) ?? string.Empty,
                        CompProcTypeMm = csv.GetField(13) ?? string.Empty,
                        CompSpTypeMm = csv.GetField(14) ?? string.Empty,
                        CompSpTypeBom = csv.GetField(15) ?? string.Empty,
                        BulkMat = csv.GetField(16) ?? string.Empty,
                        CostRelevancyBom = csv.GetField(17) ?? string.Empty,
                        PhantomItem = csv.GetField(18) ?? string.Empty,
                        DeletionIndicator = csv.GetField(19) ?? string.Empty,
                        MatProvisionIndicator = csv.GetField(20) ?? string.Empty,
                        CompEffectDtFrom = csv.GetField(21) ?? string.Empty,
                        CompEffectDtTo = csv.GetField(22) ?? string.Empty,
                        Unit = csv.GetField(23) ?? string.Empty,
                        MSize = csv.GetField(24) ?? string.Empty,
                        QuantityModel = csv.GetField(25) ?? string.Empty,
                        Scrap = csv.GetField(26) ?? string.Empty,
                        TotalScrap = decimal.TryParse(csv.GetField(27), out var scrapQtyValue) ? scrapQtyValue : 0,
                        TotalQuantity = decimal.TryParse(csv.GetField(28), out var totalQuantityValue) ? totalQuantityValue : 0,
                        QuantityUnit = csv.GetField(29) ?? string.Empty,
                        Import = csv.GetField(30) ?? string.Empty,
                        TaxExp = csv.GetField(31) ?? string.Empty,
                        Local = csv.GetField(32) ?? string.Empty,
                        StdPrice = csv.GetField(33) ?? string.Empty,
                        PriceQtyUnit = csv.GetField(34) ?? string.Empty,
                        NoPrice = csv.GetField(35) ?? string.Empty,
                        NoCost = csv.GetField(36) ?? string.Empty,
                        SumValue = decimal.TryParse(csv.GetField(37), out var value) ? value : 0,
                        SumTotalValue = decimal.TryParse(csv.GetField(38), out var totalValue) ? totalValue : 0,
                        PurPriceCur = csv.GetField(39) ?? string.Empty,
                        ItemClass = csv.GetField(40) ?? string.Empty
                    };

                    records.Add(record);
                }
            }

        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error processing file: {ex.Message}");
        }

        return records;
    }

    private List<TotalCostAs400> ParseExcelFileTotalSTDCostAS400(string filePath)
    {
        var records = new List<TotalCostAs400>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null) throw new Exception("No worksheet found.");

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Skip header row
            {
                var record = new TotalCostAs400
                {
                    FiscalYear = worksheet.Cells[row, 1].Text.Trim(),
                    Model = worksheet.Cells[row, 2].Text.Trim(),
                    MaterialCost = decimal.TryParse(worksheet.Cells[row, 3].Text.Trim(), out var MaterialCost) ? MaterialCost : 0,
                    ProcessCost = decimal.TryParse(worksheet.Cells[row, 4].Text.Trim(), out var ProcessCost) ? ProcessCost : 0,
                    Ohcost = decimal.TryParse(worksheet.Cells[row, 5].Text.Trim(), out var Ohcost) ? Ohcost : 0,
                    Srvpackingpercent = worksheet.Cells[row, 6].Text.Trim(),
                    SrvpackingCost = worksheet.Cells[row, 7].Text.Trim(),
                    TotalStdcost = decimal.TryParse(worksheet.Cells[row, 8].Text.Trim(), out var TotalStdcost) ? TotalStdcost : 0,
                    TotalStdcostRound = decimal.TryParse(worksheet.Cells[row, 9].Text.Trim(), out var TotalStdcostRound) ? TotalStdcostRound : 0,
                    Unit = worksheet.Cells[row, 10].Text.Trim(),
                    Priceperunit = worksheet.Cells[row, 11].Text.Trim(),
                    TotalTs = decimal.TryParse(worksheet.Cells[row, 12].Text.Trim(), out var TotalTs) ? TotalTs : 0,
                    Tsunit = worksheet.Cells[row, 13].Text.Trim()
                };

                records.Add(record);
            }
        }

        return records;
    }

    public List<TotalCostAs400> ParseCsvFileTotalSTDCostAS400(string filePath)
    {
        var records = new List<TotalCostAs400>();

        try
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",  // Use comma delimiter
                BadDataFound = null  // Ignore bad data rows
            }))
            {
                if (!csv.Read() || !csv.ReadHeader())
                throw new Exception("CSV file does not contain a valid header.");
                // Skip the header row
                csv.Read();
                csv.ReadHeader();

                // Read each record
                while (csv.Read())
                {

                    var record = new TotalCostAs400
                    {
                        FiscalYear = csv.GetField(0) ?? string.Empty,
                        Model = csv.GetField(1) ?? string.Empty,
                        MaterialCost = decimal.TryParse(csv.GetField(2), out var MaterialCost) ? MaterialCost : 0,
                        ProcessCost = decimal.TryParse(csv.GetField(3), out var ProcessCost) ? ProcessCost : 0,
                        Ohcost = int.TryParse(csv.GetField(4), out var lv) ? lv : 0,
                        Srvpackingpercent = csv.GetField(5) ?? string.Empty,
                        SrvpackingCost = csv.GetField(6) ?? string.Empty,
                        TotalStdcost = decimal.TryParse(csv.GetField(7), out var TotalStdcost) ? TotalStdcost : 0,
                        TotalStdcostRound = decimal.TryParse(csv.GetField(8), out var TotalStdcostRound) ? TotalStdcostRound : 0,
                        Unit = csv.GetField(9) ?? string.Empty,
                        Priceperunit = csv.GetField(10) ?? string.Empty,
                        TotalTs = decimal.TryParse(csv.GetField(11), out var TotalTs) ? TotalTs : 0,
                        Tsunit = csv.GetField(12) ?? string.Empty,
                    };

                    records.Add(record);
                }
            }

        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error processing file: {ex.Message}");
        }

        return records;
    }

    private List<TotalCostSap> ParseExcelFileTotalSTDCostSAP(string filePath)
    {
        var records = new List<TotalCostSap>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null) throw new Exception("No worksheet found.");

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Skip header row
            {
                var record = new TotalCostSap
                {
                    FiscalYear = worksheet.Cells[row, 1].Text.Trim(),
                    Model = worksheet.Cells[row, 2].Text.Trim(),
                    MaterialCost = decimal.TryParse(worksheet.Cells[row, 3].Text.Trim(), out var MaterialCost) ? MaterialCost : 0,
                    ProcessCost = decimal.TryParse(worksheet.Cells[row, 4].Text.Trim(), out var ProcessCost) ? ProcessCost : 0,
                    Ohcost = decimal.TryParse(worksheet.Cells[row, 5].Text.Trim(), out var Ohcost) ? Ohcost : 0,
                    Srvpackingpercent = worksheet.Cells[row, 6].Text.Trim(),
                    SrvpackingCost = worksheet.Cells[row, 7].Text.Trim(),
                    TotalStdcost = decimal.TryParse(worksheet.Cells[row, 8].Text.Trim(), out var TotalStdcost) ? TotalStdcost : 0,
                    TotalStdcostRound = decimal.TryParse(worksheet.Cells[row, 9].Text.Trim(), out var TotalStdcostRound) ? TotalStdcostRound : 0,
                    Unit = worksheet.Cells[row, 10].Text.Trim(),
                    Priceperunit = worksheet.Cells[row, 11].Text.Trim(),
                    TotalTs = decimal.TryParse(worksheet.Cells[row, 12].Text.Trim(), out var TotalTs) ? TotalTs : 0,
                    Tsunit = worksheet.Cells[row, 13].Text.Trim()
                };

                records.Add(record);
            }
        }

        return records;
    }

    public List<TotalCostSap> ParseCsvFileTotalSTDCostSAP(string filePath)
    {
        var records = new List<TotalCostSap>();

        try
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",  // Use comma delimiter
                BadDataFound = null  // Ignore bad data rows
            }))
            {
                if (!csv.Read() || !csv.ReadHeader())
                throw new Exception("CSV file does not contain a valid header.");
                // Skip the header row
                csv.Read();
                csv.ReadHeader();

                // Read each record
                while (csv.Read())
                {

                    var record = new TotalCostSap
                    {
                        FiscalYear = csv.GetField(0) ?? string.Empty,
                        Model = csv.GetField(1) ?? string.Empty,
                        MaterialCost = decimal.TryParse(csv.GetField(2), out var MaterialCost) ? MaterialCost : 0,
                        ProcessCost = decimal.TryParse(csv.GetField(3), out var ProcessCost) ? ProcessCost : 0,
                        Ohcost = int.TryParse(csv.GetField(4), out var lv) ? lv : 0,
                        Srvpackingpercent = csv.GetField(5) ?? string.Empty,
                        SrvpackingCost = csv.GetField(6) ?? string.Empty,
                        TotalStdcost = decimal.TryParse(csv.GetField(7), out var TotalStdcost) ? TotalStdcost : 0,
                        TotalStdcostRound = decimal.TryParse(csv.GetField(8), out var TotalStdcostRound) ? TotalStdcostRound : 0,
                        Unit = csv.GetField(9) ?? string.Empty,
                        Priceperunit = csv.GetField(10) ?? string.Empty,
                        TotalTs = decimal.TryParse(csv.GetField(11), out var TotalTs) ? TotalTs : 0,
                        Tsunit = csv.GetField(12) ?? string.Empty,
                    };

                    records.Add(record);
                }
            }

        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error processing file: {ex.Message}");
        }

        return records;
    }

     private List<CostBomSap> ParseExcelFileMaterialCostSAP(string filePath)
    {
        var records = new List<CostBomSap>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[0];
            if (worksheet == null) throw new Exception("No worksheet found.");

            for (int row = 2; row <= worksheet.Dimension.End.Row; row++) // Skip header row
            {
                var record = new CostBomSap
                {
                    Model = worksheet.Cells[row, 1].Text.Trim(),
                    Plant = worksheet.Cells[row, 2].Text.Trim(),
                    CostingRun = worksheet.Cells[row, 3].Text.Trim(),
                    CostingRundt = worksheet.Cells[row, 4].Text.Trim(),
                    RunningNo = int.TryParse(worksheet.Cells[row, 5].Text.Trim(), out var runningNo) ? runningNo : 0,
                    Lv = int.TryParse(worksheet.Cells[row, 6].Text.Trim(), out var lv) ? lv : 0,
                    ParentMat = worksheet.Cells[row, 7].Text.Trim(),
                    ParentMatDesc = worksheet.Cells[row, 8].Text.Trim(),
                    ParentProcTypeMm = worksheet.Cells[row, 9].Text.Trim(),
                    ParentSpTypeMm = worksheet.Cells[row, 10].Text.Trim(),
                    Component = worksheet.Cells[row, 11].Text.Trim(),
                    ComponentDesc = worksheet.Cells[row, 12].Text.Trim(),
                    ItemType = worksheet.Cells[row, 13].Text.Trim(),
                    CompProcTypeMm = worksheet.Cells[row, 14].Text.Trim(),
                    CompSpTypeMm = worksheet.Cells[row, 15].Text.Trim(),
                    CompSpTypeBom = worksheet.Cells[row, 16].Text.Trim(),
                    BulkMat = worksheet.Cells[row, 17].Text.Trim(),
                    CostRelevancyBom = worksheet.Cells[row, 18].Text.Trim(),
                    PhantomItem = worksheet.Cells[row, 19].Text.Trim(),
                    DeletionIndicator = worksheet.Cells[row, 20].Text.Trim(),
                    MatProvisionIndicator = worksheet.Cells[row, 21].Text.Trim(),
                    CompEffectDtFrom = worksheet.Cells[row, 22].Text.Trim(),
                    CompEffectDtTo = worksheet.Cells[row, 23].Text.Trim(),
                    Unit = worksheet.Cells[row, 24].Text.Trim(),
                    MSize = worksheet.Cells[row, 25].Text.Trim(),
                    QuantityModel = worksheet.Cells[row, 26].Text.Trim(),
                    Scrap = worksheet.Cells[row, 27].Text.Trim(),
                    TotalScrap = decimal.TryParse(worksheet.Cells[row, 28].Text.Trim(), out var totalScrapValue) ? totalScrapValue : 0,
                    TotalQuantity = decimal.TryParse(worksheet.Cells[row, 29].Text.Trim(), out var totalQuantityValue) ? totalQuantityValue : 0,
                    QuantityUnit = worksheet.Cells[row, 30].Text.Trim(),
                    Import = worksheet.Cells[row, 31].Text.Trim(),
                    TaxExp = worksheet.Cells[row, 32].Text.Trim(),
                    Local = worksheet.Cells[row, 33].Text.Trim(),
                    StdPrice = worksheet.Cells[row, 34].Text.Trim(),
                    PriceQtyUnit = worksheet.Cells[row, 35].Text.Trim(),
                    NoPrice = worksheet.Cells[row, 36].Text.Trim(),
                    NoCost = worksheet.Cells[row, 37].Text.Trim(),
                    SumValue = decimal.TryParse(worksheet.Cells[row, 38].Text.Trim(), out var value) ? value : 0,
                    SumTotalValue = decimal.TryParse(worksheet.Cells[row, 39].Text.Trim(), out var totalValue) ? totalValue : 0,
                    PurPriceCur = worksheet.Cells[row, 40].Text.Trim(),
                    ItemClass = worksheet.Cells[row, 41].Text.Trim()
                };

                records.Add(record);
            }
        }

        return records;
    }

    public List<CostBomSap> ParseCsvFileMaterialCostSAP(string filePath)
    {
        var records = new List<CostBomSap>();

        try
        {
            using (var reader = new StreamReader(filePath))
            using (var csv = new CsvReader(reader, new CsvConfiguration(CultureInfo.InvariantCulture)
            {
                Delimiter = ",",  // Use comma delimiter
                BadDataFound = null  // Ignore bad data rows
            }))
            {
                 if (!csv.Read() || !csv.ReadHeader())
                throw new Exception("CSV file does not contain a valid header.");

                // Skip the header row
                csv.Read();
                csv.ReadHeader();

                // Read each record
                while (csv.Read())
                {
                      if (csv.Parser.Count != 41)
                    throw new Exception($"Error at line {csv.Parser.Row}: Expected 41 columns but found {csv.Parser.Count} columns.");

                    var record = new CostBomSap
                    {
                        Model = csv.GetField(0) ?? string.Empty,
                        Plant = csv.GetField(1) ?? string.Empty,
                        CostingRun = csv.GetField(2) ?? string.Empty,
                        CostingRundt = csv.GetField(3) ?? string.Empty,
                        RunningNo = int.TryParse(csv.GetField(4), out var runningNo) ? runningNo : 0,
                        Lv = int.TryParse(csv.GetField(5), out var lv) ? lv : 0,
                        ParentMat = csv.GetField(6) ?? string.Empty,
                        ParentMatDesc = csv.GetField(7) ?? string.Empty,
                        ParentProcTypeMm = csv.GetField(8) ?? string.Empty,
                        ParentSpTypeMm = csv.GetField(9) ?? string.Empty,
                        Component = csv.GetField(10) ?? string.Empty,
                        ComponentDesc = csv.GetField(11) ?? string.Empty,
                        ItemType = csv.GetField(12) ?? string.Empty,
                        CompProcTypeMm = csv.GetField(13) ?? string.Empty,
                        CompSpTypeMm = csv.GetField(14) ?? string.Empty,
                        CompSpTypeBom = csv.GetField(15) ?? string.Empty,
                        BulkMat = csv.GetField(16) ?? string.Empty,
                        CostRelevancyBom = csv.GetField(17) ?? string.Empty,
                        PhantomItem = csv.GetField(18) ?? string.Empty,
                        DeletionIndicator = csv.GetField(19) ?? string.Empty,
                        MatProvisionIndicator = csv.GetField(20) ?? string.Empty,
                        CompEffectDtFrom = csv.GetField(21) ?? string.Empty,
                        CompEffectDtTo = csv.GetField(22) ?? string.Empty,
                        Unit = csv.GetField(23) ?? string.Empty,
                        MSize = csv.GetField(24) ?? string.Empty,
                        QuantityModel = csv.GetField(25) ?? string.Empty,
                        Scrap = csv.GetField(26) ?? string.Empty,
                        TotalScrap = decimal.TryParse(csv.GetField(27), out var scrapQtyValue) ? scrapQtyValue : 0,
                        TotalQuantity = decimal.TryParse(csv.GetField(28), out var totalQuantityValue) ? totalQuantityValue : 0,
                        QuantityUnit = csv.GetField(29) ?? string.Empty,
                        Import = csv.GetField(30) ?? string.Empty,
                        TaxExp = csv.GetField(31) ?? string.Empty,
                        Local = csv.GetField(32) ?? string.Empty,
                        StdPrice = csv.GetField(33) ?? string.Empty,
                        PriceQtyUnit = csv.GetField(34) ?? string.Empty,
                        NoPrice = csv.GetField(35) ?? string.Empty,
                        NoCost = csv.GetField(36) ?? string.Empty,
                        SumValue = decimal.TryParse(csv.GetField(37), out var value) ? value : 0,
                        SumTotalValue = decimal.TryParse(csv.GetField(38), out var totalValue) ? totalValue : 0,
                        PurPriceCur = csv.GetField(39) ?? string.Empty,
                        ItemClass = csv.GetField(40) ?? string.Empty
                    };

                    records.Add(record);
                }
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error processing file: {ex.Message}");
        }

        return records;
    }

    // Change string? to string (non-nullable)
    public class DataPayload
    {
        public List<Dictionary<string, object>>? Data { get; set; }  // Non-nullable string as key
        public string? Radio { get; set; }
    }

    [HttpPost]
    public IActionResult SentData([FromBody] DataPayload payload)
    {
        var data = payload.Data;  // Table data
        var radio = payload.Radio; // Selected radio button value (SAP or AS400)

        if (radio != null && data != null && data.Count > 0)
        {
            // Process the data and log the system type (SAP or AS400)
            Console.WriteLine($"System Type: {radio}"); // Log or use the systemType for processing

            foreach (var row in data)
            {
                // Iterate over each row in the data list
                foreach (var kvp in row)
                {
                    string columnName = kvp.Key;  // Column name (key)
                    object value = kvp.Value;     // Value of the column (value)

                    // For example, log or process the data
                    Console.WriteLine($"Column: {columnName}, Value: {value}");
                }
            }

            // If data processing is successful
            return Json(new { success = true });
        }
        else
        {
            // If no data is received
            return Json(new { success = false, message = "No data received" });
        }
    }


    [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
    public IActionResult Error()
    {
        return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
    }
}
