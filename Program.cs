using TemplateSTD.Models;
using Microsoft.AspNetCore.Identity;
using Microsoft.EntityFrameworkCore;
using Microsoft.AspNetCore.Authentication.Cookies;
using TemplateSTD.Models.STD;
using OfficeOpenXml;
using Microsoft.AspNetCore.Http.Features;


var builder = WebApplication.CreateBuilder(args);

ExcelPackage.LicenseContext = LicenseContext.Commercial; 

// Add services to the container.
var connectionString = builder.Configuration.GetConnectionString("DefaultConnection") ?? throw new InvalidOperationException("Connection string 'DefaultConnection' not found.");
builder.Services.AddDbContext<STDContext>(options =>
    options.UseSqlServer(connectionString));
    

// Add MVC services to the services container.
builder.Services.AddMvc();
builder.Services.AddDistributedMemoryCache(); // Adds a default in-memory implementation of IDistributedCache
builder.Services.AddSession();
builder.Services.AddSession(options => {
    options.IdleTimeout = TimeSpan.FromMinutes(120);
});

builder.Services.Configure<FormOptions>(options =>
{
    options.MultipartBodyLengthLimit = 524288000; // 500MB
});

builder.WebHost.ConfigureKestrel(options =>
{
    options.Limits.MaxRequestBodySize = 524288000; // 500MB
});

var app = builder.Build();

app.Use(async (context, next) =>
{
    var maxRequestBodySizeFeature = context.Features.Get<IHttpMaxRequestBodySizeFeature>();
    if (maxRequestBodySizeFeature != null)
    {
        maxRequestBodySizeFeature.MaxRequestBodySize = 524288000; // 500MB
    }
    await next();
});


app.UseSession();
app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();

app.UseAuthentication();
app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=UploadDB}/{id?}");
app.MapRazorPages();

app.Run();
