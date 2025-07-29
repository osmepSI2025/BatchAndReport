using BatchAndReport.DAO;
using BatchAndReport.Entities;
using BatchAndReport.Repository;
using BatchAndReport.Services;
using Microsoft.EntityFrameworkCore;
using QuestPDF.Infrastructure;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.
builder.Services.AddControllersWithViews();
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddRazorPages();


// ตั้งค่า LicenseType
QuestPDF.Settings.License = LicenseType.Community;

// ตั้งค่า DbContext และการเชื่อมต่อฐานข้อมูล
builder.Services.AddDbContext<BatchDBContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("DefaultConnection")));


// Inside the builder.Services section:
builder.Services.AddDbContext<K2DBContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("K2DBContext")));

// Inside the builder.Services section:
builder.Services.AddDbContext<K2DBContext_SME>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("K2DBContext_SME")));

// Inside the builder.Services section:
builder.Services.AddDbContext<K2DBContext_Workflow>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("K2DBContext_Workflow")));

// Inside the builder.Services section:
builder.Services.AddDbContext<K2DBContext>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("K2DBContext")));

builder.Services.AddDbContext<K2DBContext_SME>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("K2DBContext_SME")));

builder.Services.AddDbContext<K2DBContext_EContract>(options =>
    options.UseSqlServer(builder.Configuration.GetConnectionString("K2DBContext_EContract")));

// Register ScheduledJobService as a singleton and hosted service
builder.Services.AddScoped<IApiInformationRepository, ApiInformationRepository>();
builder.Services.AddScoped<ICallAPIService, CallAPIService>();
builder.Services.AddScoped<SqlConnectionDAO>();
builder.Services.AddScoped<HrDAO>();
builder.Services.AddScoped<SmeDAO>();
builder.Services.AddScoped<WorkflowDAO>();
builder.Services.AddScoped<EContractDAO>();
builder.Services.AddScoped<IPdfService, PdfService>();
builder.Services.AddScoped<IWordService, WordService>();
builder.Services.AddScoped<IWordWFService, WordWFService>();
builder.Services.AddScoped<IWordEContractService, WordEContractService>();
builder.Services.AddHttpClient<ICallAPIService, CallAPIService>();

builder.Services.AddSingleton<WordServiceSetting>();
builder.Services.AddScoped< WordEContract_AllowanceService>();

builder.Services.AddScoped<WordEContract_ContactToDoThingService>();
builder.Services.AddScoped<WordEContract_BorrowMoneyService>();
builder.Services.AddScoped<WordEContract_HireEmployee>();
builder.Services.AddScoped<WordEContract_Test_HeaderLOGOService>();
builder.Services.AddSingleton<ScheduledJobService>();



//service for Word EContract
builder.Services.AddScoped<WordEContract_LoanPrinterService>();
builder.Services.AddScoped<WordEContract_MaintenanceComputerService>(); 
builder.Services.AddScoped<WordEContract_LoanComputerService>();

builder.Services.AddScoped<WordEContract_BuyAgreeProgram>();
builder.Services.AddScoped<WordEContract_BuyOrSellComputerService>();
builder.Services.AddScoped<WordEContract_BuyOrSellService>();
builder.Services.AddScoped<WordEContract_DataSecretService>();
builder.Services.AddScoped<WordEContract_MemorandumService>();
builder.Services.AddScoped<WordEContract_PersernalProcessService>();
builder.Services.AddScoped<WordEContract_SupportSMEsService>();
builder.Services.AddScoped<WordEContract_JointOperationService>();
builder.Services.AddScoped<WordEContract_ControlDataService>();
builder.Services.AddScoped<WordEContract_DataPersonalService>();
builder.Services.AddScoped<WordEContract_ConsultantService>();
builder.Services.AddScoped<WordEContract_ContactToDoThingService>();


//Impoert EContract Report
builder.Services.AddScoped<E_ContractReportDAO>();
builder.Services.AddScoped<Econtract_Report_SPADAO>();
builder.Services.AddScoped<Econtract_Report_CPADAO>();
builder.Services.AddScoped<Econtract_Report_SLADAO>();
builder.Services.AddScoped<Econtract_Report_SMCDAO>();

builder.Services.AddScoped<Econtract_Report_CLADAO>();
builder.Services.AddScoped<Econtract_Report_PMLDAO>();
builder.Services.AddScoped<Econtract_Report_ECDAO>();
builder.Services.AddScoped<Econtract_Report_CTRDAO>();
builder.Services.AddScoped<Econtract_Report_PDSADAO>();
builder.Services.AddScoped<Econtract_Report_CWADAO>();


builder.Services.AddHostedService(provider => provider.GetRequiredService<ScheduledJobService>());

var app = builder.Build();

// Configure the HTTP request pipeline.
if (!app.Environment.IsDevelopment())
{
    app.UseExceptionHandler("/Home/Error");
    app.UseHsts();
}
else 
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

  
app.UseHttpsRedirection();
app.UseStaticFiles();
app.UseRouting();

app.UseAuthorization();

app.MapControllerRoute(
    name: "default",
    pattern: "{controller=Home}/{action=Index}/{id?}");

app.MapGet("/", context =>
{
    context.Response.Redirect("/report/export");
    return Task.CompletedTask;
});

app.MapRazorPages();

app.Run();
