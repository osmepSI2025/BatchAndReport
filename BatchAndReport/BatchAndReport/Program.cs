using BatchAndReport.Entities;
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

// Register ScheduledJobService as a singleton and hosted service
builder.Services.AddSingleton<ScheduledJobService>();
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
