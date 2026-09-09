using TeamTracker.API.Repositories;
using Microsoft.AspNetCore.Authentication.JwtBearer;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using OfficeOpenXml;

ExcelPackage.License.SetNonCommercialPersonal(
    "TeamTracker");

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddOpenApi();

var dataPath = Path.Combine(
    Directory.GetCurrentDirectory(),
    "..", "..", "Data", "TeamTracker.xlsx");

Console.WriteLine($"TeamTracker path: {dataPath}");
Console.WriteLine($"TeamTracker exists: {File.Exists(dataPath)}");

builder.Services.AddSingleton<ExcelRepository>(
    new ExcelRepository(dataPath));

builder.Services.AddSingleton<TeamTrackerRepository>(
    new TeamTrackerRepository(dataPath));

builder.Services.AddCors(options =>
{
    options.AddPolicy("AllowAll", policy =>
    {
        policy.AllowAnyOrigin()
              .AllowAnyMethod()
              .AllowAnyHeader();
    });
});

builder.Services.AddAuthentication(
    JwtBearerDefaults.AuthenticationScheme)
    .AddJwtBearer(options =>
    {
        options.TokenValidationParameters =
            new TokenValidationParameters
            {
                ValidateIssuer = true,
                ValidateAudience = true,
                ValidateLifetime = true,
                ValidateIssuerSigningKey = true,
                ValidIssuer = builder
                    .Configuration["Jwt:Issuer"],
                ValidAudience = builder
                    .Configuration["Jwt:Audience"],
                IssuerSigningKey =
                    new SymmetricSecurityKey(
                        Encoding.UTF8.GetBytes(
                            builder.Configuration[
                                "Jwt:Key"] ?? ""))
            };
    });

builder.Services.AddControllers();

var app = builder.Build();

if (app.Environment.IsDevelopment())
{
    app.MapOpenApi();
}

//app.UseHttpsRedirection();
app.UseCors("AllowAll");
app.UseAuthentication();
app.UseAuthorization();
app.UseDefaultFiles();
app.UseStaticFiles();
app.MapControllers();

app.Run();
Console.ReadLine();
