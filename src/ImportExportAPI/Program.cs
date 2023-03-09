using Microsoft.OpenApi.Models;
using Swashbuckle.AspNetCore.SwaggerGen;
using Swashbuckle.AspNetCore.SwaggerUI;

var builder = WebApplication.CreateBuilder(args);

builder.Services.AddControllers();

builder.Services.AddAWSLambdaHosting(LambdaEventSource.HttpApi);

builder.Services.AddSwaggerGen(ConfigureSwaggerGenOptions);

var app = builder.Build();

app.UseSwagger();
app.UseSwaggerUI(ConfigureSwaggerUiOptions);
app.UseHttpsRedirection();
app.UseAuthorization();
app.MapControllers();
app.UseHttpLogging();
//app.MapGet("/swagger/index.html", () => "Welcome to running ASP.NET Core Minimal API on AWS Lambda");

app.Run();


void ConfigureSwaggerGenOptions(SwaggerGenOptions options)
{
    options.SwaggerDoc("v1", new OpenApiInfo
    {
        Version = "v1",
        Title = ".NET 6",
        Description = "AWS Lambda Minimal API",
    });
}

void ConfigureSwaggerUiOptions(SwaggerUIOptions c)
{
    c.SwaggerEndpoint("/swagger/v1/swagger.json", "Era API");
    c.RoutePrefix = "swagger";
}
