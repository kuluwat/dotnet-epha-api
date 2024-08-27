using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.FileProviders;
using System.Reflection;

var builder = WebApplication.CreateBuilder(args);

// การตั้งค่า CORS เพื่ออนุญาตเฉพาะโดเมนที่ระบุ
var MyAllowSpecificOrigins = "_myAllowSpecificOrigins";
builder.Services.AddCors(options =>
{
    options.AddPolicy(MyAllowSpecificOrigins,
                      policy =>
                      {
                          policy.WithOrigins(
                              "*"
                          )
                          .AllowAnyHeader()
                          .AllowAnyMethod();
                      });
});

// เพิ่มบริการต่างๆ ลงในคอนเทนเนอร์
builder.Services.AddControllersWithViews(options =>
{
    // เพิ่มการตรวจสอบ CSRF สำหรับทุกคำขอ POST, PUT, DELETE
    options.Filters.Add(new AutoValidateAntiforgeryTokenAttribute());
});

// เพิ่มการตั้งค่า JSON ใน Controllers
builder.Services.AddControllers(options => { options.AllowEmptyInputInBodyModelBinding = true; })
    .AddJsonOptions(opt =>
    {
        opt.JsonSerializerOptions.PropertyNameCaseInsensitive = true;
        opt.JsonSerializerOptions.PropertyNamingPolicy = null;
    });

// เพิ่มบริการสำหรับการตรวจสอบ Anti-Forgery
builder.Services.AddAntiforgery(options =>
{
    options.HeaderName = "Authorization"; // กำหนดชื่อของ header ที่ใช้ส่ง token
    options.Cookie.Name = "Authorization"; // ส่ง token ผ่าน cookie ด้วย
});

// เพิ่มการตั้งค่า Swagger
//builder.Services.AddSwaggerGen(c =>
//{
//    try
//    {
//        var xmlFilename = $"{Assembly.GetExecutingAssembly().GetName().Name}.xml";
//        c.IncludeXmlComments(Path.Combine(AppContext.BaseDirectory, xmlFilename));
//    }
//    catch { }
//});
builder.Services.AddSwaggerGen(c =>
 {
     c.ResolveConflictingActions(apiDescriptions => apiDescriptions.First());
     c.AddSecurityDefinition("Bearer",
         new Microsoft.OpenApi.Models.OpenApiSecurityScheme()
         {
             In = Microsoft.OpenApi.Models.ParameterLocation.Header,
             Description = "Please enter into field the word 'Bearer' following by space and JWT",
             Name = "Authorization",
             Type = Microsoft.OpenApi.Models.SecuritySchemeType.ApiKey,
             Scheme = "Bearer"
         });
     c.AddSecurityRequirement(new Microsoft.OpenApi.Models.OpenApiSecurityRequirement()
            {
                {
                    new Microsoft.OpenApi.Models.OpenApiSecurityScheme()
                    {
                        Reference = new Microsoft.OpenApi.Models.OpenApiReference() { Type = Microsoft.OpenApi.Models.ReferenceType.SecurityScheme, Id = "Bearer" },
                        Scheme = "oauth2",
                        Name = "Bearer",
                        In = Microsoft.OpenApi.Models.ParameterLocation.Header,
                    },
                    new List<string>()
                }
            });
 });

// เพิ่มการตั้งค่า Directory Browser และการจัดการไฟล์
builder.Services.AddDirectoryBrowser();

var app = builder.Build();

// เปิดใช้งาน CORS ตาม Policy ที่กำหนด
app.UseCors(MyAllowSpecificOrigins);

// ตรวจสอบว่าโฟลเดอร์ Logs มีอยู่หรือไม่ และตั้งค่าให้บริการไฟล์
string logPath = app.Configuration["appsettings:folder_Logs"] ?? "";
if (Directory.Exists(logPath))
{
    // ให้บริการไฟล์จากโฟลเดอร์ "folder_Log" ผ่าน URL "/log"
    app.UseFileServer(new FileServerOptions
    {
        FileProvider = new PhysicalFileProvider(Path.Combine(logPath, "folder_Log")),
        RequestPath = "/log",
        EnableDirectoryBrowsing = false // ปิดการแสดงรายการไฟล์
    });

    // ให้บริการไฟล์จากโฟลเดอร์ "pic" ผ่าน URL "/pic"
    app.UseFileServer(new FileServerOptions
    {
        FileProvider = new PhysicalFileProvider(Path.Combine(logPath, "pic")),
        RequestPath = "/pic",
        EnableDirectoryBrowsing = false // ปิดการแสดงรายการไฟล์
    });
}

// การตั้งค่า Pipeline ของ HTTP requests
//if (app.Environment.IsDevelopment())
//{
//    //เปิดใช้ Swagger เฉพาะในสภาพแวดล้อมการพัฒนา
//    //app.UseSwagger();
//    //app.UseSwaggerUI(c =>
//    //{
//    //    c.SwaggerEndpoint("/swagger/v1/swagger.json", "My API V1");
//    //});

app.UseSwagger(o => { o.RouteTemplate = "swagger/{documentName}/swagger.json"; });
app.UseSwaggerUI(c =>
{
    c.RoutePrefix = "swagger";
    c.DefaultModelsExpandDepth(-1);
});
//}

app.UseHttpsRedirection();
app.UseStaticFiles();

app.UseRouting();  // ใช้ Routing ก่อน Authorization

// ถ้ามีการใช้ Authentication ให้เปิดใช้ตรงนี้
app.UseAuthentication();

app.UseAuthorization();

app.MapDefaultControllerRoute(); // เพิ่มเส้นทางเริ่มต้นของ Controller

app.Run();
