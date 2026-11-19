using OpenXmlDemo;

// مرحله ۱: برنامه وب را با تنظیمات پیش‌فرض ASP.NET Core می‌سازیم.
var builder = WebApplication.CreateBuilder(args);
var app = builder.Build();

// مرحله ۲: صفحه index.html و فایل‌های پوشه wwwroot را در دسترس قرار می‌دهیم.
app.UseDefaultFiles();
app.UseStaticFiles();

// مرحله ۳: در هر درخواست، یک فایل مستقل در حافظه ساخته و برای دانلود ارسال می‌شود.
// Content-Type نوع فایل و نام آخر، نام پیشنهادی فایل در مرورگر را مشخص می‌کند.
// ASP.NET Core پس از ارسال پاسخ، جریان را Dispose می‌کند؛ اینجا نباید از using استفاده کنیم.
app.MapGet("/downloads/excel", () => Results.File(
    ExcelSample.Create(),
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    "sample.xlsx"));

app.MapGet("/downloads/word", () => Results.File(
    WordSample.Create(),
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    "sample.docx"));

// مرحله ۴: وب‌سرور تا زمان توقف برنامه به درخواست‌ها پاسخ می‌دهد.
app.Run();
