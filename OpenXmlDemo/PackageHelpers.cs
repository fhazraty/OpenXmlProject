using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;

namespace OpenXmlDemo;

/// <summary>
/// کارهای مشترک بسته‌های Office: نوشتن XML و بررسی ساختار فایل خروجی.
/// </summary>
internal static class PackageHelpers
{
    internal static void WriteXml(OpenXmlPart part, string xml)
    {
        // UTF-8 متن فارسی را بدون وابستگی به زبان سیستم ذخیره می‌کند.
        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        part.FeedData(stream);
    }

    internal static void Validate(OpenXmlPackage package)
    {
        // اعتبارسنج، ساختار بخش‌ها و رابطه‌های آن‌ها را بررسی می‌کند؛ فرمول را اجرا نمی‌کند.
        var errors = new OpenXmlValidator().Validate(package).Take(20).ToList();
        if (errors.Count == 0)
        {
            return;
        }

        // برای پیدا کردن محل خطا، نام بخش و مسیر عنصر مشکل‌دار را هم گزارش می‌دهیم.
        var details = errors.Select(error =>
            $"بخش: {error.Part?.Uri} | مسیر: {error.Path?.XPath} | خطا: {error.Description}");

        throw new InvalidOperationException(
            "ساختار فایل Office معتبر نیست:" + Environment.NewLine +
            string.Join(Environment.NewLine, details));
    }
}
