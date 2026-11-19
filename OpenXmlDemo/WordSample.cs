using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OpenXmlDemo;

/// <summary>
/// ساخت یک گزارش فارسی با متن راست‌به‌چپ، تیترهای استاندارد و تصویر داخلی.
/// </summary>
public static class WordSample
{
    private const string BodyFontName = "B Nazanin";
    private const string HeadingFontName = "B Titr";
    private const string NormalStyleId = "Normal";
    private const string TitleStyleId = "Title";
    private const string HeadingStyleId = "Heading1";
    private const string ImageResourceName = "OpenXmlDemo.Assets.report.png";

    public static MemoryStream Create()
    {
        // مرحله ۱: فایل را در حافظه می‌سازیم تا بتوان آن را مستقیماً دانلود کرد.
        var output = new MemoryStream();

        try
        {
            WriteDocument(output);

            // پس از نهایی شدن بسته، خواندن پاسخ باید از ابتدای جریان شروع شود.
            output.Position = 0;
            return output;
        }
        catch
        {
            // در صورت خطا، جریان به وب‌سرور تحویل نمی‌شود و خودمان آن را آزاد می‌کنیم.
            output.Dispose();
            throw;
        }
    }

    private static void WriteDocument(MemoryStream output)
    {
        using (var document = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document))
        {
            var mainPart = document.AddMainDocumentPart();

            // مرحله ۲: فونت متن، فونت تیترها و قواعد نوشتار فارسی را تعریف می‌کنیم.
            AddStyles(mainPart);

            // مرحله ۳: متن و تیترهای سند را به ترتیب نمایش اضافه می‌کنیم.
            var body = CreateBody();
            mainPart.Document = new Document(body);

            // مرحله ۴: تصویر در خود فایل Word ذخیره می‌شود؛ لینک بیرونی نیست.
            AddReportImage(mainPart, body);
            body.Append(CreateParagraph("شکل ۱ ـ تصویر تعبیه‌شده در سند"));
            body.Append(CreateParagraph("نحوه تولید سند", HeadingStyleId));
            body.Append(CreateParagraph(
                "ابتدا بسته سند و بخش اصلی ساخته می‌شود. سپس سبک تیترها و تنظیمات فارسی افزوده می‌شوند. " +
                "تصویر در یک بخش مستقل ذخیره و با یک رابطه به متن سند متصل می‌شود. " +
                "در پایان، ساختار فایل با اعتبارسنج Open XML بررسی می‌شود."));

            // مرحله ۵: تنظیمات بخش باید پس از محتوای بدنه قرار بگیرند.
            body.Append(CreatePageSettings());

            // مرحله ۶: ساختار سند را پیش از تحویل بررسی می‌کنیم.
            mainPart.Document.Save();
            PackageHelpers.Validate(document);
        }

        // بستن بسته، فایل ZIP را نهایی می‌کند؛ خود MemoryStream برای ارسال پاسخ باز می‌ماند.
    }

    private static void AddStyles(MainDocumentPart mainPart)
    {
        var stylesPart = mainPart.AddNewPart<StyleDefinitionsPart>();

        // اندازه فونت در Word بر حسب نیم‌پوینت است؛ ۲۸ یعنی ۱۴ پوینت.
        // ComplexScript مشخصاً برای متن فارسی و عربی اهمیت دارد.
        var defaults = new DocDefaults(
            new RunPropertiesDefault(
                new RunPropertiesBaseStyle(
                    CreateFonts(BodyFontName),
                    new FontSize { Val = "28" },
                    new FontSizeComplexScript { Val = "28" },
                    new Languages { Val = "fa-IR", Bidi = "fa-IR" })),
            new ParagraphPropertiesDefault(
                new ParagraphPropertiesBaseStyle(
                    new BiDi(),
                    CreateRightToLeftAlignment())));

        var normalStyle = new Style(
            new StyleName { Val = NormalStyleId },
            new StyleParagraphProperties(
                new BiDi(),
                new SpacingBetweenLines
                {
                    After = "180",
                    Line = "360",
                    LineRule = LineSpacingRuleValues.Auto
                },
                CreateRightToLeftAlignment()),
            new StyleRunProperties(CreateFonts(BodyFontName)))
        {
            Type = StyleValues.Paragraph,
            StyleId = NormalStyleId,
            Default = true
        };

        var titleStyle = CreateHeadingStyle(TitleStyleId, "Title", "40", isSectionHeading: false);
        var headingStyle = CreateHeadingStyle(HeadingStyleId, "heading 1", "30", isSectionHeading: true);

        stylesPart.Styles = new Styles(defaults, normalStyle, titleStyle, headingStyle);
    }

    private static Style CreateHeadingStyle(
        string styleId,
        string displayName,
        string fontSize,
        bool isSectionHeading)
    {
        var paragraphProperties = new StyleParagraphProperties(
            new KeepNext(),
            new BiDi(),
            new SpacingBetweenLines { Before = "240", After = "160" },
            CreateRightToLeftAlignment());

        // سطح تیتر باعث می‌شود بخش‌ها در پنل Navigation ورد هم قابل دسترسی باشند.
        if (isSectionHeading)
        {
            paragraphProperties.Append(new OutlineLevel { Val = 0 });
        }

        return new Style(
            new StyleName { Val = displayName },
            new BasedOn { Val = NormalStyleId },
            new NextParagraphStyle { Val = NormalStyleId },
            paragraphProperties,
            new StyleRunProperties(
                CreateFonts(HeadingFontName),
                new Color { Val = "173D56" },
                new FontSize { Val = fontSize },
                new FontSizeComplexScript { Val = fontSize }))
        {
            Type = StyleValues.Paragraph,
            StyleId = styleId
        };
    }

    private static RunFonts CreateFonts(string fontName)
    {
        // هر چهار گروه نویسه یک فونت دارند تا حروف فارسی و لاتین هماهنگ نمایش داده شوند.
        return new RunFonts
        {
            Ascii = fontName,
            HighAnsi = fontName,
            EastAsia = fontName,
            ComplexScript = fontName
        };
    }

    private static Body CreateBody()
    {
        return new Body(
            CreateParagraph("گزارش نمونه فارسی", TitleStyleId),
            CreateParagraph(
                "این سند با کتابخانه Open XML در یک برنامه وب دات‌نت ۱۱ ساخته شده است. " +
                "متن فارسی، تیترهای استاندارد و تصویرِ تعبیه‌شده، بدون نیاز به نصب Microsoft Word در سرور تولید می‌شوند."),
            CreateParagraph("معرفی گزارش", HeadingStyleId),
            CreateParagraph(
                "هدف این نمونه، نمایش ساخت یک سند خوانا با پشتیبانی از زبان فارسی است. " +
                "جهت پاراگراف‌ها راست‌به‌چپ، تراز متن راست و زبان متن فارسی تنظیم شده است. " +
                "متن با فونت بی نازنین و تیترها با فونت بی تیتر نمایش داده می‌شوند."),
            CreateParagraph("تصویر نمونه", HeadingStyleId),
            CreateParagraph(
                "تصویر زیر یک طرح ساده از گزارش و نمودار است. " +
                "فایل تصویر داخل سند ذخیره می‌شود و برای نمایش به اینترنت وابسته نیست."));
    }

    private static Paragraph CreateParagraph(string text, string styleId = NormalStyleId)
    {
        // هم جهت خواندن و هم تراز ابتدای پاراگراف باید مشخص باشند.
        // قالب علامت پایان پاراگراف، جهت متن تازه را هنگام ادامه تایپ حفظ می‌کند.
        return new Paragraph(
            new ParagraphProperties(
                new ParagraphStyleId { Val = styleId },
                new BiDi { Val = true },
                CreateRightToLeftAlignment(),
                new ParagraphMarkRunProperties(
                    new RightToLeftText { Val = true },
                    new Languages { Val = "fa-IR", Bidi = "fa-IR" })),
            new Run(
                new RunProperties(new RightToLeftText { Val = true }),
                new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
    }

    private static Justification CreateRightToLeftAlignment()
    {
        // نکته مهم: Word مقدار left/right را نسبت به جهت پاراگراف تفسیر می‌کند.
        // وقتی BiDi فعال است، Left به معنی ابتدای پاراگراف و در نتیجه لبه راست است.
        // استفاده از Right در این حالت، متن را به لبه چپ می‌برد.
        // مرجع: https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/26ecf09a-0f0b-4574-9907-ebd1ddf3015f
        return new Justification { Val = JustificationValues.Left };
    }

    private static void AddReportImage(MainDocumentPart mainPart, Body body)
    {
        var imagePart = mainPart.AddImagePart(ImagePartType.Png);
        using (var imageStream = typeof(WordSample).Assembly.GetManifestResourceStream(ImageResourceName)
            ?? throw new InvalidOperationException("تصویر نمونه در منابع برنامه پیدا نشد."))
        {
            imagePart.FeedData(imageStream);
        }

        // اندازه تصویر در DrawingML بر حسب EMU است؛ این ابعاد نسبت ۲ به ۱ را حفظ می‌کنند.
        const long imageWidth = 5_029_200;
        const long imageHeight = 2_514_600;

        var picture = new PIC.Picture(
            new PIC.NonVisualPictureProperties(
                new PIC.NonVisualDrawingProperties { Id = 0U, Name = "report.png" },
                new PIC.NonVisualPictureDrawingProperties()),
            new PIC.BlipFill(
                new A.Blip { Embed = mainPart.GetIdOfPart(imagePart) },
                new A.Stretch(new A.FillRectangle())),
            new PIC.ShapeProperties(
                new A.Transform2D(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = imageWidth, Cy = imageHeight }),
                new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }));

        // Inline یعنی تصویر مانند یک نویسه در جریان متن قرار می‌گیرد و روی متن نمی‌افتد.
        var drawing = new Drawing(
            new DW.Inline(
                new DW.Extent { Cx = imageWidth, Cy = imageHeight },
                new DW.DocProperties
                {
                    Id = 1U,
                    Name = "تصویر گزارش",
                    Description = "طرح گزارش با سه ستون رنگی و خطوط متن"
                },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = true }),
                new A.Graphic(
                    new A.GraphicData(picture)
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U
            });

        body.Append(new Paragraph(
            new ParagraphProperties(
                new BiDi(),
                new Justification { Val = JustificationValues.Center }),
            new Run(drawing)));
    }

    private static SectionProperties CreatePageSettings()
    {
        // ابعاد صفحه و حاشیه‌ها در Word بر حسب یک‌بیستم پوینت (twip) هستند.
        return new SectionProperties(
            new PageSize { Width = 11906, Height = 16838 },
            new PageMargin
            {
                Top = 1000,
                Right = 1134,
                Bottom = 1000,
                Left = 1134,
                Header = 400,
                Footer = 400,
                Gutter = 0
            },
            new BiDi());
    }
}
