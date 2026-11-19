using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlDemo;

/// <summary>
/// تعریف قالب‌های مشترک Excel؛ مقدار سلول از ظاهر آن جدا نگه داشته می‌شود.
/// </summary>
internal static class ExcelStyleBuilder
{
    internal static void AddStyles(WorkbookPart workbookPart)
    {
        var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();

        // شناسه‌های قالب سفارشی از ۱۶۴ شروع می‌شوند؛ ۹ و ۲۰ قالب‌های داخلی Excel هستند.
        var numberFormats = new NumberingFormats(
            new NumberingFormat { NumberFormatId = 164, FormatCode = "#,##0\" ریال\"" },
            new NumberingFormat { NumberFormatId = 165, FormatCode = "yyyy-mm-dd" },
            new NumberingFormat { NumberFormatId = 166, FormatCode = "#,##0.000" })
        {
            Count = 3
        };

        var fonts = new Fonts(
            new Font(new FontSize { Val = 11 }, new FontName { Val = "Tahoma" }),
            new Font(
                new Bold(),
                new FontSize { Val = 11 },
                new Color { Rgb = "FFFFFFFF" },
                new FontName { Val = "Tahoma" }))
        {
            Count = 2
        };

        // دو Fill اول ساختار متعارف Excel هستند؛ Fill سوم رنگ سرستون‌ها است.
        var fills = new Fills(
            new Fill(new PatternFill { PatternType = PatternValues.None }),
            new Fill(new PatternFill { PatternType = PatternValues.Gray125 }),
            new Fill(new PatternFill(
                new ForegroundColor { Rgb = "FF173D56" },
                new BackgroundColor { Indexed = 64 })
            {
                PatternType = PatternValues.Solid
            }))
        {
            Count = 3
        };

        var borders = new Borders(new Border(
            new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder()))
        {
            Count = 1
        };

        var baseFormats = new CellStyleFormats(new CellFormat
        {
            NumberFormatId = 0,
            FontId = 0,
            FillId = 0,
            BorderId = 0
        })
        {
            Count = 1
        };

        // ترتیب قالب‌ها: معمولی، سرستون، پول، درصد، تاریخ، اعشار و زمان.
        // StyleIndex هر سلول، شماره جایگاه آن در همین فهرست است.
        var cellFormats = new CellFormats(
            CreateTextFormat(isHeader: false),
            CreateTextFormat(isHeader: true),
            CreateNumberFormat(164),
            CreateNumberFormat(9),
            CreateNumberFormat(165),
            CreateNumberFormat(166),
            CreateNumberFormat(20))
        {
            Count = 7
        };

        stylesPart.Stylesheet = new Stylesheet(
            numberFormats,
            fonts,
            fills,
            borders,
            baseFormats,
            cellFormats,
            new CellStyles(new CellStyle { Name = "Normal", FormatId = 0, BuiltinId = 0 })
            {
                Count = 1
            });
    }

    private static CellFormat CreateTextFormat(bool isHeader)
    {
        return new CellFormat(
            new Alignment
            {
                Vertical = VerticalAlignmentValues.Center,
                WrapText = true,
                ReadingOrder = 2
            })
        {
            NumberFormatId = 0,
            FontId = isHeader ? 1U : 0U,
            FillId = isHeader ? 2U : 0U,
            BorderId = 0,
            FormatId = 0,
            ApplyFont = isHeader,
            ApplyFill = isHeader,
            ApplyAlignment = true
        };
    }

    private static CellFormat CreateNumberFormat(uint numberFormatId)
    {
        return new CellFormat
        {
            NumberFormatId = numberFormatId,
            FontId = 0,
            FillId = 0,
            BorderId = 0,
            FormatId = 0,
            ApplyNumberFormat = true
        };
    }
}
