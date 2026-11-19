using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlDemo;

/// <summary>
/// ساخت نمونه Excel؛ ترتیب متدهای Create همان مراحل ساخت فایل است.
/// </summary>
public static class ExcelSample
{
    // هر عدد به جای یک شماره نامفهوم، نام قالب مربوط به خود را دارد.
    // ترتیب این شماره‌ها با قالب‌های تعریف‌شده در ExcelStyleBuilder هماهنگ است.
    private const uint DefaultStyle = 0;
    private const uint HeaderStyle = 1;
    private const uint MoneyStyle = 2;
    private const uint PercentStyle = 3;
    private const uint DateStyle = 4;
    private const uint DecimalStyle = 5;
    private const uint TimeStyle = 6;
    private const int FirstProductRow = 5;

    // یک رکورد، اطلاعات مرتبط هر محصول را کنار هم نگه می‌دارد.
    private sealed record Product(string Name, int Quantity, decimal UnitPrice, decimal Discount)
    {
        public decimal NetAmount => decimal.Round(
            Quantity * UnitPrice * (1 - Discount), 0, MidpointRounding.AwayFromZero);
    }

    public static MemoryStream Create()
    {
        var output = new MemoryStream();

        try
        {
            WriteWorkbook(output);

            // جریان پس از تولید در انتهای فایل است؛ دانلود باید از ابتدا خوانده شود.
            output.Position = 0;
            return output;
        }
        catch
        {
            // اگر ساخت فایل شکست بخورد، جریان هنوز به وب‌سرور تحویل نشده است.
            output.Dispose();
            throw;
        }
    }

    private static void WriteWorkbook(MemoryStream output)
    {
        using (var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook))
        {
            // مرحله ۱: دفتر کار و قالب‌های مشترک سلول‌ها را می‌سازیم.
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            ExcelStyleBuilder.AddStyles(workbookPart);
            var sheets = workbookPart.Workbook.AppendChild(new Sheets());

            // مرحله ۲: داده‌ها را در شیت اصلی قرار می‌دهیم.
            Product[] products =
            [
                new("دفتر", 3, 125000, 0.10m),
                new("خودکار", 10, 25000, 0),
                new("کتاب", 2, 350000, 0.05m),
                new("کیف", 1, 850000, 0.20m)
            ];

            var dataSheet = AddWorksheet(workbookPart, sheets, "Data", 1, frozenRows: 4);
            var worksheet = dataSheet.Worksheet!;
            var sheetData = worksheet.GetFirstChild<SheetData>()!;
            AddProductRows(sheetData, products);
            AddSummaryRows(sheetData, products);
            AddDataTypeExamples(sheetData);

            // مرحله ۳: فیلتر و محدودیت‌های ورود اطلاعات را اضافه می‌کنیم.
            // ترتیب اجزای Worksheet در استاندارد Open XML اهمیت دارد.
            worksheet.Append(new AutoFilter { Reference = "A4:G8" });
            AddValidationRules(worksheet);

            // مرحله ۴: نمودار به سلول‌ها متصل می‌شود؛ تصویر ثابت نیست.
            ExcelChartBuilder.AddChart(
                dataSheet,
                products.Select(product => product.Name).ToArray(),
                products.Select(product => product.NetAmount).ToArray());

            // مرحله ۵: روش استفاده و معنی فرمول‌ها را در شیت جداگانه می‌نویسیم.
            AddGuideSheet(workbookPart, sheets);

            // مرحله ۶: SDK فرمول را اجرا نمی‌کند؛ Excel هنگام بازکردن دوباره محاسبه می‌کند.
            workbookPart.Workbook.Append(new CalculationProperties
            {
                CalculationId = 191029,
                FullCalculationOnLoad = true,
                ForceFullCalculation = true,
                CalculationMode = CalculateModeValues.Auto
            });

            workbookPart.Workbook.Save();
            PackageHelpers.Validate(document);
        }

        // بسته نهایی شده، اما جریان تا پایان ارسال پاسخ باز می‌ماند.
    }

    private static void AddProductRows(SheetData sheetData, Product[] products)
    {
        sheetData.Append(CreateRow(1, CreateTextCell("A1", "نمونه Excel با Open XML", HeaderStyle)));
        sheetData.Append(CreateRow(2, CreateTextCell("A2", "ورودی‌ها: ردیف ۵ تا ۸؛ نمودار: پایین جدول")));
        sheetData.Append(CreateRow(4,
            CreateTextCell("A4", "محصول", HeaderStyle),
            CreateTextCell("B4", "تعداد", HeaderStyle),
            CreateTextCell("C4", "قیمت واحد (ریال)", HeaderStyle),
            CreateTextCell("D4", "تخفیف", HeaderStyle),
            CreateTextCell("E4", "مبلغ خالص (ریال)", HeaderStyle),
            CreateTextCell("F4", "وضعیت", HeaderStyle),
            CreateTextCell("G4", "تاریخ", HeaderStyle)));

        var firstDate = new DateTime(2026, 11, 19);
        for (var index = 0; index < products.Length; index++)
        {
            var rowNumber = FirstProductRow + index;
            var product = products[index];

            // تاریخ در Excel یک عدد سریال است؛ قالب سلول آن را به شکل تاریخ نشان می‌دهد.
            var dateSerial = (decimal)firstDate.AddDays(index).ToOADate();
            sheetData.Append(CreateRow((uint)rowNumber,
                CreateTextCell($"A{rowNumber}", product.Name),
                CreateNumberCell($"B{rowNumber}", product.Quantity),
                CreateNumberCell($"C{rowNumber}", product.UnitPrice, MoneyStyle),
                CreateNumberCell($"D{rowNumber}", product.Discount, PercentStyle),
                CreateFormulaCell(
                    $"E{rowNumber}",
                    $"ROUND(B{rowNumber}*C{rowNumber}*(1-D{rowNumber}),0)",
                    product.NetAmount,
                    MoneyStyle),
                CreateTextCell($"F{rowNumber}", "فعال"),
                CreateNumberCell($"G{rowNumber}", dateSerial, DateStyle)));
        }
    }

    private static void AddSummaryRows(SheetData sheetData, Product[] products)
    {
        var total = products.Sum(product => product.NetAmount);

        sheetData.Append(CreateRow(10,
            CreateTextCell("D10", "جمع", HeaderStyle),
            CreateFormulaCell("E10", "SUM(E5:E8)", total, MoneyStyle)));
        sheetData.Append(CreateRow(11,
            CreateTextCell("D11", "میانگین", HeaderStyle),
            CreateFormulaCell("E11", "AVERAGE(E5:E8)", products.Average(product => product.NetAmount), MoneyStyle)));
        sheetData.Append(CreateRow(12,
            CreateTextCell("D12", "بیشترین", HeaderStyle),
            CreateFormulaCell("E12", "MAX(E5:E8)", products.Max(product => product.NetAmount), MoneyStyle)));
        sheetData.Append(CreateRow(13,
            CreateTextCell("D13", "تعداد اعداد", HeaderStyle),
            CreateFormulaCell("E13", "COUNT(B5:B8)", products.Length)));

        // نتیجه IF متن است؛ بنابراین نوع سلول با فرمول‌های عددی تفاوت دارد.
        var conditionCell = new Cell
        {
            CellReference = "E14",
            DataType = CellValues.String,
            CellFormula = new CellFormula("IF(E10>1000000,\"بله\",\"خیر\")"),
            CellValue = new CellValue(total > 1000000 ? "بله" : "خیر")
        };
        sheetData.Append(CreateRow(14, CreateTextCell("D14", "سفارش بزرگ؟", HeaderStyle), conditionCell));
    }

    private static void AddDataTypeExamples(SheetData sheetData)
    {
        sheetData.Append(CreateRow(17,
            CreateTextCell("A17", "نوع", HeaderStyle),
            CreateTextCell("B17", "نمونه", HeaderStyle),
            CreateTextCell("C17", "توضیح", HeaderStyle)));
        sheetData.Append(CreateRow(18,
            CreateTextCell("A18", "عدد منفی"),
            CreateNumberCell("B18", -42),
            CreateTextCell("C18", "عدد صحیح")));
        sheetData.Append(CreateRow(19,
            CreateTextCell("A19", "اعشاری"),
            CreateNumberCell("B19", 1234.567m, DecimalStyle),
            CreateTextCell("C19", "سه رقم اعشار")));

        // مقدار منطقی به صورت ۱ یا ۰ ذخیره و به شکل TRUE یا FALSE نمایش داده می‌شود.
        sheetData.Append(CreateRow(20,
            CreateTextCell("A20", "بولی"),
            new Cell
            {
                CellReference = "B20",
                DataType = CellValues.Boolean,
                CellValue = new CellValue("1")
            },
            CreateTextCell("C20", "TRUE / FALSE")));
        sheetData.Append(CreateRow(21,
            CreateTextCell("A21", "کاراکتر Unicode"),
            CreateTextCell("B21", "ژ")));
        sheetData.Append(CreateRow(22,
            CreateTextCell("A22", "شناسه متنی"),
            CreateTextCell("B22", "00123456789012345678"),
            CreateTextCell("C22", "حفظ صفر اول و ارقام بلند")));
        sheetData.Append(CreateRow(23,
            CreateTextCell("A23", "زمان"),
            CreateNumberCell("B23", 0.5m, TimeStyle),
            CreateTextCell("C23", "۰٫۵ روز یعنی ساعت ۱۲")));
    }

    private static void AddValidationRules(Worksheet worksheet)
    {
        var firstDate = FormatNumber((decimal)new DateTime(2026, 1, 1).ToOADate());
        var lastDate = FormatNumber((decimal)new DateTime(2026, 12, 31).ToOADate());

        var rules = new DataValidations(
            CreateRangeValidation("B5:B8", DataValidationValues.Whole, "1", "100", "عدد صحیح از ۱ تا ۱۰۰"),
            CreateRangeValidation("C5:C8", DataValidationValues.Decimal, "0", "100000000", "قیمت از صفر تا صد میلیون"),
            CreateRangeValidation("D5:D8", DataValidationValues.Decimal, "0", "1", "تخفیف از صفر تا یک"),
            CreateRangeValidation("A5:A8", DataValidationValues.TextLength, "1", "30", "متن از ۱ تا ۳۰ کاراکتر"),
            CreateRangeValidation("G5:G8", DataValidationValues.Date, firstDate, lastDate, "تاریخ در سال ۲۰۲۶"),
            new DataValidation
            {
                Type = DataValidationValues.List,
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "F5:F8" },
                AllowBlank = false,
                ShowErrorMessage = true,
                ErrorStyle = DataValidationErrorStyleValues.Stop,
                ErrorTitle = "ورودی نامعتبر",
                Error = "از فهرست انتخاب کنید",
                Formula1 = new Formula1("\"فعال,غیرفعال\"")
            });

        rules.Count = (uint)rules.ChildElements.Count;
        worksheet.Append(rules);
    }

    private static DataValidation CreateRangeValidation(
        string range,
        DataValidationValues type,
        string minimum,
        string maximum,
        string message)
    {
        // Between برای عدد، تاریخ و طول متن قابل استفاده است.
        // Stop ورود مستقیم مقدار نامعتبر را در رابط Excel متوقف می‌کند.
        return new DataValidation
        {
            Type = type,
            Operator = DataValidationOperatorValues.Between,
            SequenceOfReferences = new ListValue<StringValue> { InnerText = range },
            AllowBlank = false,
            ShowInputMessage = true,
            PromptTitle = "راهنما",
            Prompt = message,
            ShowErrorMessage = true,
            ErrorStyle = DataValidationErrorStyleValues.Stop,
            ErrorTitle = "ورودی نامعتبر",
            Error = message,
            Formula1 = new Formula1(minimum),
            Formula2 = new Formula2(maximum)
        };
    }

    private static WorksheetPart AddWorksheet(
        WorkbookPart workbookPart,
        Sheets sheets,
        string name,
        uint sheetId,
        uint frozenRows)
    {
        var view = new SheetView { WorkbookViewId = 0, RightToLeft = true };
        if (frozenRows > 0)
        {
            view.Append(new Pane
            {
                VerticalSplit = frozenRows,
                TopLeftCell = $"A{frozenRows + 1}",
                ActivePane = PaneValues.BottomLeft,
                State = PaneStateValues.Frozen
            });
        }

        var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
        worksheetPart.Worksheet = new Worksheet(
            new SheetViews(view),
            new Columns(new Column { Min = 1, Max = 7, Width = 26, CustomWidth = true }),
            new SheetData());

        // این رابطه، نام نمایشی شیت را به بخش واقعی Worksheet متصل می‌کند.
        sheets.Append(new Sheet
        {
            Id = workbookPart.GetIdOfPart(worksheetPart),
            SheetId = sheetId,
            Name = name
        });
        return worksheetPart;
    }

    private static void AddGuideSheet(WorkbookPart workbookPart, Sheets sheets)
    {
        var guideSheet = AddWorksheet(workbookPart, sheets, "Guide", 2, frozenRows: 1);
        var worksheet = guideSheet.Worksheet!;
        var columns = worksheet.GetFirstChild<Columns>()!;
        columns.RemoveAllChildren();
        columns.Append(
            new Column { Min = 1, Max = 1, Width = 28, CustomWidth = true },
            new Column { Min = 2, Max = 2, Width = 110, CustomWidth = true });

        (string Topic, string Description)[] guideRows =
        [
            ("راهنما", "توضیح"),
            ("ویرایش", "ورودی‌ها در Data، ردیف‌های ۵ تا ۸ قرار دارند."),
            ("فرمول سطری", "=ROUND(B5*C5*(1-D5),0) تعداد × قیمت × (۱ − تخفیف)"),
            ("جمع", "=SUM(E5:E8) جمع مبلغ محصولات"),
            ("میانگین", "=AVERAGE(E5:E8) میانگین مبلغ محصولات"),
            ("بیشترین", "=MAX(E5:E8) بزرگ‌ترین مبلغ"),
            ("شمارش", "=COUNT(B5:B8) تعداد سلول‌های عددی"),
            ("شرط", "=IF(E10>1000000,\"بله\",\"خیر\") بررسی مبلغ سفارش"),
            ("محاسبه فرمول", "Open XML موتور محاسبه ندارد؛ نتیجه اولیه ذخیره شده و Excel پس از بازشدن دوباره محاسبه می‌کند."),
            ("دقت عددی", "Excel حدود ۱۵ رقم معنادار نگه می‌دارد؛ شناسه بلند و دارای صفر ابتدایی را متن ذخیره کنید."),
            ("قالب عدد", "درصد، پول، تاریخ و ساعت همگی عدد هستند؛ قالب، نحوه نمایش را تعیین می‌کند."),
            ("متن و کاراکتر", "متن فارسی و کاراکتر با InlineString ذخیره می‌شوند؛ سقف متن سلول ۳۲۷۶۷ کاراکتر است."),
            ("محدودیت عددی", "تعداد: صحیح ۱ تا ۱۰۰؛ قیمت: صفر تا صد میلیون؛ تخفیف: صفر تا یک."),
            ("سایر محدودیت‌ها", "نام محصول: ۱ تا ۳۰ کاراکتر؛ وضعیت: فهرست انتخابی؛ تاریخ: سال ۲۰۲۶."),
            ("اعتبارسنجی", "Data Validation ورود معمول در Excel را کنترل می‌کند؛ paste و ابزارهای دیگر ممکن است آن را دور بزنند."),
            ("نمودار", "نمودار ستونی به Data!A5:A8 و Data!E5:E8 متصل است و همراه محاسبات به‌روز می‌شود."),
            ("ابعاد شیت", "حداکثر ۱۰۴۸۵۷۶ ردیف و ۱۶۳۸۴ ستون در هر شیت."),
            ("تاریخ", "تاریخ میلادی، سریال با سیستم ۱۹۰۰؛ تاریخ شمسی نیازمند تبدیل جداگانه است."),
            ("متن شبیه فرمول", "InlineString حتی با شروع = متن می‌ماند؛ فرمول‌ها فقط از کد ثابت برنامه تولید می‌شوند.")
        ];

        var sheetData = worksheet.GetFirstChild<SheetData>()!;
        for (var index = 0; index < guideRows.Length; index++)
        {
            var rowNumber = (uint)index + 1;
            var style = index == 0 ? HeaderStyle : DefaultStyle;
            sheetData.Append(CreateRow(rowNumber,
                CreateTextCell($"A{rowNumber}", guideRows[index].Topic, style),
                CreateTextCell($"B{rowNumber}", guideRows[index].Description, style)));
        }
    }

    private static Row CreateRow(uint rowNumber, params Cell[] cells)
    {
        return new Row(cells) { RowIndex = rowNumber, Height = 34, CustomHeight = true };
    }

    private static Cell CreateTextCell(string address, string value, uint style = DefaultStyle)
    {
        // InlineString متن را مستقیم در سلول ذخیره می‌کند؛ حتی متنِ شروع‌شده با = فرمول نمی‌شود.
        return new Cell
        {
            CellReference = address,
            StyleIndex = style,
            DataType = CellValues.InlineString,
            InlineString = new InlineString(new Text(value))
        };
    }

    private static Cell CreateNumberCell(string address, decimal value, uint style = DefaultStyle)
    {
        return new Cell
        {
            CellReference = address,
            StyleIndex = style,
            DataType = CellValues.Number,
            CellValue = new CellValue(FormatNumber(value))
        };
    }

    private static Cell CreateFormulaCell(
        string address,
        string formula,
        decimal cachedValue,
        uint style = DefaultStyle)
    {
        // فرمول بدون علامت = نوشته می‌شود؛ مقدار ذخیره‌شده برای نمایش اولیه است.
        return new Cell
        {
            CellReference = address,
            StyleIndex = style,
            CellFormula = new CellFormula(formula),
            CellValue = new CellValue(FormatNumber(cachedValue))
        };
    }

    private static string FormatNumber(decimal value)
    {
        // جداکننده اعشار فایل باید نقطه باشد و به زبان ویندوز وابسته نباشد.
        return value.ToString(CultureInfo.InvariantCulture);
    }
}
