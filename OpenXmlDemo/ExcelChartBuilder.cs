using System.Globalization;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OpenXmlDemo;

/// <summary>
/// ایجاد نمودار ستونی و اتصال آن به سلول‌های شیت Data.
/// </summary>
internal static class ExcelChartBuilder
{
    internal static void AddChart(WorksheetPart worksheetPart, string[] productNames, decimal[] amounts)
    {
        // مرحله ۱: بخش نگهدارنده اشیای ترسیمی و بخش نمودار را ایجاد می‌کنیم.
        var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
        var chartPart = drawingsPart.AddNewPart<ChartPart>();

        // مرحله ۲: داده اولیه نمودار را کنار آدرس سلول‌ها نگه می‌داریم.
        // این cache برای نمایش اولیه است؛ منبع اصلی نمودار همان سلول‌های Excel است.
        XNamespace chartNamespace = "http://schemas.openxmlformats.org/drawingml/2006/chart";
        var labelCache = new XElement(chartNamespace + "strCache",
            new XElement(chartNamespace + "ptCount", new XAttribute("val", productNames.Length)));
        var valueCache = new XElement(chartNamespace + "numCache",
            new XElement(chartNamespace + "formatCode", "#,##0"),
            new XElement(chartNamespace + "ptCount", new XAttribute("val", amounts.Length)));

        for (var index = 0; index < productNames.Length; index++)
        {
            // XElement نویسه‌های خاص متن را به شکل صحیح XML تبدیل می‌کند.
            labelCache.Add(new XElement(chartNamespace + "pt",
                new XAttribute("idx", index),
                new XElement(chartNamespace + "v", productNames[index])));
            valueCache.Add(new XElement(chartNamespace + "pt",
                new XAttribute("idx", index),
                new XElement(chartNamespace + "v", amounts[index].ToString(CultureInfo.InvariantCulture))));
        }

        // مرحله ۳: سری داده، محور افقی و محور عمودی را تعریف می‌کنیم.
        // این XML صریح نوشته شده تا ارتباط بخش‌های نمودار قابل مشاهده باشد.
        var chartXml = $$"""
            <c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart">
              <c:chart>
                <c:autoTitleDeleted val="1" />
                <c:plotArea>
                  <c:layout />
                  <c:barChart>
                    <c:barDir val="col" />
                    <c:grouping val="clustered" />
                    <c:ser>
                      <c:idx val="0" />
                      <c:order val="0" />
                      <c:tx><c:v>مبلغ خالص (ریال)</c:v></c:tx>
                      <c:cat>
                        <c:strRef>
                          <c:f>Data!$A$5:$A$8</c:f>
                          {{labelCache}}
                        </c:strRef>
                      </c:cat>
                      <c:val>
                        <c:numRef>
                          <c:f>Data!$E$5:$E$8</c:f>
                          {{valueCache}}
                        </c:numRef>
                      </c:val>
                    </c:ser>
                    <c:axId val="1" />
                    <c:axId val="2" />
                  </c:barChart>
                  <c:catAx>
                    <c:axId val="1" />
                    <c:scaling><c:orientation val="minMax" /></c:scaling>
                    <c:axPos val="b" />
                    <c:tickLblPos val="nextTo" />
                    <c:crossAx val="2" />
                    <c:crosses val="autoZero" />
                  </c:catAx>
                  <c:valAx>
                    <c:axId val="2" />
                    <c:scaling><c:orientation val="minMax" /></c:scaling>
                    <c:axPos val="l" />
                    <c:majorGridlines />
                    <c:numFmt formatCode="#,##0" sourceLinked="0" />
                    <c:tickLblPos val="nextTo" />
                    <c:crossAx val="1" />
                    <c:crosses val="autoZero" />
                  </c:valAx>
                </c:plotArea>
                <c:legend><c:legendPos val="b" /><c:layout /></c:legend>
                <c:plotVisOnly val="1" />
              </c:chart>
            </c:chartSpace>
            """;
        PackageHelpers.WriteXml(chartPart, chartXml);

        // مرحله ۴: جای نمودار را از ردیف ۲۶ مشخص می‌کنیم.
        // شماره سطر و ستون در DrawingML از صفر شروع می‌شود؛ row=25 یعنی ردیف ۲۶.
        var drawingXml = $$"""
            <xdr:wsDr
                xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
                xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <xdr:twoCellAnchor>
                <xdr:from>
                  <xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff>
                  <xdr:row>25</xdr:row><xdr:rowOff>0</xdr:rowOff>
                </xdr:from>
                <xdr:to>
                  <xdr:col>7</xdr:col><xdr:colOff>0</xdr:colOff>
                  <xdr:row>40</xdr:row><xdr:rowOff>0</xdr:rowOff>
                </xdr:to>
                <xdr:graphicFrame macro="">
                  <xdr:nvGraphicFramePr>
                    <xdr:cNvPr id="2" name="نمودار مبلغ محصولات" />
                    <xdr:cNvGraphicFramePr />
                  </xdr:nvGraphicFramePr>
                  <xdr:xfrm><a:off x="0" y="0" /><a:ext cx="0" cy="0" /></xdr:xfrm>
                  <a:graphic>
                    <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
                      <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"
                          r:id="{{drawingsPart.GetIdOfPart(chartPart)}}" />
                    </a:graphicData>
                  </a:graphic>
                </xdr:graphicFrame>
                <xdr:clientData />
              </xdr:twoCellAnchor>
            </xdr:wsDr>
            """;
        PackageHelpers.WriteXml(drawingsPart, drawingXml);

        // مرحله ۵: شیت را به بخش ترسیمی متصل می‌کنیم تا نمودار در فایل دیده شود.
        worksheetPart.Worksheet!.Append(new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
    }
}
