using System.Globalization;
using System.Security.Cryptography;
using System.Text.Json;
using ClosedXML.Excel;
using ExcelLite.Core;
using ExcelLite.Foundation;
using ExcelLite.Persistence;

CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("en-US");
if (args.Length == 2 && args[0] == "--locales")
{
    string Pattern(string value) => value.Replace("tt", "AM/PM").Replace("'", "\"");
    var locales = TemporalFormatOptions.Locales.Select(c => new object[] {
        c.LCID, c.Name, Pattern(TemporalFormats.DatePattern(c)), Pattern(c.DateTimeFormat.LongDatePattern),
        Pattern(c.DateTimeFormat.LongTimePattern), c.NumberFormat.NumberGroupSeparator, c.NumberFormat.NumberDecimalSeparator });
    File.WriteAllText(args[1], "[\n" + string.Join(",\n", locales.Select(locale => JsonSerializer.Serialize(locale,
        new JsonSerializerOptions { Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }))) + "\n]\n");
    return;
}
if (args.Length == 2 && args[0] is "--capture-1904" or "--check-1904")
{
    var path = args[1];
    if (args[0] == "--capture-1904")
    {
        using var book = new XLWorkbook { Use1904DateSystem = true };
        var sheet = book.AddWorksheet("Epoch");
        var formats = new[] { "[$-405]dd.mm.yyyy", "h:mm:ss", "[h]:mm:ss", "0.00", "yyyy-mm-dd hh:mm" };
        for (var row = 1; row <= formats.Length; row++)
        {
            sheet.Cell(row, 1).Value = 1.5;
            sheet.Cell(row, 1).Style.NumberFormat.Format = formats[row - 1];
        }
        book.SaveAs(path);
    }
    var expected = new[] { (1463.5, "02.01.1904"), (1.5, "12:00:00"), (1.5, "36:00:00"),
        (1.5, "1.50"), (1463.5, "1904-01-02 12:00") };
    var imported = XlsxExchange.Import(path);
    var cells = imported.Document.Sheets.Single().Cells.OrderBy(cell => cell.Row).ToArray();
    if (cells.Length != expected.Length || cells.Select(cell =>
        (((CellContent.Number)cell.Content).Value, EverydayCommands.Display(cell))).SequenceEqual(expected) == false)
        throw new Exception("The desktop 1904 import differs from the stored epoch expectations: " +
            string.Join(", ", cells.Select(cell => (EverydayCommands.Value(cell), EverydayCommands.Display(cell)))));
    Console.WriteLine("Verified 5 desktop XLSX 1904 date, time, elapsed and numeric expectations.");
    return;
}

var cases = new List<object>();
void Capture(string code, params double[] values)
{
    foreach (var value in values)
    {
        var cell = new CellDocument(Guid.NewGuid(), 1, 1, new CellContent.Number(value), new(false, code, null), null);
        cases.Add(new { code, value, display = EverydayCommands.Display(cell), color = NumberFormats.DisplayColor(cell),
            entry = TemporalFormats.EditText(value, code) });
    }
}
foreach (var preset in NumberFormats.Presets)
    Capture(EverydayCommands.NumberFormat(preset), preset == "Scientific" ? [1234.5, -1234.5] : [1234.5, -1234.5, 0]);
foreach (var option in NumberFormats.FractionTypes) Capture(option, 2.5, 0.33333, -1.25, 2, 0);
foreach (var culture in SpecialFormatOptions.Locales)
    foreach (var option in SpecialFormatOptions.Choices(culture))
        Capture(option.Code, 18100, 123456789, 1234567890);
foreach (var name in new[] { "cs-CZ", "en-US", "sk-SK", "en-GB", "de-DE", "fr-FR", "pl-PL" })
{
    var culture = CultureInfo.GetCultureInfo(name);
    foreach (var option in TemporalFormatOptions.Choices(false, culture).Concat(TemporalFormatOptions.Choices(true, culture)))
        Capture(option.Code, 40982.563136574074);
    foreach (var option in NumberFormatDialogOptions.Negatives("Number", 3, true, "USD", culture))
        Capture(option.Code, 1234.125, -1234.125, 0);
}
foreach (var currency in new[] { "None", "USD", "CZK", "EUR", "GBP", "EUR-SK", "PLN" })
    Capture(NumberFormats.Numeric("Accounting", currency: currency), 1234.5, -1234.5, 0);

Capture(NumberFormats.Comma, 1234.5, -1234.5, 0);
Capture("[$-405]dd.mm.yyyy", 10, 59, 60, 60.5, 61);
Capture("[$-405]h:mm:ss", 10, 10.123456789, .5);
Capture("yyyy-mm-dd hh:mm:ss.000", 60.5, 40982.563136574074);
Capture("0.00\" days\";[Red](0.00\" days\");\"none\";@", 1.125, -1.25, 0);
Capture("#,##0.0,,\"M\"", 1250000, -1250000);
Capture("0.##", 1, 1.25, 1.125);
Capture("000\\ 0000", 123, 12345678);
Capture("0.0E-00", .0025, 2500);
Capture("##0.0E+0", .0025, 2500);
Capture("dddd, yyyy-mm-dd", 0, 1, 59, 60, 61);
Capture("[h]:mm:ss", -1.25, 0, 1.25);
Capture("0.0;[Red]-0.0;\"zero\"", -1, 0, 1);
Capture("0.##\"x\"", 2, .01, 100);
Capture("0.0\\;\" semicolon\"", 1.25);
Capture("0.0%;[Red]0.0%", -.125, .125);
Capture("mm:ss.0", .00001234, .0000129, .999999);
Capture("# ?/?", .999, .999999, 1.99999);
Capture("?/??", .999, 1.333, -1.333);
Capture("0.0E+00", 9.99, .0999, -.0999);
Capture("[>=10]0;[<0]0", -5, 0, 5, 10);
Capture("[$€-407]#,##0.00", -1234.5, 1234.5);
Capture(";;;",-12.345, 0, 12.345);
Capture("0;;", -12.345, 0, 12.345);
Capture(".00", -12.345, .2345, 12.345);
Capture("0.00,00", 12.345);
Capture("General 0", 12.345);
Capture("yyyy 000", 12.345);
Capture("[DBNum1]0", 12.345);
Capture("[ms]", 12.345);
Capture("0/00", -12.345, 0, 12.345);
Capture("0 0/0", 0, 2);
Capture("[<0]0;0", -12.345, 0, 12.345);
Capture("[>=10]0;[>=0]0;(0)", -12.345, 0, 12.345);
Capture("0.0E+", -12.345, .2345, 12.345);
Capture("0.0e+00", -12.345, .2345, 12.345);
Capture("General", 1e-5, 1e15, 10.220000000000001);
Capture("[>=1000]#,##0;[<0](0);0.00", 1234.5, -12.5, 0.5);

var assembly = typeof(ExcelNumberFormat.NumberFormat).Assembly;
var file = new { source = "ExcelLite.Foundation.EverydayCommands and ExcelLite.Core.TemporalFormats",
    runtime = ".NET " + Environment.Version, assembly = assembly.GetName().Version!.ToString(),
    sha256 = Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(assembly.Location))).ToLowerInvariant(), cases };
var json = JsonSerializer.Serialize(file, new JsonSerializerOptions { WriteIndented = true,
    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping }) + "\n";
if (args.Length == 2 && args[0] == "--capture")
    File.WriteAllText(args[1], json);
else if (args.Length != 1 || File.ReadAllText(args[0]).Replace("\r\n", "\n") != json.Replace("\r\n", "\n"))
    throw new Exception("Desktop format fixture differs.");
Console.WriteLine($"Verified {cases.Count} desktop display and edit cases.");
