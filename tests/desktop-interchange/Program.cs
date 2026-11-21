using ExcelLite.Core;
using ExcelLite.Persistence;

if (args.Length >= 3 && args[0] == "compare" && (args.Length - 1) % 2 == 0)
{
    for (var index = 1; index < args.Length; index += 2)
    {
        var original = NativePackage.Load(args[index]);
        var rewritten = NativePackage.Load(args[index + 1]);
        if (original != rewritten)
            throw new InvalidOperationException($"Desktop native document differs after web round-trip: {Path.GetFileName(args[index])}.");
    }
    Console.WriteLine($"Desktop native round-trip equality passed for {(args.Length - 1) / 2} fixtures.");
    return;
}

if (args is not [var operation, var path] || operation is not ("export" or "check"))
    throw new ArgumentException("Use export or check followed by a fixture directory.");

Directory.CreateDirectory(path);
var xlsx = Path.Combine(path, "DesktopStyles.xlsx");
var native = Path.Combine(path, "DesktopStyles.exlt");
var style = CellStyle.Default with
{
    FillRgb = "FFFFFF",
    BorderDetails = new CellBorders(
        Top: new CellBorder("slantDashDot", "112233"),
        Right: new CellBorder("double"),
        Bottom: new CellBorder("mediumDashDotDot", "445566"),
        DiagonalUp: new CellBorder("dotted", "778899"),
        DiagonalDown: new CellBorder("dotted", "778899")),
    TextRotation = -45,
    Alignment = "distributed",
    Indent = 3,
    ReadingOrder = "rtl",
    JustifyLastLine = true,
};
var a = style with { FillDetails = new CellFill(GradientRgb: "2288AA", Shading: "diagonalDown", Variant: 3) };
var b = style with
{
    TextRotation = 255,
    FillDetails = new CellFill(GradientRgb: "2288AA", Shading: "center", Variant: 2),
};
var sheet = new SheetDocument(
    Guid.Parse("b71420c7-ad56-4000-8000-000000000001"), "Styles",
    [
        new CellDocument(Guid.Parse("b71420c7-ad56-4000-8000-000000000002"), 1, 1, new CellContent.Number(60), a with { NumberFormat = "yyyy-mm-dd" }, null),
        new CellDocument(Guid.Parse("b71420c7-ad56-4000-8000-000000000003"), 2, 1, new CellContent.Text("stacked"), b, null),
    ]).WithFrozenPanes(new FrozenPaneSplit(1, 0));
var expected = new WorkbookDocument(Guid.Parse("b71420c7-ad56-4000-8000-000000000004"), 1, "1900", [sheet]);

if (operation == "export")
{
    NativePackage.Save(expected, native);
    XlsxExchange.Export(expected, xlsx);
}
else
{
    var rewritten = NativePackage.Load(native);
    if (rewritten != expected)
        throw new InvalidOperationException("Native workbook differs from the original.");
    var imported = XlsxExchange.Import(xlsx);
    var cells = imported.Document.Sheets.Single().Cells.OrderBy(cell => cell.Row).ToArray();
    if (cells.Length != 2 || cells[0].Style != sheet.Cells[0].Style ||
        cells[1].Style != sheet.Cells[1].Style || cells[0].Content is not CellContent.Number { Value: 60 } ||
        cells[1].Content is not CellContent.Text { Value: "stacked" })
        throw new InvalidOperationException("XLSX values or styles differ after desktop reimport: " +
            string.Join(" / ", cells.Select(cell => cell.Style)));
    Console.WriteLine("Desktop native equality and XLSX style reimport passed.");
}
