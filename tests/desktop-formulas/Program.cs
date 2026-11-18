using System.Collections.Immutable;
using System.Globalization;
using System.Text.Json;
using ExcelLite.Calculation;
using ExcelLite.Core;

CultureInfo.CurrentCulture = CultureInfo.GetCultureInfo("cs-CZ");
using var input = JsonDocument.Parse(File.ReadAllText(args[0]));
var count = 0;
foreach (var fixture in input.RootElement.EnumerateArray())
{
    var id = fixture.GetProperty("id").GetString();
    var main = Guid.NewGuid();
    var other = Guid.NewGuid();
    ImmutableArray<CellDocument> ReadCells(string key) =>
        fixture.GetProperty(key).EnumerateArray().Select(cell =>
        {
            var value = cell[2];
            CellContent content = value.ValueKind switch
            {
                JsonValueKind.Number => new CellContent.Number(value.GetDouble()),
                JsonValueKind.True or JsonValueKind.False => new CellContent.Boolean(value.GetBoolean()),
                _ when value.GetString()!.StartsWith('=') => new CellContent.Formula(value.GetString()!),
                _ => new CellContent.Text(value.GetString()!)
            };
            return new CellDocument(Guid.NewGuid(), cell[0].GetInt32(), cell[1].GetInt32(), content, null, null);
        }).ToImmutableArray();

    var sheet = new SheetDocument(main, "Sheet1", ReadCells("cells"));
    if (id == "subtotal-hidden")
    {
        sheet = sheet with { Rows = [new(4, Hidden: true)], Filter = new(1, 1, 4, 2, [new(1, ["a"])]) };
    }
    var document = new WorkbookDocument(Guid.NewGuid(), 1, "1900", [sheet, new(other, "O'Brien data", ReadCells("other"))]);
    foreach (var test in fixture.GetProperty("cases").EnumerateArray())
    {
        var expression = test.GetProperty("expression").GetString()!;
        var formula = new CellDocument(Guid.NewGuid(), 1, 20, new CellContent.Formula(expression), null, null);
        var withFormula = document with { Sheets = document.Sheets.SetItem(0, sheet with { Cells = sheet.Cells.Add(formula) }) };
        var cache = new NativeCalculator(withFormula).Calculate().Sheets[0].Cells.Last().Cache!;
        var kind = test.GetProperty("kind").GetString();
        var value = test.GetProperty("value").GetString();
        if (cache.Kind != kind || cache.Text != value)
        {
            throw new InvalidOperationException($"{id} {expression}: expected {kind} {value}, got {cache.Kind} {cache.Text}");
        }
        count++;
    }
}
Console.WriteLine($"Verified {count} cases with the desktop NativeCalculator (cs-CZ).");
