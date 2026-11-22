using ExcelLite.Core;
using ExcelLite.Persistence;

if (args.Length != 1) throw new ArgumentException("Pass the directory containing showcase-web.xlsx and styles-web.xlsx.");
var directory = Path.GetFullPath(args[0]);

var showcase = XlsxExchange.Import(Path.Combine(directory, "showcase-web.xlsx")).Document;
if (showcase.Sheets.Length != 3) throw new InvalidOperationException("Sheet count differs.");
if (showcase.Sheets[0].Name != "Customers" || showcase.Sheets[1].Name != "Products" || showcase.Sheets[2].Name != "Sales")
    throw new InvalidOperationException("Sheet names differ.");
var sales = showcase.Sheets[2];
var total = sales.Cells.Single(c => c.Row == 29 && c.Column == 4);
if (total.Content is not CellContent.Formula formula || formula.Source != "=SUM(D4:D28)")
    throw new InvalidOperationException("Formula source differs.");
if (total.Cache is not { Kind: CellCache.NumberKind, Text: "333" })
    throw new InvalidOperationException("Formula cache differs.");
var product = showcase.Sheets[1].Cells.Single(c => c.Row == 4 && c.Column == 3);
if (product.Style?.NumberFormat != "\\$#,##0.00") throw new InvalidOperationException("Number format differs.");
if (product.Content is not CellContent.Number { Value: 1499.99 }) throw new InvalidOperationException("Number differs.");

var styles = XlsxExchange.Import(Path.Combine(directory, "styles-web.xlsx")).Document;
var sheet = styles.Sheets[0];
var styled = sheet.Cells.Single(c => c.Row == 1 && c.Column == 1);
if (styled.Content is not CellContent.Number { Value: 60 } || styled.Style?.NumberFormat != "yyyy-mm-dd")
    throw new InvalidOperationException("Date serial and format differ.");
if (styled.Style?.TextRotation != -45 || styled.Style.Alignment != "distributed" || styled.Style.Indent != 3 ||
    styled.Style.ReadingOrder != "rtl" || !styled.Style.JustifyLastLine)
    throw new InvalidOperationException("Alignment differs.");
if (styled.Style.FillDetails?.Shading != "diagonalDown" || styled.Style.FillDetails.Variant != 3 ||
    styled.Style.FillDetails.GradientRgb != "2288AA")
    throw new InvalidOperationException("Gradient differs.");
if (styled.Style.BorderDetails?.Top?.Style != "slantDashDot" || styled.Style.BorderDetails.Right?.Style != "double" ||
    styled.Style.BorderDetails.DiagonalUp?.Style != "dotted")
    throw new InvalidOperationException("Borders differ.");
if (sheet.FrozenPanes.Rows != 1) throw new InvalidOperationException("Freeze differs.");

Console.WriteLine("Desktop ClosedXML accepted both web-exported XLSX files with matching values, formulas, caches, styles and freeze.");
