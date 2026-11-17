using System.Collections.Immutable;
namespace ExcelLite.Core;

public static class FormulaSubset
{
    public static readonly ImmutableArray<string> Functions = [
        "SUM", "AVERAGE", "MIN", "MAX", "COUNT", "COUNTA",
        "IF", "AND", "OR", "NOT", "IFERROR",
        "ABS", "ROUND", "ROUNDUP", "ROUNDDOWN",
        "SUMIF", "COUNTIF",
        "LEFT", "RIGHT", "MID", "LEN", "TRIM", "UPPER", "LOWER",
        "DATE", "YEAR", "MONTH", "DAY", "TODAY",
        "VLOOKUP", "SUBTOTAL"
    ];
}
