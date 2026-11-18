var passed = 0;
void Check(bool ok, string scenario)
{
    if (!ok) throw new InvalidOperationException(scenario);
    passed++;
}
FontFillChecks.Run(Check);
FormatPainterChecks.Run(Check);
Console.WriteLine($"Passed {passed} desktop Font/Fill and Format Painter assertions.");
