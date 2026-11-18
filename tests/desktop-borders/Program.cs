var passed = 0;
void Check(bool ok, string scenario)
{
    if (!ok) throw new InvalidOperationException(scenario);
    passed++;
}
BorderChecks.Run(Check);
BorderToolChecks.Run(Check);
Console.WriteLine($"Passed {passed} desktop border and border-tool assertions.");
