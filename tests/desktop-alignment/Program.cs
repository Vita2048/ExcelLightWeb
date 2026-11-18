var passed = 0;
void Check(bool ok, string scenario)
{
    if (!ok) throw new InvalidOperationException(scenario);
    passed++;
}
AlignmentChecks.Run(Check);
Console.WriteLine($"Passed {passed} desktop alignment and merge assertions.");
