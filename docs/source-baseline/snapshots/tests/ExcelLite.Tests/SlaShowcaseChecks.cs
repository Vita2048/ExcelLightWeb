using System.Globalization;
using ExcelLite.Application;
using ExcelLite.Core;
using ExcelLite.Foundation;
using ExcelLite.Persistence;
using ExcelLite.Python;

static class SlaShowcaseChecks
{
    internal static async Task RunAsync(Action<bool, string> check, string? output = null, bool useArrowInputs = false)
    {
        var root = new DirectoryInfo(AppContext.BaseDirectory);
        while (!Directory.Exists(Path.Combine(root.FullName, "docs", "showcase", "sla")))
            root = root.Parent ?? throw new DirectoryNotFoundException("SLA showcase source folder");
        var source = Path.Combine(root.FullName, "docs", "showcase", "sla");
        (int Row, int Column)[] sources = [(6, 2), (SlaShowcase.TrendRow, 2), (SlaShowcase.HeatmapRow, 2)];
        var tickets = SlaShowcase.GenerateTickets();
        var fixture = SlaShowcase.Create(source, tickets);
        var runtime = PythonRuntime.Locate();
        using var worker = new PythonWorkerClient(runtime, useArrowInputs: useArrowInputs);

        SheetDocument Sheet(WorkbookDocument document, string name) => document.Sheets.Single(s => s.Name == name);
        CellDocument Cell(WorkbookDocument document, string sheet, int row, int col) =>
            Sheet(document, sheet).Cells.Single(c => c.Row == row && c.Column == col);
        PythonBlock Block(WorkbookDocument document, int row, int column) =>
            ((CellContent.Python)Cell(document, "Analysis", row, column).Content).Block;
        double? Number(WorkbookDocument document, string sheet, int row, int col)
        {
            var cell = Cell(document, sheet, row, col);
            if (cell.Content is CellContent.Number n) return n.Value;
            if (cell.Cache is { Kind: CellCache.NumberKind } cache)
                return double.Parse(cache.Text, CultureInfo.InvariantCulture);
            return null;
        }
        static bool Near(double? actual, double? expected) =>
            actual is double a && expected is double e ? Math.Abs(a - e) < 1e-9 : actual == expected;
        static double? P90(IEnumerable<double?> values)
        {
            var sorted = values.OfType<double>().Order().ToArray();
            if (sorted.Length == 0) return null;
            var position = .9 * (sorted.Length - 1);
            var lo = (int)position;
            return sorted[lo] + (position - lo) * (sorted[(int)Math.Ceiling(position)] - sorted[lo]);
        }
        static bool? Breached(SlaShowcase.Ticket ticket, double[] targets)
        {
            var limit = targets[ticket.Priority];
            if (ticket.Reply is double reply) return reply > limit;
            return (SlaShowcase.Snapshot - ticket.Opened).TotalHours > limit ? true : null;
        }
        static double? Rate(IEnumerable<bool?> values)
        {
            var assessed = values.OfType<bool>().ToArray();
            return assessed.Length == 0 ? null : (double)assessed.Count(x => x) / assessed.Length;
        }

        async Task<PythonScalarOutcome> Execute(CellReference _, string script, IReadOnlyDictionary<string, object?> input,
            CancellationToken cancellationToken)
        {
            var response = await worker.RunAsync(script, input, TimeSpan.FromSeconds(30), cancellationToken);
            return new(response.Status, Error: response.Error, Table: response.Table, Figure: response.Figure);
        }
        Task<WorkbookDocument> Calculate(WorkbookDocument document, HashSet<CellReference> cells) =>
            MixedPythonCalculation.CalculateAsync(document, cells, false, runtime.Fingerprint, Execute);

        void Verify(WorkbookDocument document, double[] targets, string label)
        {
            foreach (var (r, c) in sources)
            {
                var block = Block(document, r, c);
                check(PythonCells.IsCurrent(document, Sheet(document, "Analysis").Id, block),
                    $"SLA {label} Analysis!{CellRectangle.ColumnName(c)}{r} is current ({block.Error ?? block.Status})");
            }
            var outcomes = tickets.Select(t => Breached(t, targets)).ToArray();
            double?[] expected =
                [tickets.Length, Rate(outcomes), tickets.Count(t => t.Resolution is null), P90(tickets.Select(t => t.Resolution))];
            check(expected.Select((n, i) => Near(Number(document, "Dashboard", 6, 2 + i * 3), n)).All(x => x),
                $"SLA {label} four linked Dashboard KPIs match the source records");
            check(Number(document, "Analysis", 11, 3) == outcomes.Count(x => x is null)
                  && Number(document, "Analysis", 12, 3) == outcomes.Count(x => x is not null),
                $"SLA {label} pending and assessed reconcile to opened");

            for (var month = 1; month <= 6; month++)
            {
                var cohort = tickets.Where(t => t.Opened.Month == month).ToArray();
                var assessed = cohort.Select(t => Breached(t, targets)).ToArray();
                check(Number(document, "Analysis", 16 + month, 3) == cohort.Length
                      && Number(document, "Analysis", 16 + month, 4) == assessed.Count(x => x is not null)
                      && Near(Number(document, "Analysis", 16 + month, 5), Rate(assessed)),
                    $"SLA {label} month {month} volume and rate reconcile");
                for (var team = 0; team < SlaShowcase.Teams.Length; team++)
                {
                    var group = cohort.Where(t => t.Team == team).ToArray();
                    if (!Near(Number(document, "Analysis", 27 + team, 2 + month), Rate(group.Select(t => Breached(t, targets))))
                        || !Near(Number(document, "Analysis", 36 + team, 2 + month), P90(group.Select(t => t.Resolution))))
                        throw new InvalidOperationException($"SLA {label} pivot mismatch for {SlaShowcase.Teams[team]}, month {month}.");
                }
            }
            check(true, $"SLA {label} both 5 × 6 pivots match independent group calculations");
            for (var priority = 0; priority < targets.Length; priority++)
            {
                var group = tickets.Where(t => t.Priority == priority).Select(t => Breached(t, targets)).ToArray();
                var row = 45 + priority;
                check(Number(document, "Analysis", row, 3) == group.Length
                      && Number(document, "Analysis", row, 4) == group.Count(x => x is not null)
                      && Number(document, "Analysis", row, 5) == group.Count(x => x == true)
                      && Near(Number(document, "Analysis", row, 6), Rate(group)),
                    $"SLA {label} P{priority + 1} diagnostic counts and rate reconcile");
            }
            var placements = PythonFigures.Placements(document).ToArray();
            check(placements.Length == 2 && placements.All(p => p.Placement.SheetId == Sheet(document, "Dashboard").Id
                && p.Block.LastFigure is { Width: >= 750, Height: >= 450 }),
                $"SLA {label} both native chart placements have rendered PNGs");
        }

        var baseline = await Calculate(fixture, new MixedPythonCalculation.Graph(fixture).Python);
        Verify(baseline, SlaShowcase.Targets, "baseline");
        check(Number(baseline, "Analysis", 29, 7) > Number(baseline, "Analysis", 29, 6) + .20
              && Number(baseline, "Analysis", 20, 3) > Number(baseline, "Analysis", 17, 3) * 1.5,
            "SLA fixture has a May Integrations hotspot and an April volume surge");

        var scenarioTimer = System.Diagnostics.Stopwatch.StartNew();
        var book = new WorkbookSession(baseline) { DeferCalculation = true };
        var target = new CellReference(Sheet(baseline, "Dashboard").Id, SlaShowcase.TargetRow, SlaShowcase.TargetColumn);
        book.Execute(new SetCellsCommand([new(target, new CellContent.Number(2))]));
        var affected = new MixedPythonCalculation.Graph(book.Active).Affected(baseline, book.Active);
        check(affected.Count == 3, "SLA target edit selects the spill and both downstream charts");
        var scenario = await Calculate(book.Active, affected);
        if (Environment.GetCommandLineArgs().Contains("--sla-showcase-only"))
            Console.WriteLine($"Warm SLA P2 scenario through both chart results ({(useArrowInputs ? "Arrow inputs" : "JSON inputs")}): {scenarioTimer.Elapsed.TotalMilliseconds:F1} ms");
        Verify(scenario, [1, 2, 12, 24], "P2 at two hours");
        check(Number(scenario, "Dashboard", 6, 5) > Number(baseline, "Dashboard", 6, 5)
              && Number(scenario, "Analysis", 11, 3) < Number(baseline, "Analysis", 11, 3)
              && Block(scenario, SlaShowcase.TrendRow, 2).LastFigure != Block(baseline, SlaShowcase.TrendRow, 2).LastFigure
              && Block(scenario, SlaShowcase.HeatmapRow, 2).LastFigure != Block(baseline, SlaShowcase.HeatmapRow, 2).LastFigure,
            "SLA tighter P2 target turns pending tickets into breaches and redraws both charts");

        book.Execute(new SetCellsCommand([new(target, new CellContent.Number(0))]));
        var invalid = await Calculate(book.Active, affected);
        check(Block(invalid, 6, 2).Error?.Contains("positive number of hours", StringComparison.Ordinal) == true
              && Block(invalid, SlaShowcase.TrendRow, 2).LastFigure == Block(baseline, SlaShowcase.TrendRow, 2).LastFigure,
            "SLA invalid targets fail visibly and retain the last successful figure");

        var path = Path.Combine(Path.GetTempPath(), $"sla-showcase-{Guid.NewGuid():N}.exlt");
        try
        {
            NativePackage.Save(baseline, path);
            var reopened = NativePackage.Load(path);
            check(reopened == baseline, "SLA native save and reopen preserve inputs, results, styles and figures");
            var requests = worker.SubmittedRequests;
            await Calculate(reopened, new MixedPythonCalculation.Graph(reopened).Python);
            check(worker.SubmittedRequests == requests, "SLA cached baseline reopens without a calculation request");
            if (output is not null)
                NativePackage.Save(reopened, Path.GetFullPath(output));
            else
            {
                var delivered = NativePackage.Load(Path.Combine(source, "Support-SLA.exlt"));
                Verify(delivered, SlaShowcase.Targets, "delivered workbook");
                check(Sheet(delivered, "Tickets").Cells.Select(c => (c.Row, c.Column, c.Content, c.Style))
                    .SequenceEqual(Sheet(baseline, "Tickets").Cells.Select(c => (c.Row, c.Column, c.Content, c.Style))),
                    "SLA delivered Tickets preserves the generated inputs and formatting");
                static IEnumerable<(int Row, int Column, CellContent Content, CellStyle? Style)> Targets(WorkbookDocument document)
                {
                    var area = WorkbookNames.Rectangle(document.Names.Single(name => name.Name == "ResponseTargets" && name.IsTable));
                    return document.Sheets.Single(sheet => sheet.Id == area.SheetId).Cells
                        .Where(cell => area.Contains(cell.Row, cell.Column))
                        .Select(cell => (cell.Row, cell.Column, cell.Content, cell.Style));
                }
                check(Targets(delivered).SequenceEqual(Targets(baseline)) &&
                    Number(delivered, "Dashboard", 4, 16) == Number(baseline, "Dashboard", 4, 16),
                    "SLA delivered Dashboard preserves the generated scenario inputs and target table formatting");
                check(sources.All(at =>
                    Block(delivered, at.Row, at.Column).Source == Block(baseline, at.Row, at.Column).Source),
                    "SLA delivered workbook embeds the checked-in Python source");
            }
        }
        finally { if (File.Exists(path)) File.Delete(path); }
        Console.WriteLine(FormattableString.Invariant(
            $"SLA baseline {tickets.Length} tickets, {Number(baseline, "Dashboard", 6, 5):P2} breached, {Number(baseline, "Dashboard", 6, 8)} open, {Number(baseline, "Dashboard", 6, 11):F2} h P90. P2 scenario {Number(scenario, "Dashboard", 6, 5):P2} breached."));
    }
}
