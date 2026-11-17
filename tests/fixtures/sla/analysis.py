import numpy as np
import pandas as pd

months = pd.period_range("2026-01", periods=6, freq="M").astype(str).tolist()
teams = ["Access", "Billing", "Integrations", "Platform", "Workspace"]
priorities = ["P1", "P2", "P3", "P4"]
as_of = pd.to_datetime(xl("Dashboard!P4"), utc=True, format="ISO8601")
targets = xl("ResponseTargets[#Data]", headers=True)
tickets = xl("SupportTickets[#Data]", headers=True).dropna(how="all").copy()

if pd.isna(as_of):
    raise ValueError("Enter a snapshot timestamp on SLA Targets.")
if sorted(targets["Priority"].dropna().tolist()) != priorities or len(targets) != len(priorities):
    raise ValueError("Set one target per priority, P1 through P4.")
target_hours = pd.to_numeric(targets["Target hrs"], errors="raise")
if not np.isfinite(target_hours).all() or not target_hours.gt(0).all():
    raise ValueError("Every response target must be a positive number of hours.")

required = ["Ticket ID", "Opened UTC", "Team", "Priority", "Issue", "Segment", "Channel", "Status"]
if tickets[required].isna().any().any() or tickets["Ticket ID"].duplicated().any():
    raise ValueError("Fill every required ticket field and use unique ticket IDs.")
if not (tickets["Team"].isin(teams) & tickets["Priority"].isin(priorities)
        & tickets["Status"].isin(["Open", "Resolved"])).all():
    raise ValueError("Check the team, priority and status on each ticket.")

opened = pd.to_datetime(tickets["Opened UTC"], utc=True, format="ISO8601")
for column in ["Reply hrs", "Resolve hrs"]:
    duration = pd.to_numeric(tickets[column], errors="raise")
    if (duration.notna() & (~np.isfinite(duration) | duration.lt(0))).any():
        raise ValueError(f"{column} must be blank or a finite, nonnegative duration.")
    tickets[column] = duration
resolved = tickets["Status"].eq("Resolved")
if (resolved != tickets["Resolve hrs"].notna()).any() or (
    resolved & (tickets["Reply hrs"].isna() | tickets["Resolve hrs"].lt(tickets["Reply hrs"]))
).any():
    raise ValueError("Resolved tickets need reply and resolution hours in chronological order.")

tickets["Month"] = opened.dt.strftime("%Y-%m")
tickets["Age"] = (as_of - opened).dt.total_seconds() / 3600
cohort = tickets[tickets["Month"].isin(months) & tickets["Age"].ge(0)].copy()
reply = cohort["Reply hrs"].where(cohort["Reply hrs"].le(cohort["Age"]))
resolution = cohort["Resolve hrs"].where(cohort["Resolve hrs"].le(cohort["Age"]))
target = cohort["Priority"].map(dict(zip(targets["Priority"], target_hours)))
decided = reply.notna() | cohort["Age"].gt(target)
cohort["Breach"] = reply.fillna(np.inf).gt(target).astype(float).where(decided)
cohort["Resolution"] = resolution

monthly = cohort.groupby("Month").agg(
    Tickets=("Ticket ID", "size"), Assessed=("Breach", "count"), Rate=("Breach", "mean"),
).reindex(months)

def pivot(metric, aggregate):
    return cohort.pivot_table(
        index="Team", columns="Month", values=metric, aggfunc=aggregate,
    ).reindex(index=teams, columns=months).reset_index()

def p90(values):
    return values.quantile(0.9)

priority = cohort.groupby("Priority").agg(
    Tickets=("Ticket ID", "size"), Assessed=("Breach", "count"),
    Missed=("Breach", "sum"), Rate=("Breach", "mean"),
).reindex(priorities)

# Keep the spill addresses stable for the native KPI formulas and both chart cells.
report = [[None] * 7 for _ in range(43)]

def section(sheet_row, rows):
    for offset, values in enumerate(rows):
        report[sheet_row - 6 + offset][:len(values)] = values

section(6, [
    ["Metric", "Value"], ["Opened", len(cohort)], ["Breach rate", cohort["Breach"].mean()],
    ["Open at snapshot", int(resolution.isna().sum())], ["P90 resolve hrs", p90(resolution)],
    ["Pending response SLA", int((~decided).sum())], ["Assessed", int(decided.sum())],
])
section(15, [["MONTHLY RESPONSE SLA"]])
section(16, [["Month", "Tickets", "Assessed", "Breach rate"]] + [
    [month, 0 if pd.isna(n) else int(n), 0 if pd.isna(a) else int(a), rate]
    for month, n, a, rate in monthly.itertuples(name=None)
])
section(25, [["BREACH RATE BY TEAM / MONTH"]])
section(34, [["P90 RESOLUTION HOURS BY TEAM / MONTH"]])
for row, matrix in [(26, pivot("Breach", "mean")), (35, pivot("Resolution", p90))]:
    section(row, [["Team", *months]] + matrix.values.tolist())
section(43, [["RESPONSE SLA BY PRIORITY"]])
section(44, [["Priority", "Tickets", "Assessed", "Breached", "Breach rate"]] + [
    [label, 0 if pd.isna(n) else int(n), 0 if pd.isna(a) else int(a),
     0 if pd.isna(m) else int(m), rate]
    for label, n, a, m, rate in priority.itertuples(name=None)
])
report
