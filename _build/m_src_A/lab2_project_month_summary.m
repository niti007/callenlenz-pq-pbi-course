// =====================================================================================
// LAB 2 - PROJECT-MONTH SUMMARY: Group By project and month
// Needs the query  LabourLines  (lab2_timesheets.m). Create this one with
// Home > New Source > Blank Query, or right-click LabourLines > Reference, then paste.
// Output: one row per Project Code x Month with Hours, OT Hours and Labour Cost (GBP, 2 dp).
// =====================================================================================
let
    Source = LabourLines,
    #"Grouped Project Month" = Table.Group(
        Source,
        {"Project Code", "Month"},
        {
            {"Hours", each List.Sum([Hours]), type number},
            {"OT Hours", each List.Sum([OT Hours]), type number},
            {"Labour Cost", each Number.Round(List.Sum([Labour Cost]), 2), type number}
        }
    ),
    #"Sorted Rows" = Table.Sort(#"Grouped Project Month", {{"Project Code", Order.Ascending}, {"Month", Order.Ascending}})
in
    #"Sorted Rows"
