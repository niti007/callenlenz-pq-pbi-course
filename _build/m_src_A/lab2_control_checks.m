// =====================================================================================
// LAB 2 - CONTROL CHECKS: loaded hours per month vs the timesheet system's control totals
// -------------------------------------------------------------------------------------
// Needs:  parameter FolderPath (Text), and a query called Timesheets_Deduped.
// Make Timesheets_Deduped like this: right-click the Lab 2 query (LabourLines) > Duplicate, rename it
// Timesheets_Deduped, then delete every step AFTER  #"Removed Duplicates"  (so it still
// contains the unknown-code rows - that is the point of the check).
// Expected result: every month OK except Feb, which is 7.5 hours OVER the control total;
// the 'Hours on Unknown Codes' column explains the whole difference (6 rows).
// =====================================================================================
let
    Source = Timesheets_Deduped,
    ReferenceFolder = Text.TrimEnd(FolderPath, "\") & "\00_Reference\",
    ControlFile = Text.TrimEnd(FolderPath, "\") & "\02_Timesheets\Control_Totals.xlsx",
    Projects = Excel.Workbook(File.Contents(ReferenceFolder & "Projects.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],
    ValidCodes = Projects[Project Code],

    #"Added Month" = Table.AddColumn(Source, "Month", each Date.StartOfMonth([Week Ending]), type date),
    #"Added Valid Flag" = Table.AddColumn(#"Added Month", "Is Valid Code", each List.Contains(ValidCodes, [Project Code]), type logical),

    // Loaded hours, lines, and what part of them sits on unknown codes
    #"Loaded Totals" = Table.Group(
        #"Added Valid Flag",
        {"Month"},
        {
            {"Loaded Hours", each List.Sum([Hours]), type number},
            {"Loaded Lines", each Table.RowCount(_), Int64.Type},
            {"Hours on Unknown Codes", each List.Sum(List.Combine({{0}, Table.SelectRows(_, each not [Is Valid Code])[Hours]})), type number},
            {"Unknown Code Lines", each Table.RowCount(Table.SelectRows(_, each not [Is Valid Code])), Int64.Type}
        }
    ),

    // Control totals sheet: 3 rows above the header (title, source note, blank line)
    ControlRaw = Excel.Workbook(File.Contents(ControlFile), false, true){[Item = "Control", Kind = "Sheet"]}[Data],
    #"Skipped Title Rows" = Table.Skip(ControlRaw, 3),
    #"Promoted Headers" = Table.PromoteHeaders(#"Skipped Title Rows", [PromoteAllScalars = true]),
    #"Typed Control" = Table.TransformColumnTypes(
        #"Promoted Headers", {{"Month", type date}, {"System Hours", type number}, {"System Lines", Int64.Type}}
    ),
    #"Control Month Start" = Table.TransformColumns(#"Typed Control", {{"Month", each Date.StartOfMonth(_), type date}}),

    Merged = Table.NestedJoin(#"Loaded Totals", {"Month"}, #"Control Month Start", {"Month"}, "Control", JoinKind.LeftOuter),
    Expanded = Table.ExpandTableColumn(Merged, "Control", {"System Hours", "System Lines"}, {"System Hours", "System Lines"}),

    #"Added Difference" = Table.AddColumn(Expanded, "Hours Difference", each [Loaded Hours] - [System Hours], type number),
    #"Added Line Difference" = Table.AddColumn(#"Added Difference", "Line Difference", each [Loaded Lines] - [System Lines], Int64.Type),
    #"Added Status" = Table.AddColumn(
        #"Added Line Difference", "Status", each if Number.Abs([Hours Difference]) < 0.001 then "OK" else "CHECK", type text
    ),
    // 0 means the unknown-code rows fully explain the difference
    #"Added Unexplained" = Table.AddColumn(
        #"Added Status", "Unexplained Hours", each Number.Round([Hours Difference] - [Hours on Unknown Codes], 3), type number
    ),
    #"Sorted Rows" = Table.Sort(#"Added Unexplained", {{"Month", Order.Ascending}})
in
    #"Sorted Rows"
