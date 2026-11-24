// =====================================================================================
// LAB 2 - TIMESHEETS: combine, clean, merge rates, overtime, group by project-month
// -------------------------------------------------------------------------------------
// BEFORE YOU PASTE: create a parameter  (Home > Manage Parameters > New Parameter)
//     FolderPath   Type = Text   Current value = the folder that holds 00_Reference, 02_Timesheets ...
//                  e.g.  C:\CallenLenz_Course\data      (no trailing backslash needed)
// Then: Home > New Source > Blank Query > Advanced Editor > paste this whole file.
//
// OVERTIME RULE (same in the check totals):
//   Weekly Hours   = sum of an employee's hours for one Week Ending (after removing unknown codes)
//   Overtime Hours = max(0, Weekly Hours - 37.5)
//   Each line gets its share of the overtime pro-rata by hours:
//        OT Hours (line) = Overtime Hours x Hours / Weekly Hours
//   Labour Cost (line)  = Hours x Standard Rate  +  OT Hours (line) x Standard Rate x (Multiplier - 1)
//   i.e. every hour is paid at standard rate, and overtime hours get the extra 0.5 premium.
//
// STAGES (useful checkpoints):  Combined -> Removed Duplicates -> Valid rows -> Final Labour Lines
// NAME THIS QUERY  LabourLines  (one row per cleaned line). The project-month Group By is a separate query:
// see lab2_project_month_summary.m, and lab5_pivot_hours.m / lab2_control_checks.m which also use it.
// =====================================================================================
let
    // ---------- 1. Where are the files? (the _Later sub-folder is deliberately ignored) ----------
    TimesheetFolder = Text.TrimEnd(FolderPath, "\") & "\02_Timesheets\",
    ReferenceFolder = Text.TrimEnd(FolderPath, "\") & "\00_Reference\",
    Source = Folder.Files(TimesheetFolder),
    #"Filtered Timesheet Files" = Table.SelectRows(
        Source,
        each [Folder Path] = TimesheetFolder
            and [Extension] = ".csv"
            and Text.StartsWith([Name], "Timesheet_")
    ),

    // ---------- 2. Small helper functions (each one is a single, teachable fix) ----------
    // Trim spaces, turn non-breaking spaces (CHAR 160) into normal spaces, null -> ""
    fnTrim = (x as any) as text =>
        if x = null then "" else Text.Trim(Text.Replace(Text.From(x), Character.FromNumber(160), " ")),
    // "Oliver   Hartley" -> "Oliver Hartley"
    fnCollapse = (x as text) as text =>
        Text.Combine(List.Select(Text.Split(x, " "), each _ <> ""), " "),
    // "cl p101" / "CLP101" / "CL-P101 " -> "CL-P101"
    fnProjectCode = (x as text) as text =>
        let
            Letters = Text.Select(Text.Upper(x), {"A".."Z", "0".."9"})
        in
            if Text.Length(Letters) >= 4 then "CL-P" & Text.End(Letters, 3) else Letters,
    // 7 -> "E007" ; "E007" stays "E007"
    fnEmployeeId = (x as text) as text =>
        let
            Upper = Text.Upper(x)
        in
            if Upper = "" then "" else if Text.StartsWith(Upper, "E") then Upper else "E" & Text.PadStart(Upper, 3, "0"),

    // ---------- 3. The 'Transform Sample File' function: loads and cleans ONE monthly file ----------
    fnLoadTimesheet = (fileContent as binary, fileName as text) as table =>
        let
            Raw = Csv.Document(fileContent, [Delimiter = ",", Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
            #"Promoted Headers" = Table.PromoteHeaders(Raw, [PromoteAllScalars = true]),
            #"Trimmed Headers" = Table.TransformColumnNames(#"Promoted Headers", Text.Trim),
            // March calls the column "Proj Code" - rename it here so every file lines up
            #"Renamed Proj Code" = Table.RenameColumns(#"Trimmed Headers", {{"Proj Code", "Project Code"}}, MissingField.Ignore),
            #"Kept Columns" = Table.SelectColumns(
                #"Renamed Proj Code",
                {"Week Ending", "Employee ID", "Employee Name", "Project Code", "Task", "Hours"}
            ),
            // Trim every column as text first (also turns nulls into "")
            #"Trimmed Text" = Table.TransformColumns(
                #"Kept Columns",
                List.Transform(Table.ColumnNames(#"Kept Columns"), each {_, fnTrim, type text})
            ),
            // Blank lines have an empty Week Ending; the footer line says "Total"
            #"Removed Blank and Total Rows" = Table.SelectRows(
                #"Trimmed Text",
                each [Week Ending] <> "" and Text.Lower([Week Ending]) <> "total"
            ),
            // April is US month/day/year text; every other file is UK day/month/year.
            // NOTE: choosing the culture from the FILE NAME is a teaching shortcut. In real life you would
            // fix the source export or detect the format from the data (e.g. a day > 12).
            Culture = if Text.Contains(fileName, "2026-04") then "en-US" else "en-GB",
            #"Changed Week Ending Type" = Table.TransformColumnTypes(
                #"Removed Blank and Total Rows", {{"Week Ending", type date}}, Culture
            ),
            #"Cleaned Fields" = Table.TransformColumns(
                #"Changed Week Ending Type",
                {
                    {"Employee ID", fnEmployeeId, type text},
                    {"Employee Name", fnCollapse, type text},
                    {"Project Code", fnProjectCode, type text},
                    {"Task", each Text.Proper(_), type text},
                    // "7.5 hrs" -> 7.5
                    {"Hours", each Number.FromText(Text.Select(_, {"0".."9", "."}), "en-GB"), type number}
                }
            )
        in
            #"Cleaned Fields",

    // ---------- 4. Combine all monthly files ----------
    #"Added Cleaned Data" = Table.AddColumn(
        #"Filtered Timesheet Files", "Data", each fnLoadTimesheet([Content], [Name])
    ),
    Combined = Table.Combine(#"Added Cleaned Data"[Data]),

    // ---------- 5. Typo year: 08/05/2062 should be 2026 ----------
    #"Fixed Year Typo" = Table.TransformColumns(
        Combined,
        {{"Week Ending", each if Date.Year(_) = 2062 then #date(2026, Date.Month(_), Date.Day(_)) else _, type date}}
    ),

    // ---------- 6. Remove exact duplicates (AFTER cleaning, so near-duplicates now match) ----------
    #"Removed Duplicates" = Table.Distinct(#"Fixed Year Typo"),

    // ---------- 7. Reference data ----------
    Projects = Excel.Workbook(File.Contents(ReferenceFolder & "Projects.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],
    Employees = Excel.Workbook(File.Contents(ReferenceFolder & "Employees.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],
    Rates = Excel.Workbook(File.Contents(ReferenceFolder & "Labour_Rates.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],

    // ---------- 8. Unknown project codes (Left Anti) - inspect this step: it is Q8 ----------
    #"Unknown Code Rows" = Table.NestedJoin(
        #"Removed Duplicates", {"Project Code"}, Projects, {"Project Code"}, "Match", JoinKind.LeftAnti
    ),
    // Valid rows only (Inner join) and bring in the project name
    #"Merged Projects" = Table.NestedJoin(
        #"Removed Duplicates", {"Project Code"}, Projects, {"Project Code"}, "Project", JoinKind.Inner
    ),
    #"Expanded Project" = Table.ExpandTableColumn(#"Merged Projects", "Project", {"Project Name"}, {"Project Name"}),

    // ---------- 9. Employee name + grade, then rate ----------
    #"Removed Typed Name" = Table.RemoveColumns(#"Expanded Project", {"Employee Name"}),
    #"Merged Employees" = Table.NestedJoin(
        #"Removed Typed Name", {"Employee ID"}, Employees, {"Employee ID"}, "Emp", JoinKind.LeftOuter
    ),
    #"Expanded Employees" = Table.ExpandTableColumn(
        #"Merged Employees", "Emp", {"Employee Name", "Grade"}, {"Employee Name", "Grade"}
    ),
    #"Merged Rates" = Table.NestedJoin(#"Expanded Employees", {"Grade"}, Rates, {"Grade"}, "Rate", JoinKind.LeftOuter),
    #"Expanded Rates" = Table.ExpandTableColumn(
        #"Merged Rates",
        "Rate",
        {"Standard Hourly Rate GBP", "Overtime Multiplier"},
        {"Standard Rate", "OT Multiplier"}
    ),

    // ---------- 10. Weekly hours per employee, merged back onto every line ----------
    #"Weekly Hours Table" = Table.Group(
        #"Expanded Rates",
        {"Employee ID", "Week Ending"},
        {{"Weekly Hours", each List.Sum([Hours]), type number}}
    ),
    #"Merged Weekly Hours" = Table.NestedJoin(
        #"Expanded Rates", {"Employee ID", "Week Ending"}, #"Weekly Hours Table", {"Employee ID", "Week Ending"}, "Weekly", JoinKind.LeftOuter
    ),
    #"Expanded Weekly Hours" = Table.ExpandTableColumn(#"Merged Weekly Hours", "Weekly", {"Weekly Hours"}, {"Weekly Hours"}),

    // ---------- 11. Overtime and labour cost per line ----------
    #"Added OT Hours" = Table.AddColumn(
        #"Expanded Weekly Hours",
        "OT Hours",
        each if [Weekly Hours] > 37.5 then ([Weekly Hours] - 37.5) * [Hours] / [Weekly Hours] else 0,
        type number
    ),
    #"Added Labour Cost" = Table.AddColumn(
        #"Added OT Hours",
        "Labour Cost",
        each [Hours] * [Standard Rate] + [OT Hours] * [Standard Rate] * ([OT Multiplier] - 1),
        type number
    ),
    #"Added Month" = Table.AddColumn(#"Added Labour Cost", "Month", each Date.StartOfMonth([Week Ending]), type date),
    // Final line-level table: one row per cleaned booking line (this is the query called LabourLines)
    #"Final Labour Lines" = Table.SelectColumns(
        #"Added Month",
        {"Week Ending", "Month", "Employee ID", "Employee Name", "Project Code", "Project Name", "Task", "Hours", "Standard Rate", "Weekly Hours", "OT Hours", "Labour Cost"}
    )
in
    #"Final Labour Lines"
