// =====================================================================================
// LAB 5 - PIVOT: hours per project (rows) by month (columns)
// Needs the query called LabourLines (= lab2_timesheets.m, which now ends at the line-level table).
// =====================================================================================
let
    Source = LabourLines,
    #"Kept Columns" = Table.SelectColumns(Source, {"Project Code", "Month", "Hours"}),
    // Month date -> "Jan", "Feb" ...
    #"Added Month Name" = Table.AddColumn(
        #"Kept Columns", "Month Name", each Date.ToText([Month], [Format = "MMM", Culture = "en-GB"]), type text
    ),
    #"Removed Month Date" = Table.RemoveColumns(#"Added Month Name", {"Month"}),
    // Pivot: new column headings come from Month Name, values are Hours added up
    // Month names found in the data, in order of appearance (Jan..Jun, more after a refresh)
    MonthNames = List.Distinct(#"Removed Month Date"[Month Name]),
    #"Pivoted Months" = Table.Pivot(#"Removed Month Date", MonthNames, "Month Name", "Hours", List.Sum),
    #"Replaced Nulls" = Table.ReplaceValue(#"Pivoted Months", null, 0, Replacer.ReplaceValue, MonthNames),
    #"Added Total" = Table.AddColumn(
        #"Replaced Nulls", "Total", each List.Sum(Record.FieldValues(Record.SelectFields(_, MonthNames))), type number
    ),
    #"Sorted Rows" = Table.Sort(#"Added Total", {{"Project Code", Order.Ascending}})
in
    #"Sorted Rows"
