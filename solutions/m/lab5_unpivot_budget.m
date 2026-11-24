// =====================================================================================
// LAB 5 - UNPIVOT THE WIDE BUDGET (Jan-26 ... Dec-26 columns -> one row per project-month)
// Needs parameter FolderPath (Text).
// =====================================================================================
let
    BudgetFile = Text.TrimEnd(FolderPath, "\") & "\05_Budget_Assets\Project_Budget_2026.xlsx",
    Source = Excel.Workbook(File.Contents(BudgetFile), false, true),
    Sheet = Source{[Item = "Budget", Kind = "Sheet"]}[Data],

    // 3 rows above the header: title, subtitle, blank
    #"Removed Title Rows" = Table.Skip(Sheet, 3),
    #"Promoted Headers" = Table.PromoteHeaders(#"Removed Title Rows", [PromoteAllScalars = true]),

    // Blank spacer column (header becomes Column3) and the Total column must go BEFORE unpivoting
    #"Removed Blank and Total Columns" = Table.RemoveColumns(#"Promoted Headers", {"Column3", "Total"}, MissingField.Ignore),

    // Clean the codes, then keep only real project rows (drops the Total row and the footnote)
    #"Cleaned Project Code" = Table.TransformColumns(
        #"Removed Blank and Total Columns",
        {{"Project Code", each if _ = null then "" else Text.Upper(Text.Trim(Text.From(_))), type text}}
    ),
    #"Removed Total and Footnote Rows" = Table.SelectRows(#"Cleaned Project Code", each Text.StartsWith([Project Code], "CL-P")),

    // Unpivot the 12 month columns
    #"Unpivoted Months" = Table.UnpivotOtherColumns(
        #"Removed Total and Footnote Rows", {"Project Code", "Project Name"}, "Month Label", "Budget Text"
    ),

    // "Jan-26" (text) -> 01/01/2026 (date)
    #"Added Month" = Table.AddColumn(
        #"Unpivoted Months",
        "Month",
        each Date.FromText("1 " & Text.Start([Month Label], 3) & " 20" & Text.End(Text.Trim([Month Label]), 2), "en-GB"),
        type date
    ),
    // One budget cell is text ("45,000.00"); blanks become 0
    #"Added Budget GBP" = Table.AddColumn(
        #"Added Month",
        "Budget GBP",
        each
            let
                s = if [Budget Text] = null then "" else Text.Select(Text.From([Budget Text]), {"0".."9", "."})
            in
                if s = "" then 0 else Number.FromText(s, "en-GB"),
        type number
    ),
    #"Removed Helper Columns" = Table.RemoveColumns(#"Added Budget GBP", {"Month Label", "Budget Text"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Helper Columns", {"Project Code", "Project Name", "Month", "Budget GBP"}),
    #"Sorted Rows" = Table.Sort(#"Reordered Columns", {{"Project Code", Order.Ascending}, {"Month", Order.Ascending}})
in
    #"Sorted Rows"
