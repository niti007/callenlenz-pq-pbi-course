// =====================================================================================
// LAB 5 - ASSET COST ROLL-FORWARD FOR ANY PERIOD (all asset classes together)
// Needs parameters:  FolderPath (Text)   Period (Decimal/Whole number, e.g. 10 then 11)
// Month 11 appears when you copy  _Later\Asset_Movements_2026-11.xlsx  into 05_Budget_Assets
// and set Period = 11 (or just refresh if Period is already 11).
// Layout of output matches the client's sheet 2:
//   Cost Bfwd = opening cost + all movements in earlier periods
//   Additions (+), Disposals (-), Write offs (-)  -> shown as signed numbers
//   Cost Cfwd = Bfwd + Additions + Disposals + Write offs
// =====================================================================================
let
    AssetFolder = Text.TrimEnd(FolderPath, "\") & "\05_Budget_Assets\",
    Source = Folder.Files(AssetFolder),
    #"Filtered Asset Files" = Table.SelectRows(
        Source,
        each [Folder Path] = AssetFolder
            and [Extension] = ".xlsx"
            and Text.StartsWith([Name], "Asset_Movements_2026")
            and not Text.StartsWith([Name], "~$")
    ),

    // Turn anything number-ish into a number: null/blank/"-" -> 0 ; "(25,000.00)" -> -25000 ; "£1,500.00" -> 1500
    fnNum = (x as any) as number =>
        if x = null then 0
        else if x is number then x
        else
            let
                s = Text.Trim(Text.From(x)),
                IsNegative = (Text.StartsWith(s, "(") and Text.EndsWith(s, ")")) or Text.StartsWith(s, "-"),
                Digits = Text.Select(s, {"0".."9", "."})
            in
                if Digits = "" then 0 else (if IsNegative then -1 else 1) * Number.FromText(Digits, "en-GB"),

    // Read the 'Movements' sheet of one file (3 title rows above the header)
    fnReadMovements = (content as binary) as table =>
        let
            Workbook = Excel.Workbook(content, false, true),
            Sheet = Workbook{[Item = "Movements", Kind = "Sheet"]}[Data],
            #"Skipped Title Rows" = Table.Skip(Sheet, 3),
            #"Promoted Headers" = Table.PromoteHeaders(#"Skipped Title Rows", [PromoteAllScalars = true]),
            #"Kept Columns" = Table.SelectColumns(#"Promoted Headers", {"Asset Class", "Period", "Additions", "Disposals", "Write Offs"})
        in
            #"Kept Columns",

    #"Added Movements" = Table.AddColumn(#"Filtered Asset Files", "Data", each fnReadMovements([Content])),
    Combined = Table.Combine(#"Added Movements"[Data]),

    // Asset Class only appears on the first row of each block (merged-cell style) -> Fill Down
    #"Filled Down Asset Class" = Table.FillDown(Combined, {"Asset Class"}),
    // Subtotal / grand total / blank rows have no Period
    #"Removed Total Rows" = Table.SelectRows(#"Filled Down Asset Class", each [Period] <> null),
    #"Cleaned Numbers" = Table.TransformColumns(
        #"Removed Total Rows",
        {
            {"Period", each Number.From(_), Int64.Type},
            {"Additions", each Number.Abs(fnNum(_)), type number},
            {"Disposals", each Number.Abs(fnNum(_)), type number},
            {"Write Offs", each Number.Abs(fnNum(_)), type number}
        }
    ),

    // Opening cost at 1 Jan 2026 (sheet in the main file only; 2 rows above the header)
    MainFile = Table.SelectRows(#"Filtered Asset Files", each [Name] = "Asset_Movements_2026.xlsx"){0}[Content],
    OpeningRaw = Excel.Workbook(MainFile, false, true){[Item = "Opening_Cost", Kind = "Sheet"]}[Data],
    OpeningTable = Table.PromoteHeaders(Table.Skip(OpeningRaw, 2), [PromoteAllScalars = true]),
    OpeningTotal = List.Sum(List.Transform(Table.Column(OpeningTable, "Opening Cost at 1 Jan 2026"), each fnNum(_))),

    // The roll-forward for the chosen Period
    EarlierMovements = Table.SelectRows(#"Cleaned Numbers", each [Period] < Period),
    ThisPeriod = Table.SelectRows(#"Cleaned Numbers", each [Period] = Period),
    CostBfwd = OpeningTotal + List.Sum(EarlierMovements[Additions]) - List.Sum(EarlierMovements[Disposals]) - List.Sum(EarlierMovements[Write Offs]),
    Additions = List.Sum(ThisPeriod[Additions]),
    Disposals = - List.Sum(ThisPeriod[Disposals]),
    WriteOffs = - List.Sum(ThisPeriod[Write Offs]),
    CostCfwd = CostBfwd + Additions + Disposals + WriteOffs,

    #"Roll Forward" = Table.FromRows(
        {
            {"Cost Bfwd", Number.Round(CostBfwd, 2)},
            {"Additions", Number.Round(Additions, 2)},
            {"Disposals", Number.Round(Disposals, 2)},
            {"Write offs", Number.Round(WriteOffs, 2)},
            {"Cost Cfwd", Number.Round(CostCfwd, 2)}
        },
        {"Line", "Month " & Text.From(Period)}
    )
in
    #"Roll Forward"
