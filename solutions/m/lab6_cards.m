// =====================================================================================
// LAB 6 - CORPORATE CARD STATEMENTS: combine, clean, dedupe, FX, match to cardholder
// Needs parameter FolderPath (Text). Name this query  Transactions  (the function and the
// Power Automate flow use it).
// RULES
//   * Only Status = Posted counts (Pending rows are removed)
//   * The same Bank Ref can appear on two statements (month-end overlap): keep one
//   * Amount GBP = the Amount on the statement, UNLESS it is blank or 0, then
//        Original Currency Amount x rate from FX_Rates.xlsx   (GBP rate = 1.00)
//   * "CR" after an amount means a refund (negative)
//   * Cardholder = Employees.xlsx matched on Last Four (4 digits, leading zeros restored)
// =====================================================================================
let
    CardFolder = Text.TrimEnd(FolderPath, "\") & "\06_Cards\",
    ReferenceFolder = Text.TrimEnd(FolderPath, "\") & "\00_Reference\",
    Source = Folder.Files(CardFolder),
    #"Filtered Statements" = Table.SelectRows(
        Source, each [Extension] = ".csv" and Text.StartsWith([Name], "Card_Statement_")
    ),

    // The client's 15 columns, in order. Headers differ slightly by file, so we rename by POSITION
    StandardNames = {
        "Date", "Posted Date", "Status", "Merchant", "Amount", "Category", "City", "Last Four",
        "Name on Card", "MCC", "Trans Type", "Bank Ref", "Original Currency Amount",
        "Merchant Currency Description", "Merchant Country"
    },

    // ---------- helper functions ----------
    fnTrim = (x as any) as text =>
        if x = null then "" else Text.Trim(Text.Replace(Text.From(x), Character.FromNumber(160), " ")),
    fnCollapse = (x as text) as text =>
        Text.Combine(List.Select(Text.Split(x, " "), each _ <> ""), " "),
    // "45.00 CR" -> -45 ; "£1,250.00" -> 1250 ; "" -> null
    fnAmount = (x as text) as nullable number =>
        let
            Upper = Text.Upper(x),
            IsRefund = Text.EndsWith(Upper, "CR") or Text.StartsWith(Upper, "-") or (Text.StartsWith(Upper, "(") and Text.EndsWith(Upper, ")")),
            Digits = Text.Select(Upper, {"0".."9", "."})
        in
            if Digits = "" then null else (if IsRefund then -1 else 1) * Number.FromText(Digits, "en-GB"),

    // ---------- one statement: skip title lines, promote headers, rename by position ----------
    fnLoadStatement = (content as binary) as table =>
        let
            Raw = Csv.Document(content, [Delimiter = ",", Columns = 15, Encoding = 65001, QuoteStyle = QuoteStyle.Csv]),
            // March has 3 title lines above the header: find the row whose first cell is "Date"
            HeaderRow = List.PositionOf(List.Transform(Table.Column(Raw, "Column1"), each Text.Lower(fnTrim(_))), "date"),
            #"Skipped Title Rows" = Table.Skip(Raw, HeaderRow),
            #"Promoted Headers" = Table.PromoteHeaders(#"Skipped Title Rows", [PromoteAllScalars = true]),
            #"Renamed Columns" = Table.RenameColumns(
                #"Promoted Headers", List.Zip({Table.ColumnNames(#"Promoted Headers"), StandardNames})
            ),
            #"Trimmed Text" = Table.TransformColumns(
                #"Renamed Columns", List.Transform(StandardNames, each {_, fnTrim, type text})
            )
        in
            #"Trimmed Text",

    #"Added Statement" = Table.AddColumn(#"Filtered Statements", "Data", each fnLoadStatement([Content])),
    Combined = Table.Combine(#"Added Statement"[Data]),

    // ---------- remove junk ----------
    #"Removed Blank Rows" = Table.SelectRows(Combined, each [Bank Ref] <> ""),
    #"Kept Posted Only" = Table.SelectRows(#"Removed Blank Rows", each Text.Upper([Status]) = "POSTED"),
    // Bank Ref is the key: make it consistent first (some are lower case / have spaces), then dedupe
    #"Cleaned Bank Ref" = Table.TransformColumns(#"Kept Posted Only", {{"Bank Ref", each Text.Upper(fnTrim(_)), type text}}),
    #"Removed Duplicate Bank Refs" = Table.Distinct(#"Cleaned Bank Ref", {"Bank Ref"}),

    // ---------- fix the columns ----------
    #"Cleaned Columns" = Table.TransformColumns(
        #"Removed Duplicate Bank Refs",
        {
            {"Date", each Date.FromText(_, "en-GB"), type date},
            {"Posted Date", each if _ = "" then null else Date.FromText(_, "en-GB"), type nullable date},
            {"City", each Text.Proper(fnCollapse(_)), type text},
            {"Merchant", each fnCollapse(_), type text},
            // "937" -> "0937"
            {"Last Four", each Text.PadStart(_, 4, "0"), type text},
            {"Original Currency Amount", each fnAmount(_), type number},
            {"Merchant Currency Description", each Text.Upper(fnCollapse(_)), type text}
        }
    ),
    #"Added Parsed Amount" = Table.AddColumn(#"Cleaned Columns", "Statement Amount", each fnAmount([Amount]), type nullable number),

    // ---------- FX rates (title rows above the header) ----------
    FxRaw = Excel.Workbook(File.Contents(CardFolder & "FX_Rates.xlsx"), false, true){[Item = "FX_Rates", Kind = "Sheet"]}[Data],
    FxTable = Table.PromoteHeaders(Table.Skip(FxRaw, 3), [PromoteAllScalars = true]),
    FxKeyed = Table.TransformColumns(FxTable, {{"Currency Description", each Text.Upper(Text.Trim(_)), type text}}),
    #"Merged FX" = Table.NestedJoin(
        #"Added Parsed Amount", {"Merchant Currency Description"}, FxKeyed, {"Currency Description"}, "FX", JoinKind.LeftOuter
    ),
    #"Expanded FX" = Table.ExpandTableColumn(#"Merged FX", "FX", {"GBP per 1 unit"}, {"FX Rate"}),

    // blank or 0 amount -> Original Currency Amount x FX rate
    #"Added Amount GBP" = Table.AddColumn(
        #"Expanded FX",
        "Amount GBP",
        each
            if [Statement Amount] <> null and [Statement Amount] <> 0
            then [Statement Amount]
            else Number.Round([Original Currency Amount] * [FX Rate], 2),
        type number
    ),

    // ---------- match the card to a person (join on Last Four, NOT on name) ----------
    Employees = Excel.Workbook(File.Contents(ReferenceFolder & "Employees.xlsx"), true){[Item = "Sheet1", Kind = "Sheet"]}[Data],
    CardHolders = Table.SelectRows(Employees, each [Card Last Four] <> null and [Card Last Four] <> ""),
    #"Merged Cardholders" = Table.NestedJoin(
        #"Added Amount GBP", {"Last Four"}, CardHolders, {"Card Last Four"}, "Holder", JoinKind.LeftOuter
    ),
    #"Expanded Cardholders" = Table.ExpandTableColumn(
        #"Merged Cardholders", "Holder", {"Employee Name", "Department", "Email"}, {"Employee Name", "Department", "Email"}
    ),
    #"Added Month" = Table.AddColumn(#"Expanded Cardholders", "Month", each Date.StartOfMonth([Date]), type date),
    #"Selected Columns" = Table.SelectColumns(
        #"Added Month",
        {"Date", "Month", "Posted Date", "Merchant", "Category", "City", "Last Four", "Employee Name", "Department", "Email", "Bank Ref", "Merchant Currency Description", "Original Currency Amount", "Amount GBP"}
    ),
    #"Sorted Rows" = Table.Sort(#"Selected Columns", {{"Date", Order.Ascending}, {"Bank Ref", Order.Ascending}})
    // To build the CardholderSummary table for Power Automate: Reference this query ->
    //   Group By Employee Name, Email, Last Four -> Count rows + Sum of Amount GBP.
in
    #"Sorted Rows"
