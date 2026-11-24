// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\CallenLenz_Course\data\   (keep the trailing backslash). No paths are hard-coded below.
// Lab 3 - Purchase Log (48 columns).  Query name: PurchaseLog_Clean
// Also needs Supplier_Name_Map.xlsx in 03_Valuation.
let
    // ---- reusable helper functions (each one is a custom function) ----
    // CleanText: line breaks and CHAR(160) -> space, collapse repeated spaces, trim
    CleanText = (t as nullable text) as nullable text =>
        if t = null then null
        else Text.Trim(Text.Combine(List.Select(Text.Split(Text.Replace(Text.Replace(Text.Replace(t, "#(00A0)", " "), "#(lf)", " "), "#(cr)", " "), " "), each _ <> ""), " ")),
    // NormPart: any spelling of a part number -> PRT-0000 ; anything that is not a PRT part -> null
    NormPart = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                Squashed = Text.Upper(Text.Remove(Text.Replace(Text.Clean(t), "#(00A0)", " "), {" ", "-"})),
                Digits = Text.Middle(Squashed, 3)
            in
                if Text.StartsWith(Squashed, "PRT") and Digits <> "" then "PRT-" & Text.PadStart(Digits, 4, "0") else null,
    // NormProject: cl-p101 / CL P101 / CLP101 -> CL-P101
    NormProject = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                S = Text.Upper(Text.Remove(Text.Replace(Text.Clean(t), "#(00A0)", " "), {" ", "-"}))
            in
                if Text.StartsWith(S, "CLP") then "CL-" & Text.Middle(S, 2) else S,
    // ParseMoney: 1234.5 / "£1,234.50" / "(150.00)" -> number ; anything unusable -> null
    ParseMoney = (v as any) as nullable number =>
        if v = null then null
        else if Value.Is(v, type number) then v
        else
            let
                S = Text.Trim(Text.Replace(Text.Replace(Text.From(v), "£", ""), ",", "")),
                IsNeg = Text.StartsWith(S, "(") and Text.EndsWith(S, ")"),
                N = try Number.FromText(Text.Remove(S, {"(", ")"}), "en-GB") otherwise null
            in
                if N = null then null else if IsNeg then -N else N,
    // ToDate: real Excel dates pass through, UK text dates (dd/mm/yyyy or dd-Mon-yyyy) are parsed with en-GB
    ToDate = (v as any) as nullable date =>
        if v = null then null
        else if Value.Is(v, type text) then Date.FromText(Text.Trim(v), [Culture = "en-GB"])
        else Date.From(v),

    // ---- 1. load; header is on row 1 ----
    Source = Excel.Workbook(File.Contents(FolderPath & "03_Valuation\Purchase_Log.xlsx"), null, true),
    LogSheet = Source{[Item = "Purchase Log", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(LogSheet, [PromoteAllScalars = true]),

    // ---- 2. headers: line breaks -> space, trim trailing spaces ----
    CleanedHeaders = Table.TransformColumnNames(PromotedHeaders, each Text.Trim(Text.Replace(Text.Replace(_, "#(lf)", " "), "#(cr)", " "))),

    // ---- 3. the 5 empty columns have blank headers, so PQ called them Column8, Column16 ... remove them ----
    BlankColumnNames = List.Select(Table.ColumnNames(CleanedHeaders), each Text.StartsWith(_, "Column")),
    RemovedBlankColumns = Table.RemoveColumns(CleanedHeaders, BlankColumnNames),
    KeptColumns = Table.SelectColumns(RemovedBlankColumns, {"PR Number", "Date", "Project Code", "Supplier", "Part Number", "Qty", "Net Unit Price", "Notes"}),

    // ---- 4. extract PO number and Jira from Notes ("PO-55123 / Jira CL-2231 / urgent") ----
    CleanedNotes = Table.TransformColumns(KeptColumns, {{"Notes", CleanText, type text}}),
    AddedPO = Table.AddColumn(CleanedNotes, "PO Number", each Text.BeforeDelimiter([Notes], " / "), type text),
    AddedJira = Table.AddColumn(AddedPO, "Jira", each Text.BetweenDelimiters([Notes], "Jira ", " / "), type text),

    // ---- 5. cancelled lines out (any capitalisation of CANCELLED) ----
    RemovedCancelled = Table.SelectRows(AddedJira, each not Text.Contains([Notes], "CANCELLED", Comparer.OrdinalIgnoreCase)),

    // ---- 6. keys, dates, prices ----
    CleanedKeys = Table.TransformColumns(RemovedCancelled, {{"Part Number", NormPart, type text}, {"Project Code", each NormProject(_), type text}, {"Supplier", CleanText, type text}, {"PR Number", CleanText, type text}}),
    ConvertedDates = Table.TransformColumns(CleanedKeys, {{"Date", ToDate, type date}}),
    ConvertedPrices = Table.TransformColumns(ConvertedDates, {{"Net Unit Price", ParseMoney, type number}}),
    ChangedTypes = Table.TransformColumnTypes(ConvertedPrices, {{"Qty", Int64.Type}}),
    AddedNetValue = Table.AddColumn(ChangedTypes, "Net Value", each Number.Round([Qty] * [Net Unit Price], 2), type number),

    // ---- 7. standard supplier name via the From/To map ----
    SupplierMap = Table.PromoteHeaders(Excel.Workbook(File.Contents(FolderPath & "03_Valuation\Supplier_Name_Map.xlsx"), null, true){[Item = "Map", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
    MergedSupplier = Table.NestedJoin(AddedNetValue, {"Supplier"}, SupplierMap, {"From"}, "SupplierMatch", JoinKind.LeftOuter),
    ExpandedSupplier = Table.ExpandTableColumn(MergedSupplier, "SupplierMatch", {"To"}, {"Supplier Standard"}),
    Result = Table.SelectColumns(ExpandedSupplier, {"PR Number", "PO Number", "Date", "Project Code", "Supplier Standard", "Part Number", "Qty", "Net Unit Price", "Net Value", "Jira"})
in
    Result
