// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\CallenLenz_Course\data\   (keep the trailing backslash). No paths are hard-coded below.
// Lab 3 - Xero payable invoice lines.  Query name: Xero_Clean
// Also needs the Supplier_Name_Map.xlsx file (From/To) in 03_Valuation.
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

    // ---- 1. load; the real header is on row 5, so skip the 4 title rows ----
    Source = Excel.Workbook(File.Contents(FolderPath & "03_Valuation\Xero_Payable_Invoice_Detail.xlsx"), null, true),
    XeroSheet = Source{[Item = "Xero", Kind = "Sheet"]}[Data],
    RemovedTopRows = Table.Skip(XeroSheet, 4),
    PromotedHeaders = Table.PromoteHeaders(RemovedTopRows, [PromoteAllScalars = true]),

    // ---- 2. remove blank rows and the 'Total' footer (it has no Item Code) ----
    RemovedBlankRows = Table.SelectRows(PromotedHeaders, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    RemovedFooter = Table.SelectRows(RemovedBlankRows, each [Item Code] <> null),

    // ---- 3. clean text columns and the keys ----
    CleanedText = Table.TransformColumns(RemovedFooter, {{"Contact", CleanText, type text}, {"Source", CleanText, type text}, {"Description", CleanText, type text}}),
    CleanedReference = Table.TransformColumns(CleanedText, {{"Reference", each Text.Upper(CleanText(_)), type text}}),
    NormalisedItemCode = Table.TransformColumns(CleanedReference, {{"Item Code", NormPart, type text}}),

    // ---- 4. credit notes must be removed BEFORE we ever look for the latest price ----
    RemovedCreditNotes = Table.SelectRows(NormalisedItemCode, each [Source] <> "Payable Credit Note"),

    // ---- 5. types: dates (mixed real/text) and prices (some are '£' text) ----
    ConvertedDates = Table.TransformColumns(RemovedCreditNotes, {{"Invoice Date", ToDate, type date}}),
    ConvertedPrices = Table.TransformColumns(ConvertedDates, {{"Unit Price (ex) (GBP)", ParseMoney, type number}}),
    ChangedTypes = Table.TransformColumnTypes(ConvertedPrices, {{"Quantity", Int64.Type}}),
    RenamedPrice = Table.RenameColumns(ChangedTypes, {{"Unit Price (ex) (GBP)", "Unit Price"}}),

    // ---- 6. near-duplicate lines: same Reference + Item Code (after cleaning) -> keep first ----
    BufferedRows = Table.Buffer(RenamedPrice),
    RemovedDuplicates = Table.Distinct(BufferedRows, {"Reference", "Item Code"}),

    // ---- 7. standard supplier name via the From/To map ----
    SupplierMap = Table.PromoteHeaders(Excel.Workbook(File.Contents(FolderPath & "03_Valuation\Supplier_Name_Map.xlsx"), null, true){[Item = "Map", Kind = "Sheet"]}[Data], [PromoteAllScalars = true]),
    MergedSupplier = Table.NestedJoin(RemovedDuplicates, {"Contact"}, SupplierMap, {"From"}, "SupplierMatch", JoinKind.LeftOuter),
    ExpandedSupplier = Table.ExpandTableColumn(MergedSupplier, "SupplierMatch", {"To"}, {"Supplier Standard"}),
    Result = Table.SelectColumns(ExpandedSupplier, {"Invoice Date", "Supplier Standard", "Source", "Reference", "Item Code", "Description", "Quantity", "Unit Price"}),
    RenamedSupplier = Table.RenameColumns(Result, {{"Supplier Standard", "Contact"}})
in
    RenamedSupplier
