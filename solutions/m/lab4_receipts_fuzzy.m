// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\CallenLenz_Course\data\   (keep the trailing backslash). No paths are hard-coded below.
// Lab 4 - Bank receipts matched to customers.  Query name: Receipts_Matched
// Deterministic answer = ExtractPayer + Fuzzy_Transform_Table. Pure fuzzy at 0.8 WITHOUT the table is approximate
// (abbreviations such as FLINTWD / DEF SYS and the typos score below 0.8) - confirm on the day.
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

    // ExtractPayer: strip 'Recd from ' / 'BACS ' (any case), then keep the text before ' - ' (removes ' - INV 2231')
    ExtractPayer = (t as nullable text) as nullable text =>
        if t = null then null
        else
            let
                C = CleanText(t),
                NoPrefix = if Text.StartsWith(C, "Recd from ", Comparer.OrdinalIgnoreCase) then Text.Middle(C, 10)
                           else if Text.StartsWith(C, "BACS ", Comparer.OrdinalIgnoreCase) then Text.Middle(C, 5)
                           else C
            in
                Text.Trim(List.First(Text.Split(NoPrefix, " - "))),

    // ---- 1. load receipts (header on row 1) ----
    Source = Excel.Workbook(File.Contents(FolderPath & "04_Receipts\Bank_Receipts_Q1.xlsx"), null, true),
    ReceiptSheet = Source{[Item = "Receipts", Kind = "Sheet"]}[Data],
    PromotedHeaders = Table.PromoteHeaders(ReceiptSheet, [PromoteAllScalars = true]),

    // ---- 2. types: mixed real/text dates, '£' text and '(150.00)' amounts ----
    ConvertedDates = Table.TransformColumns(PromotedHeaders, {{"Date", ToDate, type date}}),
    ConvertedAmounts = Table.TransformColumns(ConvertedDates, {{"Amount", ParseMoney, type number}}),
    ChangedRef = Table.TransformColumnTypes(ConvertedAmounts, {{"Ref", Int64.Type}}),

    // ---- 3. same Ref twice = duplicate receipt, keep the first ----
    BufferedRows = Table.Buffer(ChangedRef),
    RemovedDuplicates = Table.Distinct(BufferedRows, {"Ref"}),

    // ---- 4. extract the payer text ----
    AddedPayer = Table.AddColumn(RemovedDuplicates, "Payer Text", each ExtractPayer([Description]), type text),

    // ---- 5. customers and transformation table ----
    CustomerSource = Excel.Workbook(File.Contents(FolderPath & "04_Receipts\Customer_List.xlsx"), null, true){[Item = "Customers", Kind = "Sheet"]}[Data],
    Customers = Table.TransformColumns(Table.PromoteHeaders(CustomerSource, [PromoteAllScalars = true]), {{"Customer", CleanText, type text}}),
    TransformSource = Excel.Workbook(File.Contents(FolderPath & "04_Receipts\Fuzzy_Transform_Table.xlsx"), null, true){[Item = "Transform", Kind = "Sheet"]}[Data],
    TransformTable = Table.PromoteHeaders(TransformSource, [PromoteAllScalars = true]),

    // ---- 6. fuzzy match Payer Text to Customer ----
    FuzzyJoined = Table.FuzzyNestedJoin(AddedPayer, {"Payer Text"}, Customers, {"Customer"}, "CustomerMatch", JoinKind.LeftOuter,
        [Threshold = 0.8, IgnoreCase = true, IgnoreSpace = true, NumberOfMatches = 1, TransformationTable = TransformTable]),
    ExpandedCustomer = Table.ExpandTableColumn(FuzzyJoined, "CustomerMatch", {"Customer"}, {"Customer Matched"}),
    LabelledUnmatched = Table.ReplaceValue(ExpandedCustomer, null, "UNMATCHED", Replacer.ReplaceValue, {"Customer Matched"}),
    Result = Table.SelectColumns(LabelledUnmatched, {"Date", "Ref", "Description", "Payer Text", "Amount", "Customer Matched"})
in
    Result
