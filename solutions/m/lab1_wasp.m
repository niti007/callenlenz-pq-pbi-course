// FolderPath is a TEXT PARAMETER (Home > Manage Parameters > New Parameter, Type = Text),
// e.g.  C:\CallenLenz_Course\data\   (keep the trailing backslash). No paths are hard-coded below.
// Lab 1 - Wasp stock clean-up.  Query name: Wasp_Clean
// Rule order matters: errors -> blank rows -> repeated header -> fill down -> normalise -> filter -> dedupe.
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

    // ---- 1. load and remove the 2 export title rows, promote the real header (row 3) ----
    Source = Excel.Workbook(File.Contents(FolderPath & "01_Wasp_Stock\Wasp_Stock_Export.xlsx"), null, true),
    StockSheet = Source{[Item = "Stock", Kind = "Sheet"]}[Data],
    RemovedTopRows = Table.Skip(StockSheet, 2),
    PromotedHeaders = Table.PromoteHeaders(RemovedTopRows, [PromoteAllScalars = true]),

    // ---- 2. Replace Errors: the #REF! cells in Item Cost become null (must happen BEFORE row filters) ----
    ReplacedErrors = Table.ReplaceErrorValues(PromotedHeaders, {{"Item Cost", null}}),

    // ---- 3. structure clean-up ----
    RemovedBlankRows = Table.SelectRows(ReplacedErrors, each not List.IsEmpty(List.RemoveMatchingItems(Record.FieldValues(_), {"", null}))),
    RemovedRepeatedHeaders = Table.SelectRows(RemovedBlankRows, each [Item Number] <> "Item Number"),
    FilledDownCategory = Table.FillDown(RemovedRepeatedHeaders, {"Category"}),

    // ---- 4. normalise part numbers; non-PRT rows (TEST-ITEM, ZZ SAMPLE, Total) become null and are dropped ----
    NormalisedItemNumber = Table.TransformColumns(FilledDownCategory, {{"Item Number", NormPart, type text}}),
    ValidPartsOnly = Table.SelectRows(NormalisedItemNumber, each [Item Number] <> null),

    // ---- 5. remove duplicates, keeping the FIRST row ----
    // NOTE: the UI's Remove Duplicates only reliably keeps the first row if the table is buffered first;
    // without Table.Buffer Power Query may evaluate the rows in another order and keep a different duplicate.
    BufferedRows = Table.Buffer(ValidPartsOnly),
    RemovedDuplicates = Table.Distinct(BufferedRows, {"Item Number"}),

    // ---- 6. tidy text, numbers and types ----
    CleanedDescription = Table.TransformColumns(RemovedDuplicates, {{"Description", CleanText, type text}}),
    // Locale-safe cost parse: numbers pass through; text has £ and , removed then is read with en-GB; N/A etc -> null
    CleanedItemCost = Table.AddColumn(CleanedDescription, "Item Cost Clean",
        each if Value.Is([Item Cost], type number) then [Item Cost]
             else try Number.From(Text.Replace(Text.Replace(Text.Trim([Item Cost]), "£", ""), ",", ""), "en-GB") otherwise null, type nullable number),
    ChangedTypes = Table.TransformColumnTypes(CleanedItemCost, {{"Total Available", Int64.Type}, {"Total In House", Int64.Type}, {"On Order", Int64.Type}, {"Committed", Int64.Type}}),

    // ---- 7. Extract: drawing number between "[Drawing " and "]" ----
    AddedDrawingNo = Table.AddColumn(ChangedTypes, "Drawing No", each Text.BetweenDelimiters([Description], "[Drawing ", "]"), type text),

    // ---- 8. Conditional column: Stock Status ----
    AddedStockStatus = Table.AddColumn(AddedDrawingNo, "Stock Status",
        each if [Total Available] = 0 then "Out of stock" else if [Total Available] <= 5 then "Reorder" else "OK", type text),

    // ---- 9. Stock Value at Wasp cost (unusable cost = blank = adds 0), then keep the useful columns ----
    AddedStockValue = Table.AddColumn(AddedStockStatus, "Stock Value", each [Total Available] * (if [Item Cost Clean] = null then 0 else [Item Cost Clean]), type number),
    SelectedColumns = Table.SelectColumns(AddedStockValue, {"Item Number", "Description", "Drawing No", "Category", "Total Available", "Total In House", "On Order", "Committed", "Item Cost Clean", "Stocking Unit", "Location", "Site", "Stock Status", "Stock Value"}),
    Renamed = Table.RenameColumns(SelectedColumns, {{"Item Number", "Part Number"}, {"Item Cost Clean", "Item Cost"}})
in
    Renamed
