// Lab 3 - Stock valuation (KR's rule).  Query name: Stock_Valuation
// Needs three queries already loaded (Connection only is fine): Wasp_Clean, Xero_Clean, PurchaseLog_Clean.
// Price = latest Xero invoice Unit Price (by Invoice Date; credit notes already removed)
//         else latest Purchase Log Net Unit Price (by Date; cancelled lines already removed)
//         else Wasp Item Cost.   Stock Value = Total Available x Price.
let
    // ---- latest Xero price per part: sort newest first (ties: higher Reference first), buffer, keep first row per Item Code ----
    XeroSorted = Table.Buffer(Table.Sort(Xero_Clean, {{"Item Code", Order.Ascending}, {"Invoice Date", Order.Descending}, {"Reference", Order.Descending}})),
    XeroLatest = Table.SelectColumns(Table.Distinct(XeroSorted, {"Item Code"}), {"Item Code", "Unit Price"}),
    XeroLatestRenamed = Table.RenameColumns(XeroLatest, {{"Unit Price", "Xero Price"}}),

    // ---- latest Purchase Log price per part ----
    LogSorted = Table.Buffer(Table.Sort(PurchaseLog_Clean, {{"Part Number", Order.Ascending}, {"Date", Order.Descending}, {"PO Number", Order.Descending}})),
    LogLatest = Table.SelectColumns(Table.Distinct(LogSorted, {"Part Number"}), {"Part Number", "Net Unit Price"}),
    LogLatestRenamed = Table.RenameColumns(LogLatest, {{"Part Number", "Log Part Number"}, {"Net Unit Price", "Log Price"}}),

    // ---- left joins onto the cleaned Wasp parts (drop Wasp_Clean's own cost-based Stock Value; it is recalculated below) ----
    WaspParts = Table.RemoveColumns(Wasp_Clean, {"Stock Value"}),
    JoinedXero = Table.NestedJoin(WaspParts, {"Part Number"}, XeroLatestRenamed, {"Item Code"}, "XeroMatch", JoinKind.LeftOuter),
    ExpandedXero = Table.ExpandTableColumn(JoinedXero, "XeroMatch", {"Xero Price"}),
    JoinedLog = Table.NestedJoin(ExpandedXero, {"Part Number"}, LogLatestRenamed, {"Log Part Number"}, "LogMatch", JoinKind.LeftOuter),
    ExpandedLog = Table.ExpandTableColumn(JoinedLog, "LogMatch", {"Log Price"}),

    // ---- the pricing waterfall ----
    AddedPrice = Table.AddColumn(ExpandedLog, "Price",
        each if [Xero Price] <> null then [Xero Price] else if [Log Price] <> null then [Log Price] else [Item Cost], type number),
    AddedSource = Table.AddColumn(AddedPrice, "Price Source",
        each if [Xero Price] <> null then "Xero" else if [Log Price] <> null then "Purchase Log" else "Wasp", type text),
    AddedStockValue = Table.AddColumn(AddedSource, "Stock Value", each Number.Round([Total Available] * [Price], 2), type number),
    Result = Table.SelectColumns(AddedStockValue, {"Part Number", "Description", "Category", "Total Available", "Item Cost", "Price", "Price Source", "Stock Value", "Stock Status"})
in
    Result
