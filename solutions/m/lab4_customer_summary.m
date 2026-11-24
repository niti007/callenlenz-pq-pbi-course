// Lab 4 (part 2) - fills the Customer_List columns '£ paid?' and '# payments'.  Query name: Customer_Summary
// Needs the query Receipts_Matched. £ paid = NET sum (refund reduces it); # payments = count of positive receipts.
let
    MatchedOnly = Table.SelectRows(Receipts_Matched, each [Customer Matched] <> "UNMATCHED"),
    Grouped = Table.Group(MatchedOnly, {"Customer Matched"}, {
        {"£ paid?", each List.Sum([Amount]), type number},
        {"# payments", each List.Count(List.Select([Amount], each _ > 0)), Int64.Type}}),
    Rounded = Table.TransformColumns(Grouped, {{"£ paid?", each Number.Round(_, 2), type number}}),
    Sorted = Table.Sort(Rounded, {{"Customer Matched", Order.Ascending}}),
    Result = Table.RenameColumns(Sorted, {{"Customer Matched", "Customer"}})
in
    Result
