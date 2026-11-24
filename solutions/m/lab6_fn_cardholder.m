// =====================================================================================
// LAB 6 - fnCardholder : returns one cardholder's transactions
// Needs the query  Transactions  (lab6_cards.m).  Create as Blank Query, paste, rename the
// query fnCardholder.  Then: Invoke on a table of names, e.g.  fnCardholder("Steve Mallory")
// =====================================================================================
(Cardholder as text) as table =>
let
    Source = Transactions,
    WantedName = Text.Lower(Text.Trim(Cardholder)),
    #"Filtered Cardholder" = Table.SelectRows(
        Source, each [Employee Name] <> null and Text.Lower([Employee Name]) = WantedName
    ),
    #"Selected Columns" = Table.SelectColumns(
        #"Filtered Cardholder",
        {"Date", "Merchant", "Category", "City", "Merchant Currency Description", "Original Currency Amount", "Amount GBP", "Bank Ref"}
    ),
    #"Sorted Rows" = Table.Sort(#"Selected Columns", {{"Date", Order.Ascending}}),
    #"Added Line No" = Table.AddIndexColumn(#"Sorted Rows", "Line No", 1, 1, Int64.Type)
in
    #"Added Line No"
