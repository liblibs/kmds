dbMemo "SQL" ="SELECT qryFintechDivisionID.Division_ID, qryFintechTransactionClubID.Invoice_Num"
    "ber, InvoLine5.SETUPDATE AS invoice_date, qryFintechTransactionClubID.Licensee A"
    "S Vendor_store_id, 0-[Quantity] AS quantity_shipped, \"EA\" AS Quantity_uom, [in"
    "voline5].[price]/[involine5].[quantity] AS unit_price, Cashier.ID AS CashierID, "
    "Cashier.Name AS CashierName, InvoLine5.CashSysID, qryFintechTransactionClubID.Na"
    "me AS ClubName, InvoLine5.TransType, 0 AS amount\015\012FROM Sysfile, qryFintech"
    "DivisionID, Cashier INNER JOIN ((qryFinTechTransactions INNER JOIN qryFintechTra"
    "nsactionClubID ON (qryFinTechTransactions.InvoHeadID = qryFintechTransactionClub"
    "ID.InvoHeadID) AND (qryFinTechTransactions.CashSysID = qryFintechTransactionClub"
    "ID.CashSysID) AND (qryFinTechTransactions.SetupDate = qryFintechTransactionClubI"
    "D.SetupDate)) INNER JOIN InvoLine5 ON (qryFinTechTransactions.InvoHeadID = InvoL"
    "ine5.InvoHeadID) AND (qryFinTechTransactions.CashSysID = InvoLine5.CashSysID) AN"
    "D (qryFinTechTransactions.SetupDate = InvoLine5.SetupDate)) ON Cashier.ID = Invo"
    "Line5.CashierID\015\012WHERE (((InvoLine5.TransType)=\"T\"));\015\012"
dbMemo "Connect" =""
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Quantity_uom"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="product_description"
        dbInteger "ColumnWidth" ="3165"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFintechDivisionID.Division_ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFintechTransactionClubID.Invoice_Number"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice_Date"
        dbInteger "ColumnWidth" ="1620"
        dbBoolean "ColumnHidden" ="0"
        dbText "Format" ="Short Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor_store_id"
        dbInteger "ColumnWidth" ="1980"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="quantity_shipped"
        dbInteger "ColumnWidth" ="2055"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="item_number"
        dbInteger "ColumnWidth" ="1695"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="unit_price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.CashSysID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryFintechTransactionClubID.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.TransType"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="amount"
        dbLong "AggregateType" ="-1"
    End
End
