Operation =1
Option =0
Where ="(((InvoLine5.TransType)=\"C\" Or (InvoLine5.TransType)=\"D\") AND ((Item.SalesDe"
    "partmentID)=1 Or (Item.SalesDepartmentID)=2))"
Begin InputTables
    Name ="qryFintechTransactionClubID"
    Name ="qryFinTechTransactions"
    Name ="InvoLine5"
    Name ="Sysfile"
    Name ="qryFintechDivisionID"
    Name ="Item"
End
Begin OutputColumns
    Expression ="qryFintechDivisionID.Division_ID"
    Expression ="qryFintechTransactionClubID.Invoice_Number"
    Alias ="invoice_date"
    Expression ="Format([qryFintechTransactionClubID].[SetupDate],\"mm/dd/yyyy\")"
    Alias ="Vendor_store_id"
    Expression ="qryFintechTransactionClubID.Licensee"
    Alias ="quantity_shipped"
    Expression ="InvoLine5.Quantity"
    Alias ="Quantity_uom"
    Expression ="\"EA\""
    Alias ="item_number"
    Expression ="CStr([Item].[id])"
    Alias ="product_description"
    Expression ="IIf(([involine5].[TransType]=\"C\"),[item].[Description],\"Discount\")"
    Alias ="unit_price"
    Expression ="[involine5].[price]/[involine5].[quantity]"
    Expression ="Item.SalesDepartmentID"
End
Begin Joins
    LeftTable ="qryFinTechTransactions"
    RightTable ="qryFintechTransactionClubID"
    Expression ="qryFinTechTransactions.InvoHeadID = qryFintechTransactionClubID.InvoHeadID"
    Flag =1
    LeftTable ="qryFinTechTransactions"
    RightTable ="qryFintechTransactionClubID"
    Expression ="qryFinTechTransactions.CashSysID = qryFintechTransactionClubID.CashSysID"
    Flag =1
    LeftTable ="qryFinTechTransactions"
    RightTable ="qryFintechTransactionClubID"
    Expression ="qryFinTechTransactions.SetupDate = qryFintechTransactionClubID.SetupDate"
    Flag =1
    LeftTable ="qryFinTechTransactions"
    RightTable ="InvoLine5"
    Expression ="qryFinTechTransactions.InvoHeadID = InvoLine5.InvoHeadID"
    Flag =1
    LeftTable ="qryFinTechTransactions"
    RightTable ="InvoLine5"
    Expression ="qryFinTechTransactions.CashSysID = InvoLine5.CashSysID"
    Flag =1
    LeftTable ="qryFinTechTransactions"
    RightTable ="InvoLine5"
    Expression ="qryFinTechTransactions.SetupDate = InvoLine5.SetupDate"
    Flag =1
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
End
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
        dbText "Name" ="unit_price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine5.SalesDept"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =25
    Right =1510
    Bottom =818
    Left =-1
    Top =-1
    Right =1478
    Bottom =362
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =56
        Top =125
        Right =200
        Bottom =366
        Top =0
        Name ="qryFintechTransactionClubID"
        Name =""
    End
    Begin
        Left =377
        Top =13
        Right =521
        Bottom =346
        Top =0
        Name ="qryFinTechTransactions"
        Name =""
    End
    Begin
        Left =661
        Top =75
        Right =803
        Bottom =437
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =1089
        Top =56
        Right =1317
        Bottom =476
        Top =0
        Name ="Sysfile"
        Name =""
    End
    Begin
        Left =856
        Top =281
        Right =1000
        Bottom =425
        Top =0
        Name ="qryFintechDivisionID"
        Name =""
    End
    Begin
        Left =862
        Top =51
        Right =1006
        Bottom =238
        Top =0
        Name ="Item"
        Name =""
    End
End
