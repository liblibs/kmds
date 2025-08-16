Operation =1
Option =0
Where ="(((InvoLine.Code)=Forms!PurchaseOrdersOnFile!PurchaseOrdersOnFileSubItemMovement"
    ".Form!Code) And ((TranTypes.Description)=Forms!PurchaseOrdersOnFile!PurchaseOrde"
    "rsOnFileSubItemMovement.Form!TransType) And ((InvoLine.SetupDate) Between CDate("
    "Forms!PurchaseOrdersOnFile!PurchaseOrdersOnFileSubItemMovement.Form!txtDateBegin"
    ") And CDate(Forms!PurchaseOrdersOnFile!PurchaseOrdersOnFileSubItemMovement.Form!"
    "txtDateEnd)))"
Begin InputTables
    Name ="InvoLine"
    Name ="TranTypes"
    Name ="Cashier"
End
Begin OutputColumns
    Expression ="InvoLine.Code"
    Expression ="TranTypes.Description"
    Expression ="InvoLine.InputCode"
    Alias ="Cashier"
    Expression ="Cashier.Name"
    Alias ="Reg#"
    Expression ="InvoLine.CashSysID"
    Alias ="Inv#"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.Quantity"
    Expression ="InvoLine.Price"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.ID"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.TransType"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="TranTypes"
    Expression ="InvoLine.TransType=TranTypes.Code"
    Flag =1
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID=Cashier.ID"
    Flag =2
End
Begin OrderBy
    Expression ="InvoLine.SetupDate"
    Flag =0
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
        dbText "Name" ="InvoLine.SetupDate"
        dbInteger "ColumnWidth" ="2535"
        dbInteger "ColumnOrder" ="7"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Quantity"
        dbInteger "ColumnWidth" ="930"
        dbInteger "ColumnOrder" ="8"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbInteger "ColumnWidth" ="2220"
        dbInteger "ColumnOrder" ="3"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Code"
        dbInteger "ColumnWidth" ="615"
        dbInteger "ColumnOrder" ="1"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TranTypes.Description"
        dbInteger "ColumnWidth" ="1575"
        dbInteger "ColumnOrder" ="2"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.Price"
        dbInteger "ColumnOrder" ="9"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Cashier"
        dbInteger "ColumnOrder" ="4"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Reg#"
        dbInteger "ColumnWidth" ="615"
        dbInteger "ColumnOrder" ="5"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Inv#"
        dbInteger "ColumnWidth" ="915"
        dbInteger "ColumnOrder" ="6"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =0
    Right =1225
    Bottom =677
    Left =-1
    Top =-1
    Right =1209
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =256
        Top =104
        Right =352
        Bottom =203
        Top =0
        Name ="TranTypes"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =120
        Top =0
        Name ="Cashier"
        Name =""
    End
End
