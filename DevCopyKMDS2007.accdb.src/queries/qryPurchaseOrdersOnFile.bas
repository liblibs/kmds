Operation =1
Option =0
Begin InputTables
    Name ="Vendor"
    Name ="PurchaseOrderLine"
End
Begin OutputColumns
    Expression ="PurchaseOrderLine.VendorID"
    Expression ="Vendor.Name"
    Expression ="PurchaseOrderLine.PurchaseOrderID"
    Alias ="DateSetup"
    Expression ="Min(PurchaseOrderLine.SetupDate)"
End
Begin Joins
    LeftTable ="Vendor"
    RightTable ="PurchaseOrderLine"
    Expression ="Vendor.ID = PurchaseOrderLine.VendorID"
    Flag =1
End
Begin Groups
    Expression ="PurchaseOrderLine.VendorID"
    GroupLevel =0
    Expression ="Vendor.Name"
    GroupLevel =0
    Expression ="PurchaseOrderLine.PurchaseOrderID"
    GroupLevel =0
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
        dbText "Name" ="Vendor.Name"
        dbInteger "ColumnWidth" ="2265"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.PurchaseOrderID"
        dbInteger "ColumnWidth" ="1770"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="DateSetup"
        dbInteger "ColumnWidth" ="2175"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PurchaseOrderLine.VendorID"
        dbInteger "ColumnWidth" ="465"
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
    ColumnsShown =543
    Begin
        Left =198
        Top =2
        Right =294
        Bottom =101
        Top =0
        Name ="Vendor"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =143
        Bottom =105
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
End
