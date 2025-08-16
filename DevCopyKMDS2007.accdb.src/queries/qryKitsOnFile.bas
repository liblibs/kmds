Operation =1
Option =0
Begin InputTables
    Name ="Kit"
    Name ="qryKitItemsTotalRetail"
    Name ="Item"
End
Begin OutputColumns
    Expression ="Kit.ID"
    Expression ="Kit.ItemID"
    Expression ="Item.Description"
    Expression ="Kit.Price"
    Alias ="RetailPrice"
    Expression ="IIf(IsNull([Retail]),0,[retail])"
    Alias ="Discount"
    Expression ="[RetailPrice]-[Kit].[Price]"
End
Begin Joins
    LeftTable ="Kit"
    RightTable ="qryKitItemsTotalRetail"
    Expression ="Kit.ID = qryKitItemsTotalRetail.KitID"
    Flag =2
    LeftTable ="Kit"
    RightTable ="Item"
    Expression ="Kit.ItemID = Item.ID"
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
        dbText "Name" ="Kit.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kit.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Discount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Kit.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="RetailPrice"
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
    Bottom =305
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="Kit"
        Name =""
    End
    Begin
        Left =332
        Top =7
        Right =476
        Bottom =151
        Top =0
        Name ="qryKitItemsTotalRetail"
        Name =""
    End
    Begin
        Left =408
        Top =200
        Right =552
        Bottom =355
        Top =0
        Name ="Item"
        Name =""
    End
End
