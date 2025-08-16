Operation =1
Option =0
Begin InputTables
    Name ="Item"
    Name ="UPCItem"
    Name ="UPCItem"
    Alias ="UPCItem_1"
    Name ="Item"
    Alias ="Item_1"
End
Begin OutputColumns
    Alias ="CartonID"
    Expression ="Item.ID"
    Alias ="CartonUPC"
    Expression ="UPCItem.ID"
    Alias ="CartonDesc"
    Expression ="Item.Description"
    Alias ="CartonPrice"
    Expression ="Item.Price"
    Alias ="PacksPerCarton"
    Expression ="Item.UnitsPerPack"
    Alias ="PackID"
    Expression ="Item_1.ID"
    Alias ="PackUPC"
    Expression ="Item.SubPack"
    Alias ="PackDesc"
    Expression ="Item_1.Description"
    Alias ="PackPrice"
    Expression ="Item_1.Price"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="UPCItem"
    Expression ="Item.ID = UPCItem.ItemID"
    Flag =1
    LeftTable ="UPCItem_1"
    RightTable ="Item_1"
    Expression ="UPCItem_1.ItemID = Item_1.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="UPCItem_1"
    Expression ="Item.SubPack = UPCItem_1.ID"
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
        dbText "Name" ="CartonDesc"
        dbInteger "ColumnWidth" ="2415"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonPrice"
        dbInteger "ColumnWidth" ="1200"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackUPC"
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackID"
        dbInteger "ColumnWidth" ="810"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PacksPerCarton"
        dbInteger "ColumnWidth" ="1605"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackDesc"
        dbInteger "ColumnWidth" ="2655"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackPrice"
        dbInteger "ColumnWidth" ="1080"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonUPC"
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
    Bottom =193
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =55
        Top =7
        Right =203
        Bottom =181
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =268
        Top =-11
        Right =364
        Bottom =118
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =362
        Top =151
        Right =458
        Bottom =265
        Top =0
        Name ="UPCItem_1"
        Name =""
    End
    Begin
        Left =564
        Top =28
        Right =660
        Bottom =142
        Top =0
        Name ="Item_1"
        Name =""
    End
End
