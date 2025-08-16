Operation =4
Option =0
Where ="(((Item.Onhand)<>0) And ((Len(upcitem.ID))>3))"
Begin InputTables
    Name ="Item"
    Name ="UPCItem"
    Name ="Item"
    Alias ="Item_1"
End
Begin OutputColumns
    Name ="Item_1.Onhand"
    Expression ="Item_1.Onhand+(Item.onhand*Item.UnitsPerPack)"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="UPCItem"
    Expression ="Item.SubPack = UPCItem.ID"
    Flag =1
    LeftTable ="UPCItem"
    RightTable ="Item_1"
    Expression ="UPCItem.ItemID = Item_1.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="PacksSold"
        dbInteger "ColumnWidth" ="1035"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonDesc"
        dbInteger "ColumnWidth" ="2730"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CartonOnhand"
        dbInteger "ColumnWidth" ="1425"
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
        dbText "Name" ="PackDesc"
        dbInteger "ColumnWidth" ="2145"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item_1.Onhand"
        dbInteger "ColumnWidth" ="840"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PackNewOnhand"
        dbInteger "ColumnWidth" ="1605"
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
    Bottom =242
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =195
        Bottom =240
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =259
        Top =14
        Right =355
        Bottom =143
        Top =0
        Name ="UPCItem"
        Name =""
    End
    Begin
        Left =393
        Top =6
        Right =587
        Bottom =240
        Top =0
        Name ="Item_1"
        Name =""
    End
End
