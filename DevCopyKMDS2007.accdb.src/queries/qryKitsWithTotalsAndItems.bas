Operation =1
Option =0
Begin InputTables
    Name ="KitItem"
    Name ="qryKitsOnFile"
    Name ="Item"
    Alias ="Item_1"
End
Begin OutputColumns
    Expression ="qryKitsOnFile.ID"
    Alias ="KitItemID"
    Expression ="qryKitsOnFile.ItemID"
    Alias ="KitDescription"
    Expression ="qryKitsOnFile.Description"
    Alias ="KitPrice"
    Expression ="qryKitsOnFile.Price"
    Expression ="qryKitsOnFile.RetailPrice"
    Alias ="KitDiscPerct"
    Expression ="IIf([RetailPrice]=0,0,[Discount]/[RetailPrice])"
    Alias ="KitDiscount"
    Expression ="qryKitsOnFile.Discount"
    Alias ="PartID"
    Expression ="KitItem.ItemID"
    Alias ="PartDescription"
    Expression ="Item_1.Description"
    Alias ="PartPrice"
    Expression ="KitItem.Price"
    Expression ="KitItem.SetupDate"
    Expression ="KitItem.ChangeDate"
End
Begin Joins
    LeftTable ="KitItem"
    RightTable ="qryKitsOnFile"
    Expression ="KitItem.KitID = qryKitsOnFile.ID"
    Flag =3
    LeftTable ="KitItem"
    RightTable ="Item_1"
    Expression ="KitItem.ItemID = Item_1.ID"
    Flag =2
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
        dbText "Name" ="qryKitsOnFile.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryKitsOnFile.RetailPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KitItem.SetupDate"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KitItem.ChangeDate"
        dbInteger "ColumnWidth" ="1965"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartID"
        dbInteger "ColumnWidth" ="1935"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartDescription"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="PartPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KitItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KitDescription"
        dbInteger "ColumnWidth" ="2985"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KitPrice"
        dbInteger "ColumnWidth" ="2205"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KitDiscount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KitDiscPerct"
        dbInteger "ColumnWidth" ="2205"
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
    Bottom =313
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =352
        Top =6
        Right =496
        Bottom =150
        Top =0
        Name ="KitItem"
        Name =""
    End
    Begin
        Left =122
        Top =13
        Right =266
        Bottom =157
        Top =0
        Name ="qryKitsOnFile"
        Name =""
    End
    Begin
        Left =698
        Top =36
        Right =842
        Bottom =227
        Top =0
        Name ="Item_1"
        Name =""
    End
End
