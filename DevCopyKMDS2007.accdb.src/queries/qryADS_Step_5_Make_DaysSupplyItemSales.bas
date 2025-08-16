Operation =2
Name ="DaysSupplyItemSales"
Option =0
Begin InputTables
    Name ="qryADS_ItemSold"
    Name ="qryADS_ItemSoldByPacks"
End
Begin OutputColumns
    Expression ="qryADS_ItemSold.Code"
    Alias ="Sold"
    Expression ="CLng([sq]+IIf(IsNull([SumOfQuantity]),0,[SumOfQuantity]/[UnitsPerPack]))"
End
Begin Joins
    LeftTable ="qryADS_ItemSold"
    RightTable ="qryADS_ItemSoldByPacks"
    Expression ="qryADS_ItemSold.Code = qryADS_ItemSoldByPacks.ID"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Sold"
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
        Left =22
        Top =18
        Right =118
        Bottom =102
        Top =0
        Name ="qryADS_ItemSold"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =528
        Bottom =105
        Top =0
        Name ="qryADS_ItemSoldByPacks"
        Name =""
    End
End
