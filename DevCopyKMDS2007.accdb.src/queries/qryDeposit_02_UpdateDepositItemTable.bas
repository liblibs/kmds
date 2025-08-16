Operation =4
Option =0
Begin InputTables
    Name ="DepositTable"
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.Onhand"
    Expression ="[DepositTotal]/[Item].[price]"
End
Begin Joins
    LeftTable ="DepositTable"
    RightTable ="Item"
    Expression ="DepositTable.DepositItemID=Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
Begin
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =422
    Top =310
    Right =1887
    Bottom =1121
    Left =-1
    Top =-1
    Right =1427
    Bottom =490
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="DepositTable"
        Name =""
    End
    Begin
        Left =358
        Top =19
        Right =780
        Bottom =436
        Top =0
        Name ="Item"
        Name =""
    End
End
