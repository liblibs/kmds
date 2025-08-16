Operation =4
Option =0
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.SoldTodayCr1"
    Expression ="0"
    Name ="Item.SoldTodayCr2"
    Expression ="0"
    Name ="Item.SoldTodayCr3"
    Expression ="0"
    Name ="Item.SoldTodayCr4"
    Expression ="0"
    Name ="Item.TransferredToday"
    Expression ="0"
    Name ="Item.RestockAmount"
    Expression ="0"
    Name ="Item.SoldTodayCr5"
    Expression ="0"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.SoldTodayCr1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SoldTodayCr5"
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
    Bottom =93
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =240
        Bottom =111
        Top =0
        Name ="Item"
        Name =""
    End
End
