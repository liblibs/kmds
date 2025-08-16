Operation =3
Name ="ItemMovement"
Option =0
Where ="(((ItemMovement.Description) Is Null))"
Begin InputTables
    Name ="ItemMovementCurrentMonth"
    Name ="ItemMovement"
End
Begin OutputColumns
    Name ="Code"
    Expression ="ItemMovementCurrentMonth.Code"
    Name ="Description"
    Expression ="ItemMovementCurrentMonth.Description"
    Alias ="CurrMM"
    Name ="CurrMM"
    Expression ="ItemMovementCurrentMonth.CurMQty"
End
Begin Joins
    LeftTable ="ItemMovementCurrentMonth"
    RightTable ="ItemMovement"
    Expression ="ItemMovementCurrentMonth.Code = ItemMovement.Code"
    Flag =2
    LeftTable ="ItemMovementCurrentMonth"
    RightTable ="ItemMovement"
    Expression ="ItemMovementCurrentMonth.Description = ItemMovement.Description"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="CurrMM"
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
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="ItemMovementCurrentMonth"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="ItemMovement"
        Name =""
    End
End
