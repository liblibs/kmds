Operation =4
Option =0
Begin InputTables
    Name ="ItemMovement"
    Name ="ItemMovementCurrentMonth"
End
Begin OutputColumns
    Name ="ItemMovement.CurrMM"
    Expression ="[CurMQty]"
End
Begin Joins
    LeftTable ="ItemMovement"
    RightTable ="ItemMovementCurrentMonth"
    Expression ="ItemMovement.Description = ItemMovementCurrentMonth.Description"
    Flag =1
    LeftTable ="ItemMovement"
    RightTable ="ItemMovementCurrentMonth"
    Expression ="ItemMovement.Code = ItemMovementCurrentMonth.Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="0"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
Begin
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
    Bottom =134
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =19
        Top =5
        Right =164
        Bottom =119
        Top =0
        Name ="ItemMovement"
        Name =""
    End
    Begin
        Left =202
        Top =6
        Right =298
        Bottom =120
        Top =0
        Name ="ItemMovementCurrentMonth"
        Name =""
    End
End
