Operation =4
Option =0
Begin InputTables
    Name ="Item"
    Name ="ItemsSoldToday"
End
Begin OutputColumns
    Name ="Item.Onhand"
    Expression ="[Onhand]-[SoldToday]"
    Name ="Item.DateLastSold"
    Expression ="Now()"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="ItemsSoldToday"
    Expression ="Item.ID = ItemsSoldToday.Code"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
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
    Bottom =76
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =199
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =105
        Top =0
        Name ="ItemsSoldToday"
        Name =""
    End
End
