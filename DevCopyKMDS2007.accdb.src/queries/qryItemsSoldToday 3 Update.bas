Operation =4
Option =0
Begin InputTables
    Name ="Item"
    Name ="ItemsSoldToday"
End
Begin OutputColumns
    Name ="Item.SoldTodayCr1"
    Expression ="IIf(([CashSysID]=1),[SumOfQuantity],[SoldTodayCR1])"
    Name ="Item.SoldTodayCr2"
    Expression ="IIf(([CashSysID]=2),[SumOfQuantity],[SoldTodayCR2])"
    Name ="Item.SoldTodayCr3"
    Expression ="IIf(([CashSysID]=3),[SumOfQuantity],[SoldTodayCR3])"
    Name ="Item.SoldTodayCr4"
    Expression ="IIf(([CashSysID]=4),[SumOfQuantity],[SoldTodayCR4])"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =90
        Top =0
        Name ="ItemsSoldToday"
        Name =""
    End
End
