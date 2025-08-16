Operation =4
Option =0
Where ="(((Item.UnitsPerOrder)>=[Unitspercase]) AND ((Item.UnitsPerCase)>1) AND ((Item.S"
    "alesDepartmentID)=1) AND ((Item.SizeID)=10 Or (Item.SizeID)=75 Or (Item.SizeID)="
    "20 Or (Item.SizeID)=50 Or (Item.SizeID)=17))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.Location"
    Expression ="\"Back\""
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.UnitsPerOrder"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.UnitsPerCase"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Location"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-6
    Top =283
    Right =1219
    Bottom =960
    Left =-1
    Top =-1
    Right =1193
    Bottom =220
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =279
        Bottom =267
        Top =0
        Name ="Item"
        Name =""
    End
End
