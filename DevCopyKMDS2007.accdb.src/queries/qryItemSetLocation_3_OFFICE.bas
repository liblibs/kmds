Operation =4
Option =0
Where ="(((Item.UnitsPerOrder)>=[Unitspercase]) AND ((Item.UnitsPerCase)>1) AND ((Item.S"
    "alesDepartmentID)=1) AND ((Item.SizeID)=20 Or (Item.SizeID)=37 Or (Item.SizeID)="
    "5))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Name ="Item.Location"
    Expression ="\"Office\""
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
    Left =-39
    Top =258
    Right =1186
    Bottom =935
    Left =-1
    Top =-1
    Right =1193
    Bottom =203
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
