Operation =4
Option =0
Where ="(((Item.UnitsPerCase)=0) AND ((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
    Name ="Size"
End
Begin OutputColumns
    Name ="Item.UnitsPerCase"
    Expression ="[CaseQty]"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="Size"
    Expression ="Item.SizeID = Size.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Item.UnitsPerCase"
        dbInteger "ColumnWidth" ="1710"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Description"
        dbInteger "ColumnWidth" ="3900"
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
    Bottom =278
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =159
        Bottom =286
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =236
        Top =0
        Right =332
        Bottom =114
        Top =0
        Name ="Size"
        Name =""
    End
End
