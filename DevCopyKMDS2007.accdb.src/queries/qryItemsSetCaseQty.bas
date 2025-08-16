Operation =4
Option =0
Where ="(((Item.UnitsPerCase)<>[CaseQty]))"
Begin InputTables
    Name ="Item"
    Name ="ItemCaseQty"
End
Begin OutputColumns
    Name ="Item.UnitsPerCase"
    Expression ="[CaseQty]"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="ItemCaseQty"
    Expression ="Item.ID = ItemCaseQty.Code"
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
    Bottom =77
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
        Name ="ItemCaseQty"
        Name =""
    End
End
