Operation =4
Option =0
Where ="(((Item.SalesDepartmentID)>2))"
Begin InputTables
    Name ="InvoLine5"
    Name ="Item"
End
Begin OutputColumns
    Name ="InvoLine5.Cost"
    Expression ="Item.Cost*InvoLine5.quantity"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
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
    Bottom =242
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =306
        Top =6
        Right =402
        Bottom =225
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =270
        Top =0
        Name ="Item"
        Name =""
    End
End
