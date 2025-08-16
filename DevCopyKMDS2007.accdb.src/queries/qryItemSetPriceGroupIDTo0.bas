Operation =4
Option =0
Where ="(((Item.PriceGroupID)<>0) AND ((PriceGroup.ID) Is Null)) OR (((Item.PriceGroupID"
    ")<>0) AND ((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
    Name ="PriceGroup"
End
Begin OutputColumns
    Name ="Item.PriceGroupID"
    Expression ="0"
End
Begin Joins
    LeftTable ="Item"
    RightTable ="PriceGroup"
    Expression ="Item.PriceGroupID = PriceGroup.ID"
    Flag =2
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
    Bottom =172
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =165
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =276
        Top =10
        Right =372
        Bottom =124
        Top =0
        Name ="PriceGroup"
        Name =""
    End
End
