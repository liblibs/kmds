Operation =4
Option =0
Where ="(((Item.SalesDepartmentID)=1) AND ((InvoLine5.TransType)=\"S\"))"
Begin InputTables
    Name ="InvoLine5"
    Name ="Item"
    Name ="SalesDepartment"
End
Begin OutputColumns
    Name ="InvoLine5.Cost"
    Expression ="(involine5.price)*(1-[TargetGP])"
End
Begin Joins
    LeftTable ="InvoLine5"
    RightTable ="Item"
    Expression ="InvoLine5.Code = Item.ID"
    Flag =1
    LeftTable ="Item"
    RightTable ="SalesDepartment"
    Expression ="Item.SalesDepartmentID = SalesDepartment.ID"
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
    Bottom =259
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =298
        Top =30
        Right =394
        Bottom =189
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =167
        Top =46
        Right =263
        Bottom =160
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =15
        Top =23
        Right =148
        Bottom =227
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
End
