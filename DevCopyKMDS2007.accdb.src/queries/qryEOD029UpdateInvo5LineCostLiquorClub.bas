Operation =4
Option =0
Where ="(((Item.SalesDepartmentID)=1) AND ((InvoLine5.TransType)=\"C\") AND ((SalesDepar"
    "tment.ID)=2))"
Begin InputTables
    Name ="SalesDepartment"
    Name ="InvoLine5"
    Name ="Item"
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
        Left =15
        Top =23
        Right =173
        Bottom =227
        Top =0
        Name ="SalesDepartment"
        Name =""
    End
    Begin
        Left =614
        Top =39
        Right =710
        Bottom =198
        Top =0
        Name ="InvoLine5"
        Name =""
    End
    Begin
        Left =339
        Top =28
        Right =435
        Bottom =142
        Top =0
        Name ="Item"
        Name =""
    End
End
