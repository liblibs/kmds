Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)>2 Or (Item.SalesDepartmentID)>2) AND ((Item.Price)>0)"
    ")"
Having ="(((Item.VendorID)>1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.VendorID"
    Alias ="count"
    Expression ="Sum(1)"
End
Begin Groups
    Expression ="Item.VendorID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbByte "RecordsetType" ="0"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Item.VendorID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-495
    Top =80
    Right =759
    Bottom =886
    Left =-1
    Top =-1
    Right =1222
    Bottom =505
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =300
        Top =0
        Name ="Item"
        Name =""
    End
End
