Operation =1
Option =0
Where ="(((Item.Price)>0))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Alias ="count"
    Expression ="Sum(1)"
End
Begin Groups
    Expression ="Item.SalesDepartmentID"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="Item.SalesDepartmentID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="count"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-214
    Top =83
    Right =1011
    Bottom =760
    Left =-1
    Top =-1
    Right =1193
    Bottom =353
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =270
        Top =0
        Name ="Item"
        Name =""
    End
End
