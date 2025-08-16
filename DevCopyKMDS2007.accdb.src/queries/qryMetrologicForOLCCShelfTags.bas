Operation =1
Option =0
Where ="(((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.SalesDepartmentID"
    Expression ="Item.ID"
    Alias ="NewUPC"
    Expression ="Format([id],\"000000\")"
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
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewUPC"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-406
    Top =141
    Right =1065
    Bottom =906
    Left =-1
    Top =-1
    Right =1439
    Bottom =482
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =287
        Bottom =307
        Top =0
        Name ="Item"
        Name =""
    End
End
