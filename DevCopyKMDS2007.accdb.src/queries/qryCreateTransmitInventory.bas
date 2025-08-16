Operation =1
Option =0
Where ="(((Item.Onhand)>0) AND ((Item.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Alias ="ID"
    Expression ="Format([ID],\"00000000000\")"
    Expression ="Item.Onhand"
End
Begin OrderBy
    Expression ="Format([ID],\"00000000000\")"
    Flag =0
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
        dbText "Name" ="ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.Onhand"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-209
    Top =41
    Right =1016
    Bottom =718
    Left =-1
    Top =-1
    Right =1193
    Bottom =201
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =191
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
