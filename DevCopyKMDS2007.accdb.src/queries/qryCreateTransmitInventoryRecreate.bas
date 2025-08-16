Operation =1
Option =0
Where ="(((Item.ID) Between 99900000000 And 99999999999) AND ((Item.AvgTurns)>0) AND ((I"
    "tem.SalesDepartmentID)=1))"
Begin InputTables
    Name ="Item"
End
Begin OutputColumns
    Expression ="Item.ID"
    Expression ="Item.AvgTurns"
End
Begin OrderBy
    Expression ="Item.ID"
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
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.AvgTurns"
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
    Right =1193
    Bottom =45
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
