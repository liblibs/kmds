Operation =1
Option =0
Where ="(((OrderLine.OrdheadID)=[parOrdheadID]) AND ((OrderLine.ItemID)=[parItemID]))"
Begin InputTables
    Name ="OrderLine"
End
Begin OutputColumns
    Expression ="OrderLine.OrdheadID"
    Expression ="OrderLine.LineID"
    Expression ="OrderLine.ItemID"
End
Begin Parameters
    Name ="parOrdheadID"
    Flag =4
    Name ="parItemID"
    Flag =16
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
        dbText "Name" ="OrderLine.OrdheadID"
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
    Bottom =213
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="OrderLine"
        Name =""
    End
End
