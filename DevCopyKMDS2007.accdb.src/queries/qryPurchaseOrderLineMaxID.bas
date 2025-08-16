Operation =1
Option =0
Where ="(((PurchaseOrderLine.PurchaseOrderID)=[parPurchaseOrderID]))"
Begin InputTables
    Name ="PurchaseOrderLine"
End
Begin OutputColumns
    Alias ="MaxLineID"
    Expression ="Max(PurchaseOrderLine.LineID)"
End
Begin Parameters
    Name ="parPurchaseOrderID"
    Flag =4
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
        dbText "Name" ="MaxLineID"
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
    Right =1209
    Bottom =150
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =56
        Top =9
        Right =152
        Bottom =123
        Top =0
        Name ="PurchaseOrderLine"
        Name =""
    End
End
