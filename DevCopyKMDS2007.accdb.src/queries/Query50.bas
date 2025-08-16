Operation =1
Option =0
Begin InputTables
    Name ="OLCC Categories"
    Name ="OLCCUPCaseOrder"
End
Begin OutputColumns
    Expression ="OLCCUPCaseOrder.SubCategory"
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
        dbText "Name" ="OLCCUPCaseOrder.SubCategory"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =136
    Right =1548
    Bottom =859
    Left =-1
    Top =-1
    Right =1438
    Bottom =439
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="OLCC Categories"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =354
        Top =0
        Name ="OLCCUPCaseOrder"
        Name =""
    End
End
