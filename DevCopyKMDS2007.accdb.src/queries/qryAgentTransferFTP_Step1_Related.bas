Operation =4
Option =0
Where ="(((AgentTransferFTPImport.ItemID) Not Between 1000 And 99999))"
Begin InputTables
    Name ="AgentTransferFTPImport"
    Name ="UPCItem"
End
Begin OutputColumns
    Name ="AgentTransferFTPImport.ItemID"
    Expression ="UPCItem.ItemID"
End
Begin Joins
    LeftTable ="AgentTransferFTPImport"
    RightTable ="UPCItem"
    Expression ="AgentTransferFTPImport.UPCCodeMax = UPCItem.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="AgentTransferFTPImport.ItemID"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="UPCItem.ItemID"
        dbInteger "ColumnWidth" ="1575"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =61
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="AgentTransferFTPImport"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="UPCItem"
        Name =""
    End
End
