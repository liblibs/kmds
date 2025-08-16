Operation =1
Option =0
Where ="(((Item.ID) Is Null))"
Begin InputTables
    Name ="AgentTransferFTPImport"
    Name ="Item"
End
Begin OutputColumns
    Expression ="AgentTransferFTPImport.ItemID"
    Expression ="AgentTransferFTPImport.Description"
    Expression ="Item.ID"
    Expression ="AgentTransferFTPImport.LineID"
    Expression ="AgentTransferFTPImport.ItemID"
    Expression ="AgentTransferFTPImport.Description"
    Expression ="AgentTransferFTPImport.SizeID"
    Expression ="AgentTransferFTPImport.QuantityShipped"
    Expression ="AgentTransferFTPImport.Price"
    Expression ="AgentTransferFTPImport.UPCCodeMax"
    Expression ="AgentTransferFTPImport.ExtPrice"
    Expression ="AgentTransferFTPImport.Status"
    Expression ="AgentTransferFTPImport.SizeDesc"
End
Begin Joins
    LeftTable ="AgentTransferFTPImport"
    RightTable ="Item"
    Expression ="AgentTransferFTPImport.ItemID = Item.ID"
    Flag =2
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
        dbText "Name" ="Expr1000"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Expr1001"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Item.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.LineID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.ItemID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.Description"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.QuantityShipped"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.Price"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.UPCCodeMax"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.ExtPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="AgentTransferFTPImport.SizeDesc"
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
    Bottom =94
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =222
        Bottom =120
        Top =0
        Name ="AgentTransferFTPImport"
        Name =""
    End
    Begin
        Left =381
        Top =-5
        Right =477
        Bottom =109
        Top =0
        Name ="Item"
        Name =""
    End
End
