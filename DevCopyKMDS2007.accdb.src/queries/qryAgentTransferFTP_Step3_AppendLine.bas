Operation =3
Name ="AgentTransferLines"
Option =0
Begin InputTables
    Name ="qryAgentTransferHeadIDFromFTP"
    Name ="AgentTransferFTPImport"
    Name ="Item"
End
Begin OutputColumns
    Name ="LineID"
    Expression ="AgentTransferFTPImport.LineID"
    Name ="ItemID"
    Expression ="AgentTransferFTPImport.ItemID"
    Name ="QuantityOrdered"
    Expression ="AgentTransferFTPImport.QuantityShipped"
    Name ="QuantityShipped"
    Expression ="AgentTransferFTPImport.QuantityShipped"
    Name ="Price"
    Expression ="AgentTransferFTPImport.Price"
    Name ="Cost"
    Expression ="AgentTransferFTPImport.Cost"
    Name ="AgentTransferHeadID"
    Expression ="qryAgentTransferHeadIDFromFTP.ID"
End
Begin Joins
    LeftTable ="AgentTransferFTPImport"
    RightTable ="Item"
    Expression ="AgentTransferFTPImport.ItemID = Item.ID"
    Flag =1
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =295
        Top =6
        Right =391
        Bottom =99
        Top =0
        Name ="qryAgentTransferHeadIDFromFTP"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =257
        Bottom =120
        Top =0
        Name ="AgentTransferFTPImport"
        Name =""
    End
    Begin
        Left =429
        Top =6
        Right =525
        Bottom =120
        Top =0
        Name ="Item"
        Name =""
    End
End
