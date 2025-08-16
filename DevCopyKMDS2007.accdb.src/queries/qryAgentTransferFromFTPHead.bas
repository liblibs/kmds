Operation =1
Option =0
Begin InputTables
    Name ="AgentTransferFTPImport"
End
Begin OutputColumns
    Expression ="AgentTransferFTPImport.AgentNumber"
    Expression ="AgentTransferFTPImport.Order"
    Expression ="AgentTransferFTPImport.OrdDate"
End
Begin Groups
    Expression ="AgentTransferFTPImport.AgentNumber"
    GroupLevel =0
    Expression ="AgentTransferFTPImport.Order"
    GroupLevel =0
    Expression ="AgentTransferFTPImport.OrdDate"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "DefaultView" ="2"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
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
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="AgentTransferFTPImport"
        Name =""
    End
End
