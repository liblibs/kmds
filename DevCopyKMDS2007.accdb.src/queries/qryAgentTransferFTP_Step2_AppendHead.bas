Operation =3
Name ="AgentTransferHead"
Option =0
Begin InputTables
    Name ="AgentTransferFTPImport"
End
Begin OutputColumns
    Name ="AgentID"
    Expression ="AgentTransferFTPImport.AgentNumber"
    Name ="Order"
    Expression ="AgentTransferFTPImport.Order"
    Name ="OrdDate"
    Expression ="AgentTransferFTPImport.OrdDate"
    Alias ="Status"
    Name ="Status"
    Expression ="\"I\""
    Alias ="OrderNotes"
    Name ="OrderNotes"
    Expression ="\"Imported from KMDS FTP\""
End
Begin Groups
    Expression ="AgentTransferFTPImport.AgentNumber"
    GroupLevel =0
    Expression ="AgentTransferFTPImport.Order"
    GroupLevel =0
    Expression ="AgentTransferFTPImport.OrdDate"
    GroupLevel =0
    Expression ="\"I\""
    GroupLevel =0
    Expression ="\"Imported from KMDS FTP\""
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="Status"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderNotes"
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
    ColumnsShown =655
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
