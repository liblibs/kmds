Operation =1
Option =0
Begin InputTables
    Name ="TimeClockPunches"
End
Begin OutputColumns
    Expression ="TimeClockPunches.PunchDate"
End
Begin OrderBy
    Expression ="TimeClockPunches.PunchDate"
    Flag =1
End
Begin Groups
    Expression ="TimeClockPunches.PunchDate"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="0"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
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
        Name ="TimeClockPunches"
        Name =""
    End
End
