Operation =5
Option =0
Where ="(((ClubOrderGuides.ClubID)=[parClubID]))"
Begin InputTables
    Name ="ClubOrderGuides"
End
Begin OutputColumns
    Expression ="ClubOrderGuides.ClubID"
    Expression ="ClubOrderGuides.*"
End
Begin Parameters
    Name ="parClubID"
    Flag =4
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
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
    ColumnsShown =771
    Begin
        Left =80
        Top =8
        Right =176
        Bottom =122
        Top =0
        Name ="ClubOrderGuides"
        Name =""
    End
End
