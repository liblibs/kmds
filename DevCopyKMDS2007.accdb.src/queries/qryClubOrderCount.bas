Operation =1
Option =0
Where ="(((Club.ID)=[parClubID]))"
Begin InputTables
    Name ="Club"
End
Begin OutputColumns
    Expression ="Club.ID"
    Expression ="Club.OrderCount"
    Expression ="Club.NonLiquor"
    Expression ="Club.ClubNotes"
    Expression ="Club.FintechAccount"
    Expression ="Club.Licensee"
End
Begin Parameters
    Name ="parClubID"
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
        dbText "Name" ="Club.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.FintechAccount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Licensee"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =-277
    Top =68
    Right =948
    Bottom =745
    Left =-1
    Top =-1
    Right =1202
    Bottom =210
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =224
        Top =6
        Right =452
        Bottom =184
        Top =0
        Name ="Club"
        Name =""
    End
End
