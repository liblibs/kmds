Operation =3
Name ="KMDSLinks"
Option =0
Where ="(((KMDSLinks.Caption) Is Null))"
Begin InputTables
    Name ="KMDSLinks"
    Name ="KMDSLinksLocal"
End
Begin OutputColumns
    Name ="Caption"
    Expression ="KMDSLinksLocal.Caption"
    Name ="URL"
    Expression ="KMDSLinksLocal.URL"
End
Begin Joins
    LeftTable ="KMDSLinks"
    RightTable ="KMDSLinksLocal"
    Expression ="KMDSLinks.Caption = KMDSLinksLocal.Caption"
    Flag =3
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
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="KMDSLinks.[Caption]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KMDSLinks.[URL]"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KMDSLinksLocal.Caption"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="KMDSLinksLocal.URL"
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
    Bottom =488
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =421
        Top =36
        Right =611
        Bottom =180
        Top =0
        Name ="KMDSLinks"
        Name =""
    End
    Begin
        Left =163
        Top =27
        Right =307
        Bottom =171
        Top =0
        Name ="KMDSLinksLocal"
        Name =""
    End
End
