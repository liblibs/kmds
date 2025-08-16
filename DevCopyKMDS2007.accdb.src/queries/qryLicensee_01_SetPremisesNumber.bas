Operation =4
Option =0
Where ="(((LocLicensee.PremisesNumberImported) Like \"000*\"))"
Begin InputTables
    Name ="LocLicensee"
End
Begin OutputColumns
    Name ="LocLicensee.PremisesNumber"
    Expression ="CLng([loclicensee].[PremisesNumberImported])"
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="LocLicensee.PremisesNumberImported"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocLicensee.PremisesNumber"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =78
    Top =136
    Right =1254
    Bottom =842
    Left =-1
    Top =-1
    Right =1144
    Bottom =422
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =48
        Top =12
        Right =322
        Bottom =335
        Top =0
        Name ="LocLicensee"
        Name =""
    End
End
