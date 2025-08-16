Operation =3
Name ="Licensee"
Option =0
Where ="(((LocLicensee.PremisesNumber) Is Not Null) AND ((Licensee.PremisesNumber) Is Nu"
    "ll))"
Begin InputTables
    Name ="LocLicensee"
    Name ="Licensee"
End
Begin OutputColumns
    Name ="PremisesNumber"
    Expression ="LocLicensee.PremisesNumber"
    Name ="TradeName"
    Expression ="LocLicensee.TradeName"
    Name ="Street"
    Expression ="LocLicensee.Street"
    Name ="City"
    Expression ="LocLicensee.City"
    Name ="Zip"
    Expression ="LocLicensee.Zip"
End
Begin Joins
    LeftTable ="LocLicensee"
    RightTable ="Licensee"
    Expression ="LocLicensee.PremisesNumber = Licensee.PremisesNumber"
    Flag =2
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="LocLicensee.PremisesNumber"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocLicensee.TradeName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocLicensee.Street"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocLicensee.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="LocLicensee.Zip"
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
    Right =1193
    Bottom =181
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =38
        Top =6
        Right =252
        Bottom =178
        Top =0
        Name ="LocLicensee"
        Name =""
    End
    Begin
        Left =313
        Top =11
        Right =409
        Bottom =171
        Top =0
        Name ="Licensee"
        Name =""
    End
End
