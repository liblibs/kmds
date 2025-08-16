Operation =4
Option =0
Where ="(((Licensee.TradeName)<>LocLicensee.TradeName)) Or (((Licensee.Street)<>LocLicen"
    "see.Street)) Or (((Licensee.City)<>LocLicensee.City)) Or (((Licensee.Zip)<>LocLi"
    "censee.Zip))"
Begin InputTables
    Name ="Licensee"
    Name ="LocLicensee"
End
Begin OutputColumns
    Name ="Licensee.TradeName"
    Expression ="LocLicensee.TradeName"
    Name ="Licensee.Street"
    Expression ="LocLicensee.Street"
    Name ="Licensee.City"
    Expression ="LocLicensee.City"
    Name ="Licensee.Zip"
    Expression ="LocLicensee.Zip"
End
Begin Joins
    LeftTable ="Licensee"
    RightTable ="LocLicensee"
    Expression ="Licensee.PremisesNumber = LocLicensee.PremisesNumber"
    Flag =1
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
    Bottom =151
    Left =0
    Top =0
    ColumnsShown =579
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Licensee"
        Name =""
    End
    Begin
        Left =172
        Top =6
        Right =268
        Bottom =120
        Top =0
        Name ="LocLicensee"
        Name =""
    End
End
