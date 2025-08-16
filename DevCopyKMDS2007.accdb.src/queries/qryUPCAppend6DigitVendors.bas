Operation =3
Name ="upcitem"
Option =0
Begin InputTables
    Name ="qryUPCVendor5"
    Name ="UPCConvert5to6"
End
Begin OutputColumns
    Name ="ItemID"
    Expression ="qryUPCVendor5.ItemID"
    Alias ="NewUPC"
    Name ="ID"
    Expression ="checkdigit11([Vendor6Digit] & Mid([id],7,5))"
End
Begin Joins
    LeftTable ="qryUPCVendor5"
    RightTable ="UPCConvert5to6"
    Expression ="qryUPCVendor5.Vendor5U = UPCConvert5to6.Vendor5Digit"
    Flag =1
End
Begin OrderBy
    Expression ="checkdigit11([Vendor6Digit] & Mid([id],7,5))"
    Flag =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbByte "Orientation" ="0"
Begin
    Begin
        dbText "Name" ="qryUPCVendor5.ItemID"
        dbInteger "ColumnWidth" ="765"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="NewUPC"
        dbInteger "ColumnWidth" ="2580"
        dbBoolean "ColumnHidden" ="0"
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
    Bottom =94
    Left =0
    Top =0
    ColumnsShown =651
    Begin
        Left =48
        Top =11
        Right =226
        Bottom =125
        Top =0
        Name ="qryUPCVendor5"
        Name =""
    End
    Begin
        Left =266
        Top =19
        Right =362
        Bottom =103
        Top =0
        Name ="UPCConvert5to6"
        Name =""
    End
End
