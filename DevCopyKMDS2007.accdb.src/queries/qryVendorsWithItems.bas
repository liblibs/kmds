Operation =1
Option =0
Where ="(((dbo_Vendor.POMethodDefault)<>-1))"
Begin InputTables
    Name ="dbo_Vendor"
    Name ="qryItemByVendorCount"
End
Begin OutputColumns
    Expression ="dbo_Vendor.ID"
    Expression ="dbo_Vendor.Name"
    Expression ="qryItemByVendorCount.Count"
    Alias ="IsNotObsolete"
    Expression ="dbo_Vendor.POMethodDefault"
End
Begin Joins
    LeftTable ="dbo_Vendor"
    RightTable ="qryItemByVendorCount"
    Expression ="dbo_Vendor.ID=qryItemByVendorCount.VendorID"
    Flag =1
End
Begin OrderBy
    Expression ="dbo_Vendor.Name"
    Flag =0
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
    Begin
        dbText "Name" ="dbo_Vendor.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="qryItemByVendorCount.count"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="IsNotObsolete"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="dbo_Vendor.Name"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =0
    Top =40
    Right =1374
    Bottom =846
    Left =-1
    Top =-1
    Right =1342
    Bottom =471
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =215
        Bottom =414
        Top =0
        Name ="dbo_Vendor"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="qryItemByVendorCount"
        Name =""
    End
End
