Operation =1
Option =0
Where ="(((Vendor.POMethodDefault)>-1))"
Begin InputTables
    Name ="Vendor"
End
Begin OutputColumns
    Expression ="Vendor.*"
End
Begin OrderBy
    Expression ="Vendor.Name"
    Flag =0
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
        dbText "Name" ="Vendor.ID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Name"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Contact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Address1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Address2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.City"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.St"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Zip"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Phone1"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Phone2"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.Phone3"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.OrderCount"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.PoAtCost"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.PrintUpcOnPo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.POReportForm"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.VendorNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.AverageDaySupplyID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.ChangeDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.upsize_ts"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Vendor.POMethodDefault"
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
    Bottom =251
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =269
        Top =0
        Name ="Vendor"
        Name =""
    End
End
