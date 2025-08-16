Operation =1
Option =0
Begin InputTables
    Name ="vw_WebLogErrorCode-1"
    Name ="OrdHead"
    Name ="Club"
End
Begin OutputColumns
    Expression ="[vw_WebLogErrorCode-1].SetupDate"
    Alias ="Invoice"
    Expression ="Trim([WebOrderID])"
    Alias ="OrderNotes"
    Expression ="Trim([WebOrderNotes])"
    Alias ="Customer"
    Expression ="Trim([LicenseeName])"
    Alias ="Error"
    Expression ="Trim([ErrorCode])"
    Alias ="Qty"
    Expression ="Trim([Source])"
    Expression ="OrdHead.InvoiceNo"
    Expression ="OrdHead.InvQty"
    Expression ="OrdHead.InvDollars"
    Alias ="eMail"
    Expression ="Trim([Subroutine])"
    Expression ="Club.Contact"
    Expression ="Club.Phone1"
End
Begin Joins
    LeftTable ="vw_WebLogErrorCode-1"
    RightTable ="OrdHead"
    Expression ="[vw_WebLogErrorCode-1].WebOrderID = OrdHead.InvoiceNo"
    Flag =1
    LeftTable ="OrdHead"
    RightTable ="Club"
    Expression ="OrdHead.ClubID = Club.ID"
    Flag =1
End
Begin OrderBy
    Expression ="[vw_WebLogErrorCode-1].SetupDate"
    Flag =1
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
        dbText "Name" ="[vw_WebLogErrorCode-1].SetupDate"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Invoice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrderNotes"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Customer"
        dbInteger "ColumnWidth" ="2595"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Error"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvoiceNo"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvQty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="OrdHead.InvDollars"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="eMail"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Contact"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Club.Phone1"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =21
    Top =80
    Right =1463
    Bottom =873
    Left =-1
    Top =-1
    Right =1419
    Bottom =459
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =244
        Top =0
        Name ="vw_WebLogErrorCode-1"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =313
        Top =0
        Name ="OrdHead"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =357
        Top =0
        Name ="Club"
        Name =""
    End
End
