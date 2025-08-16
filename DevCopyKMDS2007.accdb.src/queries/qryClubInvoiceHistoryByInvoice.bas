Operation =1
Option =0
Where ="(((Club.Name) Between [Forms]![PrintInventoryReport]![cboClubIDBegin] And [Forms"
    "]![PrintInventoryReport]![cboClubIDEnd]))"
Having ="(((InvoLine.TransType)=\"C\") And ((InvoLine.SetupDate) Between Forms!PrintInven"
    "toryReport!txtDateRangeFrom And Forms!PrintInventoryReport!txtDateRangeTo))"
Begin InputTables
    Name ="InvoLine"
    Name ="Club"
End
Begin OutputColumns
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.SetupDate"
    Alias ="ClubID"
    Expression ="InvoLine.SalesDept"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Club"
    Expression ="InvoLine.SalesDept = Club.ID"
    Flag =1
End
Begin Groups
    Expression ="InvoLine.TransType"
    GroupLevel =0
    Expression ="InvoLine.InvoHeadID"
    GroupLevel =0
    Expression ="InvoLine.CashSysID"
    GroupLevel =0
    Expression ="InvoLine.SetupDate"
    GroupLevel =0
    Expression ="InvoLine.SalesDept"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbByte "RecordsetType" ="0"
dbBoolean "FilterOnLoad" ="0"
dbBoolean "OrderByOnLoad" ="-1"
dbBoolean "TotalsRow" ="0"
Begin
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubID"
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
    Bottom =110
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =89
        Top =3
        Right =185
        Bottom =117
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =251
        Top =-9
        Right =347
        Bottom =105
        Top =0
        Name ="Club"
        Name =""
    End
End
