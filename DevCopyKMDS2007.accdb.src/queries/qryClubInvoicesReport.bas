Operation =1
Option =0
Begin InputTables
    Name ="qryClubInvoices"
    Name ="InvoLineAllRegisters"
    Name ="Club"
End
Begin OutputColumns
    Expression ="Club.NonLiquor"
    Expression ="qryClubInvoices.ClubID"
    Alias ="ClubName"
    Expression ="Club.Name"
    Expression ="InvoLineAllRegisters.InvoHeadID"
    Alias ="Qty"
    Expression ="Sum(IIf(InvoLineAllRegisters.TransType='C',[Quantity],0))"
    Alias ="Gross"
    Expression ="Sum(IIf(InvoLineAllRegisters.TransType='C',[Price],0))"
    Alias ="Discount"
    Expression ="Sum(IIf(InvoLineAllRegisters.TransType='D',[Price],0))"
End
Begin Joins
    LeftTable ="qryClubInvoices"
    RightTable ="InvoLineAllRegisters"
    Expression ="qryClubInvoices.InvoHeadID = InvoLineAllRegisters.InvoHeadID"
    Flag =1
    LeftTable ="qryClubInvoices"
    RightTable ="Club"
    Expression ="qryClubInvoices.ClubID = Club.ID"
    Flag =1
End
Begin Groups
    Expression ="Club.NonLiquor"
    GroupLevel =0
    Expression ="qryClubInvoices.ClubID"
    GroupLevel =0
    Expression ="Club.Name"
    GroupLevel =0
    Expression ="InvoLineAllRegisters.InvoHeadID"
    GroupLevel =0
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
        dbText "Name" ="ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Gross"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Discount"
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
        Left =231
        Top =9
        Right =327
        Bottom =108
        Top =0
        Name ="qryClubInvoices"
        Name =""
    End
    Begin
        Left =470
        Top =-1
        Right =566
        Bottom =113
        Top =0
        Name ="InvoLineAllRegisters"
        Name =""
    End
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =120
        Top =0
        Name ="Club"
        Name =""
    End
End
