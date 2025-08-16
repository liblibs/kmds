Operation =2
Name ="HourlySales"
Option =0
Where ="(((Item.SalesDepartmentID)=1 Or (Item.SalesDepartmentID)=2) AND ((InvoLine.Trans"
    "Type)=\"S\" Or (InvoLine.TransType)=\"C\" Or (InvoLine.TransType)=\"D\") AND ((I"
    "nvoLine.SetupDate) Between [BeginDate] And [EndDate]))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
End
Begin OutputColumns
    Alias ="Date"
    Expression ="Format(InvoLine.SetupDate,\"mm/dd/yyyy\")"
    Alias ="SumOfPrice"
    Expression ="Sum(InvoLine.Price)"
    Alias ="10 AM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=10,InvoLine.PRICE,0))"
    Alias ="11 AM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=11,InvoLine.PRICE,0))"
    Alias ="12 AM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=12,InvoLine.PRICE,0))"
    Alias ="1 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=13,InvoLine.PRICE,0))"
    Alias ="2 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=14,InvoLine.PRICE,0))"
    Alias ="3 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=15,InvoLine.PRICE,0))"
    Alias ="4 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=16,InvoLine.PRICE,0))"
    Alias ="5 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=17,InvoLine.PRICE,0))"
    Alias ="6 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=18,InvoLine.PRICE,0))"
    Alias ="7 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=19,InvoLine.PRICE,0))"
    Alias ="8 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=20,InvoLine.PRICE,0))"
    Alias ="9 PM"
    Expression ="Sum(IIf(Format(InvoLine.SetupDate,\"hh\")=21,InvoLine.PRICE,0))"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code = Item.ID"
    Flag =1
End
Begin Groups
    Expression ="Format(InvoLine.SetupDate,\"mm/dd/yyyy\")"
    GroupLevel =0
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbByte "RecordsetType" ="0"
dbBoolean "OrderByOn" ="0"
dbByte "Orientation" ="0"
dbByte "DefaultView" ="2"
dbBoolean "UseTransaction" ="-1"
Begin
    Begin
        dbText "Name" ="Date"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SumOfPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="10 AM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="11 AM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="12 AM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="1 PM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="2 PM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="3 PM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="4 PM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="5 PM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="6 PM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="7 PM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="8 PM"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="9 PM"
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
    Bottom =261
    Left =0
    Top =0
    ColumnsShown =543
    Begin
        Left =38
        Top =6
        Right =144
        Bottom =240
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =178
        Top =33
        Right =405
        Bottom =147
        Top =0
        Name ="Item"
        Name =""
    End
End
