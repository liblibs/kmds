Operation =1
Option =0
Where ="(((InvoLine.InvoHeadID)=Forms!ItemTransactionsByMonth!InvoHeadID) And ((InvoLine"
    ".CashSysID)=Forms!ItemTransactionsByMonth!CashSysID) And ((InvoLine.SetupDate)=F"
    "orms!ItemTransactionsByMonth!SetupDate))"
Begin InputTables
    Name ="InvoLine"
    Name ="Item"
    Name ="Cashier"
    Name ="Club"
    Name ="Tender"
    Name ="Agents"
    Name ="TranTypes"
End
Begin OutputColumns
    Expression ="InvoLine.InvoHeadID"
    Expression ="InvoLine.CashSysID"
    Expression ="InvoLine.SetupDate"
    Expression ="InvoLine.ID"
    Expression ="InvoLine.Code"
    Alias ="CodeDesc"
    Expression ="IIf([TransType]=\"T\",Tender.Description,Item.Description)"
    Alias ="Qty"
    Expression ="InvoLine.Quantity"
    Alias ="UPrice"
    Expression ="CCur(InvoLine.Price/IIf([quantity]=0,1,[quantity]))"
    Alias ="Extend"
    Expression ="IIf([TransType]=\"T\",0-InvoLine.Price,InvoLine.Price)"
    Expression ="InvoLine.SalesDept"
    Expression ="InvoLine.InputCode"
    Expression ="InvoLine.SalesGroup"
    Expression ="InvoLine.Drawer"
    Expression ="InvoLine.TransType"
    Expression ="InvoLine.UpdateOnhand"
    Alias ="CashierName"
    Expression ="Cashier.Name"
    Alias ="ClubName"
    Expression ="IIf([TransType]=\"C\",Club.Name,\"\")"
    Alias ="Transaction"
    Expression ="TranTypes.Description"
    Alias ="TransferNumber"
    Expression ="InvoLine.SalesGroup"
    Expression ="Agents.FirstName"
    Expression ="Agents.LastName"
End
Begin Joins
    LeftTable ="InvoLine"
    RightTable ="Item"
    Expression ="InvoLine.Code=Item.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Cashier"
    Expression ="InvoLine.CashierID=Cashier.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Club"
    Expression ="InvoLine.SalesDept=Club.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Tender"
    Expression ="InvoLine.Code=Tender.ID"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="Agents"
    Expression ="InvoLine.SalesDept=Agents.AgentNumber"
    Flag =2
    LeftTable ="InvoLine"
    RightTable ="TranTypes"
    Expression ="InvoLine.TransType=TranTypes.Code"
    Flag =2
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
        dbText "Name" ="UPrice"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Extend"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CodeDesc"
        dbInteger "ColumnWidth" ="2895"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Qty"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="InvoLine.InputCode"
        dbInteger "ColumnWidth" ="1755"
        dbBoolean "ColumnHidden" ="0"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="ClubName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="CashierName"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="Transaction"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="TransferNumber"
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
    Bottom =316
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =38
        Top =6
        Right =134
        Bottom =180
        Top =0
        Name ="InvoLine"
        Name =""
    End
    Begin
        Left =283
        Top =38
        Right =379
        Bottom =122
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =649
        Top =-11
        Right =745
        Bottom =103
        Top =0
        Name ="Cashier"
        Name =""
    End
    Begin
        Left =288
        Top =189
        Right =384
        Bottom =303
        Top =0
        Name ="Club"
        Name =""
    End
    Begin
        Left =284
        Top =102
        Right =380
        Bottom =186
        Top =0
        Name ="Tender"
        Name =""
    End
    Begin
        Left =414
        Top =147
        Right =510
        Bottom =261
        Top =0
        Name ="Agents"
        Name =""
    End
    Begin
        Left =158
        Top =191
        Right =254
        Bottom =290
        Top =0
        Name ="TranTypes"
        Name =""
    End
End
