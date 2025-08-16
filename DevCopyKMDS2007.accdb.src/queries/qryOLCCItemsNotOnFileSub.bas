Operation =1
Option =0
Where ="(((Item.ID) Is Null))"
Begin InputTables
    Name ="dbo_OLCCitem"
    Name ="Item"
    Name ="OLCC Categories"
    Name ="qrydbo_OLCCScanCode_12Or8"
    Name ="SubCategory"
End
Begin OutputColumns
    Expression ="dbo_OLCCitem.Code"
    Expression ="dbo_OLCCitem.Description"
    Expression ="dbo_OLCCitem.Price"
    Expression ="dbo_OLCCitem.UnitsPerCase"
    Expression ="[OLCC Categories].Category"
    Expression ="dbo_OLCCitem.StockStatus"
    Alias ="SizeID"
    Expression ="Right([dbo_OLCCitem.Code],2)"
    Expression ="dbo_OLCCitem.SetupDate"
    Expression ="qrydbo_OLCCScanCode_12Or8.UPCCodeKMDS"
    Alias ="SubCategoryID"
    Expression ="SubCategory.ID"
End
Begin Joins
    LeftTable ="dbo_OLCCitem"
    RightTable ="Item"
    Expression ="dbo_OLCCitem.Code = Item.ID"
    Flag =2
    LeftTable ="dbo_OLCCitem"
    RightTable ="OLCC Categories"
    Expression ="dbo_OLCCitem.Code = [OLCC Categories].Code"
    Flag =2
    LeftTable ="dbo_OLCCitem"
    RightTable ="qrydbo_OLCCScanCode_12Or8"
    Expression ="dbo_OLCCitem.Code = qrydbo_OLCCScanCode_12Or8.Code"
    Flag =2
    LeftTable ="OLCC Categories"
    RightTable ="SubCategory"
    Expression ="[OLCC Categories].Category = SubCategory.Description"
    Flag =2
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
        dbText "Name" ="dbo_OLCCitem.Code"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SizeID"
        dbLong "AggregateType" ="-1"
    End
    Begin
        dbText "Name" ="SubCategoryID"
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
    Bottom =186
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="dbo_OLCCitem"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Item"
        Name =""
    End
    Begin
        Left =432
        Top =12
        Right =576
        Bottom =156
        Top =0
        Name ="OLCC Categories"
        Name =""
    End
    Begin
        Left =624
        Top =12
        Right =768
        Bottom =156
        Top =0
        Name ="qrydbo_OLCCScanCode_12Or8"
        Name =""
    End
    Begin
        Left =816
        Top =12
        Right =960
        Bottom =156
        Top =0
        Name ="SubCategory"
        Name =""
    End
End
