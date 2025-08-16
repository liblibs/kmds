Operation =1
Option =0
Where ="(((ClubOrderGuides.ClubID)=[Forms]![ClubOrderWebImport]![cmbLicensee]))"
Begin InputTables
    Name ="ClubOrderGuides"
    Name ="Club"
End
Begin OutputColumns
    Expression ="ClubOrderGuides.ClubID"
    Expression ="ClubOrderGuides.OrderID"
    Expression ="ClubOrderGuides.LineID"
    Expression ="ClubOrderGuides.ItemID"
    Expression ="ClubOrderGuides.Description"
    Expression ="ClubOrderGuides.Par"
    Expression ="ClubOrderGuides.QuotedPrice"
    Expression ="Club.NonLiquor"
    Expression ="ClubOrderGuides.LiquorItem"
    Expression ="Club.Contact"
    Expression ="Club.MaintainParLevels"
End
Begin Joins
    LeftTable ="ClubOrderGuides"
    RightTable ="Club"
    Expression ="ClubOrderGuides.ClubID = Club.ID"
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
        dbText "Name" ="ClubOrderGuides.ClubID"
        dbLong "AggregateType" ="-1"
    End
End
Begin
    State =0
    Left =42
    Top =120
    Right =1495
    Bottom =873
    Left =-1
    Top =-1
    Right =1430
    Bottom =467
    Left =0
    Top =0
    ColumnsShown =539
    Begin
        Left =48
        Top =12
        Right =192
        Bottom =156
        Top =0
        Name ="ClubOrderGuides"
        Name =""
    End
    Begin
        Left =240
        Top =12
        Right =384
        Bottom =156
        Top =0
        Name ="Club"
        Name =""
    End
End
