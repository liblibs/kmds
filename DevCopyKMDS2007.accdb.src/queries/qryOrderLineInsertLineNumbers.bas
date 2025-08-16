Operation =4
Option =0
Where ="(((qryOrderLineByLineIDDescending.LineID)>=[parLineID]) AND ((qryOrderLineByLine"
    "IDDescending.OrdheadID)=[parOrdheadID]))"
Begin InputTables
    Name ="qryOrderLineByLineIDDescending"
End
Begin OutputColumns
    Name ="qryOrderLineByLineIDDescending.LineID"
    Expression ="[LineID]+1"
End
Begin Parameters
    Name ="parOrdheadID"
    Flag =4
    Name ="parLineID"
    Flag =3
End
dbBoolean "ReturnsRecords" ="-1"
dbInteger "ODBCTimeout" ="60"
dbBoolean "UseTransaction" ="-1"
dbBoolean "FailOnError" ="0"
dbByte "Orientation" ="0"
Begin
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
    ColumnsShown =579
    Begin
        Left =48
        Top =6
        Right =144
        Bottom =105
        Top =0
        Name ="qryOrderLineByLineIDDescending"
        Name =""
    End
End
