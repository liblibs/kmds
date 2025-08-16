Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9240
    DatasheetFontHeight =10
    ItemSuffix =29
    Left =7095
    Top =2685
    Right =18000
    Bottom =8925
    DatasheetBackColor =65535
    Filter ="[ItemID]=99900010875"
    RecSrcDt = Begin
        0xe3b7e7649bf5e340
    End
    RecordSource ="qryItemOnTransfer"
    Caption ="Items on order for "
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
    DatasheetBackColor12 =65535
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextAlign =3
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin Rectangle
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Trebuchet MS"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontSize =9
            FontName ="Trebuchet MS"
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontSize =9
            FontName ="Trebuchet MS"
        End
        Begin Subform
            OldBorderStyle =0
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin ToggleButton
            FontSize =9
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Trebuchet MS"
            BorderLineStyle =0
        End
        Begin Tab
            FontSize =9
            FontName ="Trebuchet MS"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =4935
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1980
                    Top =180
                    ColumnWidth =975
                    ColumnOrder =2
                    Name ="VendorID"
                    ControlSource ="AgentNumber"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =540
                            Top =180
                            Width =885
                            Height =315
                            Name ="Label16"
                            Caption ="Agent"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =1980
                    Top =480
                    ColumnWidth =1995
                    ColumnOrder =3
                    TabIndex =1
                    Name ="Name"
                    ControlSource ="LastName"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =0
                            Left =540
                            Top =480
                            Width =585
                            Height =315
                            Name ="Label17"
                            Caption ="Name:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1980
                    Top =840
                    ColumnWidth =630
                    ColumnOrder =4
                    TabIndex =2
                    Name ="PurchaseOrderID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =240
                            Top =840
                            Width =1545
                            Height =315
                            Name ="Label18"
                            Caption ="PO #"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1980
                    Top =2100
                    ColumnWidth =1335
                    ColumnOrder =5
                    TabIndex =3
                    Name ="OrderDate"
                    ControlSource ="OrdDate"
                    Format ="Short Date"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =540
                            Top =2100
                            Width =1005
                            Height =315
                            Name ="Label19"
                            Caption ="Date:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =2520
                    ColumnWidth =810
                    ColumnOrder =6
                    FontSize =8
                    TabIndex =4
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =600
                            Top =2520
                            Width =645
                            Height =315
                            Name ="Label20"
                            Caption ="ItemID:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =2820
                    ColumnWidth =495
                    ColumnOrder =1
                    TabIndex =5
                    Name ="QuantityShipped"
                    ControlSource ="QuantityOrdered"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =0
                            Left =600
                            Top =2820
                            Width =1515
                            Height =315
                            Name ="Label21"
                            Caption ="Qty"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =3120
                    ColumnWidth =690
                    ColumnOrder =7
                    TabIndex =6
                    Name ="Price"
                    ControlSource ="Price"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =0
                            Left =600
                            Top =3120
                            Width =570
                            Height =315
                            Name ="Label22"
                            Caption ="Price:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =3420
                    ColumnWidth =765
                    ColumnOrder =8
                    TabIndex =7
                    Name ="Cost"
                    ControlSource ="Cost"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =0
                            Left =600
                            Top =3420
                            Width =495
                            Height =315
                            Name ="Label23"
                            Caption ="Cost:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =4320
                    ColumnOrder =10
                    TabIndex =8
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =0
                            Left =600
                            Top =4320
                            Width =825
                            Height =315
                            Name ="Label26"
                            Caption ="ExtPrice:"
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =4620
                    ColumnWidth =0
                    ColumnOrder =11
                    TabIndex =9
                    Name ="ExtCost"
                    ControlSource ="ExtCost"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =0
                            Left =600
                            Top =4620
                            Width =750
                            Height =315
                            Name ="Label27"
                            Caption ="ExtCost:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1980
                    Top =1620
                    ColumnWidth =915
                    ColumnOrder =0
                    TabIndex =10
                    BackColor =16711935
                    Name ="InOrOut"
                    ControlSource ="InOrOut"
                    ConditionalFormat = Begin
                        0x010000006c000000010000000000000002000000000000000500000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x220049006e00220000000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000000000002000000010000000000000000ff0000040000002200 ,
                        0x49006e002200000000000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =540
                            Top =1620
                            Width =810
                            Height =315
                            Name ="Label28"
                            Caption ="InOrOut:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =375
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2040
                    Top =60
                    ColumnWidth =1215
                    ColumnOrder =9
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="Short Time"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =60
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =0
                            Left =600
                            Top =60
                            Width =1110
                            Height =315
                            Name ="Label25"
                            Caption ="ChangeTim:"
                            LayoutCachedLeft =600
                            LayoutCachedTop =60
                            LayoutCachedWidth =1710
                            LayoutCachedHeight =375
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ItemOnTransfer.cls"
