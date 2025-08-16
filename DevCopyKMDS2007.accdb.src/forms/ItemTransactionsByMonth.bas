Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    ViewsAllowed =1
    TabularFamily =119
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12060
    DatasheetFontHeight =10
    ItemSuffix =28
    Left =-480
    Top =270
    Right =13305
    Bottom =11640
    Filter ="((qryInvoLineByItemByTransByMonth.[Reg#]=3))"
    RecSrcDt = Begin
        0xb20afd40b142e340
    End
    RecordSource ="qryInvoLineByItemByTransByMonth"
    Caption ="Transactions By Month"
    DatasheetFontName ="Trebuchet MS"
    AllowDatasheetView =0
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
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
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            Width =900
            Height =255
            LabelX =-1620
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =255
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Width =960
                    Height =255
                    Name ="Code_Label"
                    Caption ="Code"
                End
                Begin Label
                    OverlapFlags =87
                    Left =960
                    Width =2310
                    Height =255
                    Name ="Description_Label"
                    Caption ="Description"
                End
                Begin Label
                    OverlapFlags =93
                    Left =3300
                    Width =1320
                    Height =240
                    Name ="InputCode_Label"
                    Caption ="InputCode"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =4620
                    Width =1380
                    Height =240
                    Name ="Cashier_Label"
                    Caption ="Cashier"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =6000
                    Width =600
                    Height =240
                    Name ="Reg#_Label"
                    Caption ="Reg#"
                    EventProcPrefix ="Reg__Label"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =6660
                    Width =900
                    Height =240
                    Name ="Inv#_Label"
                    Caption ="Inv#"
                    EventProcPrefix ="Inv__Label"
                End
                Begin Label
                    OverlapFlags =87
                    Left =7560
                    Width =1860
                    Height =240
                    Name ="SetupDate_Label"
                    Caption ="SetupDate"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =9600
                    Width =900
                    Height =240
                    Name ="Quantity_Label"
                    Caption ="Quantity"
                    LayoutCachedLeft =9600
                    LayoutCachedWidth =10500
                    LayoutCachedHeight =240
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =3
                    Left =10320
                    Width =1260
                    Height =240
                    Name ="Price_Label"
                    Caption ="Price"
                End
            End
        End
        Begin Section
            Height =270
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Width =1320
                    Height =240
                    ColumnWidth =900
                    Name ="Code"
                    ControlSource ="Code"
                    BeforeUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Click Here to see the source transaction"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =1380
                    Width =2310
                    Height =240
                    ColumnWidth =1350
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =1380
                    LayoutCachedWidth =3690
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =3720
                    Width =1320
                    Height =240
                    ColumnWidth =1710
                    TabIndex =2
                    Name ="InputCode"
                    ControlSource ="InputCode"

                    LayoutCachedLeft =3720
                    LayoutCachedWidth =5040
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =95
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5040
                    Width =1380
                    Height =240
                    ColumnWidth =840
                    TabIndex =3
                    Name ="Cashier"
                    ControlSource ="Cashier"

                    LayoutCachedLeft =5040
                    LayoutCachedWidth =6420
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6420
                    Width =600
                    Height =240
                    ColumnWidth =615
                    TabIndex =4
                    Name ="Reg#"
                    ControlSource ="Reg#"
                    EventProcPrefix ="Reg_"

                    LayoutCachedLeft =6420
                    LayoutCachedWidth =7020
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7080
                    Height =240
                    ColumnWidth =810
                    TabIndex =5
                    Name ="Inv#"
                    ControlSource ="Inv#"
                    EventProcPrefix ="Inv_"

                    LayoutCachedLeft =7080
                    LayoutCachedWidth =7980
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =1
                    IMESentenceMode =3
                    Left =7980
                    Width =2160
                    Height =240
                    ColumnWidth =2385
                    TabIndex =6
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="General Date"

                    LayoutCachedLeft =7980
                    LayoutCachedWidth =10140
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10200
                    Width =720
                    Height =240
                    ColumnWidth =930
                    TabIndex =7
                    BackColor =13434828
                    Name ="Quantity"
                    ControlSource ="Quantity"

                    LayoutCachedLeft =10200
                    LayoutCachedWidth =10920
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10920
                    Width =1080
                    Height =240
                    ColumnWidth =990
                    TabIndex =8
                    Name ="Price"
                    ControlSource ="Price"

                    LayoutCachedLeft =10920
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =240
                End
            End
        End
        Begin FormFooter
            Height =606
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    Width =1440
                    Height =240
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Width =990
                            Height =240
                            Name ="Label18"
                            Caption ="InvoHeadID:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Width =1440
                    Height =240
                    TabIndex =1
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =2580
                            Width =285
                            Height =240
                            Name ="Label19"
                            Caption ="ID:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =7080
                    Width =300
                    Height =240
                    TabIndex =2
                    Name ="CashSysID"
                    ControlSource ="CashSysID"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =6120
                            Width =900
                            Height =240
                            Name ="Label20"
                            Caption ="CashSysID:"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5340
                    Width =720
                    Height =240
                    TabIndex =3
                    Name ="TransType"
                    ControlSource ="TransType"

                    Begin
                        Begin Label
                            Visible = NotDefault
                            OverlapFlags =85
                            Left =4440
                            Width =885
                            Height =240
                            Name ="Label21"
                            Caption ="TransType:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =9420
                    Top =60
                    Height =240
                    FontWeight =700
                    TabIndex =4
                    BackColor =13434828
                    Name ="txtQ\015\012txtQTotal"
                    ControlSource ="=Sum([Quantity])"
                    EventProcPrefix ="txtQ__txtQTotal"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =8280
                            Top =60
                            Width =1080
                            Height =300
                            FontWeight =700
                            Name ="Label27"
                            Caption ="Total *"
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =2
                    OverlapFlags =87
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10320
                    Top =60
                    Width =1260
                    Height =240
                    FontWeight =700
                    TabIndex =5
                    Name ="Text26"
                    ControlSource ="=Sum([Price])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =7560
                    Width =666
                    Height =606
                    TabIndex =6
                    Name ="cmdClose"
                    Caption ="Command64"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x07070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f90707f9f9f9f9fffffffff9f9f9fffff9f9f9fffffffff9f9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9fffffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9fffffff9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfffffffff9f9f90707f9f9f9f9fffffff9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffffff9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9f9fffffffff9fffffffffffff9fffffffff9f9ff ,
                        0xfffffffff9f9f90707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"

                End
            End
        End
    End
End
CodeBehindForm
' See "ItemTransactionsByMonth.cls"
