Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =7920
    DatasheetFontHeight =10
    ItemSuffix =34
    RecSrcDt = Begin
        0xbe84cf94f69de240
    End
    RecordSource ="Cashier"
    Caption ="Cashier1"
    DatasheetFontName ="Arial"
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
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
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
            Height =360
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =240
                    Height =240
                    Name ="ID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =360
                    Top =60
                    Width =480
                    Height =240
                    Name ="Name_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =900
                    Top =60
                    Width =420
                    Height =240
                    Name ="AllowCancelSale_Label"
                    Caption ="AllowCancelSale"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1380
                    Top =60
                    Width =420
                    Height =240
                    Name ="AllowCashDrops_Label"
                    Caption ="AllowCashDrops"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =1860
                    Top =60
                    Width =420
                    Height =240
                    Name ="AllowCashLoans_Label"
                    Caption ="AllowCashLoans"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2340
                    Top =60
                    Width =360
                    Height =240
                    Name ="AllowCoupons_Label"
                    Caption ="AllowCoupons"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =2760
                    Top =60
                    Width =480
                    Height =240
                    Name ="AllowDiscountItem_Label"
                    Caption ="AllowDiscountItem"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3300
                    Top =60
                    Width =480
                    Height =240
                    Name ="AllowLotteryPaidOut_Label"
                    Caption ="AllowLotteryPaidOut"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =3840
                    Top =60
                    Width =360
                    Height =240
                    Name ="AllowNegSales_Label"
                    Caption ="AllowNegSales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4260
                    Top =60
                    Width =360
                    Height =240
                    Name ="AllowNewItems_Label"
                    Caption ="AllowNewItems"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =4680
                    Top =60
                    Width =360
                    Height =240
                    Name ="AllowNoSales_Label"
                    Caption ="AllowNoSales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5100
                    Top =60
                    Width =360
                    Height =240
                    Name ="AllowPaidouts_Label"
                    Caption ="AllowPaidouts"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5520
                    Top =60
                    Width =360
                    Height =240
                    Name ="AllowSellClub_Label"
                    Caption ="AllowSellClub"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =5940
                    Top =60
                    Width =420
                    Height =240
                    Name ="AllowShutDown_Label"
                    Caption ="AllowShutDown"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6420
                    Top =60
                    Width =420
                    Height =240
                    Name ="AllowViewJournal_Label"
                    Caption ="AllowViewJournal"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =6900
                    Top =60
                    Width =360
                    Height =240
                    Name ="AllowXReports_Label"
                    Caption ="AllowXReports"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =7320
                    Top =60
                    Width =540
                    Height =240
                    Name ="AllowZReports_Label"
                    Caption ="AllowZReports"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =375
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =240
                    Height =255
                    ColumnWidth =600
                    Name ="ID"
                    ControlSource ="ID"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =360
                    Top =60
                    Width =480
                    Height =255
                    ColumnWidth =1560
                    TabIndex =1
                    Name ="Name"
                    ControlSource ="Name"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001010000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010100000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =900
                    Top =60
                    TabIndex =2
                    Name ="AllowCancelSale"
                    ControlSource ="AllowCancelSale"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1380
                    Top =60
                    TabIndex =3
                    Name ="AllowCashDrops"
                    ControlSource ="AllowCashDrops"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1860
                    Top =60
                    TabIndex =4
                    Name ="AllowCashLoans"
                    ControlSource ="AllowCashLoans"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2340
                    Top =60
                    TabIndex =5
                    Name ="AllowCoupons"
                    ControlSource ="AllowCoupons"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2760
                    Top =60
                    TabIndex =6
                    Name ="AllowDiscountItem"
                    ControlSource ="AllowDiscountItem"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3300
                    Top =60
                    TabIndex =7
                    Name ="AllowLotteryPaidOut"
                    ControlSource ="AllowLotteryPaidOut"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3840
                    Top =60
                    TabIndex =8
                    Name ="AllowNegSales"
                    ControlSource ="AllowNegSales"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4260
                    Top =60
                    TabIndex =9
                    Name ="AllowNewItems"
                    ControlSource ="AllowNewItems"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4680
                    Top =60
                    TabIndex =10
                    Name ="AllowNoSales"
                    ControlSource ="AllowNoSales"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5100
                    Top =60
                    TabIndex =11
                    Name ="AllowPaidouts"
                    ControlSource ="AllowPaidouts"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5520
                    Top =60
                    TabIndex =12
                    Name ="AllowSellClub"
                    ControlSource ="AllowSellClub"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5940
                    Top =60
                    TabIndex =13
                    Name ="AllowShutDown"
                    ControlSource ="AllowShutDown"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6420
                    Top =60
                    TabIndex =14
                    Name ="AllowViewJournal"
                    ControlSource ="AllowViewJournal"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6900
                    Top =60
                    TabIndex =15
                    Name ="AllowXReports"
                    ControlSource ="AllowXReports"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7320
                    Top =60
                    TabIndex =16
                    Name ="AllowZReports"
                    ControlSource ="AllowZReports"

                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
