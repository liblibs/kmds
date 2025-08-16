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
    Width =13980
    DatasheetFontHeight =10
    ItemSuffix =34
    Left =630
    Top =1800
    Right =630
    Bottom =1800
    OnUnload ="[Event Procedure]"
    RecSrcDt = Begin
        0xfe78d6a8f69de240
    End
    RecordSource ="Cashier"
    Caption ="Cashier Register Flags"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
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
            Height =660
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =60
                    Top =120
                    Width =420
                    Height =480
                    FontWeight =700
                    Name ="ID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =540
                    Top =120
                    Width =900
                    Height =480
                    FontWeight =700
                    Name ="Name_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =2
                    Left =1620
                    Top =120
                    Width =840
                    Height =480
                    FontWeight =700
                    Name ="AllowCancelSale_Label"
                    Caption ="Cancel Sale"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =2460
                    Top =120
                    Width =840
                    Height =480
                    FontWeight =700
                    Name ="AllowCashDrops_Label"
                    Caption ="Cash Drops"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =3300
                    Top =120
                    Width =840
                    Height =480
                    FontWeight =700
                    Name ="AllowCashLoans_Label"
                    Caption ="Cash Loans"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =4140
                    Top =120
                    Width =840
                    Height =480
                    FontWeight =700
                    Name ="AllowCoupons_Label"
                    Caption ="Coupons"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =4980
                    Top =120
                    Width =900
                    Height =480
                    FontWeight =700
                    Name ="AllowDiscountItem_Label"
                    Caption ="Discount Item"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =5880
                    Top =120
                    Width =900
                    Height =480
                    FontWeight =700
                    Name ="AllowLotteryPaidOut_Label"
                    Caption ="Lottery Paid Out"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =6780
                    Top =120
                    Width =780
                    Height =480
                    FontWeight =700
                    Name ="AllowNegSales_Label"
                    Caption ="Return Sales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =7560
                    Top =120
                    Width =780
                    Height =480
                    FontWeight =700
                    Name ="AllowNewItems_Label"
                    Caption ="Setup Items"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =8340
                    Top =120
                    Width =780
                    Height =480
                    FontWeight =700
                    Name ="AllowNoSales_Label"
                    Caption ="No Sales"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =9120
                    Top =120
                    Width =780
                    Height =480
                    FontWeight =700
                    Name ="AllowPaidouts_Label"
                    Caption ="Paid Outs"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =9900
                    Top =120
                    Width =780
                    Height =480
                    FontWeight =700
                    Name ="AllowSellClub_Label"
                    Caption ="Sell Club"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =10680
                    Top =120
                    Width =840
                    Height =480
                    FontWeight =700
                    Name ="AllowShutDown_Label"
                    Caption ="F2 Item Lookup"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =11520
                    Top =120
                    Width =840
                    Height =480
                    FontWeight =700
                    Name ="AllowViewJournal_Label"
                    Caption ="View Journal"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =95
                    TextAlign =2
                    Left =12360
                    Top =120
                    Width =780
                    Height =480
                    FontWeight =700
                    Name ="AllowXReports_Label"
                    Caption ="X Reports"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =2
                    Left =13140
                    Top =120
                    Width =840
                    Height =480
                    FontWeight =700
                    Name ="AllowZReports_Label"
                    Caption ="Z Reports"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin Section
            Height =315
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =240
                    Height =255
                    ColumnWidth =600
                    BackColor =13408767
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =480
                    Top =60
                    Width =1140
                    Height =255
                    ColumnWidth =1560
                    TabIndex =1
                    BackColor =13408767
                    Name ="Name"
                    ControlSource ="Name"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1920
                    Top =60
                    Width =365
                    TabIndex =2
                    Name ="AllowCancelSale"
                    ControlSource ="AllowCancelSale"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2760
                    Top =60
                    Width =365
                    TabIndex =3
                    Name ="AllowCashDrops"
                    ControlSource ="AllowCashDrops"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3660
                    Top =60
                    Width =365
                    TabIndex =4
                    Name ="AllowCashLoans"
                    ControlSource ="AllowCashLoans"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =4500
                    Top =60
                    Width =365
                    TabIndex =5
                    Name ="AllowCoupons"
                    ControlSource ="AllowCoupons"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5340
                    Top =60
                    Width =365
                    TabIndex =6
                    Name ="AllowDiscountItem"
                    ControlSource ="AllowDiscountItem"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =6240
                    Top =60
                    Width =365
                    TabIndex =7
                    Name ="AllowLotteryPaidOut"
                    ControlSource ="AllowLotteryPaidOut"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7140
                    Top =60
                    Width =365
                    TabIndex =8
                    Name ="AllowNegSales"
                    ControlSource ="AllowNegSales"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =7920
                    Top =60
                    Width =365
                    TabIndex =9
                    Name ="AllowNewItems"
                    ControlSource ="AllowNewItems"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =8700
                    Top =60
                    Width =365
                    TabIndex =10
                    Name ="AllowNoSales"
                    ControlSource ="AllowNoSales"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =9480
                    Top =60
                    Width =365
                    TabIndex =11
                    Name ="AllowPaidouts"
                    ControlSource ="AllowPaidouts"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =10260
                    Top =60
                    Width =365
                    TabIndex =12
                    Name ="AllowSellClub"
                    ControlSource ="AllowSellClub"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =11040
                    Top =60
                    Width =365
                    TabIndex =13
                    Name ="AllowShutDown"
                    ControlSource ="AllowShutDown"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =11873
                    Top =60
                    Width =290
                    TabIndex =14
                    Name ="AllowViewJournal"
                    ControlSource ="AllowViewJournal"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =12660
                    Top =60
                    Width =290
                    TabIndex =15
                    Name ="AllowXReports"
                    ControlSource ="AllowXReports"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =13440
                    Top =60
                    Width =290
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
CodeBehindForm
' See "CashierRegisterFlags.cls"
