Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12420
    DatasheetFontHeight =10
    ItemSuffix =93
    Left =1140
    Top =1980
    Right =15345
    Bottom =9900
    HelpContextId =7300
    OnUnload ="[Event Procedure]"
    AfterInsert ="[Event Procedure]"
    RecSrcDt = Begin
        0x7dbf2123fc92e240
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="Cashier"
    Caption ="Cashier KMDS Software ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =7740
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =1980
                    Top =180
                    TabIndex =1
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="MS Sans Serif"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1020
                            Top =180
                            Width =900
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label1"
                            Caption ="Cashier #"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =3600
                    Top =180
                    Width =2490
                    TabIndex =2
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="MS Sans Serif"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =1980
                    Top =480
                    TabIndex =5
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Password"
                    ControlSource ="Password"
                    FontName ="MS Sans Serif"
                    InputMask ="Password"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1155
                            Top =480
                            Width =765
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label5"
                            Caption ="Password"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =1980
                    Top =780
                    TabIndex =7
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="Drawer"
                    ControlSource ="Drawer"
                    FontName ="MS Sans Serif"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1320
                            Top =780
                            Width =600
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label9"
                            Caption ="Drawer"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =3
                    Left =4800
                    Top =780
                    Width =1260
                    TabIndex =8
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="ValidRegisters"
                    ControlSource ="ValidRegisters"
                    FontName ="MS Sans Serif"
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =3660
                            Top =780
                            Width =1080
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label11"
                            Caption ="ValidRegisters"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin Tab
                    OverlapFlags =85
                    Left =300
                    Top =1200
                    Width =11370
                    Height =6000
                    TabIndex =9
                    Name ="TabCtl50"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =435
                            Top =1605
                            Width =11100
                            Height =5460
                            Name ="Page52"
                            Caption ="Load Programs"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =495
                                    Top =2265
                                    Width =6105
                                    Height =4800
                                    Name ="CashiersAndPrograms"
                                    SourceObject ="Form.CashiersAndProgramsSub"
                                    LinkChildFields ="CashierID"
                                    LinkMasterFields ="ID"

                                End
                                Begin Subform
                                    OverlapFlags =247
                                    Left =6900
                                    Top =2820
                                    Width =4620
                                    Height =4095
                                    TabIndex =1
                                    Name ="CashiersAndProgramsDatasheet"
                                    SourceObject ="Form.CashiersAndProgramsDatasheet"

                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =2
                                            Left =6908
                                            Top =1605
                                            Width =4620
                                            Height =330
                                            FontSize =12
                                            Name ="lblCurrentProgramName"
                                            Caption ="4. PSC Functions"
                                        End
                                    End
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    AccessKey =65
                                    Left =480
                                    Top =1725
                                    Width =6120
                                    Height =480
                                    TabIndex =2
                                    Name ="cmdSetAllowed"
                                    Caption ="Set &ALL Allowed Flags for #99 VANCE"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =65

                                    Overlaps =1
                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    AccessKey =84
                                    Left =7080
                                    Top =2265
                                    Width =4140
                                    Height =480
                                    TabIndex =3
                                    Name ="cmdSetAllowedForThisProgram"
                                    Caption ="Se&t ALL Allowed Flags for PSC"
                                    OnClick ="[Event Procedure]"
                                    Tag ="PSC"
                                    UnicodeAccessKey =116

                                    Overlaps =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =215
                            Left =435
                            Top =1605
                            Width =11100
                            Height =5460
                            Name ="Page51"
                            Caption ="Cash Register Functions"
                            Begin
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =2570
                                    Top =2615
                                    Name ="AllowCancelSale"
                                    ControlSource ="AllowCancelSale"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =770
                                            Top =2615
                                            Width =1230
                                            Height =240
                                            Name ="Label15"
                                            Caption ="Cancel Sale"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =2570
                                    Top =2975
                                    TabIndex =1
                                    Name ="AllowCashDrops"
                                    ControlSource ="AllowCashDrops"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =770
                                            Top =2975
                                            Width =1230
                                            Height =240
                                            Name ="Label17"
                                            Caption ="Cash Drops"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =2570
                                    Top =3335
                                    TabIndex =2
                                    Name ="AllowCashLoans"
                                    ControlSource ="AllowCashLoans"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =770
                                            Top =3335
                                            Width =1230
                                            Height =240
                                            Name ="Label19"
                                            Caption ="Cash Loans"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =2570
                                    Top =3695
                                    TabIndex =3
                                    Name ="AllowCoupons"
                                    ControlSource ="AllowCoupons"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =770
                                            Top =3695
                                            Width =1080
                                            Height =240
                                            Name ="Label21"
                                            Caption ="Coupons"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =2570
                                    Top =4055
                                    TabIndex =4
                                    Name ="AllowDiscountItem"
                                    ControlSource ="AllowDiscountItem"
                                    ShortcutMenuBar ="Allow cashier to give a percent discount at the register."

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =770
                                            Top =4055
                                            Width =1395
                                            Height =240
                                            Name ="Label23"
                                            Caption ="Discount Item"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =2570
                                    Top =4415
                                    TabIndex =5
                                    Name ="AllowLotteryPaidOut"
                                    ControlSource ="AllowLotteryPaidOut"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =770
                                            Top =4415
                                            Width =1545
                                            Height =240
                                            Name ="Label25"
                                            Caption ="Lottery PaidOuts"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =4790
                                    Top =2255
                                    TabIndex =6
                                    Name ="AllowNegSales"
                                    ControlSource ="AllowNegSales"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =2990
                                            Top =2255
                                            Width =1140
                                            Height =240
                                            Name ="Label27"
                                            Caption ="Sales Return"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =4790
                                    Top =2615
                                    TabIndex =7
                                    Name ="AllowNewItems"
                                    ControlSource ="AllowNewItems"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =2985
                                            Top =2610
                                            Width =1305
                                            Height =240
                                            Name ="Label29"
                                            Caption ="Setup New Items"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =4790
                                    Top =2975
                                    TabIndex =8
                                    Name ="AllowNoSales"
                                    ControlSource ="AllowNoSales"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =2990
                                            Top =2975
                                            Width =1020
                                            Height =240
                                            Name ="Label31"
                                            Caption ="No Sales"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =4790
                                    Top =3335
                                    TabIndex =9
                                    Name ="AllowPaidouts"
                                    ControlSource ="AllowPaidouts"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =2990
                                            Top =3335
                                            Width =1065
                                            Height =240
                                            Name ="Label33"
                                            Caption ="Paidouts"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =4790
                                    Top =3695
                                    TabIndex =10
                                    Name ="AllowSellClub"
                                    ControlSource ="AllowSellClub"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =2990
                                            Top =3695
                                            Width =1245
                                            Height =240
                                            Name ="Label35"
                                            Caption ="Sell Club Invoice"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =4790
                                    Top =4055
                                    TabIndex =11
                                    Name ="AllowShutDown"
                                    ControlSource ="AllowShutDown"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =2985
                                            Top =4050
                                            Width =1290
                                            Height =240
                                            Name ="Label37"
                                            Caption ="F2 - Item Lookup"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =4790
                                    Top =4415
                                    TabIndex =12
                                    Name ="AllowViewJournal"
                                    ControlSource ="AllowViewJournal"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =2990
                                            Top =4415
                                            Width =1305
                                            Height =240
                                            Name ="Label39"
                                            Caption ="View Journal"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =2570
                                    Top =4775
                                    TabIndex =13
                                    Name ="AllowXReports"
                                    ControlSource ="AllowXReports"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =770
                                            Top =4775
                                            Width =1485
                                            Height =240
                                            Name ="Label41"
                                            Caption ="X Report (readings)"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =223
                                    Left =4790
                                    Top =4775
                                    TabIndex =14
                                    Name ="AllowZReports"
                                    ControlSource ="AllowZReports"

                                    Begin
                                        Begin Label
                                            OverlapFlags =223
                                            Left =2990
                                            Top =4775
                                            Width =1725
                                            Height =240
                                            Name ="Label43"
                                            Caption ="Z Reports (end of day)"
                                        End
                                    End
                                End
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =495
                                    Top =1950
                                    Width =4620
                                    Height =3180
                                    Name ="Box48"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =530
                                    Top =1955
                                    Width =1500
                                    Height =210
                                    Name ="Label49"
                                    Caption ="Cashier Rights"
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =435
                            Top =1605
                            Width =11100
                            Height =5460
                            Name ="Page69"
                            Caption ="Back Office Functions"
                            Begin
                                Begin Rectangle
                                    OverlapFlags =255
                                    Left =435
                                    Top =1950
                                    Width =4620
                                    Height =3180
                                    Name ="Box55"
                                End
                                Begin Label
                                    OverlapFlags =247
                                    Left =470
                                    Top =1955
                                    Width =1500
                                    Height =210
                                    Name ="Label56"
                                    Caption ="BackOffice  Rights"
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =735
                                    Top =2580
                                    Name ="AllowChangePrice"
                                    ControlSource ="AllowChangePrice"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =960
                                            Top =2550
                                            Width =1530
                                            Height =240
                                            Name ="Label63"
                                            Caption ="Change Price  "
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =735
                                    Top =2940
                                    TabIndex =1
                                    Name ="AllowChangeOnhand"
                                    ControlSource ="AllowChangeOnhand"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =965
                                            Top =2910
                                            Width =1575
                                            Height =240
                                            Name ="Label64"
                                            Caption ="Change Onhands"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =735
                                    Top =3300
                                    TabIndex =3
                                    Name ="AllowChangeCost"
                                    ControlSource ="AllowChangeCost"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =965
                                            Top =3270
                                            Width =1335
                                            Height =240
                                            Name ="Label65"
                                            Caption ="Change Cost"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =735
                                    Top =3660
                                    TabIndex =4
                                    Name ="AllowChangeOrderAt"
                                    ControlSource ="AllowChangeOrderAt"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =965
                                            Top =3630
                                            Width =1590
                                            Height =240
                                            Name ="Label66"
                                            Caption ="Change Order At"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =735
                                    Top =4020
                                    TabIndex =5
                                    Name ="AllowChangeOrderTo"
                                    ControlSource ="AllowChangeOrderTo"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =965
                                            Top =3990
                                            Width =1605
                                            Height =240
                                            Name ="Label67"
                                            Caption ="Change Order To"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =735
                                    Top =4380
                                    TabIndex =6
                                    Name ="AllowChangeDescription"
                                    ControlSource ="AllowChangeDescription"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =965
                                            Top =4350
                                            Width =1800
                                            Height =240
                                            Name ="Label68"
                                            Caption =" Change Description"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =720
                                    Top =4710
                                    TabIndex =9
                                    Name ="AllowDeleteDocument"
                                    ControlSource ="AllowDeleteDocument"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =950
                                            Top =4680
                                            Width =1800
                                            Height =240
                                            Name ="Label92"
                                            Caption =" Delete Document "
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =7380
                                    Top =2145
                                    TabIndex =7
                                    Name ="PayrollID"
                                    ControlSource ="PayrollID"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =6540
                                            Top =2145
                                            Width =780
                                            Height =240
                                            Name ="Label84"
                                            Caption ="PayrollID:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =7380
                                    Top =2445
                                    TabIndex =8
                                    Name ="HourlyWage"
                                    ControlSource ="HourlyWage"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =6300
                                            Top =2445
                                            Width =1020
                                            Height =240
                                            Name ="Label85"
                                            Caption ="HourlyWage:"
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =7380
                                    Top =2970
                                    TabIndex =2
                                    Name ="AllowHourlyWageView"
                                    ControlSource ="AllowHourlyWageView"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =7605
                                            Top =2940
                                            Width =3120
                                            Height =240
                                            Name ="Label88"
                                            Caption =" Allow this cashier to view all hourly wages"
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =82
                    Left =6300
                    Top =180
                    Width =1140
                    Height =960
                    TabIndex =3
                    Name ="cmdAppendCashiersAndPrograms"
                    Caption ="&Refresh Cashier and Program Tables"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =82

                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =67
                    Left =7560
                    Top =180
                    Width =2295
                    Height =405
                    TabIndex =4
                    Name ="cmdCashRegisterFlags"
                    Caption ="&Cash Register Functions - All"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =67

                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =66
                    Left =7560
                    Top =720
                    Width =2295
                    Height =405
                    TabIndex =6
                    Name ="cmdCashiersBackOfficeFlags"
                    Caption ="&Back Office Functions - All"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =66

                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =735
                    Top =7380
                    Width =1830
                    Height =255
                    TabIndex =10
                    BackColor =12632256
                    ForeColor =-2147483640
                    Name ="Text73"
                    ControlSource ="SetupDate"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =180
                            Top =7380
                            Width =495
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="SetupDate_Label"
                            Caption ="Setup"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =2
                    Left =9810
                    Top =7380
                    Width =1830
                    Height =255
                    TabIndex =11
                    BackColor =12632256
                    ForeColor =-2147483640
                    Name ="Text74"
                    ControlSource ="ChangeDate"
                    Format ="General Date"
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =9120
                            Top =7380
                            Width =630
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="ChangeDate_Label"
                            Caption ="Change"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =120
                    Width =11940
                    Height =7620
                    Name ="boxScreen"
                End
                Begin Rectangle
                    OverlapFlags =247
                    Left =120
                    Top =7260
                    Width =11940
                    Height =480
                    Name ="Box13"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    ColumnCount =2
                    ListRows =25
                    ListWidth =1785
                    Left =1980
                    Top =180
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cboID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Cashier.ID, Cashier.Name FROM Cashier ORDER BY Cashier.ID; "
                    ColumnWidths ="345;1440"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="MS Sans Serif"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =960
                            Top =180
                            Width =960
                            Height =245
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="ID1_Label"
                            Caption ="Cashier #"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =10080
                    Top =180
                    Width =1905
                    Height =405
                    TabIndex =12
                    Name ="cmdCRViewer"
                    Caption ="Monitor Cash Registers"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =960
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =4740
                    Top =60
                    Width =666
                    Height =606
                    ForeColor =0
                    Name ="cmdfrmDeleteRecord"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =101

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =4778
                            Top =660
                            Width =555
                            Height =285
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =3300
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =1
                    ForeColor =0
                    Name ="cmdfrmSaveRecord"
                    Caption ="Command34"
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
                        0xffff0000ffffff00070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8f8070700000000000000000000000000000000000000000000 ,
                        0x0000000000f8f8f8070007070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700000000000000000000000000000000 ,
                        0x000007070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070000000000000000000000000000000000000000 ,
                        0x000007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8070007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff00070707000707070000000000000000000000000000000000000000000000 ,
                        0x0000000000070707
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save Changes to Record"
                    UnicodeAccessKey =118

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =3367
                            Top =660
                            Width =420
                            Height =285
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =1860
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =2
                    ForeColor =0
                    Name ="cmdfrmUndo"
                    Caption ="Command35"
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
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070700 ,
                        0xfb01fb01000707070707070707070707070707070707070707070707070700fb ,
                        0x01fb01000107070707070707070707070707070707070707070707070700fb01 ,
                        0xfb010001fb070707070707070707070707070707070707070707070700fb01fb ,
                        0x010001fb010707070707070707070707070707070707070707070700fb01fb01 ,
                        0x0001fb01fb07070707070707070707070707070707070707070700fb01fb0100 ,
                        0x01fb01fb00070707070707070707070707070707070707070700fb01fb010001 ,
                        0xfb01fb00fb0707070707070707070707070707070707070700fb01fb010001fb ,
                        0x01fb00fb0007070707070707070707070707070707070700000000010001fb01 ,
                        0xfb00fb0007070707070707070707070707070707070700020202000001fb01fb ,
                        0x00fb00070707070707070707070707070707070707000000000202020001fb00 ,
                        0xfb0007070707070707070707070707070707070700fb01fb00000202000000fb ,
                        0x0007070707070707070707070707070707070700fb01fb01fb0100020200fb00 ,
                        0x07070707070707070707070707070707070700fb01fb01fb01fb000002000007 ,
                        0x07070707070707070707070707070707070000000001fb01fb01fb0002000707 ,
                        0x0707070707070707070707070707070700020202000001fb01fb010000070707 ,
                        0x070707070707070707070707070707000000000202020001fb01fb0007070707 ,
                        0x07070707070707070707070707070007fd0700000202000001fb000707070707 ,
                        0x070707070707070707070707070007fd07fd07fd00020200fb00070707070707 ,
                        0x070707070707070707070707000000000007fd07000002000007070707070707 ,
                        0x07070707070707070707070700fdfffdff0007fd070002000707070707070707 ,
                        0x070707070707070707070700fdfffdfffd00fd07fd0000070707070707070707 ,
                        0x070707070707070707070700fffdfffdff0000fd070007070707070707070707 ,
                        0x070707070707070707070700fdfffdfffdff0007000707070707070707070707 ,
                        0x070707070707070707070700fffdfffdfffd0000070707070707070707070707 ,
                        0x07070707070707070707070700fffdfffdff0007070707070707070707070707 ,
                        0x07070707070707070707070700fdfffdff000707070707070707070707070707 ,
                        0x0707070707070707070707070700000000070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    FontName ="Trebuchet MS"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Undo Changes Made to This Record"
                    UnicodeAccessKey =110

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =1927
                            Top =660
                            Width =450
                            Height =285
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =77
                    Left =6120
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =3
                    ForeColor =0
                    Name ="cmdfrmClose"
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
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =109

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =6120
                            Top =660
                            Width =885
                            Height =285
                            Name ="Label78"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =480
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =4
                    ForeColor =0
                    Name ="cmdfrmAddNew"
                    Caption ="Command67"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddcdadadadadadadaaccdadadada4adaddcccdadad4d4d4da ,
                        0xaccccdadad444daddcccccda4444444aaccccdadad444daddcccdadad4d4d4da ,
                        0xaccdadadada4adaddcdadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add a New Record"
                    UnicodeAccessKey =65

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =488
                            Top =660
                            Width =705
                            Height =285
                            Name ="Label79"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =215
                    Left =120
                    Width =11940
                    Height =960
                    Name ="Box81"
                End
            End
        End
    End
End
CodeBehindForm
' See "CashierSAVE.cls"
