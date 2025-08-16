Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11100
    DatasheetFontHeight =11
    ItemSuffix =21
    Left =1350
    Top =3465
    Right =13515
    Bottom =20355
    OrderBy ="[PricingWorkTable].[NewGP], [PricingWorkTable].[NewPrice], [PricingWorkTable].[P"
        "rice]"
    RecSrcDt = Begin
        0xb5855d40f2eae540
    End
    RecordSource ="PricingWorkTable"
    Caption ="PricingWorkTable subform"
    DatasheetFontName ="Calibri"
    OnDblClick ="[Event Procedure]"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =15921906
    DatasheetGridlinesColor12 =0
    FitToScreen =1
    DatasheetBackThemeColorIndex =1
    BorderThemeColorIndex =3
    ThemeFontIndex =1
    ForeThemeColorIndex =0
    AlternateBackThemeColorIndex =1
    AlternateBackShade =95.0
    Begin
        Begin Label
            BackStyle =0
            FontSize =11
            FontName ="Calibri"
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =0
            BorderTint =50.0
            ForeThemeColorIndex =0
            ForeTint =50.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin TextBox
            AddColon = NotDefault
            FELineBreak = NotDefault
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
            LeftPadding =30
            TopPadding =30
            RightPadding =30
            BottomPadding =30
            GridlineStyleLeft =0
            GridlineStyleTop =0
            GridlineStyleRight =0
            GridlineStyleBottom =0
            GridlineWidthLeft =1
            GridlineWidthTop =1
            GridlineWidthRight =1
            GridlineWidthBottom =1
            ShowDatePicker =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ThemeFontIndex =1
            ForeThemeColorIndex =0
            ForeTint =75.0
            GridlineThemeColorIndex =1
            GridlineShade =65.0
        End
        Begin FormHeader
            Height =960
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin TextBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    Left =2850
                    Top =240
                    Width =1530
                    Height =330
                    ColumnWidth =1530
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2850
                    LayoutCachedTop =240
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =570
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =240
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="VendorID_Label"
                            Caption ="VendorID"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =240
                            LayoutCachedWidth =2760
                            LayoutCachedHeight =570
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6150
                    Top =120
                    Width =1530
                    Height =330
                    ColumnWidth =1530
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SalesDeptID"
                    ControlSource ="SalesDeptID"
                    GridlineColor =10921638

                    LayoutCachedLeft =6150
                    LayoutCachedTop =120
                    LayoutCachedWidth =7680
                    LayoutCachedHeight =450
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3600
                            Top =120
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="SalesDeptID_Label"
                            Caption ="SalesDeptID"
                            GridlineColor =10921638
                            LayoutCachedLeft =3600
                            LayoutCachedTop =120
                            LayoutCachedWidth =6060
                            LayoutCachedHeight =450
                        End
                    End
                End
            End
        End
        Begin Section
            Height =5235
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =780
                    Width =3660
                    Height =330
                    ColumnWidth =2550
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Description"
                    ControlSource ="Description"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =780
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =780
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Description_Label"
                            Caption ="Description"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1110
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1200
                    Width =3660
                    Height =330
                    ColumnWidth =1125
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1200
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =1530
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1200
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ID_Label"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1530
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =2040
                    Width =3660
                    Height =330
                    ColumnWidth =1710
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SalesDept"
                    ControlSource ="SalesDept"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =2040
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =2370
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2040
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="SalesDept_Label"
                            Caption ="SalesDept"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2040
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2370
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =2460
                    Width =3690
                    Height =360
                    ColumnWidth =1455
                    ColumnOrder =3
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Category"
                    ControlSource ="Category"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =2460
                    LayoutCachedWidth =6600
                    LayoutCachedHeight =2820
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2460
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Category_Label"
                            Caption ="Category"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2460
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2790
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =3180
                    Width =3660
                    Height =330
                    ColumnWidth =1095
                    ColumnOrder =4
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Price"
                    ControlSource ="Price"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =3180
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =3510
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3180
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Price_Label"
                            Caption ="Price"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3180
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =3510
                        End
                    End
                End
                Begin TextBox
                    DecimalPlaces =4
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =4020
                    Width =3660
                    Height =330
                    ColumnWidth =1635
                    ColumnOrder =6
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="NewPrice"
                    ControlSource ="NewPrice"
                    Format ="$#,##0.00;($#,##0.00)"
                    AfterUpdate ="[Event Procedure]"
                    OnClick ="[Event Procedure]"
                    ConditionalFormat = Begin
                        0x0100000072000000010000000000000003000000000000000800000001000000 ,
                        0x00000000fff20000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x5b00500052004900430045005d0000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =4020
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =4350
                    ConditionalFormat14 = Begin
                        0x01000100000000000000030000000100000000000000fff20000070000005b00 ,
                        0x500052004900430045005d000000000000000000000000000000000000000000 ,
                        0x00
                    End
                    CurrencySymbol ="$"
                    ColLCID =1033
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =4020
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="NewPrice_Label"
                            Caption ="NewPrice"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =4020
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =4350
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    DecimalPlaces =0
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3000
                    Top =4860
                    Height =315
                    ColumnWidth =1110
                    ColumnOrder =7
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="NewGP"
                    ControlSource ="NewGP"
                    Format ="Percent"
                    GridlineColor =10921638

                    LayoutCachedLeft =3000
                    LayoutCachedTop =4860
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =5175
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1200
                            Top =4860
                            Width =765
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label20"
                            Caption ="NewGP"
                            GridlineColor =10921638
                            LayoutCachedLeft =1200
                            LayoutCachedTop =4860
                            LayoutCachedWidth =1965
                            LayoutCachedHeight =5175
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2970
                    Top =4440
                    Width =3660
                    Height =330
                    ColumnWidth =870
                    ColumnOrder =5
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Cost"
                    ControlSource ="Cost"
                    GridlineColor =10921638

                    LayoutCachedLeft =2970
                    LayoutCachedTop =4440
                    LayoutCachedWidth =6630
                    LayoutCachedHeight =4770
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =4440
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Cost_Label"
                            Caption ="Cost"
                            GridlineColor =10921638
                            LayoutCachedLeft =420
                            LayoutCachedTop =4440
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =4770
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =750
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2970
                    Width =3660
                    Height =330
                    ColumnWidth =1260
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="NewCost"
                    ControlSource ="NewCost"
                    Format ="$#,##0.00;($#,##0.00)"
                    GridlineColor =10921638

                    LayoutCachedLeft =2970
                    LayoutCachedWidth =6630
                    LayoutCachedHeight =330
                    CurrencySymbol ="$"
                    ColLCID =1033
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="NewCost_Label"
                            Caption ="NewCost"
                            GridlineColor =10921638
                            LayoutCachedLeft =420
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =330
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PricingToolSub.cls"
