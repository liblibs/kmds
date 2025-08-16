Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    ViewsAllowed =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6630
    DatasheetFontHeight =11
    ItemSuffix =26
    Left =1635
    Top =2610
    Right =14925
    Bottom =9060
    RecSrcDt = Begin
        0xf22d0f534a98e440
    End
    RecordSource ="qryMetrologicScanBatchVariance"
    Caption ="MetrologicScanBatchSubVariance"
    DatasheetFontName ="Calibri"
    AllowFormView =0
    FilterOnLoad =0
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
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AsianLineBreak =1
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
            Height =0
            Name ="FormHeader"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
        End
        Begin Section
            Height =3900
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =900
                    Width =1530
                    Height =330
                    ColumnWidth =615
                    ColumnOrder =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =900
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1230
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =900
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ID_Label"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =900
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1230
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1320
                    Width =3660
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =3
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Description"
                    ControlSource ="Description"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1320
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =1650
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1320
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Description_Label"
                            Caption ="Description"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1320
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1650
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1740
                    Width =3060
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =4
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PkgSize"
                    ControlSource ="PkgSize"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1740
                    LayoutCachedWidth =5970
                    LayoutCachedHeight =2070
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1740
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="PkgSize_Label"
                            Caption ="PkgSize"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1740
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2070
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =2160
                    Width =3660
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =5
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Count"
                    ControlSource ="Count"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =2160
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =2490
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2160
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Count_Label"
                            Caption ="Count"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2160
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2490
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =2580
                    Width =1530
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =6
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="OnShelf"
                    ControlSource ="OnShelf"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =2580
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =2910
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2580
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="OnShelf_Label"
                            Caption ="OnShelf"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2580
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2910
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =3000
                    Width =3660
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =7
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtVariance"
                    ControlSource ="Variance"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ffcccc00000000000400000003000000050000000100000000000000 ,
                        0xccffcc0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =3000
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =3330
                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ffcccc00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000ccffcc000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3000
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Variance_Label"
                            Caption ="Variance"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3000
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =3330
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =3420
                    Width =3660
                    Height =330
                    ColumnWidth =1065
                    ColumnOrder =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtVarianceDollars"
                    ControlSource ="Variance Dollars"
                    Format ="$#,##0.00;($#,##0.00)"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ffcccc00000000000400000003000000050000000100000000000000 ,
                        0xccffcc0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =3420
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =3750
                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ffcccc00010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000ccffcc000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    CurrencySymbol ="$"
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3420
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Variance Dollars_Label"
                            Caption ="Dollars"
                            EventProcPrefix ="Variance_Dollars_Label"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3420
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =3750
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =60
                    Width =1530
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =0
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SalesDptDescr"
                    ControlSource ="SalesDptDescr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =390
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =60
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label23"
                            Caption ="Department"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =60
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =390
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =510
                    Width =1530
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =1
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="CategoryDescr"
                    ControlSource ="CategoryDescr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =510
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =840
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =510
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label25"
                            Caption ="Category"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =510
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =840
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =990
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =240
                    Height =315
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSumVariance"
                    ControlSource ="=Sum([Variance])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3900
                    LayoutCachedTop =240
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =555
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =240
                            Width =690
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label15"
                            Caption ="Text14"
                            GridlineColor =10921638
                            LayoutCachedLeft =2100
                            LayoutCachedTop =240
                            LayoutCachedWidth =2790
                            LayoutCachedHeight =555
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3900
                    Top =675
                    Height =315
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtSumVarianceDollars"
                    ControlSource ="=Sum([Variance Dollars])"
                    GridlineColor =10921638

                    LayoutCachedLeft =3900
                    LayoutCachedTop =675
                    LayoutCachedWidth =5340
                    LayoutCachedHeight =990
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =2100
                            Top =675
                            Width =690
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label21"
                            Caption ="Text14"
                            GridlineColor =10921638
                            LayoutCachedLeft =2100
                            LayoutCachedTop =675
                            LayoutCachedWidth =2790
                            LayoutCachedHeight =990
                        End
                    End
                End
            End
        End
    End
End
