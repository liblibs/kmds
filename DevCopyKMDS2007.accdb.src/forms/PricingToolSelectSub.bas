Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10230
    DatasheetFontHeight =11
    ItemSuffix =19
    Left =-5550
    Top =1755
    Right =5400
    Bottom =8265
    RecSrcDt = Begin
        0x5cf135abb2f3e540
    End
    RecordSource ="qryItemsByVendor"
    Caption ="PricingToolSub"
    DatasheetFontName ="Calibri"
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
            Height =1275
            Name ="FormHeader"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =2
            BackTint =20.0
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6960
                    Top =180
                    Height =315
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="txtRecordCount"
                    ControlSource ="=Count(*)"
                    GridlineColor =10921638

                    LayoutCachedLeft =6960
                    LayoutCachedTop =180
                    LayoutCachedWidth =8400
                    LayoutCachedHeight =495
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =5760
                    Top =540
                    Width =1530
                    Height =330
                    ColumnWidth =1530
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    GridlineColor =10921638

                    LayoutCachedLeft =5760
                    LayoutCachedTop =540
                    LayoutCachedWidth =7290
                    LayoutCachedHeight =870
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1140
                            Top =540
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="VendorID_Label"
                            Caption ="VendorID"
                            GridlineColor =10921638
                            LayoutCachedLeft =1140
                            LayoutCachedTop =540
                            LayoutCachedWidth =3600
                            LayoutCachedHeight =870
                        End
                    End
                End
            End
        End
        Begin Section
            Height =4050
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =780
                    Width =3660
                    Height =330
                    ColumnWidth =3000
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1200
                    Width =3660
                    Height =330
                    ColumnWidth =1125
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2940
                    Top =2460
                    Width =3630
                    Height =300
                    ColumnWidth =1815
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Category"
                    ControlSource ="Category"
                    GridlineColor =10921638

                    LayoutCachedLeft =2940
                    LayoutCachedTop =2460
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =2760
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =3180
                    Width =3660
                    Height =330
                    ColumnWidth =1290
                    TabIndex =3
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
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =3600
                    Width =3660
                    Height =330
                    ColumnWidth =1050
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Cost"
                    ControlSource ="Cost"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =3600
                    LayoutCachedWidth =6570
                    LayoutCachedHeight =3930
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3600
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Cost_Label"
                            Caption ="Cost"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3600
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =3930
                        End
                    End
                End
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1620
                    Width =1530
                    Height =330
                    ColumnWidth =0
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SalesDeptID"
                    ControlSource ="SalesDeptID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1620
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =1950
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1620
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="SalesDeptID_Label"
                            Caption ="SalesDeptID"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1620
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1950
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =2040
                    Width =3660
                    Height =330
                    ColumnWidth =1950
                    TabIndex =6
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
            End
        End
        Begin FormFooter
            Height =0
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
        End
    End
End
