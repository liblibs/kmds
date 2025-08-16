Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    DatasheetFontHeight =11
    ItemSuffix =18
    Left =1695
    Top =2370
    Right =14865
    Bottom =8940
    RecSrcDt = Begin
        0xd3f84a96aa98e440
    End
    RecordSource ="qryMetroScanBatchDetail"
    Caption ="MetrologicOptimiserSubAuditLog"
    DatasheetFontName ="Calibri"
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
            Height =4290
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =360
                    Width =2130
                    Height =330
                    ColumnWidth =615
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =360
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =690
                    ThemeFontIndex =-1
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ID_Label"
                            Caption ="ID"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =780
                    Width =3660
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Description"
                    ControlSource ="Description"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =780
                    LayoutCachedWidth =6330
                    LayoutCachedHeight =1110
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =780
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Description_Label"
                            Caption ="Description"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1110
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =1980
                    Width =6154
                    Height =585
                    ColumnWidth =-2
                    ColumnOrder =4
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ScanCode"
                    ControlSource ="ScanCode"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =1980
                    LayoutCachedWidth =8824
                    LayoutCachedHeight =2565
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1980
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ScanCode_Label"
                            Caption ="ScanCode"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1980
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =2310
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =2700
                    Width =1530
                    Height =330
                    ColumnWidth =975
                    ColumnOrder =5
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Quantity"
                    ControlSource ="Quantity"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =2700
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =3030
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2700
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Quantity_Label"
                            Caption ="Count"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2700
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =3030
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =3120
                    Width =3660
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =6
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="CurPrice"
                    ControlSource ="Price"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =3120
                    LayoutCachedWidth =6330
                    LayoutCachedHeight =3450
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3120
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Price_Label"
                            Caption ="Price"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3120
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =3450
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =3540
                    Width =1620
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =7
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ScanDate"
                    ControlSource ="ScanDate"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =3540
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =3870
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3540
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ScanDate_Label"
                            Caption ="ScanDate"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3540
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =3870
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =3960
                    Width =1620
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =8
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ScanTime"
                    ControlSource ="ScanTime"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =3960
                    LayoutCachedWidth =4290
                    LayoutCachedHeight =4290
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3960
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ScanTime_Label"
                            Caption ="ScanTime"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3960
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =4290
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =1140
                    Width =3660
                    Height =330
                    ColumnWidth =1545
                    ColumnOrder =2
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SalesDptDescr"
                    ControlSource ="SalesDptDescr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =1140
                    LayoutCachedWidth =6330
                    LayoutCachedHeight =1470
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1140
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label15"
                            Caption ="Department"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1140
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1470
                        End
                    End
                End
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2670
                    Top =1590
                    Width =3660
                    Height =330
                    ColumnWidth =-2
                    ColumnOrder =3
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="CategoryDescr"
                    ControlSource ="CategoryDescr"
                    GridlineColor =10921638

                    LayoutCachedLeft =2670
                    LayoutCachedTop =1590
                    LayoutCachedWidth =6330
                    LayoutCachedHeight =1920
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =1590
                            Width =2220
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label17"
                            Caption ="Category"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1590
                            LayoutCachedWidth =2580
                            LayoutCachedHeight =1920
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
CodeBehindForm
' See "MetrologicOptimiserSubAuditLog.cls"
