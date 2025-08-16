Version =20
VersionRequired =20
Begin Form
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10815
    RowHeight =264
    DatasheetFontHeight =10
    ItemSuffix =30
    Top =4464
    Right =8964
    Bottom =7668
    OrderBy ="[InventoryReportParameters].[ReportName], [InventoryReportParameters].[QueryName"
        "] DESC"
    RecSrcDt = Begin
        0x7d4bb059e8a8e240
    End
    RecordSource ="InventoryReportParameters"
    Caption ="InventoryReportParameters"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
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
        End
        Begin CheckBox
            SpecialEffect =2
            LabelX =230
            LabelY =-30
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
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
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
        End
        Begin ComboBox
            SpecialEffect =2
            FontName ="Tahoma"
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
        End
        Begin Subform
            SpecialEffect =2
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
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            CanGrow = NotDefault
            Height =4560
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =120
                    Width =2475
                    ColumnWidth =3495
                    ColumnOrder =0
                    Name ="ReportName"
                    ControlSource ="ReportName"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT MSysObjects.Name, MSysObjects.Type FROM MSysObjects WHERE (((MSysObjects."
                        "Type)=-32764)) ORDER BY MSysObjects.Name; "

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =120
                            Width =975
                            Height =240
                            Name ="Label1"
                            Caption ="ReportName"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1920
                    Top =480
                    Width =2475
                    ColumnWidth =3810
                    ColumnOrder =2
                    TabIndex =1
                    Name ="QueryName"
                    ControlSource ="QueryName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =120
                            Top =480
                            Width =930
                            Height =240
                            Name ="Label3"
                            Caption ="QueryName"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =1080
                    Width =176
                    Height =180
                    ColumnWidth =1410
                    ColumnOrder =3
                    TabIndex =3
                    Name ="RangeSalesDepartments"
                    ControlSource ="RangeSalesDepartments"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =1080
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =1260
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1290
                            Top =1080
                            Width =555
                            Height =240
                            Name ="Label5"
                            Caption ="SalesD"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =1380
                    Width =176
                    Height =180
                    ColumnWidth =795
                    ColumnOrder =5
                    TabIndex =4
                    Name ="RangeVendor"
                    ControlSource ="RangeVendor"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =1380
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =1560
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1260
                            Top =1380
                            Width =585
                            Height =240
                            Name ="Label7"
                            Caption ="Vendor"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =1620
                    Width =176
                    Height =180
                    ColumnWidth =990
                    ColumnOrder =6
                    TabIndex =5
                    Name ="RangeLocations"
                    ControlSource ="RangeLocations"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =1620
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =1800
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =1530
                            Top =1620
                            Width =315
                            Height =240
                            Name ="Label9"
                            Caption ="Loc"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =2940
                    Width =176
                    Height =180
                    ColumnWidth =705
                    ColumnOrder =10
                    TabIndex =10
                    Name ="PrintOnhands"
                    ControlSource ="PrintOnhands"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =2940
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =3120
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1545
                            Top =2940
                            Width =300
                            Height =240
                            Name ="Label11"
                            Caption ="OH"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =3240
                    Width =176
                    Height =180
                    ColumnWidth =1650
                    ColumnOrder =11
                    TabIndex =11
                    Name ="PrintCost"
                    ControlSource ="PrintCost"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =3240
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =3420
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1440
                            Top =3240
                            Width =405
                            Height =240
                            Name ="Label13"
                            Caption ="Cost"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =3510
                    Width =176
                    Height =180
                    ColumnWidth =1080
                    ColumnOrder =12
                    TabIndex =13
                    Name ="PrintZeroOnhands"
                    ControlSource ="PrintZeroOnhands"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =3510
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =3690
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =1215
                            Top =3480
                            Width =630
                            Height =240
                            Name ="Label18"
                            Caption ="ZeroOH"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =2160
                    Width =176
                    Height =180
                    ColumnWidth =735
                    ColumnOrder =8
                    TabIndex =7
                    Name ="RangeItemID"
                    ControlSource ="RangeItemID"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =2160
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =2340
                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1275
                            Top =2160
                            Width =570
                            Height =240
                            Name ="Label19"
                            Caption ="ItemID"
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =87
                    Left =6120
                    Top =3360
                    Width =4440
                    Height =720
                    TabIndex =12
                    Name ="InventoryReportOrderBys"
                    SourceObject ="Form.InventoryReportOrderBys"
                    LinkChildFields ="ReportName"
                    LinkMasterFields ="ReportName"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =6120
                            Top =3120
                            Width =1965
                            Height =240
                            Name ="Label20"
                            Caption ="InventoryReportOrderBys"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =2430
                    Width =176
                    Height =180
                    ColumnWidth =1485
                    ColumnOrder =9
                    TabIndex =8
                    Name ="RangeDates"
                    ControlSource ="RangeDates"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =2430
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =2610
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =3
                            Left =1350
                            Top =2400
                            Width =495
                            Height =240
                            Name ="Label21"
                            Caption ="Dates"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1860
                    Top =780
                    ColumnWidth =1920
                    ColumnOrder =1
                    TabIndex =2
                    Name ="ChartQueryName"
                    ControlSource ="ChartQueryName"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =780
                            Width =1395
                            Height =240
                            Name ="Label22"
                            Caption ="ChartQueryName:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =1860
                    Width =176
                    Height =180
                    ColumnWidth =780
                    ColumnOrder =7
                    TabIndex =6
                    Name ="RangePriceGroupID"
                    ControlSource ="RangePriceGroupID"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =1860
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =2040
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =1320
                            Top =1920
                            Width =525
                            Height =240
                            Name ="Label23"
                            Caption ="PriceG"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =3780
                    Width =176
                    Height =180
                    ColumnWidth =3045
                    ColumnOrder =14
                    TabIndex =14
                    Name ="NumericRange"
                    ControlSource ="NumericRange"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =3780
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =3960
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =720
                            Top =3780
                            Width =1110
                            Height =240
                            Name ="Label24"
                            Caption ="NumericRange"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    Left =2280
                    Top =3780
                    ColumnWidth =3690
                    ColumnOrder =15
                    TabIndex =15
                    Name ="NumericRangeLeadThru"
                    ControlSource ="NumericRangeLeadThru"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2280
                            Top =3540
                            Width =1845
                            Height =240
                            Name ="Label25"
                            Caption ="NumericRangeLeadThru:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =1980
                    Top =2640
                    Width =176
                    Height =180
                    ColumnWidth =2130
                    ColumnOrder =13
                    TabIndex =9
                    Name ="RangeClubs"
                    ControlSource ="RangeClubs"

                    LayoutCachedLeft =1980
                    LayoutCachedTop =2640
                    LayoutCachedWidth =2156
                    LayoutCachedHeight =2820
                    Begin
                        Begin Label
                            OverlapFlags =87
                            TextAlign =3
                            Left =900
                            Top =2640
                            Width =930
                            Height =240
                            Name ="Label26"
                            Caption ="Clubs"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3270
                    Top =1080
                    Width =176
                    Height =180
                    ColumnWidth =960
                    ColumnOrder =4
                    TabIndex =16
                    Name ="RangeSubCategories"
                    ControlSource ="RangeSubCategories"

                    LayoutCachedLeft =3270
                    LayoutCachedTop =1080
                    LayoutCachedWidth =3446
                    LayoutCachedHeight =1260
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2580
                            Top =1080
                            Width =555
                            Height =240
                            Name ="Label28"
                            Caption ="Category"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =6780
                    Top =1200
                    Height =768
                    ColumnWidth =4710
                    TabIndex =17
                    Name ="Description"
                    ControlSource ="Description"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =5340
                            Top =1200
                            Width =930
                            Height =240
                            Name ="Label29"
                            Caption ="Description:"
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =360
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =4380
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3360
                            Width =975
                            Height =240
                            Name ="Label17"
                            Caption ="ChangeDate"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1080
                    TabIndex =1
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =180
                            Width =840
                            Height =240
                            Name ="Label15"
                            Caption ="SetupDate"
                        End
                    End
                End
            End
        End
    End
End
