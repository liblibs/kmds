Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    OrderByOn = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13320
    DatasheetFontHeight =11
    ItemSuffix =38
    Left =2640
    Top =2490
    Right =16470
    Bottom =9810
    Filter ="(((CompareScanItemToPurchaseOrder.FullCase)=True))"
    OrderBy ="[qryCompareScanItemToPurchaseOrderSub].[LineID]"
    RecSrcDt = Begin
        0x7a5472320e1ae440
    End
    RecordSource ="qryCompareScanItemToPurchaseOrderSub"
    Caption ="qryCompareScanItemToPurchaseOrder_05 subform"
    DatasheetFontName ="Calibri"
    OnLoad ="[Event Procedure]"
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
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
            BorderThemeColorIndex =1
            BorderShade =65.0
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
            Height =5490
            Name ="Detail"
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    ColumnHidden = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =360
                    Width =1530
                    Height =330
                    ColumnWidth =0
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="PurchaseOrderID"
                    ControlSource ="PurchaseOrderID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =360
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =690
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="PurchaseOrderID_Label"
                            Caption ="PurchaseOrderID"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =690
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =780
                    Width =1530
                    Height =330
                    ColumnWidth =650
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="LineID"
                    ControlSource ="LineID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =780
                    LayoutCachedWidth =4440
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
                            Name ="LineID_Label"
                            Caption ="1st Line"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =780
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1110
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1200
                    Width =2520
                    Height =330
                    ColumnOrder =2
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1200
                    LayoutCachedWidth =5430
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
                            Name ="ItemID_Label"
                            Caption ="ItemID"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1200
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1530
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =1620
                    Width =3660
                    Height =330
                    ColumnWidth =2760
                    ColumnOrder =3
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Description"
                    ControlSource ="Description"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1620
                    LayoutCachedWidth =6570
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
                            Name ="Description_Label"
                            Caption ="Description"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =1620
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =1950
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =2520
                    Width =1530
                    Height =330
                    ColumnWidth =650
                    ColumnOrder =5
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="QuantityShipped"
                    ControlSource ="QuantityShipped"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =2520
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =2850
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2520
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="QuantityShipped_Label"
                            Caption ="1st "
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2520
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2850
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =2940
                    Width =1530
                    Height =330
                    ColumnWidth =650
                    ColumnOrder =6
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="QuantityScanned"
                    ControlSource ="QuantityScanned"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =2940
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =3270
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2940
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="QuantityScanned_Label"
                            Caption ="2nd "
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2940
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =3270
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2910
                    Top =3360
                    Width =1530
                    Height =330
                    ColumnWidth =650
                    ColumnOrder =7
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Diff"
                    ControlSource ="Diff"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ed1c2400000000000400000003000000050000000100000000000000 ,
                        0xa7da4e0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =3360
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =3690
                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ed1c2400010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000a7da4e000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3360
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Diff_Label"
                            Caption ="Diff"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3360
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =3690
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2940
                    Top =2040
                    Height =315
                    ColumnWidth =885
                    ColumnOrder =4
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Price"
                    ControlSource ="Price"
                    GridlineColor =10921638

                    LayoutCachedLeft =2940
                    LayoutCachedTop =2040
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =2355
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1140
                            Top =2040
                            Width =540
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label14"
                            Caption ="Price"
                            GridlineColor =10921638
                            LayoutCachedLeft =1140
                            LayoutCachedTop =2040
                            LayoutCachedWidth =1680
                            LayoutCachedHeight =2355
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3000
                    Top =3900
                    ColumnWidth =345
                    TabIndex =8
                    BorderColor =10921638
                    Name ="chkFullCase"
                    ControlSource ="FullCase"
                    GridlineColor =10921638

                    LayoutCachedLeft =3000
                    LayoutCachedTop =3900
                    LayoutCachedWidth =3260
                    LayoutCachedHeight =4140
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =3225
                            Top =3870
                            Width =900
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label16"
                            Caption ="Full Case"
                            GridlineColor =10921638
                            LayoutCachedLeft =3225
                            LayoutCachedTop =3870
                            LayoutCachedWidth =4125
                            LayoutCachedHeight =4185
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =4320
                    Width =1530
                    Height =330
                    TabIndex =9
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ExtShipped"
                    ControlSource ="ExtShipped"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =4320
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =4650
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =4320
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label19"
                            Caption ="1st Order $"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =4320
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =4650
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =4740
                    Width =1530
                    Height =330
                    TabIndex =10
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ExtScanned"
                    ControlSource ="ExtScanned"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =4740
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =5070
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =4740
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label21"
                            Caption ="2nd Order $"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =4740
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =5070
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2910
                    Top =5160
                    Width =1530
                    Height =330
                    TabIndex =11
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ExtDiff"
                    ControlSource ="ExtDiff"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ed1c2400000000000400000003000000050000000100000000000000 ,
                        0xa7da4e0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =5160
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =5490
                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ed1c2400010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000a7da4e000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =5160
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label23"
                            Caption ="Diff"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =5160
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =5490
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =3030
            Name ="FormFooter"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2550
                    Width =1530
                    Height =330
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SumQuantityShipped"
                    ControlSource ="=Sum([QuantityShipped])"
                    GridlineColor =10921638

                    LayoutCachedLeft =2550
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =330
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label25"
                            Caption ="Shipped"
                            GridlineColor =10921638
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =330
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2550
                    Top =480
                    Width =1530
                    Height =330
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SumQuantityScanned"
                    ControlSource ="=Sum([QuantityScanned])"
                    GridlineColor =10921638

                    LayoutCachedLeft =2550
                    LayoutCachedTop =480
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =810
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =480
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label29"
                            Caption ="Scanned"
                            GridlineColor =10921638
                            LayoutCachedTop =480
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =810
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2550
                    Top =900
                    Width =1530
                    Height =330
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SumDiff"
                    ControlSource ="=Sum([Diff])"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ed1c2400000000000400000003000000050000000100000000000000 ,
                        0xa7da4e0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2550
                    LayoutCachedTop =900
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =1230
                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ed1c2400010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000a7da4e000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =900
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label31"
                            Caption ="Diff"
                            GridlineColor =10921638
                            LayoutCachedTop =900
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =1230
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2550
                    Top =1860
                    Width =1530
                    Height =330
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SumExtShipped"
                    ControlSource ="=Sum([ExtShipped])"
                    GridlineColor =10921638

                    LayoutCachedLeft =2550
                    LayoutCachedTop =1860
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =2190
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =1860
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label33"
                            Caption ="ExtShipped"
                            GridlineColor =10921638
                            LayoutCachedTop =1860
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =2190
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2550
                    Top =2280
                    Width =1530
                    Height =330
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SumExtScanned"
                    ControlSource ="=Sum([ExtScanned])"
                    GridlineColor =10921638

                    LayoutCachedLeft =2550
                    LayoutCachedTop =2280
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =2610
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =2280
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label35"
                            Caption ="ExtScanned"
                            GridlineColor =10921638
                            LayoutCachedTop =2280
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =2610
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2550
                    Top =2700
                    Width =1530
                    Height =330
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SumExtDiff"
                    ControlSource ="=Sum([ExtDiff])"
                    ConditionalFormat = Begin
                        0x010000006c000000020000000000000005000000000000000200000001000000 ,
                        0x00000000ed1c2400000000000400000003000000050000000100000000000000 ,
                        0xa7da4e0000000000000000000000000000000000000000000000000000000000 ,
                        0x300000000000300000000000
                    End
                    GridlineColor =10921638

                    LayoutCachedLeft =2550
                    LayoutCachedTop =2700
                    LayoutCachedWidth =4080
                    LayoutCachedHeight =3030
                    ConditionalFormat14 = Begin
                        0x01000200000000000000050000000100000000000000ed1c2400010000003000 ,
                        0x0000000000000000000000000000000000000000000000000004000000010000 ,
                        0x0000000000a7da4e000100000030000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Top =2700
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label37"
                            Caption ="Diff"
                            GridlineColor =10921638
                            LayoutCachedTop =2700
                            LayoutCachedWidth =2460
                            LayoutCachedHeight =3030
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "CompareScanItemToPurchaseOrderSub.cls"
