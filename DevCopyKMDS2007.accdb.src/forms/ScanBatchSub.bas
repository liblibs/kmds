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
    Width =6630
    DatasheetFontHeight =11
    ItemSuffix =30
    Left =3150
    Top =5085
    Right =16515
    Bottom =8025
    RecSrcDt = Begin
        0x81d80ee58f09e440
    End
    RecordSource ="SELECT ScanBatchLineCounts.ItemID, Item.Description, Item.Price, ScanBatchLineCo"
        "unts.CountQuantity, ScanBatchLineCounts.OnShelf, ScanBatchLineCounts.Location, S"
        "canBatchLineCounts.ID, ScanBatchLineCounts.ScanCode, ScanBatchLineCounts.ChangeD"
        "ate, ScanBatchLineCounts.SetupDate, ScanBatchLineCounts.OnShelfDate, ScanBatchLi"
        "neCounts.InputMethod FROM ScanBatchLineCounts INNER JOIN Item ON ScanBatchLineCo"
        "unts.ItemID = Item.ID WHERE (((ScanBatchLineCounts.InputMethod)<>\"Removed\")) O"
        "RDER BY ScanBatchLineCounts.ID DESC; "
    Caption ="ScanBatchLineCountsSub2"
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
        Begin ComboBox
            AddColon = NotDefault
            BorderLineStyle =0
            LabelX =-1800
            FontSize =11
            FontName ="Calibri"
            AllowValueListEdits =1
            InheritValueList =1
            ThemeFontIndex =1
            BackThemeColorIndex =1
            BorderThemeColorIndex =1
            BorderShade =65.0
            ForeThemeColorIndex =2
            ForeShade =50.0
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
            Height =5850
            Name ="Detail"
            AutoHeight =1
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
                    Height =315
                    ColumnWidth =1440
                    ColumnOrder =0
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ID"
                    ControlSource ="ID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =360
                    LayoutCachedWidth =4350
                    LayoutCachedHeight =675
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =360
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ID_Label"
                            Caption ="ID"
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
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2910
                    Top =1200
                    Width =1530
                    Height =330
                    ColumnWidth =855
                    ColumnOrder =1
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =1200
                    LayoutCachedWidth =4440
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
                    Top =2460
                    Width =2460
                    Height =330
                    ColumnWidth =2460
                    ColumnOrder =8
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Location"
                    ControlSource ="Location"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =2460
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =2790
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2460
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Location_Label"
                            Caption ="Location"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2460
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =2790
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =3
                    IMESentenceMode =3
                    Left =2910
                    Top =2880
                    Width =1530
                    Height =330
                    ColumnWidth =2000
                    ColumnOrder =3
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="ScanCode"
                    ControlSource ="ScanCode"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =2880
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =3210
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2880
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="ScanCode_Label"
                            Caption ="ScanCode"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2880
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =3210
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =3300
                    Width =1530
                    Height =330
                    ColumnWidth =1000
                    ColumnOrder =4
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="CountQuantity"
                    ControlSource ="CountQuantity"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =3300
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =3630
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3300
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="CountQuantity_Label"
                            Caption ="Count"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3300
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =3630
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2910
                    Top =3720
                    Width =1530
                    Height =330
                    ColumnWidth =1000
                    ColumnOrder =5
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="OnShelf"
                    ControlSource ="OnShelf"
                    GridlineColor =10921638

                    LayoutCachedLeft =2910
                    LayoutCachedTop =3720
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =4050
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =3720
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="OnShelf_Label"
                            Caption ="Shelf"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =3720
                            LayoutCachedWidth =2820
                            LayoutCachedHeight =4050
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2970
                    Top =4200
                    Width =1620
                    Height =330
                    ColumnWidth =1000
                    ColumnOrder =6
                    TabIndex =6
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="OnShelfDate"
                    ControlSource ="OnShelfDate"
                    Format ="Short Time"
                    GridlineColor =10921638

                    LayoutCachedLeft =2970
                    LayoutCachedTop =4200
                    LayoutCachedWidth =4590
                    LayoutCachedHeight =4530
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =420
                            Top =4200
                            Width =2460
                            Height =330
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="OnShelfDate_Label"
                            Caption ="Time"
                            GridlineColor =10921638
                            LayoutCachedLeft =420
                            LayoutCachedTop =4200
                            LayoutCachedWidth =2880
                            LayoutCachedHeight =4530
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2880
                    Top =1560
                    Height =315
                    ColumnWidth =2500
                    ColumnOrder =2
                    TabIndex =7
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Description"
                    ControlSource ="Description"
                    GridlineColor =10921638

                    LayoutCachedLeft =2880
                    LayoutCachedTop =1560
                    LayoutCachedWidth =4320
                    LayoutCachedHeight =1875
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1080
                            Top =1560
                            Width =1140
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label26"
                            Caption ="Description"
                            GridlineColor =10921638
                            LayoutCachedLeft =1080
                            LayoutCachedTop =1560
                            LayoutCachedWidth =2220
                            LayoutCachedHeight =1875
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3060
                    Top =4620
                    Height =315
                    ColumnWidth =2000
                    ColumnOrder =7
                    TabIndex =8
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="InputMethod"
                    ControlSource ="InputMethod"
                    GridlineColor =10921638

                    LayoutCachedLeft =3060
                    LayoutCachedTop =4620
                    LayoutCachedWidth =4500
                    LayoutCachedHeight =4935
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =4620
                            Width =1305
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label27"
                            Caption ="InputMethod"
                            GridlineColor =10921638
                            LayoutCachedLeft =1260
                            LayoutCachedTop =4620
                            LayoutCachedWidth =2565
                            LayoutCachedHeight =4935
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
' See "ScanBatchSub.cls"
