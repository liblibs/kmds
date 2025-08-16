Version =20
VersionRequired =20
Begin Form
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowEdits = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8884
    RowHeight =525
    DatasheetFontHeight =11
    ItemSuffix =7
    Left =765
    Top =2685
    Right =15690
    Bottom =10215
    RecSrcDt = Begin
        0xfe32b61b79e3e440
    End
    RecordSource ="SELECT [vw_WebLogErrorCode-1].SetupDate, Trim([WebOrderID]) AS Invoice, Trim([We"
        "bOrderNotes]) AS OrderNotes, Trim([LicenseeName]) AS Customer, Trim([ErrorCode])"
        " AS Error, Trim([Source]) AS Qty, Trim([Subroutine]) AS Sub, Trim([Subroutine]) "
        "AS eMail FROM [vw_WebLogErrorCode-1] ORDER BY [vw_WebLogErrorCode-1].SetupDate D"
        "ESC;"
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
        Begin Section
            Height =7560
            Name ="Detail"
            AutoHeight =1
            AlternateBackThemeColorIndex =1
            AlternateBackShade =95.0
            BackThemeColorIndex =1
            Begin
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2100
                    Top =180
                    Height =315
                    ColumnWidth =1455
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="mm/dd hh:nn am/pm"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =180
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =495
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =180
                            Width =1050
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label0"
                            Caption ="Time"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =180
                            LayoutCachedWidth =1350
                            LayoutCachedHeight =495
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =600
                    Height =315
                    ColumnWidth =2595
                    TabIndex =1
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Customer"
                    ControlSource ="Customer"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =600
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =915
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =600
                            Width =1440
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label1"
                            Caption ="Customer"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =600
                            LayoutCachedWidth =1740
                            LayoutCachedHeight =915
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =1440
                    Height =315
                    ColumnWidth =-2
                    TabIndex =3
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="OrderNotes"
                    ControlSource ="OrderNotes"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1440
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =1755
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1440
                            Width =1605
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label2"
                            Caption ="Message"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =1440
                            LayoutCachedWidth =1905
                            LayoutCachedHeight =1755
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2100
                    Top =1860
                    Height =315
                    ColumnWidth =1140
                    TabIndex =4
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Invoice"
                    ControlSource ="Invoice"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1860
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =2175
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =300
                            Top =1860
                            Width =1245
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label3"
                            Caption ="Invoice "
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =1860
                            LayoutCachedWidth =1545
                            LayoutCachedHeight =2175
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2100
                    Top =1020
                    Width =840
                    Height =315
                    ColumnWidth =900
                    TabIndex =2
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="Qty"
                    ControlSource ="Qty"
                    GridlineColor =10921638

                    LayoutCachedLeft =2100
                    LayoutCachedTop =1020
                    LayoutCachedWidth =2940
                    LayoutCachedHeight =1335
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =300
                            Top =1020
                            Width =1245
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label5"
                            Caption ="Qty"
                            GridlineColor =10921638
                            LayoutCachedLeft =300
                            LayoutCachedTop =1020
                            LayoutCachedWidth =1545
                            LayoutCachedHeight =1335
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2160
                    Top =2580
                    Height =315
                    ColumnWidth =2370
                    TabIndex =5
                    BorderColor =10921638
                    ForeColor =4210752
                    Name ="eMail"
                    ControlSource ="eMail"
                    GridlineColor =10921638

                    LayoutCachedLeft =2160
                    LayoutCachedTop =2580
                    LayoutCachedWidth =3600
                    LayoutCachedHeight =2895
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =360
                            Top =2580
                            Width =600
                            Height =315
                            BorderColor =8355711
                            ForeColor =8355711
                            Name ="Label6"
                            Caption ="eMail"
                            GridlineColor =10921638
                            LayoutCachedLeft =360
                            LayoutCachedTop =2580
                            LayoutCachedWidth =960
                            LayoutCachedHeight =2895
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubOrderWebImportSib.cls"
