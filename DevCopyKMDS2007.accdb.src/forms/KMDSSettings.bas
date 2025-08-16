Version =20
VersionRequired =20
Begin Form
    DividingLines = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =2
    TabularCharSet =238
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11527
    DatasheetFontHeight =11
    ItemSuffix =12
    Left =3990
    Top =4365
    Right =15795
    Bottom =7365
    OrderBy ="[KMDSSettings].[KeyName], [KMDSSettings].[Value]"
    RecSrcDt = Begin
        0x633e8257ea74e340
    End
    RecordSource ="SELECT Settings.ID, Settings.KeyName, Settings.Value FROM Settings WHERE (((Sett"
        "ings.KeyName) Not Like \"Ad*\")) ORDER BY Settings.KeyName; "
    OnCurrent ="[Event Procedure]"
    DatasheetFontName ="Calibri"
    AllowFormView =0
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    ShowPageMargins =0
    DisplayOnSharePointSite =1
    DatasheetAlternateBackColor =16053492
    DatasheetGridlinesColor12 =15062992
    FitToScreen =1
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
        End
        Begin TextBox
            FELineBreak = NotDefault
            LabelX =-1800
            FontSize =11
            BorderColor =12632256
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
        End
        Begin Section
            Height =2994
            Name ="Detail"
            AutoHeight =1
            Begin
                Begin TextBox
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2250
                    Top =360
                    Width =3975
                    Height =359
                    ColumnWidth =4785
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="KeyName"
                    ControlSource ="KeyName"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2250
                    LayoutCachedTop =360
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =719
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =1
                            Left =375
                            Top =360
                            Width =1815
                            Height =359
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Label3"
                            Caption ="Key Name"
                            BottomPadding =38
                            LayoutCachedLeft =375
                            LayoutCachedTop =360
                            LayoutCachedWidth =2190
                            LayoutCachedHeight =719
                            LayoutGroup =1
                            GroupTable =7
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2250
                    Top =795
                    Width =3975
                    Height =330
                    ColumnWidth =5775
                    TabIndex =1
                    LeftMargin =22
                    TopMargin =22
                    RightMargin =22
                    BottomMargin =22
                    Name ="Value"
                    ControlSource ="Value"
                    RightPadding =38
                    BottomPadding =38

                    LayoutCachedLeft =2250
                    LayoutCachedTop =795
                    LayoutCachedWidth =6225
                    LayoutCachedHeight =1125
                    RowStart =1
                    RowEnd =1
                    ColumnStart =1
                    ColumnEnd =1
                    LayoutGroup =1
                    GroupTable =7
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextFontCharSet =238
                            TextAlign =1
                            Left =375
                            Top =795
                            Width =1815
                            Height =330
                            LeftMargin =22
                            TopMargin =22
                            RightMargin =22
                            BottomMargin =22
                            ForeColor =-2147483615
                            Name ="Label5"
                            Caption ="Value"
                            BottomPadding =38
                            LayoutCachedLeft =375
                            LayoutCachedTop =795
                            LayoutCachedWidth =2190
                            LayoutCachedHeight =1125
                            RowStart =1
                            RowEnd =1
                            LayoutGroup =1
                            GroupTable =7
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "KMDSSettings.cls"
