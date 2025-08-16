Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =1800
    DatasheetFontHeight =10
    ItemSuffix =3
    Left =315
    Top =1200
    RecSrcDt = Begin
        0x9bbb21fecbcae240
    End
    RecordSource ="qryInvoLine5SalesByCashierSummary"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetBackColor12 =16777215
    DisplayOnSharePointSite =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            TextFontFamily =2
            FontName ="Arial"
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
            OldBorderStyle =0
            TextFontFamily =2
            FontName ="Arial"
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
            ShowDatePicker =0
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =0
            Name ="ReportHeader"
        End
        Begin PageHeader
            Visible = NotDefault
            Height =0
            Name ="PageHeaderSection"
        End
        Begin Section
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =360
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Width =480
                    ColumnWidth =1755
                    Name ="TotalTransactions"
                    ControlSource ="TotalTransactions"

                    Begin
                        Begin Label
                            Width =1395
                            Height =225
                            Name ="Label0"
                            Caption ="TotalTransactions:"
                        End
                    End
                End
            End
        End
        Begin PageFooter
            Visible = NotDefault
            Height =0
            Name ="PageFooterSection"
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =720
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    BorderWidth =2
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =0
                    IMESentenceMode =3
                    Top =420
                    Width =1800
                    Height =300
                    FontSize =10
                    Name ="TotalTransactions Grand Total Sum"
                    ControlSource ="=Sum([TotalTransactions])"
                    Format ="Standard"
                    FontName ="Times New Roman"
                    EventProcPrefix ="TotalTransactions_Grand_Total_Sum"

                    Begin
                        Begin Label
                            FontItalic = NotDefault
                            OldBorderStyle =1
                            TextAlign =2
                            TextFontFamily =18
                            Top =60
                            Width =1800
                            Height =660
                            FontSize =11
                            FontWeight =700
                            ForeColor =8388608
                            Name ="Label51"
                            Caption ="Customer Count"
                            FontName ="Times New Roman"
                            Tag ="DetachedLabel"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOS040CustomerCount.cls"
