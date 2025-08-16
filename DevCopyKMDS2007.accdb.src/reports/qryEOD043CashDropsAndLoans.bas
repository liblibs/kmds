Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoResize = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =8
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =78
    Top =600
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0xf868844196b7e240
    End
    RecordSource ="qryEOCashDropsAndLoans"
    Caption ="qryEOCashDropsAndLoans"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="CashierID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="TenderID"
        End
        Begin BreakLevel
            ControlSource ="Time"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1095
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =540
                    Width =10620
                    Height =510
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =67
                    Width =10620
                    Height =510
                    FontSize =20
                    Name ="Label28"
                    Caption ="EOD Drops and Loans by Cashier"
                End
                Begin Line
                    BorderWidth =3
                    Width =10740
                    BorderColor =12632256
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =3
                    Top =1080
                    Width =10740
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line34"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =10800
                    Name ="Line35"
                End
                Begin Label
                    Left =540
                    Top =60
                    Width =1020
                    Height =300
                    Name ="Label36"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1740
                    Top =60
                    Width =555
                    Height =300
                    Name ="Label38"
                    Caption ="Reg#"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =2340
                    Top =60
                    Width =750
                    Height =300
                    Name ="Label39"
                    Caption ="Invoice"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3120
                    Top =60
                    Width =540
                    Height =300
                    Name ="Label40"
                    Caption ="Time"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8220
                    Top =60
                    Width =810
                    Height =300
                    Name ="Label59"
                    Caption ="Loans"
                End
                Begin Label
                    TextAlign =3
                    Left =6720
                    Top =60
                    Width =810
                    Height =300
                    Name ="Label58"
                    Caption ="Drops"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =420
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =420
                    Top =60
                    Width =540
                    Height =300
                    FontSize =11
                    ForeColor =8388608
                    Name ="Text57"
                    ControlSource ="CashierID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =1020
                    Top =60
                    Width =1872
                    Height =300
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="CashierName"
                    ControlSource ="CashierName"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =270
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1200
                    Width =540
                    Height =270
                    FontSize =10
                    Name ="TenderID"
                    ControlSource ="TenderID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1800
                    Width =2640
                    Height =270
                    FontSize =10
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    Left =1200
                    Width =360
                    Height =225
                    Name ="Text43"
                    ControlSource ="CR"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1200
                    LayoutCachedWidth =1560
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =1
                    Left =1680
                    Width =1380
                    Height =225
                    TabIndex =1
                    Name ="Text44"
                    ControlSource ="Invoice"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1680
                    LayoutCachedWidth =3060
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =1
                    Left =3120
                    Width =780
                    Height =225
                    TabIndex =2
                    Name ="Text45"
                    ControlSource ="Time"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =1
                    Left =3900
                    Width =1140
                    Height =225
                    TabIndex =3
                    Name ="Text46"
                    ControlSource ="Transact"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6585
                    Width =1260
                    TabIndex =4
                    Name ="Drop"
                    ControlSource ="Drop"
                    Format ="$#,##0.00;($#,##0.00);\"\";\\l"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6585
                    LayoutCachedWidth =7845
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8160
                    Width =1260
                    TabIndex =5
                    Name ="Loan"
                    ControlSource ="Loan"
                    Format ="$#,##0.00;($#,##0.00);\"\";\\l"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8160
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupFooter2"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6360
                    Top =60
                    Width =1485
                    Height =225
                    FontWeight =700
                    Name ="Text72"
                    ControlSource ="=Sum([Drop])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6360
                    LayoutCachedTop =60
                    LayoutCachedWidth =7845
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =5220
                    Top =60
                    Width =1080
                    Height =225
                    TabIndex =1
                    Name ="Text73"
                    ControlSource ="=Sum([TranCount])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =3540
                    Top =60
                    Width =1560
                    Height =225
                    TabIndex =2
                    Name ="Text74"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =2160
                    Top =60
                    Width =8400
                    Name ="Line75"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7935
                    Top =60
                    Width =1485
                    Height =225
                    FontWeight =700
                    TabIndex =3
                    Name ="Text76"
                    ControlSource ="=Sum([Loan])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7935
                    LayoutCachedTop =60
                    LayoutCachedWidth =9420
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9420
                    Top =60
                    Width =1185
                    Height =225
                    FontWeight =700
                    TabIndex =4
                    Name ="Text77"
                    ControlSource ="=Sum([Drop])-Sum([Loan])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9420
                    LayoutCachedTop =60
                    LayoutCachedWidth =10605
                    LayoutCachedHeight =285
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6330
                    Top =120
                    Width =1365
                    Height =225
                    FontWeight =700
                    Name ="Text47"
                    ControlSource ="=Sum([Drop])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6330
                    LayoutCachedTop =120
                    LayoutCachedWidth =7695
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =5220
                    Top =120
                    Width =1080
                    Height =225
                    TabIndex =1
                    Name ="Text51"
                    ControlSource ="=Sum([TranCount])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    Left =840
                    Top =120
                    Width =1560
                    Height =225
                    TabIndex =2
                    Name ="Text34"
                    ControlSource ="CashierName"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =900
                    Top =60
                    Width =9660
                    Name ="Line36"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7905
                    Top =120
                    Width =1365
                    Height =225
                    FontWeight =700
                    TabIndex =3
                    Name ="Text60"
                    ControlSource ="=Sum([Loan])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7905
                    LayoutCachedTop =120
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9270
                    Top =120
                    Width =1365
                    Height =225
                    FontWeight =700
                    TabIndex =4
                    Name ="Text61"
                    ControlSource ="=Sum([Drop])-Sum([Loan])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9270
                    LayoutCachedTop =120
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =345
                End
            End
        End
        Begin PageFooter
            Height =660
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =360
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5760
                    Top =360
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin Line
                    BorderWidth =3
                    Width =10800
                    BorderColor =12632256
                    Name ="Line59"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10800
                    BorderColor =12632256
                    Name ="Line60"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =120
                    Top =60
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text63"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5760
                    Top =60
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =120
                    Top =360
                    Width =10680
                    Height =300
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =120
                    Top =60
                    Width =10680
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =330
            Name ="ReportFooter"
            Begin
                Begin Label
                    TextFontFamily =34
                    Left =3780
                    Top =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label53"
                    Caption ="Grand Total"
                    FontName ="Arial Narrow"
                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                    FontName ="Arial Narrow"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =6330
                    Top =60
                    Width =1365
                    Height =225
                    FontWeight =700
                    Name ="Text62"
                    ControlSource ="=Sum([Drop])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6330
                    LayoutCachedTop =60
                    LayoutCachedWidth =7695
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =5220
                    Top =60
                    Width =1080
                    Height =225
                    TabIndex =1
                    Name ="Text66"
                    ControlSource ="=Sum([TranCount])"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =7905
                    Top =60
                    Width =1365
                    Height =225
                    FontWeight =700
                    TabIndex =2
                    Name ="Text67"
                    ControlSource ="=Sum([Loan])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7905
                    LayoutCachedTop =60
                    LayoutCachedWidth =9270
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9270
                    Top =60
                    Width =1365
                    Height =225
                    FontWeight =700
                    TabIndex =3
                    Name ="Text69"
                    ControlSource ="=Sum([Drop])-Sum([Loan])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9270
                    LayoutCachedTop =60
                    LayoutCachedWidth =10635
                    LayoutCachedHeight =285
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD043CashDropsAndLoans.cls"
