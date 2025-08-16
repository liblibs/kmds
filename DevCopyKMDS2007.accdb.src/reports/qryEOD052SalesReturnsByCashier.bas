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
    Width =10680
    DatasheetFontHeight =10
    ItemSuffix =75
    Top =600
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x946eaeabe9c8e240
    End
    RecordSource ="qryEOSalesReturnsByCashierByItem"
    Caption ="EOD Sales Returns By Cashier By Item"
    DatasheetFontName ="Trebuchet MS"
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
            GroupFooter = NotDefault
            ControlSource ="CashierID"
        End
        Begin FormHeader
            Height =1620
            Name ="ReportHeader"
            Begin
                Begin Label
                    Left =7560
                    Top =900
                    Width =780
                    Height =300
                    Name ="Name_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =1207
                    Top =900
                    Width =1230
                    Height =300
                    Name ="TotalTransactions_Label"
                    Caption ="Invoice Date"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =540
                    Width =10560
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =2
                    Top =1260
                    Width =10560
                    Name ="Line42"
                End
                Begin Label
                    Left =4500
                    Top =900
                    Width =1395
                    Height =300
                    Name ="Label46"
                    Caption ="Returned Item"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-30
                    Width =9225
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="EOD Sales Returns By Cashier By Item"
                End
                Begin Label
                    TextAlign =3
                    Left =8760
                    Top =600
                    Width =780
                    Height =600
                    Name ="Label57"
                    Caption ="Return Qty "
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9750
                    Top =600
                    Width =810
                    Height =600
                    Name ="Label58"
                    Caption ="Return Amount"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =2520
                    Top =900
                    Width =1575
                    Height =300
                    Name ="Label74"
                    Caption ="Invoice Number"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin PageHeader
            Height =540
            Name ="PageHeaderSection"
        End
        Begin Section
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =2520
                    Width =240
                    Height =225
                    Name ="CashSysID"
                    ControlSource ="CashSysID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2520
                    LayoutCachedWidth =2760
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =2820
                    Height =225
                    TabIndex =1
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2820
                    LayoutCachedWidth =4260
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4380
                    Width =1260
                    Height =225
                    TabIndex =2
                    Name ="Code"
                    ControlSource ="Code"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4380
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5700
                    Width =2820
                    Height =225
                    ColumnWidth =3150
                    TabIndex =3
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5700
                    LayoutCachedWidth =8520
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8580
                    Width =660
                    Height =225
                    TabIndex =4
                    Name ="SumOfQuantity"
                    ControlSource ="SumOfQuantity"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8580
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Width =1020
                    Height =225
                    TabIndex =5
                    Name ="SumOfPrice"
                    ControlSource ="SumOfPrice"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9540
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =240
                    Width =1920
                    ColumnWidth =2100
                    TabIndex =6
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =240
                    LayoutCachedWidth =2160
                    LayoutCachedHeight =240
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6480
                    Top =120
                    Width =300
                    FontWeight =700
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =120
                    FontWeight =700
                    TabIndex =1
                    Name ="CashierName"
                    ControlSource ="CashierName"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =120
                    Width =1020
                    FontWeight =700
                    TabIndex =2
                    Name ="Text67"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8580
                    Top =120
                    Width =660
                    FontWeight =700
                    TabIndex =3
                    Name ="Text68"
                    ControlSource ="=Sum([SumOfQuantity])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8580
                    LayoutCachedTop =120
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =360
                End
                Begin Line
                    BorderWidth =1
                    Top =420
                    Width =10560
                    Name ="Line69"
                End
                Begin Line
                    BorderWidth =1
                    Left =6480
                    Top =60
                    Width =4080
                    Name ="Line73"
                End
            End
        End
        Begin PageFooter
            Height =600
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Top =300
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5640
                    Top =300
                    Width =5040
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text50"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
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
                    Left =5640
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
                    Top =300
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
            Height =480
            Name ="ReportFooter"
            Begin
                Begin Line
                    Top =60
                    Width =10560
                    Name ="Line45"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =120
                    Width =1020
                    Height =300
                    FontWeight =700
                    Name ="Text71"
                    ControlSource ="=Sum([SumOfPrice])"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8580
                    Top =120
                    Width =660
                    Height =300
                    FontWeight =700
                    TabIndex =1
                    Name ="Text72"
                    ControlSource ="=Sum([SumOfQuantity])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =8580
                    LayoutCachedTop =120
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =420
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD052SalesReturnsByCashier.cls"
