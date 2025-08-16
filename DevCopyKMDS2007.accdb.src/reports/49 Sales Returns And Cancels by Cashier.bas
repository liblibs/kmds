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
    Width =10710
    DatasheetFontHeight =10
    ItemSuffix =90
    Left =1515
    RecSrcDt = Begin
        0x95fbcfd51612e340
    End
    RecordSource ="qrySalesReturnsAndCancelsByCashierByItem"
    Caption ="EOD  Returns  and Cancels By Cashier By Item"
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
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="Name"
        End
        Begin BreakLevel
            ControlSource ="ID"
        End
        Begin FormHeader
            Height =540
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =-135
                    Width =10755
                    Height =540
                    FontSize =20
                    Name ="Label47"
                    Caption ="Returns and Cancels By Cashier By Invoice"
                End
            End
        End
        Begin PageHeader
            Height =435
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =120
                    Top =60
                    Width =780
                    Height =300
                    Name ="Name_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =8100
                    Top =60
                    Width =495
                    Height =300
                    Name ="TotalTransactions_Label"
                    Caption ="Date"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Width =10560
                    Name ="Line41"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =10560
                    Name ="Line42"
                End
                Begin Label
                    Left =1200
                    Top =60
                    Width =540
                    Height =300
                    Name ="Label46"
                    Caption =" Item"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =8520
                    Top =60
                    Width =1335
                    Height =300
                    Name ="Label57"
                    Caption ="Qty"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =8520
                    LayoutCachedTop =60
                    LayoutCachedWidth =9855
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =3
                    Left =9900
                    Top =60
                    Width =810
                    Height =300
                    Name ="Label58"
                    Caption ="Amount"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9900
                    LayoutCachedTop =60
                    LayoutCachedWidth =10710
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =5700
                    Top =60
                    Width =1665
                    Height =300
                    Name ="Label74"
                    Caption ="Register Invoice "
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =7365
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =2
                    Left =4725
                    Top =60
                    Width =645
                    Height =300
                    Name ="Label84"
                    Caption ="UPC"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4725
                    LayoutCachedTop =60
                    LayoutCachedWidth =5370
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =3540
                    Top =60
                    Width =1005
                    Height =300
                    Name ="Label89"
                    Caption =" ON Shelf"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =240
            Name ="GroupHeader2"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    FontWeight =700
                    Name ="Text75"
                    ControlSource ="Name"

                End
            End
        End
        Begin Section
            Height =248
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =1
                    IMESentenceMode =3
                    Left =90
                    Width =1290
                    Height =225
                    Name ="Code"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =90
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1440
                    Width =2520
                    Height =225
                    ColumnWidth =3150
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1440
                    LayoutCachedWidth =3960
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3900
                    Width =660
                    Height =225
                    FontWeight =700
                    TabIndex =2
                    Name ="SumOfQuantity"
                    ControlSource ="OnShelf"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4680
                    Height =225
                    TabIndex =3
                    Name ="InputCode"
                    ControlSource ="InputCode"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =7320
                    Width =1920
                    ColumnWidth =2100
                    TabIndex =4
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    Format ="General Date"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7320
                    LayoutCachedWidth =9240
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6180
                    Width =240
                    Height =225
                    TabIndex =5
                    Name ="Text85"
                    ControlSource ="CashSysID"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6375
                    Width =900
                    Height =225
                    TabIndex =6
                    Name ="Text86"
                    ControlSource ="InvoHeadID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6375
                    LayoutCachedWidth =7275
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9047
                    Top =23
                    Width =660
                    Height =225
                    TabIndex =7
                    Name ="Text87"
                    ControlSource ="Quantity"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9047
                    LayoutCachedTop =23
                    LayoutCachedWidth =9707
                    LayoutCachedHeight =248
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9780
                    Width =840
                    Height =225
                    TabIndex =8
                    Name ="Text88"
                    ControlSource ="Price"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9780
                    LayoutCachedWidth =10620
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =480
            Name ="GroupFooter0"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6840
                    Top =120
                    FontWeight =700
                    Name ="CashierName"
                    ControlSource ="Name"

                End
                Begin Line
                    BorderWidth =1
                    Top =420
                    Width =10560
                    Name ="Line69"
                End
                Begin Line
                    BorderWidth =1
                    Left =6900
                    Top =60
                    Width =3660
                    Name ="Line73"
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9540
                    Top =120
                    Width =1020
                    FontWeight =700
                    TabIndex =1
                    Name ="Text80"
                    ControlSource ="=Sum([Price])"
                    Format ="Standard"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =120
                    Width =660
                    FontWeight =700
                    TabIndex =2
                    Name ="Text81"
                    ControlSource ="=Sum([Quantity])"

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
                    FontWeight =700
                    Name ="Text82"
                    ControlSource ="=Sum([Price])"
                    Format ="Standard"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8340
                    Top =120
                    Width =660
                    FontWeight =700
                    TabIndex =1
                    Name ="Text83"
                    ControlSource ="=Sum([Quantity])"

                End
            End
        End
    End
End
CodeBehindForm
' See "49 Sales Returns And Cancels by Cashier.cls"
