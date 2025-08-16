Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
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
    ItemSuffix =38
    Left =5340
    Top =1305
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x9188db240b7de540
    End
    RecordSource ="qryEO128LiquorItemsNotScanned"
    Caption ="qryEOD128RLiquorItemsNotScanned"
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
            ControlSource ="ItemID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1155
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    Left =60
                    Top =600
                    Width =10620
                    Height =510
                    ColumnOrder =0
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
                    Left =-15
                    Top =60
                    Width =10755
                    Height =510
                    FontSize =20
                    Name ="Label28"
                    Caption ="EOD Liquor Items Not Scanned"
                    LayoutCachedLeft =-15
                    LayoutCachedTop =60
                    LayoutCachedWidth =10740
                    LayoutCachedHeight =570
                End
                Begin Line
                    BorderWidth =3
                    Top =60
                    Width =10740
                    BorderColor =12632256
                    Name ="Line31"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10740
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =480
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =120
                    Top =60
                    Width =1020
                    Height =300
                    Name ="CashierID_Label"
                    Caption ="Cashier"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =3
                    Left =9900
                    Top =60
                    Width =810
                    Height =300
                    Name ="Amount_Label"
                    Caption ="Amount"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1380
                    Top =60
                    Width =360
                    Height =300
                    Name ="CR_Label"
                    Caption ="CR"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1380
                    LayoutCachedTop =60
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =1800
                    Top =60
                    Width =1110
                    Height =300
                    Name ="Invoice_Label"
                    Caption ="Invoice"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =2910
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =3240
                    Top =60
                    Width =540
                    Height =300
                    Name ="Time_Label"
                    Caption ="Time"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =3240
                    LayoutCachedTop =60
                    LayoutCachedWidth =3780
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =4740
                    Top =60
                    Width =3270
                    Height =300
                    Name ="CustomerID_Label"
                    Caption ="Item -       Entered - Description"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4740
                    LayoutCachedTop =60
                    LayoutCachedWidth =8010
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =8820
                    Top =60
                    Width =630
                    Height =300
                    Name ="TranCount_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line32"
                End
                Begin Line
                    BorderWidth =2
                    Top =420
                    Width =10800
                    Name ="Line33"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =540
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextFontFamily =18
                    Left =60
                    Top =120
                    Width =540
                    Height =300
                    ColumnWidth =1065
                    FontSize =11
                    ForeColor =8388608
                    Name ="CashierID"
                    ControlSource ="CashierID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    Left =660
                    Top =120
                    Width =1872
                    Height =300
                    ColumnWidth =705
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="CashierName"
                    ControlSource ="CashierName"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9720
                    Width =1035
                    Height =225
                    ColumnWidth =885
                    Name ="Amount"
                    ControlSource ="Amount"
                    Format ="Standard"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Left =1140
                    Width =360
                    Height =225
                    TabIndex =1
                    Name ="CR"
                    ControlSource ="Reigister"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =1560
                    Width =1080
                    Height =225
                    TabIndex =2
                    Name ="Invoice"
                    ControlSource ="Invoice"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1560
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Left =2700
                    Width =1980
                    Height =225
                    ColumnWidth =630
                    TabIndex =3
                    Name ="Time"
                    ControlSource ="Date-Time"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2700
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =1
                    Left =4740
                    Width =1020
                    Height =225
                    ColumnWidth =930
                    TabIndex =4
                    Name ="CustomerID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4740
                    LayoutCachedWidth =5760
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =1
                    Left =5700
                    Width =1020
                    Height =225
                    ColumnWidth =2070
                    TabIndex =5
                    Name ="Customer"
                    ControlSource ="UPCCode"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5700
                    LayoutCachedWidth =6720
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    Left =8880
                    Width =540
                    Height =225
                    ColumnWidth =1140
                    TabIndex =6
                    Name ="TranCount"
                    ControlSource ="Quantity"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6840
                    Width =1860
                    Height =225
                    TabIndex =7
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6840
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =225
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    TextAlign =3
                    Left =9720
                    Top =240
                    Width =1035
                    Height =225
                    Name ="Sum Of Price"
                    ControlSource ="=Sum([Amount])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_Price"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =8340
                    Top =240
                    Width =1080
                    Height =225
                    TabIndex =1
                    Name ="Sum Of TranCount"
                    ControlSource ="=Sum([quantity])"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Sum_Of_TranCount"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    Left =6660
                    Top =240
                    Width =1560
                    Height =225
                    TabIndex =2
                    Name ="Text34"
                    ControlSource ="CashierName"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Left =6660
                    Top =120
                    Width =4080
                    Name ="Line36"
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
            Height =420
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =2
                    OldBorderStyle =1
                    TextAlign =3
                    Left =9720
                    Top =60
                    Width =1035
                    Height =285
                    FontSize =10
                    Name ="Price Grand Total Sum"
                    ControlSource ="=Sum([Amount])"
                    Format ="Standard"
                    FontName ="Arial Narrow"
                    EventProcPrefix ="Price_Grand_Total_Sum"

                End
                Begin Label
                    TextFontFamily =34
                    Left =6360
                    Top =60
                    Width =1095
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label23"
                    Caption ="Grand Total"
                    FontName ="Arial Narrow"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =3
                    Left =8280
                    Top =60
                    Width =1155
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text35"
                    ControlSource ="=Sum([quantity])"
                    FontName ="Arial Narrow"

                End
                Begin Label
                    Visible = NotDefault
                    TextFontFamily =34
                    Left =1020
                    Top =120
                    Width =3240
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="lblNoData"
                    Caption ="No Data for this report"
                    FontName ="Arial Narrow"
                End
            End
        End
    End
End
CodeBehindForm
' See "qryEOD128RLiquorItemsNotScanned.cls"
