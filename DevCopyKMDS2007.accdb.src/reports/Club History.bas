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
    Width =10740
    DatasheetFontHeight =10
    ItemSuffix =53
    Left =1920
    Top =30
    RecSrcDt = Begin
        0xf15edc60d9fae340
    End
    RecordSource ="qryClubOrderCreateFromSalesHistoryReport"
    Caption ="Club History"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1140
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =45
                    Top =60
                    Width =10635
                    Height =510
                    FontSize =20
                    Name ="Label20"
                    Caption ="Club History"
                    LayoutCachedLeft =45
                    LayoutCachedTop =60
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =570
                End
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =9360
                    BorderColor =12632256
                    Name ="Line23"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Top =600
                    Width =10680
                    Height =510
                    ColumnOrder =0
                    FontSize =20
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="Times New Roman"

                    LayoutCachedTop =600
                    LayoutCachedWidth =10680
                    LayoutCachedHeight =1110
                End
            End
        End
        Begin PageHeader
            Height =540
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =10020
                    Name ="Line24"
                    LayoutCachedWidth =10020
                End
                Begin Line
                    BorderWidth =2
                    Top =360
                    Width =10020
                    Name ="Line25"
                    LayoutCachedTop =360
                    LayoutCachedWidth =10020
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =0
                    Left =540
                    Width =630
                    Height =300
                    Name ="Label42"
                    Caption ="Code"
                    LayoutCachedLeft =540
                    LayoutCachedWidth =1170
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =0
                    Left =1320
                    Width =1215
                    Height =300
                    Name ="Label43"
                    Caption ="Description"
                    LayoutCachedLeft =1320
                    LayoutCachedWidth =2535
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =3
                    Left =3720
                    Width =1050
                    Height =330
                    Name ="Label44"
                    Caption ="Purchases"
                    LayoutCachedLeft =3720
                    LayoutCachedWidth =4770
                    LayoutCachedHeight =330
                End
                Begin Label
                    TextAlign =0
                    Left =6180
                    Width =1455
                    Height =300
                    Name ="Label45"
                    Caption ="First Purchase"
                    LayoutCachedLeft =6180
                    LayoutCachedWidth =7635
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =0
                    Left =7860
                    Width =2175
                    Height =300
                    Name ="Label46"
                    Caption ="Most Recent Purchase"
                    LayoutCachedLeft =7860
                    LayoutCachedWidth =10035
                    LayoutCachedHeight =300
                End
                Begin Label
                    TextAlign =3
                    Left =4980
                    Width =960
                    Height =300
                    Name ="Label47"
                    Caption ="Price"
                    LayoutCachedLeft =4980
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =300
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =300
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Width =1020
                    Height =225
                    Name ="Code"
                    ControlSource ="Code"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1260
                    Width =2280
                    Height =225
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =1260
                    LayoutCachedWidth =3540
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3720
                    Width =960
                    Height =225
                    TabIndex =2
                    Name ="SumOfQuantity"
                    ControlSource ="SumOfQuantity"

                    LayoutCachedLeft =3720
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6180
                    Height =225
                    ColumnWidth =2250
                    TabIndex =3
                    Name ="FirstOfSetupDate"
                    ControlSource ="MinOfSetupDate"
                    Format ="Medium Date"

                    LayoutCachedLeft =6180
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8400
                    Height =225
                    ColumnWidth =2250
                    TabIndex =4
                    Name ="LastOfSetupDate"
                    ControlSource ="MaxOfSetupDate"
                    Format ="Medium Date"

                    LayoutCachedLeft =8400
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4980
                    Width =960
                    Height =225
                    TabIndex =5
                    Name ="ExtPrice"
                    ControlSource ="ExtPrice"

                    LayoutCachedLeft =4980
                    LayoutCachedWidth =5940
                    LayoutCachedHeight =225
                End
            End
        End
        Begin PageFooter
            Height =555
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =300
                    Width =1920
                    Height =255
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
                    Left =5700
                    Top =300
                    Width =5040
                    Height =255
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =5340
                    Height =255
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
                    Left =5700
                    Width =5040
                    Height =255
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
                    Height =255
                    FontSize =6
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
                    Left =60
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
            Height =780
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3720
                    Top =120
                    Width =960
                    Height =285
                    FontSize =10
                    Name ="Text50"
                    ControlSource ="=Sum([SumOfQuantity])"
                    Format ="Standard"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3720
                    LayoutCachedTop =120
                    LayoutCachedWidth =4680
                    LayoutCachedHeight =405
                    Begin
                        Begin Label
                            Left =240
                            Top =120
                            Width =1050
                            Height =270
                            FontSize =10
                            ForeColor =0
                            Name ="Label49"
                            Caption ="Report total"
                            LayoutCachedLeft =240
                            LayoutCachedTop =120
                            LayoutCachedWidth =1290
                            LayoutCachedHeight =390
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4740
                    Top =120
                    Width =1275
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text51"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="$#,##0.00;($#,##0.00)"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4740
                    LayoutCachedTop =120
                    LayoutCachedWidth =6015
                    LayoutCachedHeight =405
                    CurrencySymbol ="$"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =2040
                    Top =120
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text41"
                    ControlSource ="=Count([Code])"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =2040
                    LayoutCachedTop =120
                    LayoutCachedWidth =3480
                    LayoutCachedHeight =405
                End
                Begin Label
                    Left =60
                    Top =420
                    Width =2655
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label52"
                    Caption ="Note: Sales history is 500 Days."
                    LayoutCachedLeft =60
                    LayoutCachedTop =420
                    LayoutCachedWidth =2715
                    LayoutCachedHeight =690
                End
            End
        End
    End
End
CodeBehindForm
' See "Club History.cls"
