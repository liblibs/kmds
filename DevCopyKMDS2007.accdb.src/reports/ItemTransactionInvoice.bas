Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AutoCenter = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =48
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8220
    DatasheetFontHeight =10
    ItemSuffix =59
    Left =5715
    Top =525
    RecSrcDt = Begin
        0x7b3aba8388fee340
    End
    RecordSource ="qryItemTransactionInvoice"
    Caption ="qryItemTransactionInvoice"
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
            ControlSource ="InvoHeadID"
        End
        Begin BreakLevel
            ControlSource ="ID"
        End
        Begin BreakLevel
            ControlSource ="TransType"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =420
            Name ="ReportHeader"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =7380
                    BorderColor =12632256
                    Name ="Line39"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =240
                    Top =180
                    ColumnOrder =0
                    Name ="CashierName"
                    ControlSource ="CashierName"

                End
            End
        End
        Begin PageHeader
            Visible = NotDefault
            Height =0
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1200
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =960
                    Top =360
                    Width =960
                    Height =300
                    FontSize =11
                    ForeColor =8388608
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =960
                    LayoutCachedTop =360
                    LayoutCachedWidth =1920
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            Left =360
                            Top =360
                            Width =495
                            Height =300
                            Name ="Label48"
                            Caption ="Inv#"
                            LayoutCachedLeft =360
                            LayoutCachedTop =360
                            LayoutCachedWidth =855
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =960
                    Top =600
                    Width =660
                    Height =300
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="CashSysID"
                    ControlSource ="CashSysID"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =960
                    LayoutCachedTop =600
                    LayoutCachedWidth =1620
                    LayoutCachedHeight =900
                    Begin
                        Begin Label
                            Left =360
                            Top =600
                            Width =540
                            Height =300
                            Name ="Label49"
                            Caption ="Reg#"
                            LayoutCachedLeft =360
                            LayoutCachedTop =600
                            LayoutCachedWidth =900
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3000
                    Top =360
                    Height =300
                    FontSize =11
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text46"
                    ControlSource ="CashierName"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =360
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            Left =2100
                            Top =360
                            Width =855
                            Height =300
                            Name ="Label47"
                            Caption ="Cashier:"
                            LayoutCachedLeft =2100
                            LayoutCachedTop =360
                            LayoutCachedWidth =2955
                            LayoutCachedHeight =660
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5580
                    Top =360
                    Width =2220
                    Height =300
                    ColumnWidth =2400
                    FontSize =11
                    TabIndex =3
                    ForeColor =8388608
                    Name ="ClubName"
                    ControlSource ="ClubName"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =3000
                    Top =660
                    Width =2475
                    Height =300
                    ColumnWidth =2280
                    FontSize =11
                    TabIndex =4
                    ForeColor =8388608
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =660
                    LayoutCachedWidth =5475
                    LayoutCachedHeight =960
                    Begin
                        Begin Label
                            Left =1860
                            Top =660
                            Width =1095
                            Height =300
                            Name ="Label50"
                            Caption ="Trans Date"
                            LayoutCachedLeft =1860
                            LayoutCachedTop =660
                            LayoutCachedWidth =2955
                            LayoutCachedHeight =960
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =1020
                    Width =8100
                    Name ="Line41"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1380
                    Top =60
                    Width =960
                    Height =300
                    FontSize =11
                    TabIndex =5
                    ForeColor =8388608
                    Name ="Text55"
                    ControlSource ="TransferNumber"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            Left =60
                            Top =60
                            Width =1200
                            Height =300
                            Name ="Label56"
                            Caption ="Reference #"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5580
                    Top =660
                    Width =2520
                    Height =300
                    FontSize =11
                    TabIndex =6
                    ForeColor =8388608
                    Name ="Text57"
                    ControlSource ="LastName"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4500
                    Top =360
                    Width =1020
                    Height =300
                    FontSize =11
                    TabIndex =7
                    ForeColor =8388608
                    Name ="Text58"
                    ControlSource ="SalesDept"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =4500
                    LayoutCachedTop =360
                    LayoutCachedWidth =5520
                    LayoutCachedHeight =660
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            OnPrint ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Width =480
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =540
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =720
                    Width =1020
                    Height =225
                    TabIndex =1
                    Name ="Code"
                    ControlSource ="Code"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =720
                    LayoutCachedWidth =1740
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7980
                    Width =240
                    Height =225
                    TabIndex =2
                    Name ="SalesDept"
                    ControlSource ="SalesDept"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3600
                    Width =600
                    Height =225
                    TabIndex =3
                    Name ="Quantity"
                    ControlSource ="Qty"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5280
                    Width =1260
                    Height =225
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="Extend"
                    Format ="#,###.00;-#,###.00;\"\";\"\""
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6600
                    Width =1620
                    Height =225
                    TabIndex =5
                    Name ="InputCode"
                    ControlSource ="InputCode"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1860
                    Width =1680
                    ColumnWidth =3150
                    TabIndex =6
                    Name ="Description"
                    ControlSource ="CodeDesc"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4260
                    Width =1020
                    TabIndex =7
                    Name ="UPrice"
                    ControlSource ="UPrice"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =7560
                    Width =360
                    Height =225
                    TabIndex =8
                    Name ="TransType"
                    ControlSource ="TransType"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Visible = NotDefault
                    Left =120
                    Top =240
                    Width =7380
                    Name ="Line5"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =0
            Name ="GroupFooter1"
        End
        Begin PageFooter
            Height =300
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Width =2280
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text37"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =6180
                    Width =1740
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text38"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "ItemTransactionInvoice.cls"
