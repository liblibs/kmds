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
    Width =10800
    DatasheetFontHeight =10
    ItemSuffix =55
    RecSrcDt = Begin
        0xd7cf26e06b0be340
    End
    RecordSource ="qryItemTransactionInvoiceForVendors"
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
            Height =930
            Name ="ReportHeader"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =60
                    Width =10800
                    BorderColor =12632256
                    Name ="Line39"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2880
                    Top =240
                    Name ="CashierName"
                    ControlSource ="CashierName"

                End
            End
        End
        Begin PageHeader
            Visible = NotDefault
            Height =900
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Width =10800
                    Name ="Line40"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =720
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    DecimalPlaces =0
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1680
                    Top =300
                    Width =960
                    Height =300
                    FontSize =11
                    ForeColor =8388608
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            Left =1680
                            Width =495
                            Height =300
                            Name ="Label48"
                            Caption ="Inv#"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =2760
                    Top =300
                    Width =660
                    Height =300
                    FontSize =11
                    TabIndex =1
                    ForeColor =8388608
                    Name ="CashSysID"
                    ControlSource ="CashSysID"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            Left =2760
                            Width =540
                            Height =300
                            Name ="Label49"
                            Caption ="Reg#"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4080
                    Top =300
                    Height =300
                    FontSize =11
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text46"
                    ControlSource ="CashierName"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            Left =4020
                            Width =855
                            Height =300
                            Name ="Label47"
                            Caption ="Cashier:"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =5700
                    Top =300
                    Width =3300
                    Height =300
                    ColumnWidth =2400
                    FontSize =11
                    TabIndex =3
                    ForeColor =8388608
                    Name ="ClubName"
                    ControlSource ="ClubName"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            Left =5700
                            Width =960
                            Height =300
                            Name ="Label45"
                            Caption ="Customer"
                        End
                    End
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =9060
                    Top =300
                    Width =1680
                    Height =300
                    ColumnWidth =2280
                    FontSize =11
                    TabIndex =4
                    ForeColor =8388608
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Times New Roman"

                    Begin
                        Begin Label
                            Left =9060
                            Width =1095
                            Height =300
                            Name ="Label50"
                            Caption ="Trans Date"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =660
                    Width =10800
                    Name ="Line41"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =1560
                    Height =300
                    ColumnWidth =2010
                    FontSize =11
                    TabIndex =5
                    ForeColor =8388608
                    Name ="Transaction"
                    ControlSource ="Transaction"
                    FontName ="Times New Roman"

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
                    TextAlign =3
                    IMESentenceMode =3
                    Left =300
                    Width =900
                    Height =225
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =1320
                    Width =660
                    Height =225
                    TabIndex =1
                    Name ="Code"
                    ControlSource ="Code"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =360
                    Width =240
                    Height =225
                    TabIndex =2
                    Name ="SalesDept"
                    ControlSource ="SalesDept"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5160
                    Width =600
                    Height =225
                    TabIndex =3
                    Name ="Quantity"
                    ControlSource ="Qty"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6900
                    Width =1260
                    Height =225
                    TabIndex =4
                    Name ="Price"
                    ControlSource ="Extend"
                    Format ="#,###.00;-#,###.00;\"\";\"\""

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =8880
                    Width =1620
                    Height =225
                    TabIndex =5
                    Name ="InputCode"
                    ControlSource ="InputCode"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2100
                    Width =3000
                    Height =225
                    ColumnWidth =3150
                    TabIndex =6
                    Name ="Description"
                    ControlSource ="CodeDesc"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6000
                    Width =780
                    Height =225
                    TabIndex =7
                    Name ="UPrice"
                    ControlSource ="UPrice"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =60
                    Width =360
                    Height =225
                    TabIndex =8
                    Name ="TransType"
                    ControlSource ="TransType"

                End
                Begin Line
                    Visible = NotDefault
                    Left =840
                    Top =240
                    Width =9660
                    Name ="Line5"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =420
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    IMESentenceMode =3
                    Left =4740
                    Top =60
                    Width =1020
                    Height =270
                    Name ="Sum Of Quantity"
                    ControlSource ="=Sum([Qty])"
                    EventProcPrefix ="Sum_Of_Quantity"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Top =60
                    Width =1560
                    Height =270
                    TabIndex =1
                    Name ="Sum Of Price"
                    ControlSource ="=Sum([Extend])"
                    Format ="#,###.00;-#,###.00;\"\";\"\""
                    EventProcPrefix ="Sum_Of_Price"

                End
                Begin Line
                    BorderWidth =2
                    Width =10740
                    Name ="Line51"
                End
            End
        End
        Begin PageFooter
            Height =540
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =60
                    Top =240
                    Width =5040
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
                    Left =5700
                    Top =240
                    Width =5040
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
' See "ItemTransactionInvoiceForVendors.cls"
