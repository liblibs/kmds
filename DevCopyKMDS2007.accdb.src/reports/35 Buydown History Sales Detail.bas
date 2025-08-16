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
    Width =10695
    DatasheetFontHeight =10
    ItemSuffix =113
    Left =525
    Top =1230
    OnNoData ="[Event Procedure]"
    RecSrcDt = Begin
        0x5fab525f50d8e240
    End
    RecordSource ="qryBuyDownReport"
    OnOpen ="[Event Procedure]"
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
            ControlSource ="BuyDownID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1740
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10621
                    Height =420
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =45
                    Top =60
                    Width =10635
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Buydowns Report"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1140
                    Width =10680
                    Height =30
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1440
                    Width =10621
                    Height =300
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="=\"Between \" & Format([BeginDate],\"mm/dd/yyyy\") & \" And \" & Format([EndDate"
                        "],\"mm/dd/yyyy\")"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =510
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =1020
                    Top =60
                    Width =1155
                    Height =300
                    Name ="Department_Label"
                    Caption ="Item"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5340
                    Top =120
                    Width =1140
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Quantity"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Width =10635
                    Height =30
                    Name ="Line54"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =450
                    Width =10635
                    Height =30
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =480
                    Width =10635
                    Height =30
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6420
                    Top =120
                    Width =900
                    Height =300
                    Name ="Label86"
                    Caption ="Amount"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            Name ="GroupHeader3"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =120
                    Top =60
                    Width =840
                    Height =225
                    FontWeight =700
                    Name ="Group"
                    ControlSource ="RepsName"

                End
                Begin TextBox
                    TextAlign =1
                    Left =1020
                    Top =60
                    Width =7119
                    Height =225
                    ColumnWidth =1185
                    FontWeight =700
                    TabIndex =1
                    Name ="GroupDesc"
                    ControlSource ="ReferenceNo"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Visible = NotDefault
            Height =570
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Left =5280
                    Width =2100
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="InvoiceDate"
                    Format ="General Date"

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7620
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Purchases"
                    ControlSource ="Quantity"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8460
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Counter"
                    ControlSource ="Price"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin TextBox
                    TextAlign =1
                    Left =540
                    Width =2100
                    Height =225
                    TabIndex =3
                    Name ="Text109"
                    ControlSource ="Description"
                    Format ="General Date"

                End
                Begin TextBox
                    TextAlign =1
                    Left =2700
                    Width =780
                    Height =225
                    TabIndex =4
                    Name ="Text110"
                    ControlSource ="ItemsCode"
                    Format ="General Date"

                End
                Begin TextBox
                    TextAlign =1
                    Left =3540
                    Width =780
                    Height =225
                    TabIndex =5
                    Name ="Text111"
                    ControlSource ="CashRegister"
                    Format ="General Date"

                End
                Begin TextBox
                    TextAlign =1
                    Left =4380
                    Width =780
                    Height =225
                    TabIndex =6
                    Name ="Text112"
                    ControlSource ="InvoiceNumber"
                    Format ="General Date"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            Name ="GroupFooter4"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    Left =2580
                    Top =60
                    Width =2880
                    Height =225
                    Name ="Text47"
                    ControlSource ="=\"Total\""

                End
                Begin TextBox
                    DecimalPlaces =0
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5760
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =1
                    Name ="Text102"
                    ControlSource ="=Sum([quantity])"
                    Format ="#,###;-#;\"\";\"\""

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6600
                    Top =60
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Text103"
                    ControlSource ="=Sum([Price])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
            End
        End
        Begin PageFooter
            Height =585
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
                    Width =4965
                    Height =285
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
                    Width =4380
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
                    Width =4965
                    Height =285
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
                    Width =10620
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "35 Buydown History Sales Detail.cls"
