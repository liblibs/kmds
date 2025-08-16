Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
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
    ItemSuffix =136
    Left =1050
    Top =1860
    OrderBy ="Description"
    RecSrcDt = Begin
        0x53a78cbef5a8e240
    End
    RecordSource ="qryBuyDowns"
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
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =660
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =15
                    Top =60
                    Width =10695
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Items on Buydowns"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10755
                    BorderColor =12632256
                    Name ="Line51"
                End
            End
        End
        Begin PageHeader
            Height =0
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =1215
            Name ="GroupHeader0"
            Begin
                Begin Label
                    TextAlign =3
                    Left =7740
                    Top =840
                    Width =1365
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Qty Sold"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =840
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =60
                    Top =840
                    Width =780
                    Height =300
                    Name ="Label85"
                    Caption ="Item"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =1440
                    Width =2340
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Name"
                    ControlSource ="RepsName"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1440
                    Top =300
                    Width =2520
                    TabIndex =1
                    Name ="Contact"
                    ControlSource ="ReferenceNo"

                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5085
                    Width =1350
                    Height =300
                    FontSize =11
                    TabIndex =2
                    Name ="BeginDate"
                    ControlSource ="BeginDate"
                    Format ="Short Date"

                    LayoutCachedLeft =5085
                    LayoutCachedWidth =6435
                    LayoutCachedHeight =300
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =3840
                            Width =1215
                            Height =300
                            Name ="Label112"
                            Caption ="Begin Date:"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5085
                    Top =360
                    Width =1470
                    Height =300
                    FontSize =11
                    TabIndex =3
                    Name ="EndDate"
                    ControlSource ="EndDate"
                    Format ="Short Date"

                    LayoutCachedLeft =5085
                    LayoutCachedTop =360
                    LayoutCachedWidth =6555
                    LayoutCachedHeight =660
                    Begin
                        Begin Label
                            TextAlign =3
                            Left =4005
                            Top =360
                            Width =1050
                            Height =300
                            Name ="Label113"
                            Caption ="End Date:"
                        End
                    End
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =540
                    Top =60
                    Width =780
                    TabIndex =4
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            TextAlign =0
                            Left =120
                            Top =60
                            Width =390
                            Height =300
                            FontSize =8
                            Name ="Label116"
                            Caption ="ID:"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7785
                    Width =720
                    Height =300
                    TabIndex =5
                    Name ="ItemsID"
                    ControlSource ="PromoItemID"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =6735
                            Width =1020
                            Height =225
                            FontSize =8
                            Name ="Label117"
                            Caption ="Discount Item"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8580
                    Width =2160
                    Height =300
                    TabIndex =6
                    Name ="Item.Description"
                    ControlSource ="PromItemDesc"
                    EventProcPrefix ="Item_Description"

                End
                Begin Line
                    BorderWidth =2
                    Top =1200
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =9240
                    Top =840
                    Width =1365
                    Height =300
                    Name ="Label125"
                    Caption ="$$$ Sold"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =3630
                    Top =840
                    Width =1065
                    Height =300
                    Name ="Label128"
                    Caption ="SalesDptD"
                End
                Begin Label
                    TextAlign =2
                    Left =6480
                    Top =840
                    Width =1095
                    Height =300
                    Name ="Label129"
                    Caption ="SizeD"
                End
                Begin Label
                    Left =1125
                    Top =840
                    Width =1695
                    Height =300
                    Name ="Label130"
                    Caption ="Description"
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8520
                    Top =360
                    ColumnWidth =1575
                    TabIndex =7
                    Name ="BuydownDollars"
                    ControlSource ="BuydownDollars"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7080
                    Top =360
                    TabIndex =8
                    Name ="BuydownQty"
                    ControlSource ="BuydownQty"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =10680
                    Name ="Line5"
                End
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Width =1080
                    Name ="ItemsCode"
                    ControlSource ="ItemsCode"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1200
                    Width =2580
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7680
                    TabIndex =2
                    Name ="QuantitySold"
                    ControlSource ="QuantitySold"
                    Format ="#,##0;-#,##0;\"\";\"Null\""

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9180
                    TabIndex =3
                    Name ="DiscountDollars"
                    ControlSource ="DiscountDollars"
                    Format ="#,##0.00;-#,##0.00;\"\";\"Null\""

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3840
                    Width =2700
                    TabIndex =4
                    Name ="SalesdptD"
                    ControlSource ="SalesdptD"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6660
                    Width =900
                    TabIndex =5
                    Name ="SizeD"
                    ControlSource ="SizeD"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            ForceNewPage =2
            Height =540
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7680
                    Top =120
                    Name ="Text131"
                    ControlSource ="=Sum([QuantitySold])"

                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9180
                    Top =120
                    TabIndex =1
                    Name ="Text132"
                    ControlSource ="=Sum([DiscountDollars])"
                    Format ="$#,##0.00;($#,##0.00)"

                End
                Begin Line
                    Left =7740
                    Top =60
                    Width =2880
                    Name ="Line133"
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
                    Width =4860
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
                    Left =120
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
            Height =0
            Name ="ReportFooter"
        End
    End
End
CodeBehindForm
' See "09 Items on Buydown.cls"
