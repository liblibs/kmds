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
    Width =10665
    DatasheetFontHeight =10
    ItemSuffix =104
    Left =585
    Top =1320
    OrderBy ="Description"
    RecSrcDt = Begin
        0x8753faa992e3e240
    End
    RecordSource ="qryInventoryCountBySubCategory"
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
            ControlSource ="SalesDepartmentID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="SubCatDesc"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1575
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10561
                    Height =300
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =30
                    Top =60
                    Width =10605
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Inventory Count Sheets By Category"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10620
                    Height =30
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1560
                    Width =10620
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =960
                    Width =10561
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboSalesDepartmentIDBegin & \" \" & F"
                        "orms!PrintInventoryReport!cboSalesDepartmentIDBegin.Column(1) & \" and \" & Form"
                        "s!PrintInventoryReport!cboSalesDepartmentIDEnd & \" \" & Forms!PrintInventoryRep"
                        "ort!cboSalesDepartmentIDEnd.Column(1)"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =1260
                    Width =10561
                    Height =300
                    ColumnOrder =2
                    TabIndex =2
                    Name ="Text79"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboCategoryIDBegin & \" and \" & Form"
                        "s!PrintInventoryReport!cboCategoryIDEnd"
                    Format ="Long Date"

                End
            End
        End
        Begin PageHeader
            Height =615
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =4890
                    Top =120
                    Width =1215
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="On the shelf"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =60
                    Width =10560
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =600
                    Width =10560
                    Name ="Line57"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =0
                    Left =60
                    Top =120
                    Width =1200
                    Height =300
                    Name ="Label77"
                    Caption ="Department"
                End
                Begin Label
                    TextAlign =3
                    Left =6270
                    Top =120
                    Width =930
                    Height =300
                    Name ="Label79"
                    Caption ="Reserved"
                End
                Begin Label
                    TextAlign =3
                    Left =7305
                    Top =120
                    Width =945
                    Height =300
                    Name ="Label80"
                    Caption ="Total"
                End
                Begin Label
                    TextAlign =3
                    Left =8445
                    Top =180
                    Width =2115
                    Height =300
                    Name ="Label82"
                    Caption ="Physical count"
                End
                Begin Label
                    TextAlign =2
                    Left =1080
                    Top =120
                    Width =960
                    Height =300
                    Name ="Label85"
                    Caption ="Item"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =300
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Width =480
                    Name ="Location"
                    ControlSource ="SalesDepartmentID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =600
                    Width =3540
                    TabIndex =1
                    Name ="Text90"
                    ControlSource ="DepartmentDesc"

                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Width =5763
                    Height =300
                    FontSize =11
                    ForeColor =8388608
                    Name ="SubCategoryDescription"
                    ControlSource ="SubCatDesc"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =435
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =540
                    Top =120
                    Width =1320
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="Item.ID"

                    LayoutCachedLeft =540
                    LayoutCachedTop =120
                    LayoutCachedWidth =1860
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    Left =1980
                    Top =120
                    Width =2460
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    TextAlign =3
                    Left =5400
                    Top =120
                    Width =720
                    Height =225
                    ColumnWidth =900
                    TabIndex =2
                    Name ="OnTheShelf"
                    ControlSource ="OnTheShelf"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4500
                    Top =120
                    Width =840
                    TabIndex =3
                    Name ="SizeD"
                    ControlSource ="SizeD"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6270
                    Top =120
                    Width =930
                    TabIndex =4
                    Name ="OnReserve"
                    ControlSource ="OnReserve"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7470
                    Top =120
                    Width =930
                    TabIndex =5
                    Name ="OnHand"
                    ControlSource ="=[OnTheShelf]-IIf(IsNull([OnReserve]),0,[OnReserve])"

                End
                Begin Line
                    Left =9000
                    Top =360
                    Width =1620
                    Name ="Line83"
                End
                Begin Line
                    Left =60
                    Top =420
                    Width =10560
                    Name ="line5"
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
                    IMESentenceMode =3
                    Left =1800
                    Top =60
                    Width =3540
                    Name ="Text100"
                    ControlSource ="SubCatDesc"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =7500
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text101"
                    ControlSource ="=Sum([OnTheShelf]-IIf(IsNull([OnReserve]),0,[OnReserve]))"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =6300
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text102"
                    ControlSource ="=Sum(IIf(IsNull([OnReserve]),0,[OnReserve]))"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =5220
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text103"
                    ControlSource ="=Sum([OnTheShelf])"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            OnFormat ="[Event Procedure]"
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =1260
                    Top =60
                    Width =480
                    Name ="Text91"
                    ControlSource ="SalesDepartmentID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1800
                    Top =60
                    Width =3540
                    TabIndex =1
                    Name ="Text92"
                    ControlSource ="DepartmentDesc"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =7500
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text95"
                    ControlSource ="=Sum([OnTheShelf]-IIf(IsNull([OnReserve]),0,[OnReserve]))"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =6300
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =3
                    Name ="Text96"
                    ControlSource ="=Sum(IIf(IsNull([OnReserve]),0,[OnReserve]))"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =5220
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =4
                    Name ="Text97"
                    ControlSource ="=Sum([OnTheShelf])"
                    FontName ="Times New Roman"

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
                    Width =4905
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
                    Width =4740
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
                    Width =4905
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
                    Width =10620
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
                    Width =10545
                    Height =285
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
            OnFormat ="[Event Procedure]"
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =7500
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=Sum([OnTheShelf]-IIf(IsNull([OnReserve]),0,[OnReserve]))"
                    FontName ="Times New Roman"

                End
                Begin Label
                    Left =600
                    Top =60
                    Width =1050
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label49"
                    Caption ="Report total"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =6300
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =1
                    Name ="Text93"
                    ControlSource ="=Sum(IIf(IsNull([OnReserve]),0,[OnReserve]))"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =5220
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    TabIndex =2
                    Name ="Text94"
                    ControlSource ="=Sum([OnTheShelf])"
                    FontName ="Times New Roman"

                End
            End
        End
    End
End
CodeBehindForm
' See "05 Inventory Count Sheets By Category.cls"
