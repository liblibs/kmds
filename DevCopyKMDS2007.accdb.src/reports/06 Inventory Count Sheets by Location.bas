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
    Width =10770
    DatasheetFontHeight =10
    ItemSuffix =90
    Left =5595
    Top =540
    OrderBy ="Location,Description"
    RecSrcDt = Begin
        0x8f27be91d066e440
    End
    RecordSource ="qryInventoryByLocation"
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
        Begin FormHeader
            KeepTogether = NotDefault
            Height =1800
            Name ="ReportHeader"
            Begin
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =600
                    Width =10621
                    Height =300
                    ColumnOrder =0
                    Name ="Text50"
                    ControlSource ="=Now()"
                    Format ="Long Date"

                End
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Left =60
                    Top =60
                    Width =10620
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Inventory Count Sheets By Location"
                End
                Begin Line
                    BorderWidth =3
                    Top =30
                    Width =10755
                    BorderColor =12632256
                    Name ="Line51"
                End
                Begin Line
                    BorderWidth =3
                    Top =1680
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
                Begin TextBox
                    TextAlign =2
                    Left =60
                    Top =960
                    Width =10621
                    Height =300
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Text84"
                    ControlSource ="=\"Between \" & Forms!PrintInventoryReport!cboLocationBegin & \" and \" & Forms!"
                        "PrintInventoryReport!cboLocationEnd"
                    Format ="Long Date"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =60
                    Top =1320
                    Width =10620
                    ColumnOrder =2
                    FontWeight =700
                    TabIndex =2
                    Name ="VendorName"
                    ControlSource ="VendorName"

                End
            End
        End
        Begin PageHeader
            Height =540
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =5280
                    Top =120
                    Width =1215
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="On the shelf"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5280
                    LayoutCachedTop =120
                    LayoutCachedWidth =6495
                    LayoutCachedHeight =420
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =60
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =450
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =0
                    Left =60
                    Top =120
                    Width =960
                    Height =300
                    Name ="Label77"
                    Caption ="Location:"
                End
                Begin Label
                    TextAlign =3
                    Left =6660
                    Top =120
                    Width =930
                    Height =300
                    Name ="Label79"
                    Caption ="Reserved"
                    LayoutCachedLeft =6660
                    LayoutCachedTop =120
                    LayoutCachedWidth =7590
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =7695
                    Top =120
                    Width =945
                    Height =300
                    Name ="Label80"
                    Caption ="Total"
                    LayoutCachedLeft =7695
                    LayoutCachedTop =120
                    LayoutCachedWidth =8640
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =3
                    Left =8760
                    Top =120
                    Width =1800
                    Height =300
                    Name ="Label82"
                    Caption ="Physical count"
                    LayoutCachedLeft =8760
                    LayoutCachedTop =120
                    LayoutCachedWidth =10560
                    LayoutCachedHeight =420
                End
                Begin Label
                    TextAlign =2
                    Left =1080
                    Top =120
                    Width =960
                    Height =300
                    Name ="Label85"
                    Caption ="Item"
                    LayoutCachedLeft =1080
                    LayoutCachedTop =120
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =420
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
                    Left =960
                    Top =120
                    Width =1320
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =960
                    LayoutCachedTop =120
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    Left =2340
                    Top =120
                    Width =2460
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2340
                    LayoutCachedTop =120
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    TextAlign =3
                    Left =5745
                    Top =120
                    Width =720
                    Height =225
                    ColumnWidth =900
                    TabIndex =2
                    Name ="OnTheShelf"
                    ControlSource ="OnTheShelf"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5745
                    LayoutCachedTop =120
                    LayoutCachedWidth =6465
                    LayoutCachedHeight =345
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Top =120
                    Width =840
                    TabIndex =3
                    Name ="Location"
                    ControlSource ="Location"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedTop =120
                    LayoutCachedWidth =900
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4860
                    Top =120
                    Width =840
                    TabIndex =4
                    Name ="SizeD"
                    ControlSource ="SizeD"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =4860
                    LayoutCachedTop =120
                    LayoutCachedWidth =5700
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6630
                    Top =120
                    Width =930
                    TabIndex =5
                    Name ="OnReserve"
                    ControlSource ="OnReserve"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6630
                    LayoutCachedTop =120
                    LayoutCachedWidth =7560
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7680
                    Top =120
                    Width =930
                    TabIndex =6
                    Name ="OnHand"
                    ControlSource ="=[OnTheShelf]-IIf(IsNull([OnReserve]),0,[OnReserve])"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =7680
                    LayoutCachedTop =120
                    LayoutCachedWidth =8610
                    LayoutCachedHeight =360
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
                    Width =10620
                    Name ="line5"
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
            Height =405
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    OldBorderStyle =1
                    TextAlign =2
                    TextFontFamily =18
                    Left =5400
                    Top =60
                    Width =900
                    Height =285
                    FontSize =10
                    Name ="Text41"
                    ControlSource ="=Sum([Onhand])"
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
            End
        End
    End
End
CodeBehindForm
' See "06 Inventory Count Sheets by Location.cls"
