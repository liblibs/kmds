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
    ItemSuffix =112
    Left =1170
    Top =2040
    OrderBy ="Description"
    RecSrcDt = Begin
        0x0836102f2db8e240
    End
    RecordSource ="qryVendorItemSupplyList"
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
            ControlSource ="VendorID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =975
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
                    Left =30
                    Top =60
                    Width =10665
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Vendor Supply List"
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
                    Top =960
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =1800
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =5355
                    Top =1380
                    Width =1365
                    Height =300
                    Name ="TotalP_Label"
                    Caption ="Avail"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =1380
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =6960
                    Top =1380
                    Width =600
                    Height =300
                    Name ="Label79"
                    Caption ="Min"
                End
                Begin Label
                    TextAlign =3
                    Left =8640
                    Top =1380
                    Width =600
                    Height =300
                    Name ="Label80"
                    Caption ="SOQ"
                End
                Begin Label
                    TextAlign =2
                    Left =60
                    Top =1380
                    Width =780
                    Height =300
                    Name ="Label85"
                    Caption ="Item"
                End
                Begin Label
                    OverlapFlags =4
                    TextAlign =2
                    Left =1140
                    Top =1380
                    Width =1140
                    Height =300
                    Name ="Label87"
                    Caption ="Stock #"
                    LayoutCachedLeft =1140
                    LayoutCachedTop =1380
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =1680
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =120
                    Width =720
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="VendorID"
                    ControlSource ="VendorID"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =900
                    Width =3300
                    Height =300
                    FontSize =11
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Name"
                    ControlSource ="Name"
                    FontName ="Times New Roman"

                End
                Begin Label
                    TextAlign =3
                    Left =7905
                    Top =1380
                    Width =495
                    Height =300
                    Name ="Label94"
                    Caption ="Max"
                End
                Begin Label
                    TextAlign =3
                    Left =9420
                    Top =1380
                    Width =555
                    Height =300
                    Name ="Label96"
                    Caption ="Cost:"
                End
                Begin Label
                    TextAlign =3
                    Left =10110
                    Top =1380
                    Width =570
                    Height =300
                    Name ="Label97"
                    Caption ="Price"
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =300
                    Width =3300
                    TabIndex =2
                    Name ="Contact"
                    ControlSource ="Contact"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =540
                    Width =3300
                    TabIndex =3
                    Name ="Address1"
                    ControlSource ="Address1"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =780
                    Width =3300
                    TabIndex =4
                    Name ="Address2"
                    ControlSource ="Address2"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =1020
                    TabIndex =5
                    Name ="City"
                    ControlSource ="City"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2340
                    Top =1020
                    Width =420
                    TabIndex =6
                    Name ="St"
                    ControlSource ="St"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2760
                    Top =1020
                    TabIndex =7
                    Name ="Zip"
                    ControlSource ="Zip"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4860
                    Top =315
                    Height =225
                    TabIndex =8
                    Name ="Phone1"
                    ControlSource ="Phone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =4320
                            Top =315
                            Width =510
                            Height =225
                            FontSize =8
                            Name ="Label105"
                            Caption ="Phone"
                        End
                    End
                End
                Begin TextBox
                    EnterKeyBehavior = NotDefault
                    ScrollBars =2
                    IMESentenceMode =3
                    Left =6960
                    Width =3360
                    Height =1263
                    TabIndex =9
                    Name ="VendorNotes"
                    ControlSource ="VendorNotes"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4875
                    Top =555
                    Height =225
                    TabIndex =10
                    Name ="Phone2"
                    ControlSource ="Phone2"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =4500
                            Top =555
                            Width =345
                            Height =225
                            FontSize =8
                            Name ="Label107"
                            Caption ="Fax"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4845
                    Top =795
                    Width =840
                    Height =225
                    TabIndex =11
                    Name ="PoAtCost"
                    ControlSource ="PoAtCost"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =4440
                            Top =795
                            Width =375
                            Height =225
                            FontSize =8
                            Name ="Label108"
                            Caption ="Cost"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6180
                    Top =60
                    Width =0
                    TabIndex =12
                    Name ="PrintUpcOnPo"
                    ControlSource ="PrintUpcOnPo"

                    Begin
                        Begin Label
                            TextAlign =3
                            TextFontFamily =34
                            Left =4860
                            Top =75
                            Width =1230
                            Height =225
                            FontSize =8
                            FontWeight =400
                            ForeColor =0
                            Name ="Label109"
                            Caption ="Supply All Items"
                            FontName ="Arial"
                        End
                    End
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4875
                    Top =1035
                    Width =1905
                    Height =225
                    TabIndex =13
                    Name ="POReportForm"
                    ControlSource ="POReportForm"

                    Begin
                        Begin Label
                            TextAlign =3
                            Left =4320
                            Top =1035
                            Width =525
                            Height =225
                            FontSize =8
                            Name ="Label110"
                            Caption ="Report"
                        End
                    End
                End
                Begin Line
                    BorderWidth =2
                    Top =1740
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            ForceNewPage =1
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    Top =60
                    Width =1140
                    Height =225
                    ColumnWidth =570
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedTop =60
                    LayoutCachedWidth =1140
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    Left =2340
                    Top =60
                    Width =2880
                    Height =225
                    ColumnWidth =3120
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5280
                    Top =60
                    Width =600
                    TabIndex =2
                    Name ="SizeD"
                    ControlSource ="SizeD"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =1020
                    Top =60
                    Width =1260
                    TabIndex =3
                    Name ="StockNumber"
                    ControlSource ="StockNumber"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1020
                    LayoutCachedTop =60
                    LayoutCachedWidth =2280
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5940
                    Top =60
                    Width =780
                    TabIndex =4
                    Name ="Avail"
                    ControlSource ="Avail"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =6780
                    Top =60
                    Width =780
                    TabIndex =5
                    Name ="OrderAt"
                    ControlSource ="OrderAt"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =7620
                    Top =60
                    Width =780
                    TabIndex =6
                    Name ="OrderTo"
                    ControlSource ="OrderTo"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8460
                    Top =60
                    Width =780
                    TabIndex =7
                    Name ="SOQ"
                    ControlSource ="SOQ"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =9375
                    Top =60
                    Width =600
                    TabIndex =8
                    Name ="Cost"
                    ControlSource ="Cost"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =10080
                    Top =60
                    Width =600
                    TabIndex =9
                    Name ="Text98"
                    ControlSource ="TodaysPrice"
                    FontName ="Arial Narrow"

                End
                Begin Line
                    Visible = NotDefault
                    Top =300
                    Width =10680
                    Name ="Line5"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =360
            Name ="GroupFooter1"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =8460
                    Top =60
                    Width =780
                    Name ="Text95"
                    ControlSource ="=Sum([SOQ])"

                End
            End
        End
        Begin PageFooter
            Height =555
            Name ="PageFooterSection"
            Begin
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
                    Left =5760
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
' See "29 Vendor Supply List.cls"
