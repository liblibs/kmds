Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =10380
    DatasheetFontHeight =10
    ItemSuffix =37
    Left =1680
    Top =345
    RecSrcDt = Begin
        0x5eeb163ad36de340
    End
    RecordSource ="qryCategory"
    Caption ="Club"
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
            ControlSource ="SalesDepartmentID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="CatDescription"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =570
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =2940
                    Height =510
                    FontSize =20
                    Name ="Label24"
                    Caption ="Category Listing"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =3000
                    LayoutCachedHeight =570
                End
            End
        End
        Begin PageHeader
            Height =375
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =360
                    Width =10260
                    Name ="Line27"
                    LayoutCachedTop =360
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =2
                    Left =5760
                    Width =1500
                    Height =360
                    Name ="Label39"
                    Caption ="SKU Count"
                    LayoutCachedLeft =5760
                    LayoutCachedWidth =7260
                    LayoutCachedHeight =360
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =360
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =300
                    Top =60
                    Width =720
                    Height =225
                    ColumnWidth =2205
                    Name ="SalesDepartmentID"
                    ControlSource ="SalesDepartmentID"

                    LayoutCachedLeft =300
                    LayoutCachedTop =60
                    LayoutCachedWidth =1020
                    LayoutCachedHeight =285
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1200
                    Top =60
                    Width =3240
                    Height =225
                    ColumnWidth =2820
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =1200
                    LayoutCachedTop =60
                    LayoutCachedWidth =4440
                    LayoutCachedHeight =285
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            CanGrow = NotDefault
            CanShrink = NotDefault
            Height =225
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =1140
                    Height =225
                    ColumnWidth =2370
                    Name ="CatDescription"
                    ControlSource ="CatDescription"

                    LayoutCachedLeft =1140
                    LayoutCachedWidth =2580
                    LayoutCachedHeight =225
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
                    Width =10200
                    Name ="Line5"
                    LayoutCachedTop =240
                    LayoutCachedWidth =10200
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2220
                    Height =225
                    Name ="SubCatDesc"
                    ControlSource ="SubCatDesc"

                    LayoutCachedLeft =2220
                    LayoutCachedWidth =3660
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4200
                    Height =225
                    TabIndex =1
                    Name ="TargetGP"
                    ControlSource ="TargetGP"

                    LayoutCachedLeft =4200
                    LayoutCachedWidth =5640
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    DecimalPlaces =0
                    IMESentenceMode =3
                    Left =6480
                    Width =720
                    Height =225
                    TabIndex =2
                    Name ="Text40"
                    ControlSource ="Items"
                    Format ="Standard"

                    LayoutCachedLeft =6480
                    LayoutCachedWidth =7200
                    LayoutCachedHeight =225
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
                    Top =360
                    Width =4560
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =360
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =4440
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
                    Top =60
                    Width =4560
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text64"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =5700
                    LayoutCachedTop =60
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =360
                    Width =10215
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                    LayoutCachedTop =360
                    LayoutCachedWidth =10215
                    LayoutCachedHeight =600
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =180
                    Top =60
                    Width =10140
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =180
                    LayoutCachedTop =60
                    LayoutCachedWidth =10320
                    LayoutCachedHeight =300
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
                    TextAlign =1
                    TextFontFamily =18
                    Left =180
                    Top =120
                    Width =1860
                    Height =300
                    FontSize =11
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text73"
                    ControlSource ="=\"Count =\" & Count([SubCatDesc])"
                    FontName ="Times New Roman"

                End
                Begin Line
                    Top =60
                    Width =10380
                    Name ="Line30"
                    LayoutCachedTop =60
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =60
                End
                Begin Line
                    Top =120
                    Width =10380
                    Name ="Line31"
                    LayoutCachedTop =120
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =120
                End
            End
        End
    End
End
CodeBehindForm
' See "61 Category List.cls"
