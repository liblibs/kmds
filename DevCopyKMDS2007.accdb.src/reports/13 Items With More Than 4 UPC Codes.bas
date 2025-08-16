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
    ItemSuffix =13
    Left =585
    Top =1320
    RecSrcDt = Begin
        0xd5175781dc27e340
    End
    RecordSource ="qryItemsWithMoreThan4UPCCodes"
    Caption ="Items With More Than 4 UPC Codes"
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
            ControlSource ="SalesDepartmentID"
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            ControlSource ="SubCategory.Description"
        End
        Begin BreakLevel
            SortOrder = NotDefault
            ControlSource ="Count"
        End
        Begin BreakLevel
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =930
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =6180
                    Height =510
                    FontSize =20
                    Name ="Label6"
                    Caption ="Items With More Than 4 UPC Codes"
                End
            End
        End
        Begin PageHeader
            Height =420
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =60
                    Top =60
                    Width =1170
                    Height =300
                    Name ="ID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =1290
                    Top =60
                    Width =2310
                    Height =300
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    OverlapFlags =12
                    TextAlign =2
                    Left =4110
                    Top =60
                    Width =1770
                    Height =300
                    Name ="Count_Label"
                    Caption ="Count"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =4110
                    LayoutCachedTop =60
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =360
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =360
                    Width =5910
                    Name ="Line9"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =360
            Name ="GroupHeader0"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =2220
                    Top =60
                    Name ="SalesDepartmentID"
                    ControlSource ="SDesc"

                    Begin
                        Begin Label
                            TextAlign =0
                            Top =60
                            Width =1950
                            Height =300
                            Name ="Label11"
                            Caption ="SalesDepartmentID:"
                        End
                    End
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =420
            BreakLevel =1
            Name ="GroupHeader1"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =180
                    Top =120
                    Name ="SubCategory.Description"
                    ControlSource ="SubCategory.Description"
                    EventProcPrefix ="SubCategory_Description"

                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =255
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =1
                    IMESentenceMode =3
                    Left =60
                    Height =225
                    Name ="ID"
                    ControlSource ="ID"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1740
                    Width =2310
                    Height =225
                    ColumnWidth =3150
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Item.Description"

                    LayoutCachedLeft =1740
                    LayoutCachedWidth =4050
                    LayoutCachedHeight =225
                End
                Begin TextBox
                    OverlapFlags =4
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4110
                    Width =1770
                    Height =225
                    TabIndex =2
                    Name ="Count"
                    ControlSource ="Count"

                    LayoutCachedLeft =4110
                    LayoutCachedWidth =5880
                    LayoutCachedHeight =225
                End
                Begin Line
                    Visible = NotDefault
                    Top =240
                    Width =6060
                    Name ="Line5"
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
                    Width =4260
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
' See "13 Items With More Than 4 UPC Codes.cls"
