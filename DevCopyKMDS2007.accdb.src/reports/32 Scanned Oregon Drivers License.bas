Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =49
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9360
    DatasheetFontHeight =10
    ItemSuffix =17
    Left =525
    Top =1230
    RecSrcDt = Begin
        0x506e16718b70e340
    End
    RecordSource ="qryScannedODL"
    Caption ="Scanned Oregon Drivers License"
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
            ControlSource ="SetupDate"
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
                    Width =5550
                    Height =510
                    FontSize =20
                    Name ="Label12"
                    Caption ="Scanned Oregon Drivers License"
                End
            End
        End
        Begin PageHeader
            Height =420
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    TextAlign =3
                    Left =180
                    Top =60
                    Width =1920
                    Height =300
                    Name ="SetupDate_Label"
                    Caption ="SetupDate"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =2160
                    Top =60
                    Width =1170
                    Height =300
                    Name ="CashSysID_Label"
                    Caption ="CashSysID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3390
                    Top =60
                    Width =1230
                    Height =300
                    Name ="InvoHeadID_Label"
                    Caption ="InvoHeadID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4680
                    Top =60
                    Width =1170
                    Height =300
                    Name ="CashierID_Label"
                    Caption ="CashierID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =5880
                    Top =60
                    Width =1455
                    Height =300
                    Name ="SalesDept_Label"
                    Caption ="ODL"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7380
                    Top =60
                    Width =1185
                    Height =300
                    Name ="InputCode_Label"
                    Caption ="DOB "
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =360
                    Width =7260
                    Name ="Line15"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =240
            Name ="Detail"
            Begin
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =180
                    Width =1920
                    ColumnWidth =2205
                    Name ="SetupDate"
                    ControlSource ="SetupDate"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2160
                    Width =1170
                    ColumnWidth =1155
                    TabIndex =1
                    Name ="CashSysID"
                    ControlSource ="CashSysID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3390
                    Width =1230
                    TabIndex =2
                    Name ="InvoHeadID"
                    ControlSource ="InvoHeadID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =4680
                    Width =1170
                    ColumnWidth =1020
                    TabIndex =3
                    Name ="CashierID"
                    ControlSource ="CashierID"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =5880
                    Width =1455
                    TabIndex =4
                    Name ="SalesDept"
                    ControlSource ="SalesDept"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7440
                    Width =1185
                    ColumnWidth =1770
                    TabIndex =5
                    Name ="InputCode"
                    ControlSource ="InputCode"

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
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text13"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    IMESentenceMode =3
                    Left =4740
                    Top =240
                    Width =4560
                    Height =300
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text14"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin Line
                    BorderWidth =3
                    Left =60
                    Top =240
                    Width =9240
                    BorderColor =12632256
                    Name ="Line16"
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
