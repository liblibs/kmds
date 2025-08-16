Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =124
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9300
    DatasheetFontHeight =10
    ItemSuffix =23
    Left =315
    Top =1200
    RecSrcDt = Begin
        0x8bf030d28b1be340
    End
    RecordSource ="qryFulPrice"
    Caption ="qryFulPrice"
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
            Height =1320
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =4695
                    Height =510
                    FontSize =20
                    Name ="Label14"
                    Caption ="Web Prices vs Local Prices"
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Left =180
                    Top =720
                    Width =2640
                    Height =375
                    FontSize =14
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Month"
                    ControlSource ="=\"Web Prices as of \""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Left =2940
                    Top =720
                    Width =480
                    Height =375
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text20"
                    ControlSource ="Month"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Left =3540
                    Top =720
                    Height =375
                    FontSize =14
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Year"
                    ControlSource ="Year"
                    FontName ="Times New Roman"

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
                    Width =780
                    Height =300
                    Name ="ID_Label"
                    Caption ="ID"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =900
                    Top =60
                    Width =1500
                    Height =300
                    Name ="Description_Label"
                    Caption ="Description"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =2460
                    Top =60
                    Width =780
                    Height =300
                    Name ="Onhand_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =3300
                    Top =60
                    Width =1500
                    Height =300
                    Name ="Price_Label"
                    Caption ="Price"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =3
                    Left =4860
                    Top =60
                    Width =1500
                    Height =300
                    Name ="WebPrice_Label"
                    Caption ="WebPrice"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =360
                    Width =9240
                    Name ="Line17"
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =390
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Top =60
                    Width =780
                    Height =270
                    ColumnWidth =705
                    Name ="ID"
                    ControlSource ="ID"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =900
                    Top =60
                    Width =1500
                    Height =270
                    ColumnWidth =2130
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2460
                    Top =60
                    Width =780
                    Height =270
                    TabIndex =2
                    Name ="Onhand"
                    ControlSource ="Onhand"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =3300
                    Top =60
                    Width =1500
                    Height =270
                    ColumnWidth =870
                    TabIndex =3
                    Name ="Price"
                    ControlSource ="Price"

                End
                Begin TextBox
                    TextAlign =3
                    IMESentenceMode =3
                    Left =4860
                    Top =60
                    Width =1500
                    Height =270
                    TabIndex =4
                    Name ="WebPrice"
                    ControlSource ="WebPrice"

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
                    Width =360
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
                    Width =3480
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
                    Width =3120
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
                    Width =3480
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
                    Width =9120
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
                    Width =9120
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
' See "39 Web Price Variance.cls"
