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
    Width =10455
    DatasheetFontHeight =10
    ItemSuffix =29
    Left =300
    Top =585
    RecSrcDt = Begin
        0xfd5dd98a9039e440
    End
    RecordSource ="qryOLCCScanCodeComparedToUPCItems"
    Caption ="qryFulPrice"
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
            Height =1500
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    Left =60
                    Top =60
                    Width =7170
                    Height =510
                    FontSize =20
                    Name ="Label14"
                    Caption ="Web UPC Codes  vs Local UPC Codes"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =7230
                    LayoutCachedHeight =570
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Top =720
                    Width =3795
                    Height =375
                    ColumnOrder =1
                    FontSize =14
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Month"
                    ControlSource ="=\"Downloaded from OLCC on \""
                    FontName ="Times New Roman"

                    LayoutCachedTop =720
                    LayoutCachedWidth =3795
                    LayoutCachedHeight =1095
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    BackStyle =1
                    IMESentenceMode =3
                    Left =3780
                    Top =720
                    Width =4035
                    Height =375
                    ColumnWidth =1965
                    ColumnOrder =0
                    FontSize =14
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Times New Roman"
                    ShowDatePicker =1

                    LayoutCachedLeft =3780
                    LayoutCachedTop =720
                    LayoutCachedWidth =7815
                    LayoutCachedHeight =1095
                End
                Begin Label
                    Left =600
                    Top =1200
                    Width =9300
                    Height =300
                    Name ="Label28"
                    Caption ="Items on this report should be scanned at the register to ensure the item is sca"
                        "nning correctly."
                    LayoutCachedLeft =600
                    LayoutCachedTop =1200
                    LayoutCachedWidth =9900
                    LayoutCachedHeight =1500
                End
            End
        End
        Begin PageHeader
            Height =480
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Top =60
                    Width =1380
                    Height =300
                    Name ="ID_Label"
                    Caption ="UPC Code"
                    Tag ="DetachedLabel"
                    LayoutCachedTop =60
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =1680
                    Top =60
                    Width =2655
                    Height =300
                    Name ="Description_Label"
                    Caption ="Local Item and Description"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =1680
                    LayoutCachedTop =60
                    LayoutCachedWidth =4335
                    LayoutCachedHeight =360
                End
                Begin Label
                    TextAlign =3
                    Left =9060
                    Top =60
                    Width =780
                    Height =300
                    Name ="Onhand_Label"
                    Caption ="Onhand"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =9060
                    LayoutCachedTop =60
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =360
                End
                Begin Line
                    BorderWidth =2
                    Left =60
                    Top =360
                    Width =10335
                    Name ="Line17"
                    LayoutCachedLeft =60
                    LayoutCachedTop =360
                    LayoutCachedWidth =10395
                    LayoutCachedHeight =360
                End
                Begin Label
                    Left =5220
                    Top =60
                    Width =3195
                    Height =300
                    Name ="Label26"
                    Caption ="OLCC Web Item and Description"
                    Tag ="DetachedLabel"
                    LayoutCachedLeft =5220
                    LayoutCachedTop =60
                    LayoutCachedWidth =8415
                    LayoutCachedHeight =360
                End
            End
        End
        Begin Section
            KeepTogether = NotDefault
            Height =315
            OnFormat ="[Event Procedure]"
            Name ="Detail"
            Begin
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Width =1020
                    Height =270
                    ColumnWidth =705
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =60
                    LayoutCachedWidth =1080
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1200
                    Width =1200
                    Height =270
                    TabIndex =2
                    Name ="Description"
                    ControlSource ="KMDSItem"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =1200
                    LayoutCachedWidth =2400
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9060
                    Width =780
                    Height =270
                    TabIndex =1
                    Name ="Onhand"
                    ControlSource ="Onhand"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =9060
                    LayoutCachedWidth =9840
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =2700
                    Width =2100
                    Height =270
                    TabIndex =3
                    Name ="Text23"
                    ControlSource ="KMDSDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =2700
                    LayoutCachedWidth =4800
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5040
                    Width =1260
                    Height =270
                    TabIndex =4
                    Name ="Text24"
                    ControlSource ="OLCCItem"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =5040
                    LayoutCachedWidth =6300
                    LayoutCachedHeight =270
                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6420
                    Width =2280
                    Height =270
                    TabIndex =5
                    Name ="Text25"
                    ControlSource ="OLCCDesc"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =6420
                    LayoutCachedWidth =8700
                    LayoutCachedHeight =270
                End
                Begin Line
                    Visible = NotDefault
                    Top =300
                    Width =10380
                    Name ="Line5"
                    LayoutCachedTop =300
                    LayoutCachedWidth =10380
                    LayoutCachedHeight =300
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
                    Left =60
                    Top =300
                    Width =4920
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                    LayoutCachedLeft =60
                    LayoutCachedTop =300
                    LayoutCachedWidth =4980
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Top =300
                    Width =3720
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
                    Width =3360
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
                    Width =3720
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
                    Left =120
                    Top =300
                    Width =9360
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text65"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                    LayoutCachedLeft =120
                    LayoutCachedTop =300
                    LayoutCachedWidth =9480
                    LayoutCachedHeight =540
                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Width =9360
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
' See "qryEOM301WebScanVariance.cls"
