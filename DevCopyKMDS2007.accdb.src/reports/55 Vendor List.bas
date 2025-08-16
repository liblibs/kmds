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
    Width =14340
    DatasheetFontHeight =10
    ItemSuffix =32
    Left =-75
    Top =600
    RecSrcDt = Begin
        0x11a5e447d851e440
    End
    RecordSource ="qryVendorList"
    Caption ="Vendors"
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
            ControlSource ="Name"
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
                    Width =2925
                    Height =510
                    FontSize =20
                    Name ="Label24"
                    Caption ="Vendors on file"
                End
            End
        End
        Begin PageHeader
            Height =420
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    Left =120
                    Top =60
                    Width =840
                    Height =300
                    Name ="Name_Label"
                    Caption ="Name"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =3060
                    Top =60
                    Width =840
                    Height =300
                    Name ="Contact_Label"
                    Caption ="Contact"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =7440
                    Top =60
                    Width =840
                    Height =300
                    Name ="Address1_Label"
                    Caption ="Address1"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =9960
                    Top =60
                    Width =840
                    Height =300
                    Name ="City_Label"
                    Caption ="City"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =5040
                    Top =60
                    Width =1140
                    Height =300
                    Name ="Phone1_Label"
                    Caption ="Phone"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    TextAlign =2
                    Left =6240
                    Top =60
                    Width =1140
                    Height =300
                    Name ="Phone2_Label"
                    Caption ="Fax"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =360
                    Width =14220
                    Name ="Line27"
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
                    Left =120
                    Top =60
                    Width =2940
                    FontWeight =700
                    Name ="Name"
                    ControlSource ="Name"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =3120
                    Top =60
                    Width =1860
                    TabIndex =1
                    Name ="Contact"
                    ControlSource ="Contact"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =7440
                    Top =60
                    Width =2460
                    TabIndex =2
                    Name ="Address1"
                    ControlSource ="Address1"
                    FontName ="Arial Narrow"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9960
                    Top =60
                    Width =1380
                    TabIndex =3
                    Name ="City"
                    ControlSource ="City"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =5040
                    Top =60
                    Width =1140
                    TabIndex =4
                    Name ="Phone1"
                    ControlSource ="Phone1"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6240
                    Top =60
                    Width =1140
                    TabIndex =5
                    Name ="Phone2"
                    ControlSource ="Phone2"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =11400
                    Top =60
                    Width =660
                    TabIndex =6
                    Name ="Licensee"
                    ControlSource ="St"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =12120
                    Top =60
                    Width =2100
                    TabIndex =7
                    Name ="EMail"
                    ControlSource ="Zip"

                End
                Begin Line
                    Visible = NotDefault
                    Top =300
                    Width =14340
                    Name ="Line5"
                End
            End
        End
        Begin PageFooter
            Height =615
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
                    Width =8460
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
                    Width =8460
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
                    Top =360
                    Width =14175
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
                    Left =180
                    Top =60
                    Width =14100
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
                    ControlSource ="=\"Count =\" & Count([Name])"
                    FontName ="Times New Roman"

                End
                Begin Line
                    Top =60
                    Width =14340
                    Name ="Line30"
                End
                Begin Line
                    Top =120
                    Width =14340
                    Name ="Line31"
                End
            End
        End
    End
End
CodeBehindForm
' See "55 Vendor List.cls"
