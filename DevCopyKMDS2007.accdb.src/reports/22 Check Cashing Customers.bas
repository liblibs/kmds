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
    Width =10860
    DatasheetFontHeight =10
    ItemSuffix =120
    OrderBy ="LastName"
    RecSrcDt = Begin
        0x9ca1691ff4a8e240
    End
    RecordSource ="qryCheckCashingCustomers"
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
            Height =735
            Name ="ReportHeader"
            Begin
                Begin Label
                    BackStyle =1
                    TextAlign =2
                    Top =60
                    Width =10725
                    Height =510
                    FontSize =20
                    Name ="Label47"
                    Caption ="Check Cashing Customers List"
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
                    Top =720
                    Width =10755
                    BorderColor =12632256
                    Name ="Line53"
                End
            End
        End
        Begin PageHeader
            Height =555
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Line
                    BorderWidth =2
                    Top =180
                    Width =10710
                    Name ="Line55"
                    Tag ="DetachedLabel"
                End
                Begin Line
                    BorderWidth =2
                    Top =540
                    Width =10710
                    Name ="Line56"
                    Tag ="DetachedLabel"
                End
                Begin Label
                    Left =60
                    Top =180
                    Width =645
                    Height =300
                    Name ="Label116"
                    Caption ="Name"
                End
                Begin Label
                    Left =3180
                    Top =180
                    Width =840
                    Height =300
                    Name ="Label117"
                    Caption ="Address"
                End
                Begin Label
                    Left =6540
                    Top =180
                    Width =1125
                    Height =300
                    Name ="Label118"
                    Caption ="City, St Zip"
                End
                Begin Label
                    Left =9420
                    Top =180
                    Width =1125
                    Height =300
                    Name ="Label119"
                    Caption ="Phone"
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
                    IMESentenceMode =3
                    Left =3240
                    Width =3300
                    Name ="Address1"
                    ControlSource ="Address"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =6600
                    TabIndex =1
                    Name ="City"
                    ControlSource ="City"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8100
                    Width =420
                    TabIndex =2
                    Name ="St"
                    ControlSource ="State"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =8640
                    Width =780
                    TabIndex =3
                    Name ="Zip"
                    ControlSource ="Zip"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =9480
                    Width =1260
                    TabIndex =4
                    Name ="TelephoneNumbr"
                    ControlSource ="TelephoneNumbr"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =1800
                    Width =1380
                    TabIndex =5
                    Name ="FirstName"
                    ControlSource ="FirstName"

                End
                Begin TextBox
                    IMESentenceMode =3
                    Left =60
                    Width =1680
                    TabIndex =6
                    Name ="Text114"
                    ControlSource ="LastName"

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
                    Width =5340
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
                    Left =180
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
' See "22 Check Cashing Customers.cls"
