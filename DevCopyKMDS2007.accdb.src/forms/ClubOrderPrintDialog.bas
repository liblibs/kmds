Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    MinButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =119
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8880
    DatasheetFontHeight =10
    ItemSuffix =69
    Left =2970
    Top =1920
    Right =12180
    Bottom =7980
    Filter ="[ID]=100"
    RecSrcDt = Begin
        0xd0b43a3d6cbbe240
    End
    RecordSource ="Club"
    Caption ="Print Club Documents KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Sof"
        "tware ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            BackColor =-2147483633
            ForeColor =-2147483630
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
        End
        Begin Line
            BorderLineStyle =0
            SpecialEffect =3
        End
        Begin Image
            BackStyle =0
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin OptionButton
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin OptionGroup
            SpecialEffect =3
            BorderLineStyle =0
        End
        Begin BoundObjectFrame
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
        End
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
            AsianLineBreak =255
        End
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            BackColor =-2147483643
            ForeColor =-2147483640
        End
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin UnboundObjectFrame
            SpecialEffect =2
            OldBorderStyle =1
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            FontName ="MS Sans Serif"
            BorderLineStyle =0
        End
        Begin Tab
            BackStyle =0
            BorderLineStyle =0
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =6060
            BackColor =-2147483633
            Name ="Detail"
            OnClick ="[Event Procedure]"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =420
                    Top =900
                    Width =600
                    Height =300
                    ColumnWidth =900
                    BackColor =10092543
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =1080
                    Top =900
                    Width =2640
                    Height =300
                    ColumnWidth =2310
                    TabIndex =1
                    BackColor =10092543
                    Name ="txtClubName"
                    ControlSource ="Name"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =1980
                    Top =1920
                    Width =465
                    Height =300
                    ColumnWidth =465
                    TabIndex =7
                    BackColor =10092543
                    Name ="OrderGuides"
                    ControlSource ="OrderGuides"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =1965
                    Top =2700
                    Width =465
                    Height =300
                    ColumnWidth =465
                    TabIndex =12
                    BackColor =10092543
                    Name ="Invoices"
                    ControlSource ="Invoices"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =7800
                            Top =4140
                            Width =735
                            Height =435
                            ForeColor =12349952
                            Name ="lblEmailCsv"
                            Caption ="eMail CSV Filie"
                            LayoutCachedLeft =7800
                            LayoutCachedTop =4140
                            LayoutCachedWidth =8535
                            LayoutCachedHeight =4575
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =1995
                    Top =4140
                    Width =465
                    Height =300
                    ColumnWidth =465
                    TabIndex =27
                    BackColor =10092543
                    Name ="Picklists"
                    ControlSource ="Picklists"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1245
                            Top =4140
                            Width =690
                            Height =285
                            Name ="Picklists_Label"
                            Caption ="Picklists"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =93
                    Left =1995
                    Top =4620
                    Width =465
                    Height =300
                    ColumnWidth =465
                    TabIndex =30
                    BackColor =10092543
                    Name ="PosPicks"
                    ControlSource ="PosPicks"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =1140
                            Top =4620
                            Width =795
                            Height =285
                            Name ="PosPicks_Label"
                            Caption ="POS Picks"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    Left =2550
                    Top =1920
                    Width =600
                    Height =300
                    TabIndex =8
                    Name ="txtOrderGuides"
                    ValidationRule ="Between 0 And 5"
                    ValidationText ="This value must be between 0 and 5"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    Left =2535
                    Top =2700
                    Width =600
                    Height =300
                    TabIndex =13
                    Name ="txtInvoices"
                    ValidationRule ="Between 0 And 5"
                    ValidationText ="This value must be between 0 and 5"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    Left =2565
                    Top =4140
                    Width =600
                    Height =300
                    TabIndex =28
                    Name ="txtPickLists"
                    ValidationRule ="Between 0 And 5"
                    ValidationText ="This value must be between 0 and 5"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    OverlapFlags =93
                    TextAlign =2
                    Left =2565
                    Top =4620
                    Width =600
                    Height =300
                    TabIndex =31
                    Name ="txtPOSPickLists"
                    ValidationRule ="Between 0 And 5"
                    ValidationText ="This value must be between 0 and 5"
                    FontName ="Trebuchet MS"

                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =120
                    Width =8820
                    Height =5940
                    Name ="Box18"
                End
                Begin Label
                    OverlapFlags =247
                    Left =180
                    Top =240
                    Width =5070
                    Height =285
                    Name ="Label19"
                    Caption ="Enter the number of copies of each document and then press print"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =3780
                    Top =900
                    Width =1380
                    Height =300
                    TabIndex =2
                    BackColor =10092543
                    Name ="txtInvoiceNo"
                    FontName ="Trebuchet MS"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =3285
                    Top =4140
                    Width =2070
                    TabIndex =29
                    Name ="PickListSort"
                    ControlSource ="PickListSort"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "
                    ValidationRule ="Is Not Null"
                    ValidationText ="Please select one of the sort methods listed in the drop down list."
                    DefaultValue ="\"LineID\""

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =3285
                    Top =2700
                    Width =2070
                    TabIndex =14
                    Name ="InvoiceListSort"
                    ControlSource ="InvoiceListSort"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "
                    ValidationRule ="Is Not Null"
                    ValidationText ="Please select one of the sort methods listed in the drop down list."
                    DefaultValue ="\"LineID\""

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =6045
                    Top =2640
                    Width =486
                    Height =426
                    TabIndex =11
                    Name ="cmdPrintPreviewInvoice"
                    Caption ="Command25"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000ad000ffffffffff0d000 ,
                        0x0fffffff0000700d0ffffff0788707da0fffff0788e770ad0fffff08888780da ,
                        0x0fffff08e88780ad0fffff07ee8770da0ffffff078870dad0fffffff0000dada ,
                        0x0ffffffffff0adad0fffffff0000dada0fffffff080dadad0fffffff00dadada ,
                        0x000000000dadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"

                    Overlaps =1
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =6045
                    Top =4080
                    Width =486
                    Height =426
                    TabIndex =26
                    Name ="cmdPrintPreviewPickList"
                    Caption ="Command25"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000ad000ffffffffff0d000 ,
                        0x0fffffff0000700d0ffffff0788707da0fffff0788e770ad0fffff08888780da ,
                        0x0fffff08e88780ad0fffff07ee8770da0ffffff078870dad0fffffff0000dada ,
                        0x0ffffffffff0adad0fffffff0000dada0fffffff080dadad0fffffff00dadada ,
                        0x000000000dadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"

                    Overlaps =1
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    IMESentenceMode =3
                    Left =3240
                    Top =1920
                    Width =2070
                    TabIndex =9
                    Name ="OrderGuideSort"
                    ControlSource ="OrderGuideSort"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "
                    ValidationRule ="Is Not Null"
                    ValidationText ="Please select one of the sort methods listed in the drop down list."
                    DefaultValue ="\"LineID\""

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =2
                            OverlapFlags =247
                            TextAlign =2
                            Left =3240
                            Top =1260
                            Width =1410
                            Height =540
                            Name ="Label38"
                            Caption ="Sort By"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =6060
                    Top =1800
                    Width =486
                    Height =426
                    TabIndex =5
                    Name ="cmdPrintPreviewOrderGuide"
                    Caption ="Command25"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000ad000ffffffffff0d000 ,
                        0x0fffffff0000700d0ffffff0788707da0fffff0788e770ad0fffff08888780da ,
                        0x0fffff08e88780ad0fffff07ee8770da0ffffff078870dad0fffffff0000dada ,
                        0x0ffffffffff0adad0fffffff0000dada0fffffff080dadad0fffffff00dadada ,
                        0x000000000dadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =6840
                            Top =4140
                            Width =735
                            Height =435
                            Name ="lblEmailPdf"
                            Caption ="eMail PDF Filie"
                            LayoutCachedLeft =6840
                            LayoutCachedTop =4140
                            LayoutCachedWidth =7575
                            LayoutCachedHeight =4575
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =1860
                    Width =8760
                    Height =3480
                    Name ="Box30"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =247
                    TextAlign =2
                    Left =540
                    Top =1260
                    Width =1515
                    Height =540
                    Name ="Label31"
                    Caption ="Document"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =247
                    TextAlign =2
                    Left =2460
                    Top =1260
                    Width =630
                    Height =540
                    Name ="Label32"
                    Caption ="Copies"
                    FontName ="Trebuchet MS"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    AccessKey =67
                    Left =3840
                    Top =5400
                    Width =666
                    Height =606
                    TabIndex =33
                    Name ="cmdClose"
                    Caption ="Command64"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x07070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f90707f9f9f9f9fffffffff9f9f9fffff9f9f9fffffffff9f9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9f9fffffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9f9f9f9f90707f9f9f9f9f9fffffff9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfffffffff9f9f90707f9f9f9f9fffffff9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffffff9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9f9f9f9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9fffff9f9fffff9f9fffff9f9fffff9f9fffff9ff ,
                        0xfff9f9fffff9f90707f9f9f9f9fffffffff9fffffffffffff9fffffffff9f9ff ,
                        0xfffffffff9f9f90707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f9f9070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f9f90707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9f907070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f9f9070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f9f90707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9f907070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0xf9070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9f9 ,
                        0x0707070707070707070707070707070707f9f9f9f9f9f9f9f9f9f9f9f9f9f907 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =67

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =4560
                            Top =5400
                            Width =585
                            Height =405
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label65"
                            Caption ="&Close Form"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    AccessKey =80
                    Left =2220
                    Top =5400
                    Width =576
                    Height =576
                    TabIndex =32
                    Name ="cmdPrint"
                    Caption ="Print"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaad00000000000dadd0888888888080da ,
                        0x000000000000080d0888888bbb88000a088888877788080d0000000000000880 ,
                        0x0888888888808080d000000000080800ad0ffffffff08080dad0f00000f0000a ,
                        0xada0ffffffff0daddada0f00000f0adaadad0ffffffff0addadad000000000da ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    UnicodeAccessKey =80

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =2820
                            Top =5400
                            Width =405
                            Height =240
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label9"
                            Caption ="&Print"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =5520
                    Top =1920
                    Width =300
                    Height =180
                    TabIndex =10
                    Name ="OrderGuideByCategory"
                    ControlSource ="OrderGuideByCategory"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Order Guide will be grouped by SubCategories"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =2
                            OverlapFlags =247
                            TextAlign =2
                            Left =4980
                            Top =1260
                            Width =855
                            Height =540
                            Name ="Label75"
                            Caption ="Group By Category"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =5505
                    Top =3720
                    Width =300
                    Height =180
                    TabIndex =25
                    Name ="InvoiceByCategory"
                    ControlSource ="InvoiceByCategory"
                    ControlTipText ="Invoices will be grouped by SubCategories"

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =6045
                    Top =3600
                    Width =486
                    Height =426
                    TabIndex =19
                    Name ="cmdPrintPreviewInvoiceCombined"
                    Caption ="Command34"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000ad000ffffffffff0d000 ,
                        0x0fffffff0000700d0ffffff0788707da0fffff0788e770ad0fffff08888780da ,
                        0x0fffff08e88780ad0fffff07ee8770da0ffffff078870dad0fffffff0000dada ,
                        0x0ffffffffff0adad0fffffff0000dada0fffffff080dadad0fffffff00dadada ,
                        0x000000000dadadad000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"

                    Overlaps =1
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    Left =1995
                    Top =3660
                    Width =465
                    Height =300
                    TabIndex =22
                    BackColor =10092543
                    Name ="InvoicesCombined"
                    ControlSource ="InvoicesCombined"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =465
                            Top =3660
                            Width =1470
                            Height =285
                            Name ="Label36"
                            Caption ="Customer copies:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =2
                    Left =2565
                    Top =3660
                    Width =600
                    Height =300
                    TabIndex =23
                    Name ="txtInvoicesCombined"
                    ValidationRule ="Between 0 And 5"
                    ValidationText ="This value must be between 0 and 5"
                    FontName ="Trebuchet MS"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3285
                    Top =3660
                    Width =2070
                    TabIndex =24
                    Name ="InvoiceCombinedListSort"
                    ControlSource ="InvoiceCombinedListSort"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "
                    ValidationRule ="Is Not Null"
                    ValidationText ="Please select one of the sort methods listed in the drop down list."
                    DefaultValue ="\"LineID\""

                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    IMESentenceMode =3
                    Left =5220
                    Top =900
                    Width =1980
                    Height =285
                    TabIndex =3
                    BackColor =10092543
                    Name ="EMail"
                    ControlSource ="EMail"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =5205
                            Top =600
                            Width =1995
                            Height =240
                            Name ="Label49"
                            Caption ="Email Address"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =6780
                    Top =3600
                    Width =786
                    Height =426
                    TabIndex =20
                    Name ="cmdMailInvoice"
                    Caption ="e-Mail Customer Invoice"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x070707070707070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000ff000000fffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffb000000000000000000000000fffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000000000000000fffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707c2686d136a13000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="e-Mail Customer Invoice to e-mail address above."

                    Overlaps =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =2
                            OverlapFlags =247
                            TextAlign =2
                            Left =6720
                            Top =1260
                            Width =810
                            Height =540
                            Name ="Label47"
                            Caption ="Send eMail"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =7755
                    Top =2700
                    Width =786
                    Height =426
                    TabIndex =21
                    Name ="cmdFax2"
                    Caption ="Fax Invoice"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707000000000000000000000000000000000000 ,
                        0x0000070707070707070707070700fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f80007070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8000000000000000000000000000000000000 ,
                        0x0000f80007070707070707070000ffffffffffffffffffffffffffffffffffff ,
                        0x07070000070707070707070700ffffffffffffff0000ff0000ff0000ffffffff ,
                        0xff07070007070707070707070700fffffffffffff800fff800fff800ffffffff ,
                        0xff07000707070707070707070700ffffffffffffffffffffffffffffffffffff ,
                        0xff07000707070707070707070700ffffffffffff0000ff0000ff0000ffffffff ,
                        0xff07000707070707070707070700fffffffffffff800fff800fff800ffffffff ,
                        0x070700070707070707070707070700ffffffffffffffffffffffffffffffffff ,
                        0x070007070707070707000000000000ffffffffff0000ff0000ff0000ffffffff ,
                        0x07000000000000070007070707f800fffffffffff800fff800fff800ffffff07 ,
                        0x0700070707f8f8000007070707f8f800ffffffffffffffffffffffffffffff07 ,
                        0x0007070707f8f8000000000000000000ffffffff0000ff0000ff0000ffffff07 ,
                        0x000000000000000000ffff0707070700fffffffff800fff800fff800ffff0707 ,
                        0x000707070707f80000ffffff0707070700ffffffffffffffffffffffffff0700 ,
                        0xf80707070707f8000700ffffff070707070000000700000000000007000000f8 ,
                        0x0707070707f800070700ffffffff07070707f8000700f8f8f8f8000700f8f807 ,
                        0x0707070707f8000707070000ffffffffff070707000707070707070007070707 ,
                        0x070707f800000707070707070000ffffffffff07070707070707070707070707 ,
                        0x070700000707070707070707070700000000ffffffffffffffffffffffff0000 ,
                        0x0000070707070707070707070707070707070000000000000000000000000707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x07070707070707070063010203022c000103011000010301ae03010301300701
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="FAX Customer Invoice"

                    LayoutCachedLeft =7755
                    LayoutCachedTop =2700
                    LayoutCachedWidth =8541
                    LayoutCachedHeight =3126
                    Overlaps =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =2
                            OverlapFlags =255
                            TextAlign =2
                            Left =7740
                            Top =360
                            Width =816
                            Height =540
                            Name ="Label48"
                            Caption ="Send  Fax"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =7740
                            LayoutCachedTop =360
                            LayoutCachedWidth =8556
                            LayoutCachedHeight =900
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =7320
                    Top =900
                    Height =285
                    TabIndex =4
                    BackColor =10092543
                    Name ="Phone2"
                    ControlSource ="Phone2"
                    InputMask ="!\\(999\") \"000\\-0000;;_"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =1
                            Left =7320
                            Top =600
                            Width =1260
                            Height =240
                            Name ="Label46"
                            Caption ="Fax Number"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =7740
                    Top =1860
                    Width =786
                    Height =426
                    TabIndex =6
                    Name ="cmdFaxOG"
                    Caption ="Fax Invoice"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707000000000000000000000000000000000000 ,
                        0x0000070707070707070707070700fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f80007070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8fff8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f800070707070707070700f8000000000000000000000000000000000000 ,
                        0x0000f80007070707070707070000ffffffffffffffffffffffffffffffffffff ,
                        0x07070000070707070707070700ffffffffffffff0000ff0000ff0000ffffffff ,
                        0xff07070007070707070707070700fffffffffffff800fff800fff800ffffffff ,
                        0xff07000707070707070707070700ffffffffffffffffffffffffffffffffffff ,
                        0xff07000707070707070707070700ffffffffffff0000ff0000ff0000ffffffff ,
                        0xff07000707070707070707070700fffffffffffff800fff800fff800ffffffff ,
                        0x070700070707070707070707070700ffffffffffffffffffffffffffffffffff ,
                        0x070007070707070707000000000000ffffffffff0000ff0000ff0000ffffffff ,
                        0x07000000000000070007070707f800fffffffffff800fff800fff800ffffff07 ,
                        0x0700070707f8f8000007070707f8f800ffffffffffffffffffffffffffffff07 ,
                        0x0007070707f8f8000000000000000000ffffffff0000ff0000ff0000ffffff07 ,
                        0x000000000000000000ffff0707070700fffffffff800fff800fff800ffff0707 ,
                        0x000707070707f80000ffffff0707070700ffffffffffffffffffffffffff0700 ,
                        0xf80707070707f8000700ffffff070707070000000700000000000007000000f8 ,
                        0x0707070707f800070700ffffffff07070707f8000700f8f8f8f8000700f8f807 ,
                        0x0707070707f8000707070000ffffffffff070707000707070707070007070707 ,
                        0x070707f800000707070707070000ffffffffff07070707070707070707070707 ,
                        0x070700000707070707070707070700000000ffffffffffffffffffffffff0000 ,
                        0x0000070707070707070707070707070707070000000000000000000000000707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="FAX Customer Order guide to fax number in yellow"

                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =3300
                    Top =4620
                    Width =1305
                    Height =285
                    Name ="Label52"
                    Caption ="Sorts by ITEM ID"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =255
                    Left =1965
                    Top =3180
                    Width =465
                    Height =300
                    TabIndex =16
                    BackColor =10092543
                    Name ="StoreInvoices"
                    ControlSource ="StoreInvoices"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =255
                            TextAlign =3
                            Left =705
                            Top =3180
                            Width =1200
                            Height =285
                            Name ="Label54"
                            Caption =" Store copies:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =255
                    TextAlign =2
                    Left =2535
                    Top =3180
                    Width =600
                    Height =300
                    TabIndex =17
                    Name ="txtStoreInvoices"
                    ValidationRule ="Between 0 And 5"
                    ValidationText ="This value must be between 0 and 5"
                    FontName ="Trebuchet MS"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    Left =3285
                    Top =3180
                    Width =2070
                    TabIndex =18
                    Name ="StoreInvoiceSort"
                    ControlSource ="StoreInvoiceSort"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "
                    ValidationRule ="Is Not Null"
                    ValidationText ="Please select one of the sort methods listed in the drop down list."
                    DefaultValue ="\"LineID\""

                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =6045
                    Top =3120
                    Width =486
                    Height =426
                    TabIndex =15
                    Name ="cmdPrintPreviewInvoiceStoreCopy"
                    Caption ="Command25"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadada000000000000ad000ffffffffff0d000 ,
                        0x0fffffff0000700d0ffffff0788707da0fffff0788e770ad0fffff08888780da ,
                        0x0fffff08e88780ad0fffff07ee8770da0ffffff078870dad0fffffff0000dada ,
                        0x0ffffffffff0adad0fffffff0000dada0fffffff080dadad0fffffff00dadada ,
                        0x000000000dadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Preview Report"

                    Overlaps =1
                End
                Begin TextBox
                    Locked = NotDefault
                    TabStop = NotDefault
                    SpecialEffect =0
                    OldBorderStyle =0
                    OverlapFlags =247
                    TextAlign =3
                    BackStyle =0
                    IMESentenceMode =3
                    Left =480
                    Top =1920
                    Height =480
                    TabIndex =34
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="ClubOrderGuideReport"
                    ControlSource ="ClubOrderGuideReport"
                    FontName ="Trebuchet MS"

                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =240
                    Top =2580
                    Width =3000
                    Height =1500
                    Name ="Box59"
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =255
                    TextAlign =4
                    Left =360
                    Top =2460
                    Width =960
                    Height =240
                    Name ="Label60"
                    Caption ="  Invoices.."
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =6765
                    Top =1860
                    Width =786
                    Height =426
                    TabIndex =35
                    Name ="cmdMailOrderGuide"
                    Caption ="Command61"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x070707070707070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000ff000000fffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffb000000000000000000000000fffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000000000000000fffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707c2686d136a13000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Mail Report"

                    Overlaps =1
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =5520
                    Top =4200
                    Width =300
                    Height =180
                    TabIndex =36
                    Name ="PickListByCategory"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="GetKMDSSettings \"PickListByCategory\""

                    LayoutCachedLeft =5520
                    LayoutCachedTop =4200
                    LayoutCachedWidth =5820
                    LayoutCachedHeight =4380
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =7740
                    Top =3600
                    Width =786
                    Height =426
                    TabIndex =37
                    Name ="cmdEmailCSV"
                    Caption ="e-Mail CSV File"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000020000000200000000100080000000000000400000000000000000000 ,
                        0x0000000000000000000000000000bf0000bf000000bfbf00bf000000bf00bf00 ,
                        0xbfbf0000c0c0c000c0dcc000f0c8a40000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0xf0fbff00a4a0a000808080000000ff0000ff000000ffff00ff000000ff00ff00 ,
                        0xffff0000ffffff00070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x070707070707070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000ff000000fffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffb000000000000000000000000fffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbff000000000000000000000000000000fffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff ,
                        0xfbfffbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb0505 ,
                        0x0505fffb00f8070700fffbfffbfffbfffbfffbfffbfffbfffbfffbfffbff0505 ,
                        0x0505fbff00f8070700fbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffbfffb ,
                        0xfffbfffb00f80707000000000000000000000000000000000000000000000000 ,
                        0x0000000000070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="e-Mail Customer Invoice to e-mail address above."

                    LayoutCachedLeft =7740
                    LayoutCachedTop =3600
                    LayoutCachedWidth =8526
                    LayoutCachedHeight =4026
                    Overlaps =1
                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =2
                            OverlapFlags =247
                            TextAlign =2
                            Left =7680
                            Top =1260
                            Width =810
                            Height =540
                            Name ="Label66"
                            Caption ="Send eMail"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =7680
                            LayoutCachedTop =1260
                            LayoutCachedWidth =8490
                            LayoutCachedHeight =1800
                        End
                    End
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =2
                    OverlapFlags =247
                    TextAlign =2
                    Left =5880
                    Top =1260
                    Width =750
                    Height =540
                    Name ="Label39"
                    Caption ="Print Preview"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =3
                    Left =735
                    Top =2700
                    Width =1170
                    Height =285
                    Name ="Invoices_Label"
                    Caption ="  OLCC copies:"
                    FontName ="Trebuchet MS"
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
            OnClick ="[Event Procedure]"
        End
    End
End
CodeBehindForm
' See "ClubOrderPrintDialog.cls"
