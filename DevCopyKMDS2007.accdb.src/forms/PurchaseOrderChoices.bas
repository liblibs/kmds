Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDeletions = NotDefault
    CloseButton = NotDefault
    DividingLines = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    ViewsAllowed =1
    TabularFamily =9
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =8715
    DatasheetFontHeight =10
    ItemSuffix =31
    Left =4260
    Top =1890
    Right =12975
    Bottom =5985
    Filter ="[ID]=1"
    RecSrcDt = Begin
        0x937722a83ea2e240
    End
    RecordSource ="Vendor"
    Caption ="Please make a choice then press continue..."
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =255
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin CommandButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
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
        Begin TextBox
            FELineBreak = NotDefault
            SpecialEffect =2
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Tahoma"
            AsianLineBreak =255
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin Section
            Height =4116
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionGroup
                    OverlapFlags =93
                    Left =600
                    Top =1920
                    Width =7622
                    Height =1501
                    Name ="optChoices"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            BackStyle =1
                            OverlapFlags =215
                            Left =660
                            Top =1800
                            Width =660
                            Height =285
                            BackColor =-2147483633
                            Name ="Label1"
                            Caption ="Choices"
                            FontName ="Trebuchet MS"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1380
                            Top =2248
                            OptionValue =1
                            Name ="Option5"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1610
                                    Top =2220
                                    Width =1725
                                    Height =285
                                    Name ="Label6"
                                    Caption ="Print  Purchase Order"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1380
                            Top =2608
                            OptionValue =2
                            Name ="Option7"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1610
                                    Top =2580
                                    Width =1680
                                    Height =285
                                    Name ="Label8"
                                    Caption ="Post  Purchase Order"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1380
                            Top =2956
                            OptionValue =3
                            Name ="Option26"

                            Begin
                                Begin Label
                                    OverlapFlags =247
                                    Left =1605
                                    Top =2925
                                    Width =2415
                                    Height =285
                                    Name ="Label27"
                                    Caption ="Print and Post  Purchase Order"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =2760
                    Top =3540
                    Width =915
                    Height =450
                    TabIndex =1
                    Name ="Command11"
                    Caption ="Continue"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2940
                    Top =420
                    Width =660
                    TabIndex =2
                    BackColor =65280
                    Name ="ID"
                    ControlSource ="ID"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2250
                            Top =420
                            Width =645
                            Height =240
                            Name ="Label18"
                            Caption ="Vendor"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =3660
                    Top =420
                    Width =2580
                    ColumnWidth =2550
                    TabIndex =3
                    BackColor =65280
                    Name ="txtName"
                    ControlSource ="Name"

                End
                Begin ComboBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3480
                    Top =2220
                    Width =4560
                    TabIndex =4
                    Name ="POReportForm"
                    ControlSource ="POReportForm"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                        "ReportForms.ReportType)=\"PurchaseOrder\")); "

                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =3480
                            Top =1980
                            Width =1335
                            Height =240
                            Name ="Label20"
                            Caption ="PO ReportFormat"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =2940
                    Top =900
                    Width =2190
                    Height =255
                    TabIndex =5
                    BackColor =65280
                    ForeColor =-2147483640
                    Name ="txtPurchaseOrderID"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2070
                            Top =900
                            Width =825
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Order_Label"
                            Caption ="Order No."
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =780
                    Top =2160
                    Width =426
                    Height =366
                    TabIndex =6
                    Name ="cmdPrintPreview"
                    Caption ="Command22"
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
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =67
                    Left =4440
                    Top =3480
                    Width =666
                    Height =606
                    TabIndex =7
                    ForeColor =0
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
                    FontName ="MS Sans Serif"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"
                    UnicodeAccessKey =67

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =5160
                            Top =3480
                            Width =585
                            Height =405
                            Name ="Label65"
                            Caption ="&Close Form"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =5160
                    Top =2580
                    Width =2400
                    TabIndex =8
                    Name ="Phone3"
                    ControlSource ="Phone3"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =4380
                            Top =2580
                            Width =675
                            Height =240
                            Name ="Label24"
                            Caption ="Sort by"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =6540
                    Top =3480
                    Width =576
                    Height =576
                    TabIndex =9
                    Name ="cmdExport"
                    Caption ="Command28"
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
                        0x0707070707070707070707070700000000000000000000000000000000000000 ,
                        0x000000000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706060607060706 ,
                        0x070607000707070707070707070000f806070600000000000000000000000000 ,
                        0x060706000707070707070707000000f807060700ffffffffffffffffffffff00 ,
                        0x070607000707070707070700fffb030000070600ffffffffffffffffffffff00 ,
                        0x060706000707070707070707000000f807060700ffffffffffffffffffffff00 ,
                        0x070607000707070707070707070000f806070600000000000000000000000000 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070700fffb030000070607060706070607060706070607 ,
                        0x060706000707070707070707000000f807060706070607060706070607060706 ,
                        0x070607000707070707070707070000f806070607060706070607060706070607 ,
                        0x0607060007070707070707070700000000000000000000000000000000000000 ,
                        0x0000000007070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Report to File"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =7260
                            Top =3540
                            Width =1335
                            Height =210
                            Name ="Label29"
                            Caption ="Send report to file"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =4620
                    Top =2940
                    Height =210
                    TabIndex =10
                    Name ="Check52"
                    ControlSource ="PoAtCost"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4860
                            Top =2880
                            Width =2415
                            Height =525
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label30"
                            Caption ="DO NOT Print Alternating Gray Lines on Purchase orders"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "PurchaseOrderChoices.cls"
