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
    Width =8700
    DatasheetFontHeight =10
    ItemSuffix =40
    Left =585
    Top =2430
    Right =9285
    Bottom =6435
    Filter ="[AgentNumber]=1181"
    RecSrcDt = Begin
        0x74838e7bb0a3e240
    End
    RecordSource ="Agents"
    Caption ="Please make a choice then press continue..."
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
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
            Height =4170
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionGroup
                    OverlapFlags =93
                    Left =540
                    Top =1920
                    Width =7682
                    Height =1261
                    TabIndex =7
                    Name ="optChoices"
                    AfterUpdate ="[Event Procedure]"
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
                                    Caption ="Print  Transfer"
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
                                    Caption ="Post Transfer"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3840
                    Top =3300
                    Width =1815
                    Height =576
                    TabIndex =13
                    Name ="Command11"
                    Caption ="Continue"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =2970
                    Top =540
                    Width =660
                    BackColor =65280
                    ForeColor =-2147483640
                    Name ="AgentNumber"
                    ControlSource ="AgentNumber"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2280
                            Top =540
                            Width =645
                            Height =240
                            Name ="Label18"
                            Caption ="Agent"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    Left =2925
                    Top =900
                    Width =1395
                    Height =255
                    TabIndex =3
                    BackColor =65280
                    ForeColor =-2147483640
                    Name ="txtAgentTransferID"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2055
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
                    TabStop = NotDefault
                    OverlapFlags =215
                    Left =780
                    Top =2160
                    Width =426
                    Height =366
                    TabIndex =8
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
                    Left =5820
                    Top =3300
                    Width =666
                    Height =606
                    TabIndex =14
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
                            Left =6540
                            Top =3300
                            Width =585
                            Height =405
                            Name ="Label65"
                            Caption ="&Close Form"
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
                    IMESentenceMode =3
                    Left =2925
                    Top =1260
                    Width =1395
                    TabIndex =5
                    BackColor =65280
                    ForeColor =-2147483640
                    Name ="Status"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =2280
                            Top =1260
                            Width =600
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label25"
                            Caption ="Status:"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =215
                    TextAlign =2
                    IMESentenceMode =3
                    Left =4980
                    Top =2580
                    Width =1080
                    TabIndex =10
                    Name ="txtNumberOfCopies"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            OverlapFlags =215
                            Left =3540
                            Top =2580
                            Width =1320
                            Height =240
                            Name ="Label29"
                            Caption ="Number of copies"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3780
                    Top =540
                    Width =1920
                    TabIndex =1
                    BackColor =65280
                    ForeColor =-2147483640
                    Name ="FirstName"
                    ControlSource ="FirstName"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5880
                    Top =540
                    Width =2280
                    TabIndex =2
                    BackColor =65280
                    ForeColor =-2147483640
                    Name ="LastName"
                    ControlSource ="LastName"
                    FontName ="Trebuchet MS"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    IMESentenceMode =3
                    Left =3540
                    Top =2220
                    Width =1380
                    TabIndex =9
                    BackColor =-2147483643
                    ForeColor =-2147483640
                    Name ="TransferSort"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT ReportForms.ReportName, ReportForms.ReportType FROM ReportForms WHERE ((("
                        "ReportForms.ReportType)=\"ClubPickList\")) ORDER BY ReportForms.ReportName; "
                    ValidationRule ="Is Not Null"
                    ValidationText ="Please select one of the sort methods listed in the drop down list."
                    DefaultValue ="\"LineID\""
                    FontName ="MS Sans Serif"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =2
                            OverlapFlags =85
                            TextAlign =2
                            Left =3540
                            Top =1560
                            Width =1380
                            Height =300
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label38"
                            Caption ="Sort By"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OldBorderStyle =1
                    OverlapFlags =85
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5880
                    Top =1200
                    Width =2280
                    TabIndex =4
                    BackColor =65280
                    ForeColor =-2147483640
                    Name ="EMail"
                    ControlSource ="EMail"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =5040
                            Top =1200
                            Width =735
                            Height =240
                            Name ="Label32"
                            Caption ="EMail:"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =5880
                    Top =1530
                    TabIndex =6
                    Name ="KMDScustomer"
                    ControlSource ="KMDScustomer"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =6110
                            Top =1500
                            Width =1575
                            Height =240
                            Name ="Label35"
                            Caption ="KMDS Customer"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =540
                    Top =3300
                    Width =1200
                    Height =576
                    TabIndex =11
                    Name ="cmdEmail"
                    Caption ="eMail Transfer"
                    OnClick ="[Event Procedure]"

                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =1800
                    Top =3300
                    Width =1506
                    Height =576
                    TabIndex =12
                    Name ="cmdKMDSFTP"
                    Caption ="Transmit KMDS FTP"
                    OnClick ="[Event Procedure]"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Find Next"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =3300
                    LayoutCachedWidth =3306
                    LayoutCachedHeight =3876
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =1800
                    Top =3930
                    TabIndex =15
                    Name ="chkUseWinSCPForOLASTransmits"

                    LayoutCachedLeft =1800
                    LayoutCachedTop =3930
                    LayoutCachedWidth =2060
                    LayoutCachedHeight =4170
                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =2025
                            Top =3900
                            Width =1200
                            Height =240
                            Name ="Label44"
                            Caption ="FTP via WinSCP"
                            LayoutCachedLeft =2025
                            LayoutCachedTop =3900
                            LayoutCachedWidth =3225
                            LayoutCachedHeight =4140
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AgentsTransfersChoices.cls"
