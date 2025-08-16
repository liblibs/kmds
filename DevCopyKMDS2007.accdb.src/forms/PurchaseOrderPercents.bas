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
    Width =13755
    DatasheetFontHeight =10
    ItemSuffix =57
    RecSrcDt = Begin
        0xeab3524a14c2e240
    End
    RecordSource ="PurchaseOrderPercents"
    Caption ="Average Daily Sales Calculator"
    DatasheetFontName ="Trebuchet MS"
    AllowDatasheetView =0
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
            CanGrow = NotDefault
            Height =6276
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Tab
                    OverlapFlags =85
                    Width =13755
                    Height =6150
                    TabIndex =1
                    Name ="TabCtl37"

                    Begin
                        Begin Page
                            OverlapFlags =87
                            Left =120
                            Top =405
                            Width =13500
                            Height =5616
                            Name ="Page38"
                            Caption ="Monthly Percentages"
                            Begin
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =420
                                    Width =600
                                    BackColor =16777164
                                    Name ="CurrMM"
                                    ControlSource ="CurrMM"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =780
                                    Width =600
                                    Height =255
                                    TabIndex =1
                                    BackColor =16777164
                                    Name ="Jan"
                                    ControlSource ="Jan"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =1140
                                    Width =600
                                    Height =255
                                    TabIndex =2
                                    BackColor =16777164
                                    Name ="Feb"
                                    ControlSource ="Feb"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =1500
                                    Width =600
                                    Height =255
                                    TabIndex =3
                                    BackColor =16777164
                                    Name ="Mar"
                                    ControlSource ="Mar"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =1860
                                    Width =600
                                    Height =255
                                    TabIndex =4
                                    BackColor =16777164
                                    Name ="Apr"
                                    ControlSource ="Apr"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =2220
                                    Width =600
                                    Height =255
                                    TabIndex =5
                                    BackColor =16777164
                                    Name ="May"
                                    ControlSource ="May"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =2580
                                    Width =600
                                    Height =255
                                    TabIndex =6
                                    BackColor =16777164
                                    Name ="Jun"
                                    ControlSource ="Jun"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =2940
                                    Width =600
                                    Height =255
                                    TabIndex =7
                                    BackColor =16777164
                                    Name ="Jul"
                                    ControlSource ="Jul"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =3300
                                    Width =600
                                    Height =255
                                    TabIndex =8
                                    BackColor =16777164
                                    Name ="Aug"
                                    ControlSource ="Aug"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =3660
                                    Width =600
                                    Height =255
                                    TabIndex =9
                                    BackColor =16777164
                                    Name ="Sep"
                                    ControlSource ="Sep"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =4020
                                    Width =600
                                    Height =255
                                    TabIndex =10
                                    BackColor =16777164
                                    Name ="Oct"
                                    ControlSource ="Oct"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =4380
                                    Width =600
                                    Height =255
                                    TabIndex =11
                                    BackColor =16777164
                                    Name ="Nov"
                                    ControlSource ="Nov"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =2
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =4740
                                    Width =600
                                    Height =255
                                    TabIndex =12
                                    BackColor =16777164
                                    Name ="Dec"
                                    ControlSource ="Dec"
                                    Format ="Percent"
                                    OnEnter ="[Event Procedure]"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                End
                                Begin TextBox
                                    Enabled = NotDefault
                                    Locked = NotDefault
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =3960
                                    Top =5100
                                    Width =630
                                    Height =255
                                    TabIndex =13
                                    BackColor =16777164
                                    Name ="Text28"
                                    ControlSource ="=[Jan]+[Feb]+[Mar]+[apr]+[may]+[jun]+[jul]+[aug]+[sep]+[oct]+[nov]+[dec]+[currmm"
                                        "]"
                                    Format ="Percent"
                                    ConditionalFormat = Begin
                                        0x010000007a000000020000000000000000000000000000000200000001000000 ,
                                        0x00000000ccffcc000000000004000000070000000c0000000100000000000000 ,
                                        0xffff000000000000000000000000000000000000000000000000000000000000 ,
                                        0x3000000030002e0039003900000031002e003000310000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000200000000000000000000000100000000000000ccffcc00010000003000 ,
                                        0x0400000030002e00390039000000000000000000000000000000000000000000 ,
                                        0x00040000000100000000000000ffff00000400000031002e0030003100000000 ,
                                        0x000000000000000000000000000000000000
                                    End
                                End
                                Begin CommandButton
                                    TabStop = NotDefault
                                    OverlapFlags =215
                                    AccessKey =82
                                    Left =5220
                                    Top =480
                                    Width =1080
                                    Height =780
                                    TabIndex =14
                                    Name ="cmdReset"
                                    Caption ="&Reset All Months to Equal"
                                    OnClick ="[Event Procedure]"
                                    UnicodeAccessKey =82

                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =5520
                                    Top =5445
                                    Width =576
                                    Height =576
                                    TabIndex =15
                                    Name ="cmdUp"
                                    Caption ="Command30"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                                        0xadadadadadadadaddadadacadadadadaadadadccadadadaddadadacccadadada ,
                                        0xadadadccccadadaddadadacccccadadaadadadccccadadaddadadacccadadada ,
                                        0xadadadccadadadaddadadacadadadadaadadadadadadadaddadadadadadadada ,
                                        0xadadadadadadadad
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Next Record"

                                End
                                Begin CommandButton
                                    OverlapFlags =215
                                    Left =4500
                                    Top =5445
                                    Width =576
                                    Height =576
                                    TabIndex =16
                                    Name ="cmdLower"
                                    Caption ="Command31"
                                    OnClick ="[Event Procedure]"
                                    PictureData = Begin
                                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                                        0xadadadadadadadaddadadadadcdadadaadadadadccadadaddadadadcccdadada ,
                                        0xadadadccccadadaddadadcccccdadadaadadadccccadadaddadadadcccdadada ,
                                        0xadadadadccadadaddadadadadcdadadaadadadadadadadaddadadadadadadada ,
                                        0xadadadadadadadad
                                    End
                                    ObjectPalette = Begin
                                        0x000301000000000000000000
                                    End
                                    ControlTipText ="Previous Record"

                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =7200
                                    Top =2460
                                    Width =5100
                                    Height =660
                                    Name ="Label34"
                                    Caption ="Average daily sales can be calculated based on the percentage of each month's sa"
                                        "le... for an average first click the \"Reset All Months to Equal Button\"  then "
                                        "click the \"Update Average Daily Sales\"."
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =7200
                                    Top =3180
                                    Width =5100
                                    Height =660
                                    Name ="Label35"
                                    Caption ="You can \"weight\" any month by adjusting the percent.  The \"Total Percent\" wi"
                                        "ll help you adjust to 100%, when satisfied with your choices click on the \"Upda"
                                        "te Average Daily Sales\"."
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3240
                                    Top =420
                                    Width =630
                                    Height =240
                                    Name ="CurrMM_Label"
                                    Caption ="CurrMM"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3540
                                    Top =780
                                    Width =330
                                    Height =255
                                    Name ="Jan_Label"
                                    Caption ="Jan"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3525
                                    Top =1140
                                    Width =345
                                    Height =255
                                    Name ="Feb_Label"
                                    Caption ="Feb"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3525
                                    Top =1500
                                    Width =345
                                    Height =255
                                    Name ="Mar_Label"
                                    Caption ="Mar"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3555
                                    Top =1860
                                    Width =315
                                    Height =255
                                    Name ="Apr_Label"
                                    Caption ="Apr"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3495
                                    Top =2220
                                    Width =375
                                    Height =255
                                    Name ="May_Label"
                                    Caption ="May"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3540
                                    Top =2580
                                    Width =330
                                    Height =255
                                    Name ="Jun_Label"
                                    Caption ="Jun"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3600
                                    Top =2940
                                    Width =270
                                    Height =255
                                    Name ="Jul_Label"
                                    Caption ="Jul"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3510
                                    Top =3300
                                    Width =360
                                    Height =255
                                    Name ="Aug_Label"
                                    Caption ="Aug"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3510
                                    Top =3660
                                    Width =360
                                    Height =255
                                    Name ="Sep_Label"
                                    Caption ="Sep"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3540
                                    Top =4020
                                    Width =330
                                    Height =255
                                    Name ="Oct_Label"
                                    Caption ="Oct"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3495
                                    Top =4380
                                    Width =375
                                    Height =255
                                    Name ="Nov_Label"
                                    Caption ="Nov"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    TextAlign =3
                                    Left =3495
                                    Top =4740
                                    Width =375
                                    Height =255
                                    Name ="Dec_Label"
                                    Caption ="Dec"
                                    Tag ="DetachedLabel"
                                End
                                Begin Label
                                    OverlapFlags =215
                                    Left =6780
                                    Top =3900
                                    Width =5100
                                    Height =660
                                    Name ="Label41"
                                    Caption ="For New items (set up within the last year) the Average Daily  Sales are based o"
                                        "n the number of days the item has been setup. "
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =420
                                    Width =1080
                                    TabIndex =17
                                    Name ="CurrentMonth"
                                    ControlSource ="CurrentMonth"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =855
                                            Top =420
                                            Width =1080
                                            Height =240
                                            Name ="Label42"
                                            Caption ="CurrentMonth:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =780
                                    Width =1080
                                    Height =255
                                    TabIndex =18
                                    Name ="Month1"
                                    ControlSource ="Month1"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =855
                                            Top =780
                                            Width =1080
                                            Height =255
                                            Name ="Label43"
                                            Caption ="Last Month"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =1140
                                    Width =1080
                                    Height =255
                                    TabIndex =19
                                    Name ="Month2"
                                    ControlSource ="Month2"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =840
                                            Top =1140
                                            Width =1095
                                            Height =240
                                            Name ="Label44"
                                            Caption ="2 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =1500
                                    Width =1080
                                    Height =255
                                    TabIndex =20
                                    Name ="Month3"
                                    ControlSource ="Month3"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =840
                                            Top =1500
                                            Width =1095
                                            Height =255
                                            Name ="Label45"
                                            Caption ="3 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =1860
                                    Width =1080
                                    Height =255
                                    TabIndex =21
                                    Name ="Month4"
                                    ControlSource ="Month4"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =840
                                            Top =1860
                                            Width =1095
                                            Height =255
                                            Name ="Label46"
                                            Caption ="4 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =2220
                                    Width =1080
                                    Height =255
                                    TabIndex =22
                                    Name ="Month5"
                                    ControlSource ="Month5"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =840
                                            Top =2220
                                            Width =1095
                                            Height =255
                                            Name ="Label47"
                                            Caption ="5 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =2580
                                    Width =1080
                                    Height =255
                                    TabIndex =23
                                    Name ="Month6"
                                    ControlSource ="Month6"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =840
                                            Top =2580
                                            Width =1095
                                            Height =255
                                            Name ="Label48"
                                            Caption ="6 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =2940
                                    Width =1080
                                    Height =255
                                    TabIndex =24
                                    Name ="Month7"
                                    ControlSource ="Month7"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =840
                                            Top =2940
                                            Width =1095
                                            Height =255
                                            Name ="Label49"
                                            Caption ="7 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =3300
                                    Width =1080
                                    Height =255
                                    TabIndex =25
                                    Name ="Month8"
                                    ControlSource ="Month8"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =840
                                            Top =3300
                                            Width =1095
                                            Height =255
                                            Name ="Label50"
                                            Caption ="8 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =3660
                                    Width =1080
                                    Height =255
                                    TabIndex =26
                                    Name ="Month9"
                                    ControlSource ="Month9"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =840
                                            Top =3660
                                            Width =1095
                                            Height =255
                                            Name ="Label51"
                                            Caption ="9 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =4020
                                    Width =1080
                                    Height =255
                                    TabIndex =27
                                    Name ="Month10"
                                    ControlSource ="Month10"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =750
                                            Top =4020
                                            Width =1185
                                            Height =255
                                            Name ="Label52"
                                            Caption ="10 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =4380
                                    Width =1080
                                    Height =255
                                    TabIndex =28
                                    Name ="Month11"
                                    ControlSource ="Month11"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =750
                                            Top =4380
                                            Width =1185
                                            Height =255
                                            Name ="Label53"
                                            Caption ="11 Months ago:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    DecimalPlaces =1
                                    OverlapFlags =215
                                    TextAlign =3
                                    IMESentenceMode =3
                                    Left =1995
                                    Top =4740
                                    Width =1080
                                    Height =255
                                    TabIndex =29
                                    Name ="Month12"
                                    ControlSource ="Month12"
                                    Format ="Percent"
                                    OnExit ="[Event Procedure]"
                                    ConditionalFormat = Begin
                                        0x0100000066000000010000000000000004000000000000000200000001000000 ,
                                        0x00000000ccffcc00000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x300000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000100000000000000040000000100000000000000ccffcc00010000003000 ,
                                        0x000000000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =120
                                            Top =4740
                                            Width =1815
                                            Height =255
                                            Name ="Label54"
                                            Caption ="This Month Last Year:"
                                        End
                                    End
                                End
                                Begin TextBox
                                    TabStop = NotDefault
                                    DecimalPlaces =0
                                    OverlapFlags =215
                                    IMESentenceMode =3
                                    Left =1980
                                    Top =5055
                                    Width =1140
                                    Height =255
                                    TabIndex =30
                                    BackColor =16777164
                                    Name ="Text55"
                                    ControlSource ="=[Month1]+[Month2]+[Month3]+[Month4]+[Month5]+[Month6]+[Month7]+[Month8]+[Month9"
                                        "]+[Month10]+[Month11]+[Month12]+[currentmonth]"
                                    Format ="Percent"
                                    ConditionalFormat = Begin
                                        0x010000007a000000020000000000000000000000000000000200000001000000 ,
                                        0x00000000ccffcc000000000004000000070000000c0000000100000000000000 ,
                                        0xffff000000000000000000000000000000000000000000000000000000000000 ,
                                        0x3000000030002e0039003900000031002e003000310000000000
                                    End

                                    ConditionalFormat14 = Begin
                                        0x01000200000000000000000000000100000000000000ccffcc00010000003000 ,
                                        0x0400000030002e00390039000000000000000000000000000000000000000000 ,
                                        0x00040000000100000000000000ffff00000400000031002e0030003100000000 ,
                                        0x000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =215
                                            TextAlign =3
                                            Left =900
                                            Top =5055
                                            Width =1035
                                            Height =255
                                            Name ="Label56"
                                            Caption ="Total Percent"
                                        End
                                    End
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =135
                            Top =405
                            Width =13485
                            Height =5610
                            Name ="Page39"
                            Caption ="Days Supply / Unit Sales Group"
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =300
                                    Top =705
                                    Width =10080
                                    Height =3300
                                    Name ="PurchaseOrderPercentsDSGroupLevelSubForm"
                                    SourceObject ="Form.PurchaseOrderPercentsDSGroupLevelSubForm"

                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =300
                                            Top =405
                                            Width =3480
                                            Height =240
                                            Name ="PurchaseOrderPercentsDSGroupLevelSubForm Label"
                                            Caption ="Day's supply can be assigned by the Unit Sales "
                                            EventProcPrefix ="PurchaseOrderPercentsDSGroupLevelSubForm_Label"
                                        End
                                    End
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =9780
                    Top =5700
                    Width =576
                    Height =576
                    Name ="cmdClose"
                    Caption ="Command36"
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
                        0x0707070707070707000000000000000000000000000000000000000000000000
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Close Form"

                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    AccessKey =85
                    Left =480
                    Top =5460
                    Width =2580
                    Height =600
                    TabIndex =2
                    Name ="Command32"
                    Caption ="&Update Average Daily Sales and Set Days Supply Group Level"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =85

                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =0
            BackColor =-2147483633
            Name ="FormFooter"
        End
    End
End
CodeBehindForm
' See "PurchaseOrderPercents.cls"
