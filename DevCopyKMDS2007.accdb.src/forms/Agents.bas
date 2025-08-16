Version =20
VersionRequired =20
Begin Form
    AutoCenter = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =9960
    DatasheetFontHeight =10
    ItemSuffix =69
    Left =4230
    Top =1455
    Right =14310
    Bottom =8850
    FrozenColumns =3
    HelpContextId =3600
    Tag ="63 OLCC Agent List"
    Filter ="([_Agents].[ActiveForTransfer]<>0)"
    OrderBy ="[_Agents].[Zip], [_Agents].[LastName], [_Agents].EMail DESC"
    RecSrcDt = Begin
        0x2775b41db06be340
    End
    OnDirty ="[Event Procedure]"
    RecordSource ="SELECT Agents.* FROM Agents ORDER BY Agents.FirstName; "
    Caption =" KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software © "
    OnCurrent ="[Event Procedure]"
    BeforeUpdate ="[Event Procedure]"
    AfterUpdate ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    OnError ="[Event Procedure]"
    OnMouseWheel ="[Event Procedure]"
    FilterOnLoad =0
    SplitFormOrientation =2
    SplitFormOrientation =2
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
            CanGrow = NotDefault
            Height =6360
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    Visible = NotDefault
                    ColumnHidden = NotDefault
                    OverlapFlags =93
                    Left =2385
                    Top =300
                    Width =1080
                    Height =300
                    ColumnWidth =3645
                    ColumnOrder =4
                    TabIndex =2
                    Name ="AgentNumber"
                    ControlSource ="AgentNumber"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =780
                            Top =300
                            Width =1485
                            Height =540
                            Name ="ID_Label"
                            Caption ="Agent Number Setting up a New Record"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =120
                    Width =9840
                    Height =6240
                    Name ="boxScreen"
                    LayoutCachedLeft =120
                    LayoutCachedTop =120
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =6360
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =5820
                    Width =9840
                    Height =540
                    Name ="Box13"
                    LayoutCachedLeft =120
                    LayoutCachedTop =5820
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =6360
                End
                Begin ComboBox
                    ColumnHidden = NotDefault
                    OverlapFlags =247
                    ColumnCount =5
                    ListRows =30
                    ListWidth =7200
                    Left =2385
                    Top =300
                    Width =4215
                    Height =300
                    ColumnWidth =0
                    ColumnOrder =2
                    BoundColumn =2
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cboID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Agents.LastName, Agents.FirstName, Agents.AgentNumber, Agents.Address, Ag"
                        "ents.City FROM Agents ORDER BY Agents.LastName, Agents.FirstName; "
                    ColumnWidths ="2160;1440;720;2160"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =900
                            Top =300
                            Width =1350
                            Height =495
                            Name ="Label15"
                            Caption ="Select an Agent or Click on Add New "
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =795
                    Top =5940
                    Width =2250
                    Height =300
                    ColumnOrder =5
                    TabIndex =3
                    BackColor =12632256
                    Name ="SetupDate"
                    ControlSource ="SetupDate"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =795
                    LayoutCachedTop =5940
                    LayoutCachedWidth =3045
                    LayoutCachedHeight =6240
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =240
                            Top =5940
                            Width =495
                            Height =285
                            Name ="SetupDate_Label"
                            Caption ="Setup"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =7470
                    Top =5940
                    Width =2250
                    Height =300
                    ColumnOrder =7
                    TabIndex =5
                    BackColor =12632256
                    Name ="ChangeDate"
                    ControlSource ="ChangeDate"
                    Format ="General Date"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =7470
                    LayoutCachedTop =5940
                    LayoutCachedWidth =9720
                    LayoutCachedHeight =6240
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =6780
                            Top =5940
                            Width =630
                            Height =285
                            Name ="ChangeDate_Label"
                            Caption ="Change"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =6780
                            LayoutCachedTop =5940
                            LayoutCachedWidth =7410
                            LayoutCachedHeight =6225
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    Left =4455
                    Top =5940
                    Width =705
                    Height =300
                    ColumnOrder =6
                    TabIndex =4
                    BackColor =12632256
                    Name ="OrderCount"
                    ControlSource ="OrderCount"
                    FontName ="Trebuchet MS"

                    LayoutCachedLeft =4455
                    LayoutCachedTop =5940
                    LayoutCachedWidth =5160
                    LayoutCachedHeight =6240
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =3
                            Left =3360
                            Top =5940
                            Width =1035
                            Height =285
                            Name ="Label31"
                            Caption ="OrderCount:"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =3360
                            LayoutCachedTop =5940
                            LayoutCachedWidth =4395
                            LayoutCachedHeight =6225
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    ColumnHidden = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    TextAlign =1
                    ColumnCount =5
                    ListRows =30
                    ListWidth =7200
                    Left =6720
                    Top =300
                    Width =2175
                    Height =300
                    ColumnWidth =0
                    ColumnOrder =3
                    TabIndex =1
                    ConditionalFormat = Begin
                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000
                    End
                    Name ="cboAgentByID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Agents.AgentNumber, Agents.LastName, Agents.FirstName, Agents.Address, Ag"
                        "ents.City FROM Agents ORDER BY Agents.AgentNumber; "
                    ColumnWidths ="2160;1440;720;2160"
                    AfterUpdate ="[Event Procedure]"
                    FontName ="Trebuchet MS"

                    ConditionalFormat14 = Begin
                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                        0x00000000000000000000000000000000000000
                    End
                End
                Begin Tab
                    OverlapFlags =247
                    Left =300
                    Top =900
                    Width =9480
                    Height =4755
                    TabIndex =6
                    Name ="TabCtl55"

                    LayoutCachedLeft =300
                    LayoutCachedTop =900
                    LayoutCachedWidth =9780
                    LayoutCachedHeight =5655
                    Begin
                        Begin Page
                            OverlapFlags =247
                            Left =435
                            Top =1305
                            Width =9210
                            Height =4215
                            Name ="Page56"
                            Caption ="Profile"
                            LayoutCachedLeft =435
                            LayoutCachedTop =1305
                            LayoutCachedWidth =9645
                            LayoutCachedHeight =5520
                            Begin
                                Begin TextBox
                                    TabStop = NotDefault
                                    OverlapFlags =247
                                    IMESentenceMode =3
                                    Left =2400
                                    Top =1440
                                    Width =1200
                                    ColumnOrder =0
                                    Name ="Text46"
                                    ControlSource ="AgentNumber"

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =1440
                                    LayoutCachedWidth =3600
                                    LayoutCachedHeight =1680
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =1245
                                            Top =1440
                                            Width =1140
                                            Height =240
                                            Name ="Label47"
                                            Caption ="Agent Number:"
                                            LayoutCachedLeft =1245
                                            LayoutCachedTop =1440
                                            LayoutCachedWidth =2385
                                            LayoutCachedHeight =1680
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2400
                                    Top =1800
                                    Width =2865
                                    Height =300
                                    ColumnOrder =1
                                    TabIndex =1
                                    Name ="LastName"
                                    ControlSource ="LastName"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =1800
                                    LayoutCachedWidth =5265
                                    LayoutCachedHeight =2100
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1305
                                            Top =1800
                                            Width =1080
                                            Height =285
                                            Name ="Desc_Label"
                                            Caption ="  Store Name"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1305
                                            LayoutCachedTop =1800
                                            LayoutCachedWidth =2385
                                            LayoutCachedHeight =2085
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2400
                                    Top =2220
                                    Width =2865
                                    Height =300
                                    TabIndex =2
                                    Name ="FirstName"
                                    ControlSource ="FirstName"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =2220
                                    LayoutCachedWidth =5265
                                    LayoutCachedHeight =2520
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1305
                                            Top =2220
                                            Width =1080
                                            Height =285
                                            Name ="Label45"
                                            Caption ="Agent Name"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1305
                                            LayoutCachedTop =2220
                                            LayoutCachedWidth =2385
                                            LayoutCachedHeight =2505
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2400
                                    Top =2628
                                    Width =2865
                                    Height =300
                                    ColumnWidth =3600
                                    TabIndex =3
                                    Name ="Address1"
                                    ControlSource ="Address"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =2628
                                    LayoutCachedWidth =5265
                                    LayoutCachedHeight =2928
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1575
                                            Top =2628
                                            Width =810
                                            Height =285
                                            Name ="Label23"
                                            Caption ="Address1:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1575
                                            LayoutCachedTop =2628
                                            LayoutCachedWidth =2385
                                            LayoutCachedHeight =2913
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =4725
                                    Top =3180
                                    Width =645
                                    Height =285
                                    TabIndex =4
                                    Name ="Zip"
                                    ControlSource ="Zip"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =4725
                                    LayoutCachedTop =3180
                                    LayoutCachedWidth =5370
                                    LayoutCachedHeight =3465
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =4320
                                            Top =3185
                                            Width =360
                                            Height =285
                                            Name ="Label27"
                                            Caption ="Zip:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =4320
                                            LayoutCachedTop =3185
                                            LayoutCachedWidth =4680
                                            LayoutCachedHeight =3470
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =3895
                                    Top =3190
                                    Width =285
                                    Height =285
                                    TabIndex =5
                                    Name ="St"
                                    ControlSource ="St"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =3895
                                    LayoutCachedTop =3190
                                    LayoutCachedWidth =4180
                                    LayoutCachedHeight =3475
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =3570
                                            Top =3195
                                            Width =285
                                            Height =285
                                            Name ="Label26"
                                            Caption ="St:"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =3570
                                            LayoutCachedTop =3195
                                            LayoutCachedWidth =3855
                                            LayoutCachedHeight =3480
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2400
                                    Top =3200
                                    Width =1125
                                    Height =315
                                    TabIndex =6
                                    Name ="City"
                                    ControlSource ="City"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =3200
                                    LayoutCachedWidth =3525
                                    LayoutCachedHeight =3515
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1935
                                            Top =3190
                                            Width =450
                                            Height =285
                                            Name ="Label25"
                                            Caption ="City"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1935
                                            LayoutCachedTop =3190
                                            LayoutCachedWidth =2385
                                            LayoutCachedHeight =3475
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2400
                                    Top =3610
                                    Width =2865
                                    Height =300
                                    TabIndex =7
                                    Name ="Phone1"
                                    ControlSource ="Phone"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =3610
                                    LayoutCachedWidth =5265
                                    LayoutCachedHeight =3910
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1710
                                            Top =3610
                                            Width =675
                                            Height =285
                                            Name ="Label28"
                                            Caption ="Phone"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1710
                                            LayoutCachedTop =3610
                                            LayoutCachedWidth =2385
                                            LayoutCachedHeight =3895
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2400
                                    Top =4020
                                    Width =2865
                                    Height =300
                                    TabIndex =8
                                    Name ="Phone2"
                                    ControlSource ="Fax/Modem"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =4020
                                    LayoutCachedWidth =5265
                                    LayoutCachedHeight =4320
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1445
                                            Top =4020
                                            Width =945
                                            Height =285
                                            Name ="Label29"
                                            Caption ="Fax"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1445
                                            LayoutCachedTop =4020
                                            LayoutCachedWidth =2390
                                            LayoutCachedHeight =4305
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2400
                                    Top =4480
                                    Width =230
                                    Height =210
                                    TabIndex =9
                                    Name ="PrintUpcOnPo"
                                    ControlSource ="ActiveForTransfer"

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =4480
                                    LayoutCachedWidth =2630
                                    LayoutCachedHeight =4690
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2625
                                            Top =4450
                                            Width =3150
                                            Height =285
                                            Name ="Label33"
                                            Caption ="Active for transfer"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =2625
                                            LayoutCachedTop =4450
                                            LayoutCachedWidth =5775
                                            LayoutCachedHeight =4735
                                        End
                                    End
                                End
                                Begin CheckBox
                                    OverlapFlags =247
                                    Left =2405
                                    Top =4830
                                    Width =230
                                    Height =210
                                    TabIndex =10
                                    Name ="KMDScustomer"
                                    ControlSource ="KMDScustomer"

                                    LayoutCachedLeft =2405
                                    LayoutCachedTop =4830
                                    LayoutCachedWidth =2635
                                    LayoutCachedHeight =5040
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            Left =2600
                                            Top =4800
                                            Width =3180
                                            Height =285
                                            Name ="Label50"
                                            Caption ="KMDS Customer (automate store transfers!)"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =2600
                                            LayoutCachedTop =4800
                                            LayoutCachedWidth =5780
                                            LayoutCachedHeight =5085
                                        End
                                    End
                                End
                                Begin TextBox
                                    OverlapFlags =247
                                    Left =2400
                                    Top =5220
                                    Width =2835
                                    Height =300
                                    TabIndex =11
                                    Name ="EMail"
                                    ControlSource ="EMail"
                                    FontName ="Trebuchet MS"
                                    ConditionalFormat = Begin
                                        0x0100000064000000010000000200000000000000000000000100000001000000 ,
                                        0x0000000000ff0000000000000000000000000000000000000000000000000000 ,
                                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                                        0x00000000
                                    End

                                    LayoutCachedLeft =2400
                                    LayoutCachedTop =5220
                                    LayoutCachedWidth =5235
                                    LayoutCachedHeight =5520
                                    ConditionalFormat14 = Begin
                                        0x0100010000000200000000000000010000000000000000ff0000000000000000 ,
                                        0x00000000000000000000000000000000000000
                                    End
                                    Begin
                                        Begin Label
                                            OverlapFlags =247
                                            TextAlign =3
                                            Left =1140
                                            Top =5220
                                            Width =1245
                                            Height =285
                                            Name ="Label52"
                                            Caption ="E-Mail Address"
                                            FontName ="Trebuchet MS"
                                            LayoutCachedLeft =1140
                                            LayoutCachedTop =5220
                                            LayoutCachedWidth =2385
                                            LayoutCachedHeight =5505
                                        End
                                    End
                                End
                                Begin Label
                                    FontUnderline = NotDefault
                                    OverlapFlags =247
                                    TextAlign =2
                                    Left =3780
                                    Top =1380
                                    Width =5820
                                    Height =300
                                    FontSize =10
                                    ForeColor =1279872587
                                    Name ="lblMapQuest"
                                    Caption ="Click here for &Google Map to Corvallis North"
                                    Tag ="http://maps.google.com/maps"
                                    ControlTipText ="http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q=935+NW+Circl"
                                        "e+Blvd,+Corvallis,+OR+97330&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
                                    HyperlinkAddress ="http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q=935+NW+Circl"
                                        "e+Blvd,+Corvallis,+OR+97330&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
                                    LayoutCachedLeft =3780
                                    LayoutCachedTop =1380
                                    LayoutCachedWidth =9600
                                    LayoutCachedHeight =1680
                                End
                                Begin CommandButton
                                    OverlapFlags =247
                                    Left =7380
                                    Top =4620
                                    Width =2205
                                    Height =330
                                    TabIndex =12
                                    Name ="cmdAgentsFromOLCCAgents"
                                    Caption ="Refresh Agent List from Web"
                                    OnClick ="[Event Procedure]"

                                    LayoutCachedLeft =7380
                                    LayoutCachedTop =4620
                                    LayoutCachedWidth =9585
                                    LayoutCachedHeight =4950
                                    Overlaps =1
                                End
                            End
                        End
                        Begin Page
                            OverlapFlags =247
                            Left =435
                            Top =1305
                            Width =9210
                            Height =4215
                            Name ="Page57"
                            Caption ="Previous Transfers"
                            LayoutCachedLeft =435
                            LayoutCachedTop =1305
                            LayoutCachedWidth =9645
                            LayoutCachedHeight =5520
                            Begin
                                Begin Subform
                                    OverlapFlags =247
                                    Left =600
                                    Top =1380
                                    Width =8940
                                    Height =4020
                                    Name ="AgentsPreviousOrdersSub"
                                    SourceObject ="Form.AgentsPreviousOrdersSub"
                                    LinkChildFields ="Store Number"
                                    LinkMasterFields ="AgentNumber"

                                    LayoutCachedLeft =600
                                    LayoutCachedTop =1380
                                    LayoutCachedWidth =9540
                                    LayoutCachedHeight =5400
                                End
                            End
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =255
                    Left =6900
                    Top =2520
                    Width =2220
                    Height =315
                    TabIndex =7
                    ForeColor =-2147483630
                    Name ="cmdActivate"
                    Caption ="Activate Agents for Transfer"
                    OnClick ="[Event Procedure]"
                    FontName ="Tahoma"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =6900
                    LayoutCachedTop =2520
                    LayoutCachedWidth =9120
                    LayoutCachedHeight =2835
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =255
                            Left =6900
                            Top =1860
                            Width =2160
                            Height =600
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label67"
                            Caption ="Click on the button below to activate agents for transfer"
                            FontName ="Tahoma"
                            LayoutCachedLeft =6900
                            LayoutCachedTop =1860
                            LayoutCachedWidth =9060
                            LayoutCachedHeight =2460
                        End
                    End
                End
            End
        End
        Begin FormFooter
            Height =1020
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =69
                    Left =3960
                    Top =60
                    Width =666
                    Height =606
                    Name ="cmdfrmDeleteRecord"
                    Caption ="Delete"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddada177adada77da1dad1177adad17ad11da7117dad71ada ,
                        0x111da1177d117dad1111d7117711dada11111d11111dadad1111da71117adada ,
                        0x111d77111177adad11d711da71177ada1dadadada71177addadadadadad11ada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Delete Record"
                    UnicodeAccessKey =101
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =3960
                    LayoutCachedTop =60
                    LayoutCachedWidth =4626
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =3998
                            Top =660
                            Width =555
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =3998
                            LayoutCachedTop =660
                            LayoutCachedWidth =4553
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =86
                    Left =2880
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =1
                    Name ="cmdfrmSaveRecord"
                    Caption ="Command34"
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
                        0xffff0000ffffff00070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8070707f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8f8 ,
                        0xf8f8f8f8f8f8f8f8070700000000000000000000000000000000000000000000 ,
                        0x0000000000f8f8f8070007070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800070700f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f800000000f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700f8f8f8f8f8f8f8f8f8f8f8f8000707 ,
                        0x070007070700f8f8000707070707070700000000000000000000000000000000 ,
                        0x000007070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070707070707070707070707070707070707070707 ,
                        0x070707070700f8f8000707070000000000000000000000000000000000000000 ,
                        0x000007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007000700f8f80007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff0007070700f8070007070700ffffffffffffffffffffffffffffffffffffff ,
                        0xff00070707000707070000000000000000000000000000000000000000000000 ,
                        0x0000000000070707
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Save Changes to Record"
                    UnicodeAccessKey =118
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =2880
                    LayoutCachedTop =60
                    LayoutCachedWidth =3546
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =2947
                            Top =660
                            Width =420
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblSaveRecord"
                            Caption ="Sa&ve"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =2947
                            LayoutCachedTop =660
                            LayoutCachedWidth =3367
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =78
                    Left =1800
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =2
                    Name ="cmdfrmUndo"
                    Caption ="Command35"
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
                        0x0707070707070707070707070707070707070707070707070707070707070700 ,
                        0xfb01fb01000707070707070707070707070707070707070707070707070700fb ,
                        0x01fb01000107070707070707070707070707070707070707070707070700fb01 ,
                        0xfb010001fb070707070707070707070707070707070707070707070700fb01fb ,
                        0x010001fb010707070707070707070707070707070707070707070700fb01fb01 ,
                        0x0001fb01fb07070707070707070707070707070707070707070700fb01fb0100 ,
                        0x01fb01fb00070707070707070707070707070707070707070700fb01fb010001 ,
                        0xfb01fb00fb0707070707070707070707070707070707070700fb01fb010001fb ,
                        0x01fb00fb0007070707070707070707070707070707070700000000010001fb01 ,
                        0xfb00fb0007070707070707070707070707070707070700020202000001fb01fb ,
                        0x00fb00070707070707070707070707070707070707000000000202020001fb00 ,
                        0xfb0007070707070707070707070707070707070700fb01fb00000202000000fb ,
                        0x0007070707070707070707070707070707070700fb01fb01fb0100020200fb00 ,
                        0x07070707070707070707070707070707070700fb01fb01fb01fb000002000007 ,
                        0x07070707070707070707070707070707070000000001fb01fb01fb0002000707 ,
                        0x0707070707070707070707070707070700020202000001fb01fb010000070707 ,
                        0x070707070707070707070707070707000000000202020001fb01fb0007070707 ,
                        0x07070707070707070707070707070007fd0700000202000001fb000707070707 ,
                        0x070707070707070707070707070007fd07fd07fd00020200fb00070707070707 ,
                        0x070707070707070707070707000000000007fd07000002000007070707070707 ,
                        0x07070707070707070707070700fdfffdff0007fd070002000707070707070707 ,
                        0x070707070707070707070700fdfffdfffd00fd07fd0000070707070707070707 ,
                        0x070707070707070707070700fffdfffdff0000fd070007070707070707070707 ,
                        0x070707070707070707070700fdfffdfffdff0007000707070707070707070707 ,
                        0x070707070707070707070700fffdfffdfffd0000070707070707070707070707 ,
                        0x07070707070707070707070700fffdfffdff0007070707070707070707070707 ,
                        0x07070707070707070707070700fdfffdff000707070707070707070707070707 ,
                        0x0707070707070707070707070700000000070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707070707070707070707070707070707070707070707070707 ,
                        0x0707070707070707
                    End
                    FontName ="Trebuchet MS"
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Undo Changes Made to This Record"
                    UnicodeAccessKey =110
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =1800
                    LayoutCachedTop =60
                    LayoutCachedWidth =2466
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =1867
                            Top =660
                            Width =450
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblUndo"
                            Caption ="U&ndo"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =1867
                            LayoutCachedTop =660
                            LayoutCachedWidth =2317
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =77
                    Left =7200
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =3
                    Name ="Command0"
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
                    UnicodeAccessKey =109
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =7200
                    LayoutCachedTop =60
                    LayoutCachedWidth =7866
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =7200
                            Top =660
                            Width =885
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label1"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =7200
                            LayoutCachedTop =660
                            LayoutCachedWidth =8085
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    AccessKey =65
                    Left =720
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =4
                    Name ="cmdfrmAddNew"
                    Caption ="Command67"
                    OnClick ="[Event Procedure]"
                    PictureData = Begin
                        0x2800000010000000100000000100040000000000800000000000000000000000 ,
                        0x0000000000000000000000000000800000800000008080008000000080008000 ,
                        0x8080000080808000c0c0c0000000ff00c0c0c00000ffff00ff000000c0c0c000 ,
                        0xffff0000ffffff00dadadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadaddcdadadadadadadaaccdadadada4adaddcccdadad4d4d4da ,
                        0xaccccdadad444daddcccccda4444444aaccccdadad444daddcccdadad4d4d4da ,
                        0xaccdadadada4adaddcdadadadadadadaadadadadadadadaddadadadadadadada ,
                        0xadadadadadadadad
                    End
                    ObjectPalette = Begin
                        0x000301000000000000000000
                    End
                    ControlTipText ="Add a New Record"
                    UnicodeAccessKey =65
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =720
                    LayoutCachedTop =60
                    LayoutCachedWidth =1386
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =728
                            Top =660
                            Width =705
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label79"
                            Caption ="&Add New"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =728
                            LayoutCachedTop =660
                            LayoutCachedWidth =1433
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =5040
                    Top =60
                    Width =666
                    Height =606
                    FontSize =9
                    TabIndex =5
                    ForeColor =-2147483630
                    Name ="cmdPrintPreview"
                    Caption ="Show All Agents"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000030507050203040ff ,
                        0x203850500000000000000000b0a090ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff504840ff305060ff4078c0ff ,
                        0x304860ff3050605000000000b0a090ffffffffffb0a090ffb0a090ffb0a090ff ,
                        0xb0a090ffb0a090ffb0a090ffb0a090ff908880ff305070ff3080d0ff4098e0ff ,
                        0x50b0f0ff506870ff00000000b0a090fffffffffffffffffffff8fffff0f0f0ff ,
                        0xd0d8d0ff909890ff706860ff606860ff505860ff407090ff40a0e0ff60c8ffff ,
                        0x7090a0ff7088905000000000b0a090ffffffffffffffffffffffffffe0e0e0ff ,
                        0x909090ffb0a8a0ffd0c0b0ffd0b0a0ff807870ff505850ff6090b0ff7098b0ff ,
                        0x708890800000000000000000b0a090ffffffffffffffffffffffffffb0b0b0ff ,
                        0xc0b8b0fffff0e0ffffe8e0fff0d8c0fff0d0b0ff807870ff6058609070889080 ,
                        0x000000000000000000000000c0a890ffffffffffffffffffffffffffa09890ff ,
                        0xf0e8e0fffff8f0fffff0f0ffffe8e0fff0d8d0ffd0b0a0ff505850e000000000 ,
                        0x000000000000000000000000c0a8a0ffffffffffffffffffffffffffa0a090ff ,
                        0xf0e8e0fffffffffffff8f0fffff0f0ffffe8e0ffe0c0b0ff606060e000000000 ,
                        0x000000000000000000000000c0b0a0ffffffffffffffffffffffffffc0c8c0ff ,
                        0xc0c0c0fffffffffffffffffffff8f0fffff0e0ffb0a090ff6060608000000000 ,
                        0x000000000000000000000000d0b0a0fffffffffffffffffffffffffff0f8ffff ,
                        0xc0b8b0ffc0c0c0fff0e8e0fff0e8e0ffb0b0a0ff707060ff8070601000000000 ,
                        0x000000000000000000000000d0b8a0ffffffffffffffffffffffffffffffffff ,
                        0xf0f8ffffc0c8c0ffa0a090ff909080ff909090ff605040ff0000000000000000 ,
                        0x000000000000000000000000d0b8b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffb0a090ff604830ff604830ff604830ff0000000000000000 ,
                        0x000000000000000000000000d0c0b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc0a890ffd0c8c0ff604830ffd0b0a0900000000000000000 ,
                        0x000000000000000000000000e0c0b0ffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffc0a8a0ff604830ffd0b0a090000000000000000000000000 ,
                        0x000000000000000000000000e0c0b0ffe0c0b0ffe0c0b0ffe0c0b0ffe0c0b0ff ,
                        0xd0c0b0ffd0b8b0ffd0b0a0ffd0b0a09000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =5040
                    LayoutCachedTop =60
                    LayoutCachedWidth =5706
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =2
                            Left =5040
                            Top =660
                            Width =1095
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label61"
                            Caption ="Print Preview"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =5040
                            LayoutCachedTop =660
                            LayoutCachedWidth =6135
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =95
                    Left =6120
                    Top =60
                    Width =666
                    Height =606
                    TabIndex =6
                    Name ="cmdDataSheet"
                    Caption ="Command59"
                    OnClick ="[Event Procedure]"
                    ControlTipText ="Open Form"
                    ImageData = Begin
                        0x2800000010000000100000000100200000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000806860ff604830ff604830ff604830ff604830ff604830ff ,
                        0x604830ff604830ff604830ff604830ff604830ff604830ff604830ff604830ff ,
                        0x0000000000000000907860ffd0c8c0ff805840fff0f0f0ffe0d8d0ffd0d0d0ff ,
                        0xd0c8c0ffd0c0c0ffc0b8b0ffc0b0b0ffb0a8a0ffb0a8a0ffa09890ff604830ff ,
                        0x0000000000000000a08070ffd0c8d0ff805840fff0f0f0ffb09080ffa08070ff ,
                        0xe0e0d0ff907050ff906850ffd0c8c0ff906850ff906850ffb09890ff604830ff ,
                        0x0000000000000000a09080ffd0d0d0ff805840fffff8f0fff0f0f0fff0f0f0ff ,
                        0xf0e8e0ffe0e0e0ffe0d8d0ffe0d0d0ffd0c8c0ffd0c0b0ffb0a8a0ff604830ff ,
                        0x0000000000000000b0a090ffd0d0d0ff906040fffff8ffffc0a8a0ffc09890ff ,
                        0xf0f0f0ffa08070ffa07860ffe0d8d0ff906850ff906850ffc0b0b0ff604830ff ,
                        0x0000000000000000c0a8a0ffe0d8e0ff906850fffffffffffff8fffffff8f0ff ,
                        0xf0f0f0fff0f0f0fff0e8e0ffe0e0e0ffe0d8d0ffe0d8d0ffc0b8b0ff604830ff ,
                        0x0000000000000000c0b0a0ffe0e0e0ffa07050ffffffffffc0a8a0ffc0a090ff ,
                        0xfff8f0ffc09890ffb09080fff0e8e0ffa08060ffa07860ffd0c8c0ff604830ff ,
                        0x0000000000000000c0b8a0ffe0e0e0ffa07860ffffffffffffffffffffffffff ,
                        0xfff8fffffff8f0fff0f0f0fff0f0f0fff0e8e0ffe0e0e0ffd0d0d0ff604830ff ,
                        0x0000000000000000d0b8b0fff0e8f0ffb08060ffffffffffc0a890ffc0a090ff ,
                        0xffffffffc0a090ffc0a090fff0f0f0ffb09080ffb09080ffe0d8d0ff604830ff ,
                        0x0000000000000000d0c0b0fff0f0f0ffb08870ffffffffffffffffffffffffff ,
                        0xfffffffffffffffffff8fffffff8f0fff0f0f0fff0f0f0ffe0e0e0ff604830ff ,
                        0x0000000000000000d0c0b0ffc09080ffb09070ffb08870ffa08060ffa07860ff ,
                        0x907050ff906850ff906040ff806040ff805840ff805840ff805840ff604830ff ,
                        0x0000000000000000d0c0b0fff0f0f0ffc09080fff0f0f0fff0e8f0ffe0e8e0ff ,
                        0xe0e0e0ffe0e0e0ffe0d8e0ffd0d8d0ffd0d0d0ffd0d0d0ffd0d0d0ff604830ff ,
                        0x0000000000000000d0c0b0ffd0c0b0ffd0c0b0ffd0c0b0ffd0b8b0ffc0b0a0ff ,
                        0xc0b0a0ffc0a8a0ffb0a090ffb09890ffa09080ffa08870ff908070ff907860ff ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End

                    LayoutCachedLeft =6120
                    LayoutCachedTop =60
                    LayoutCachedWidth =6786
                    LayoutCachedHeight =666
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =2
                            Left =6120
                            Top =660
                            Width =915
                            Height =285
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label60"
                            Caption ="Datasheet"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =6120
                            LayoutCachedTop =660
                            LayoutCachedWidth =7035
                            LayoutCachedHeight =945
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =215
                    Left =120
                    Width =9840
                    Height =960
                    Name ="Box81"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =9960
                    LayoutCachedHeight =960
                End
            End
        End
    End
End
CodeBehindForm
' See "Agents.cls"
