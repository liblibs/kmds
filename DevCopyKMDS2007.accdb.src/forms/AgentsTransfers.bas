Version =20
VersionRequired =20
Begin Form
    MaxButton = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowAdditions = NotDefault
    FilterOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =2
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =13020
    DatasheetFontHeight =10
    ItemSuffix =122
    Left =3315
    Top =1650
    Right =16665
    Bottom =8865
    Filter ="[ID]=1689"
    ShortcutMenuBar ="KMDSPopup"
    RecSrcDt = Begin
        0xc5a09cf70ae5e240
    End
    RecordSource ="SELECT AgentTransferHead.* FROM AgentTransferHead; "
    Caption ="Agent Transfers KMDS Software © "
    OnCurrent ="[Event Procedure]"
    OnDelete ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnResize ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =255
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
            Visible = NotDefault
            CanGrow = NotDefault
            Height =375
            BackColor =-2147483633
            Name ="FormHeader"
            Begin
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =85
                    Left =1470
                    Top =60
                    Width =1290
                    Height =255
                    ColumnOrder =0
                    Name ="AgentID"
                    ControlSource ="AgentID"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =660
                            Top =60
                            Width =750
                            Height =240
                            Name ="VendorID_Label"
                            Caption ="VendorID"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =5820
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin OptionGroup
                    TabStop = NotDefault
                    BackStyle =1
                    OverlapFlags =93
                    Left =300
                    Top =2460
                    Width =10622
                    Height =845
                    TabIndex =11
                    BackColor =13434828
                    Name ="optSortBy"
                    AfterUpdate ="[Event Procedure]"
                    DefaultValue ="1"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =900
                            Top =2760
                            Width =570
                            Height =240
                            Name ="Label52"
                            Caption ="Sort by"
                            FontName ="Trebuchet MS"
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =1727
                            Top =2820
                            Width =245
                            OptionValue =1
                            Name ="Option54"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =1560
                                    Top =2520
                                    Width =585
                                    Height =285
                                    Name ="Label55"
                                    Caption ="Line Id"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =2541
                            Top =2820
                            Width =245
                            OptionValue =4
                            Name ="Option60"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =2220
                                    Top =2520
                                    Width =750
                                    Height =285
                                    Name ="Label61"
                                    Caption ="Quantity"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =3266
                            Top =2820
                            Width =245
                            OptionValue =2
                            Name ="Option56"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =3060
                                    Top =2520
                                    Width =630
                                    Height =285
                                    Name ="Label57"
                                    Caption ="Item ID"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =4055
                            Top =2820
                            Width =245
                            OptionValue =3
                            Name ="Option58"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =3780
                                    Top =2520
                                    Width =945
                                    Height =285
                                    Name ="Label59"
                                    Caption ="Description"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =5040
                            Top =2820
                            Width =245
                            OptionValue =11
                            Name ="Option113"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    TextAlign =2
                                    Left =4800
                                    Top =2520
                                    Width =690
                                    Height =285
                                    Name ="Label114"
                                    Caption ="Size"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =5813
                            Top =2820
                            Width =245
                            OptionValue =7
                            Name ="Option71"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =5640
                                    Top =2520
                                    Width =465
                                    Height =285
                                    Name ="Label72"
                                    Caption ="Price"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =6467
                            Top =2820
                            Width =245
                            OptionValue =8
                            Name ="Option66"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =6240
                                    Top =2520
                                    Width =765
                                    Height =285
                                    Name ="Label67"
                                    Caption ="Ext Price"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =7136
                            Top =2820
                            Width =245
                            OptionValue =5
                            Name ="Option62"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =7020
                                    Top =2520
                                    Width =420
                                    Height =285
                                    Name ="Label63"
                                    Caption ="Cost"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =7760
                            Top =2820
                            Width =245
                            OptionValue =6
                            Name ="Option64"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =7500
                                    Top =2520
                                    Width =720
                                    Height =285
                                    Name ="Label65"
                                    Caption ="Ext Cost"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =8340
                            Top =2820
                            Width =245
                            OptionValue =10
                            Name ="Option92"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =8280
                                    Top =2520
                                    Width =270
                                    Height =285
                                    Name ="Label93"
                                    Caption ="GP"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                        Begin OptionButton
                            OverlapFlags =87
                            Left =8954
                            Top =2820
                            Width =245
                            OptionValue =9
                            Name ="Option68"

                            Begin
                                Begin Label
                                    OverlapFlags =87
                                    Left =8640
                                    Top =2520
                                    Width =690
                                    Height =285
                                    Name ="Label69"
                                    Caption ="Repacks"
                                    FontName ="Trebuchet MS"
                                End
                            End
                        End
                    End
                End
                Begin Subform
                    OverlapFlags =247
                    Left =300
                    Top =3480
                    Width =12480
                    Height =2150
                    Name ="AgentsTransfersSub"
                    SourceObject ="Form.AgentsTransfersSub"
                    LinkChildFields ="AgentTransferHeadID"
                    LinkMasterFields ="ID"

                    LayoutCachedLeft =300
                    LayoutCachedTop =3480
                    LayoutCachedWidth =12780
                    LayoutCachedHeight =5630
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    ColumnCount =11
                    ListWidth =1920
                    Left =420
                    Top =1080
                    Width =3360
                    TabIndex =1
                    BackColor =8454143
                    Name ="cmbAgentID"
                    ControlSource ="AgentID"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Agents.AgentNumber, Agents.LastName AS Name, Agents.FirstName AS Contact,"
                        " Agents.Address, Agents.City, Agents.ST, Agents.Zip, Agents.Phone FROM Agents OR"
                        "DER BY Agents.AgentNumber; "
                    ColumnWidths ="0;1919"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =1
                            Left =300
                            Top =780
                            Width =1245
                            Height =255
                            Name ="Label74"
                            Caption ="OLCC  Agent"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =95
                    TextAlign =1
                    Left =420
                    Top =1320
                    Width =3300
                    Height =255
                    TabIndex =2
                    BackColor =8454143
                    Name ="txtContact"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    Left =420
                    Top =1560
                    Width =3300
                    Height =255
                    TabIndex =3
                    BackColor =8454143
                    Name ="txtAdd1"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    Left =420
                    Top =1800
                    Width =3300
                    Height =255
                    TabIndex =4
                    BackColor =8454143
                    Name ="txtAdd2"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =1
                    Left =420
                    Top =2040
                    Width =3300
                    Height =255
                    TabIndex =5
                    BackColor =8454143
                    Name ="txtAdd3"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =1
                    Left =3960
                    Top =1800
                    Width =3300
                    Height =255
                    TabIndex =6
                    BackColor =8454143
                    Name ="txtPhone1"
                    FontName ="Trebuchet MS"

                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =0
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =3240
                    Top =360
                    Width =840
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =7
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalQ"
                    ControlSource ="=AgentsTransfersSub.Form!txtSumQuantity"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =3240
                            Top =120
                            Width =840
                            Height =240
                            BackColor =13750737
                            Name ="Label41"
                            Caption ="Quantity"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =1440
                    Top =360
                    Width =1020
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =8
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalCost"
                    ControlSource ="=AgentsTransfersSub.Form!txtSumCost"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =93
                            TextAlign =2
                            Left =1440
                            Top =120
                            Width =1020
                            Height =240
                            BackColor =13750737
                            Name ="Label43"
                            Caption ="Cost"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =223
                    Left =7140
                    Top =3060
                    TabIndex =9
                    Name ="chkDisplayCost"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =223
                    Left =5820
                    Top =3060
                    TabIndex =10
                    Name ="chkDisplayPrice"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =223
                    Left =360
                    Top =3000
                    Width =1125
                    Height =240
                    Name ="lblShowColumns"
                    Caption ="Show Columns"
                    FontName ="Trebuchet MS"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =300
                    Top =780
                    Width =3540
                    Height =1620
                    Name ="boxPurchasedFrom"
                End
                Begin TextBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =5130
                    Top =1200
                    Width =2190
                    Height =255
                    TabIndex =12
                    BackColor =8454143
                    Name ="ID"
                    ControlSource ="ID"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4260
                            Top =1200
                            Width =825
                            Height =240
                            Name ="Order_Label"
                            Caption ="Order No."
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =255
                    TextAlign =2
                    Left =5130
                    Top =1440
                    Width =2190
                    Height =255
                    TabIndex =13
                    BackColor =8454143
                    Name ="OrdDate"
                    ControlSource ="OrdDate"
                    Format ="Short Date"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            OverlapFlags =95
                            TextAlign =3
                            Left =4170
                            Top =1440
                            Width =915
                            Height =240
                            Name ="OrdDate_Label"
                            Caption ="Order Date"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =3840
                    Top =780
                    Width =3540
                    Height =1620
                    Name ="Box35"
                End
                Begin Label
                    OverlapFlags =255
                    Left =3840
                    Top =780
                    Width =870
                    Height =285
                    Name ="Label36"
                    Caption ="Order Info"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    OverlapFlags =93
                    TextAlign =1
                    Left =6300
                    Top =60
                    Width =6300
                    Height =540
                    FontSize =20
                    FontWeight =700
                    BackColor =32768
                    ForeColor =32768
                    Name ="lblTransferINOUT"
                    Caption ="Transfer IN"
                    FontName ="Trebuchet MS"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =12960
                    Height =5700
                    Name ="BoxBody"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =13020
                    LayoutCachedHeight =5760
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    SpecialEffect =0
                    BorderWidth =1
                    OverlapFlags =247
                    TextAlign =2
                    Left =240
                    Top =360
                    Width =1185
                    Height =300
                    FontSize =10
                    FontWeight =700
                    TabIndex =14
                    BackColor =8453888
                    BorderColor =4194368
                    ForeColor =0
                    Name ="txtTotalPrice"
                    ControlSource ="=AgentsTransfersSub.Form!txtSumPrice"
                    Format ="Standard"
                    FontName ="Trebuchet MS"

                    Begin
                        Begin Label
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =255
                            TextAlign =2
                            Left =240
                            Top =120
                            Width =1200
                            Height =240
                            BackColor =13750737
                            Name ="Label79"
                            Caption ="Price"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =4920
                    Top =210
                    Width =185
                    TabIndex =15
                    Name ="Check81"
                    DefaultValue ="False"
                    OnClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            SpecialEffect =2
                            BackStyle =1
                            OldBorderStyle =1
                            OverlapFlags =247
                            TextAlign =1
                            Left =5145
                            Top =180
                            Width =810
                            Height =285
                            BackColor =65280
                            ForeColor =-2147483640
                            Name ="Label82"
                            Caption ="Lines only"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =3360
                    Width =12840
                    Height =2400
                    Name ="boxSubFormNormal"
                    LayoutCachedLeft =60
                    LayoutCachedTop =3360
                    LayoutCachedWidth =12900
                    LayoutCachedHeight =5760
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =1740
                    Top =3060
                    TabIndex =16
                    Name ="chkDisplayLine"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =2520
                    Top =3060
                    TabIndex =17
                    Name ="chkDisplayQuantity"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =3300
                    Top =3060
                    TabIndex =18
                    Name ="chkDisplayItemID"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =4080
                    Top =3060
                    TabIndex =19
                    Name ="chkDisplayDescription"
                    StatusBarText ="d"
                    BeforeUpdate ="r"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =6480
                    Top =3060
                    TabIndex =20
                    Name ="chkDisplayExtPrice"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =7740
                    Top =3060
                    TabIndex =21
                    Name ="chkDisplayExtCost"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =8940
                    Top =3060
                    TabIndex =22
                    Name ="chkDisplayRepack"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =8340
                    Top =3060
                    TabIndex =23
                    Name ="chkDisplayGP"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin CheckBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    Left =3960
                    Top =2130
                    TabIndex =24
                    Name ="chkSupplyAllItems"

                    Begin
                        Begin Label
                            OverlapFlags =247
                            Left =4185
                            Top =2100
                            Width =2370
                            Height =240
                            Name ="Label109"
                            Caption ="This vendor can supply any Item"
                        End
                    End
                End
                Begin CheckBox
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =5040
                    Top =3060
                    TabIndex =25
                    Name ="chkDisplaySize"
                    DefaultValue ="True"
                    OnClick ="[Event Procedure]"

                End
                Begin TextBox
                    Visible = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =247
                    TextAlign =2
                    IMESentenceMode =3
                    Left =6720
                    Top =2100
                    Width =240
                    TabIndex =26
                    BackColor =8454143
                    Name ="Status"
                    ControlSource ="Status"

                End
            End
        End
        Begin FormFooter
            CanGrow = NotDefault
            Height =1020
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Rectangle
                    OverlapFlags =93
                    Left =60
                    Top =60
                    Width =12600
                    Height =660
                    Name ="Box75"
                End
                Begin CommandButton
                    OverlapFlags =215
                    Left =7080
                    Top =180
                    Width =1380
                    Height =405
                    Name ="cmdVendorPOHistory"
                    Caption ="Previous PO's"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =67
                    Left =10680
                    Top =120
                    Width =666
                    Height =606
                    TabIndex =1
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
                            OverlapFlags =215
                            TextAlign =2
                            Left =11400
                            Top =120
                            Width =585
                            Height =405
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label99"
                            Caption ="&Close Form"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =80
                    Left =9120
                    Top =120
                    Width =576
                    Height =576
                    TabIndex =2
                    Name ="cmdPrintDocs"
                    Caption ="cmdPreview"
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
                    UnicodeAccessKey =80

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =9720
                            Top =120
                            Width =780
                            Height =405
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label9"
                            Caption ="&Print / Post"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =215
                    AccessKey =69
                    Left =1080
                    Top =120
                    Width =576
                    Height =576
                    TabIndex =3
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

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =1755
                            Top =300
                            Width =630
                            Height =300
                            BackColor =16777215
                            ForeColor =0
                            Name ="lblDeleteRecord"
                            Caption ="D&elete"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =215
                    AccessKey =82
                    Left =3420
                    Top =120
                    Width =795
                    Height =525
                    TabIndex =4
                    Name ="cmdRefresh"
                    Caption ="Refresh"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =82

                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =215
                            TextAlign =2
                            Left =4305
                            Top =240
                            Width =765
                            Height =300
                            BackColor =16777215
                            ForeColor =0
                            Name ="Label121"
                            Caption ="&Refresh"
                            FontName ="Trebuchet MS"
                        End
                    End
                End
            End
        End
    End
End
CodeBehindForm
' See "AgentsTransfers.cls"
