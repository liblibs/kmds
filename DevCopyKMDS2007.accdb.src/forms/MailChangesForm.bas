Version =20
VersionRequired =20
Begin Form
    PopUp = NotDefault
    RecordSelectors = NotDefault
    MaxButton = NotDefault
    MinButton = NotDefault
    ControlBox = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =119
    BorderStyle =3
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =11400
    DatasheetFontHeight =10
    ItemSuffix =17
    Left =2985
    Top =2685
    Right =15735
    Bottom =8520
    RecSrcDt = Begin
        0x97cb95dbfbaee240
    End
    RecordSource ="Sysfile"
    Caption ="Cash register database update"
    DatasheetFontName ="Trebuchet MS"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Tahoma"
        End
        Begin Rectangle
            SpecialEffect =3
            BackStyle =0
            BorderLineStyle =0
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
        Begin ListBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin ToggleButton
            FontSize =8
            FontWeight =400
            ForeColor =-2147483630
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Tab
            FontName ="Tahoma"
            BorderLineStyle =0
        End
        Begin Section
            Height =4740
            BackColor =-2147483633
            Name ="Detail"
            OnClick ="[Event Procedure]"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =120
                    Top =180
                    Width =5325
                    Height =330
                    FontSize =12
                    Name ="Label0"
                    Caption ="Sending Database changes to the cash registers..."
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =93
                    TextAlign =1
                    Left =60
                    Top =900
                    Width =11220
                    Height =840
                    FontSize =14
                    FontWeight =700
                    BackColor =13434828
                    ForeColor =16711680
                    Name ="lblProcess"
                    Caption ="s"
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =85
                    AccessKey =77
                    Left =9540
                    Top =3600
                    Width =666
                    Height =606
                    ForeColor =0
                    Name ="cmdfrmClose"
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
                    UnicodeAccessKey =109

                    LayoutCachedLeft =9540
                    LayoutCachedTop =3600
                    LayoutCachedWidth =10206
                    LayoutCachedHeight =4206
                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =2
                            Left =10252
                            Top =3600
                            Width =510
                            Height =525
                            Name ="Label65"
                            Caption ="Close For&m"
                            FontName ="Trebuchet MS"
                            LayoutCachedLeft =10252
                            LayoutCachedTop =3600
                            LayoutCachedWidth =10762
                            LayoutCachedHeight =4125
                        End
                    End
                End
                Begin Label
                    BackStyle =1
                    OldBorderStyle =1
                    OverlapFlags =95
                    TextAlign =2
                    Left =60
                    Top =660
                    Width =11220
                    Height =240
                    BackColor =65535
                    Name ="Label8"
                    Caption ="Current Process"
                End
                Begin Rectangle
                    SpecialEffect =0
                    BorderWidth =2
                    OverlapFlags =215
                    Top =600
                    Width =11340
                    Height =1200
                    Name ="Box9"
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =2580
                    TabIndex =1
                    Name ="Check17"
                    ControlSource ="Cr1Installed"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =2580
                    LayoutCachedWidth =4820
                    LayoutCachedHeight =2820
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =2580
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Cr1Installed_Label"
                            Caption ="1"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3180
                            LayoutCachedTop =2580
                            LayoutCachedWidth =3345
                            LayoutCachedHeight =2820
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =2940
                    TabIndex =2
                    Name ="Check18"
                    ControlSource ="Cr2Installed"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =2940
                    LayoutCachedWidth =4820
                    LayoutCachedHeight =3180
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =2940
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Cr2Installed_Label"
                            Caption ="2"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3180
                            LayoutCachedTop =2940
                            LayoutCachedWidth =3345
                            LayoutCachedHeight =3180
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =3300
                    TabIndex =3
                    Name ="Check19"
                    ControlSource ="Cr3Installed"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =3300
                    LayoutCachedWidth =4820
                    LayoutCachedHeight =3540
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =3300
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Cr3Installed_Label"
                            Caption ="3"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3180
                            LayoutCachedTop =3300
                            LayoutCachedWidth =3345
                            LayoutCachedHeight =3540
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =3660
                    TabIndex =4
                    Name ="Check20"
                    ControlSource ="Cr4Installed"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =3660
                    LayoutCachedWidth =4820
                    LayoutCachedHeight =3900
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =3660
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Cr4Installed_Label"
                            Caption ="4"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3180
                            LayoutCachedTop =3660
                            LayoutCachedWidth =3345
                            LayoutCachedHeight =3900
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =5760
                    Top =2580
                    TabIndex =5
                    Name ="Check21"
                    ControlSource ="Cr1Online"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =2580
                    LayoutCachedWidth =6020
                    LayoutCachedHeight =2820
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =5760
                    Top =2940
                    TabIndex =6
                    Name ="Check22"
                    ControlSource ="Cr2Online"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =2940
                    LayoutCachedWidth =6020
                    LayoutCachedHeight =3180
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =5760
                    Top =3300
                    TabIndex =7
                    Name ="Check23"
                    ControlSource ="Cr3Online"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =3300
                    LayoutCachedWidth =6020
                    LayoutCachedHeight =3540
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =5760
                    Top =3660
                    TabIndex =8
                    Name ="Check24"
                    ControlSource ="Cr4Online"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =3660
                    LayoutCachedWidth =6020
                    LayoutCachedHeight =3900
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6900
                    Top =2580
                    TabIndex =9
                    Name ="Check25"
                    ControlSource ="Cr1CheckOut"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =2580
                    LayoutCachedWidth =7160
                    LayoutCachedHeight =2820
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6900
                    Top =2940
                    TabIndex =10
                    Name ="Check26"
                    ControlSource ="Cr2CheckOut"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =2940
                    LayoutCachedWidth =7160
                    LayoutCachedHeight =3180
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6900
                    Top =3300
                    TabIndex =11
                    Name ="Check27"
                    ControlSource ="Cr3CheckOut"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =3300
                    LayoutCachedWidth =7160
                    LayoutCachedHeight =3540
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6900
                    Top =3660
                    TabIndex =12
                    Name ="Check28"
                    ControlSource ="Cr4CheckOut"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =3660
                    LayoutCachedWidth =7160
                    LayoutCachedHeight =3900
                End
                Begin Label
                    OverlapFlags =85
                    Left =3180
                    Top =2220
                    Width =885
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label95"
                    Caption ="Cash Reg#"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =3180
                    LayoutCachedTop =2220
                    LayoutCachedWidth =4065
                    LayoutCachedHeight =2460
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =4260
                    Top =2220
                    Width =660
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label96"
                    Caption ="Installed"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =4260
                    LayoutCachedTop =2220
                    LayoutCachedWidth =4920
                    LayoutCachedHeight =2460
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =5400
                    Top =2220
                    Width =675
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label97"
                    Caption ="On Line"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =5400
                    LayoutCachedTop =2220
                    LayoutCachedWidth =6075
                    LayoutCachedHeight =2460
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =6360
                    Top =2220
                    Width =990
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label98"
                    Caption ="Checked out"
                    FontName ="MS Sans Serif"
                    LayoutCachedLeft =6360
                    LayoutCachedTop =2220
                    LayoutCachedWidth =7350
                    LayoutCachedHeight =2460
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =3480
                    Top =2580
                    Width =900
                    TabIndex =13
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="cmbCR1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=1)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="1"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =2580
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =2820
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =3480
                    Top =2940
                    Width =900
                    TabIndex =14
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="cmbCR2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=2)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="2"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =2940
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =3180
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =3480
                    Top =3300
                    Width =900
                    TabIndex =15
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="cmbCR3"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=3)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="3"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =3300
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =3540
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =3480
                    Top =3660
                    Width =900
                    TabIndex =16
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="cmbCR4"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=4)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="4"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =3660
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =3900
                End
                Begin Label
                    OverlapFlags =85
                    Left =5640
                    Top =180
                    Width =3540
                    Height =330
                    FontSize =12
                    Name ="Label11"
                    Caption ="Please stand by..."
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =4560
                    Top =4020
                    TabIndex =17
                    Name ="Check12"
                    ControlSource ="Cr5Installed"

                    LayoutCachedLeft =4560
                    LayoutCachedTop =4020
                    LayoutCachedWidth =4820
                    LayoutCachedHeight =4260
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =3180
                            Top =4020
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label13"
                            Caption ="5"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =3180
                            LayoutCachedTop =4020
                            LayoutCachedWidth =3345
                            LayoutCachedHeight =4260
                        End
                    End
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =5760
                    Top =4020
                    TabIndex =18
                    Name ="Check14"
                    ControlSource ="Cr5Online"

                    LayoutCachedLeft =5760
                    LayoutCachedTop =4020
                    LayoutCachedWidth =6020
                    LayoutCachedHeight =4260
                End
                Begin CheckBox
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    Left =6900
                    Top =4020
                    TabIndex =19
                    Name ="Check15"
                    ControlSource ="Cr5CheckOut"

                    LayoutCachedLeft =6900
                    LayoutCachedTop =4020
                    LayoutCachedWidth =7160
                    LayoutCachedHeight =4260
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =3480
                    Top =4020
                    Width =900
                    TabIndex =20
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="Combo16"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=5)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="5"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =3480
                    LayoutCachedTop =4020
                    LayoutCachedWidth =4380
                    LayoutCachedHeight =4260
                End
            End
        End
    End
End
CodeBehindForm
' See "MailChangesForm.cls"
