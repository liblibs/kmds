Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =119
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =6780
    DatasheetFontHeight =10
    ItemSuffix =35
    Top =600
    Right =8565
    Bottom =6855
    RecSrcDt = Begin
        0x61edee242999e240
    End
    RecordSource ="Sysfile"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Trebuchet MS"
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
            Height =4200
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =960
                    Width =3180
                    Name ="txtCashRegisterName"
                    BeforeUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1380
                            Top =960
                            Width =1140
                            Height =240
                            Name ="Label1"
                            Caption ="Cash Register:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2700
                    Top =1440
                    TabIndex =1
                    Name ="NetworkMailDir"
                    ControlSource ="NetworkMailDir"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1260
                            Top =1440
                            Width =1200
                            Height =240
                            Name ="Label2"
                            Caption ="NetworkMailDir:"
                        End
                    End
                End
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =2640
                    Top =600
                    Width =780
                    TabIndex =2
                    Name ="txtCashRegisterNumber"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1380
                            Top =600
                            Width =1140
                            Height =240
                            Name ="Label4"
                            Caption ="Cash Register:"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2700
                    Top =2340
                    TabIndex =3
                    Name ="Cr1Installed"
                    ControlSource ="Cr1Installed"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1320
                            Top =2340
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Cr1Installed_Label"
                            Caption ="1"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2700
                    Top =2700
                    TabIndex =4
                    Name ="Cr2Installed"
                    ControlSource ="Cr2Installed"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1320
                            Top =2700
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Cr2Installed_Label"
                            Caption ="2"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2700
                    Top =3060
                    TabIndex =5
                    Name ="Cr3Installed"
                    ControlSource ="Cr3Installed"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1320
                            Top =3060
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Cr3Installed_Label"
                            Caption ="3"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2700
                    Top =3420
                    TabIndex =6
                    Name ="Cr4Installed"
                    ControlSource ="Cr4Installed"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1320
                            Top =3420
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Cr4Installed_Label"
                            Caption ="4"
                            FontName ="MS Sans Serif"
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3900
                    Top =2340
                    TabIndex =7
                    Name ="Cr1Online"
                    ControlSource ="Cr1Online"

                End
                Begin CheckBox
                    OverlapFlags =93
                    Left =3900
                    Top =2700
                    TabIndex =8
                    Name ="Check22"
                    ControlSource ="Cr2Online"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3900
                    Top =3060
                    TabIndex =9
                    Name ="Cr3Online"
                    ControlSource ="Cr3Online"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3900
                    Top =3420
                    TabIndex =10
                    Name ="Cr4Online"
                    ControlSource ="Cr4Online"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5040
                    Top =2340
                    TabIndex =11
                    Name ="Cr1CheckOut"
                    ControlSource ="Cr1CheckOut"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5040
                    Top =2700
                    TabIndex =12
                    Name ="Cr2CheckOut"
                    ControlSource ="Cr2CheckOut"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5040
                    Top =3060
                    TabIndex =13
                    Name ="Cr3CheckOut"
                    ControlSource ="Cr3CheckOut"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5040
                    Top =3420
                    TabIndex =14
                    Name ="Check28"
                    ControlSource ="Cr4CheckOut"

                End
                Begin Label
                    OverlapFlags =85
                    Left =1320
                    Top =1980
                    Width =885
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label95"
                    Caption ="Cash Reg#"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =2400
                    Top =1980
                    Width =660
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label96"
                    Caption ="Installed"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =3540
                    Top =1980
                    Width =675
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label97"
                    Caption ="On Line"
                    FontName ="MS Sans Serif"
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =4500
                    Top =1980
                    Width =990
                    Height =240
                    BackColor =-2147483633
                    ForeColor =-2147483630
                    Name ="Label98"
                    Caption ="Checked out"
                    FontName ="MS Sans Serif"
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1620
                    Top =2340
                    Width =900
                    TabIndex =15
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="cmbCR1"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=1)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="1"
                    FontName ="MS Sans Serif"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1620
                    Top =2700
                    Width =900
                    TabIndex =16
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="cmbCR2"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=2)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="2"
                    FontName ="MS Sans Serif"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1620
                    Top =3060
                    Width =900
                    TabIndex =17
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="cmbCR3"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=3)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="3"
                    FontName ="MS Sans Serif"

                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1620
                    Top =3420
                    Width =900
                    TabIndex =18
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="cmbCR4"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=4)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="4"
                    FontName ="MS Sans Serif"

                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =2700
                    Top =3780
                    TabIndex =19
                    Name ="Cr5Installed"
                    ControlSource ="Cr5Installed"

                    LayoutCachedLeft =2700
                    LayoutCachedTop =3780
                    LayoutCachedWidth =2960
                    LayoutCachedHeight =4020
                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =1320
                            Top =3780
                            Width =165
                            Height =240
                            BackColor =-2147483633
                            ForeColor =-2147483630
                            Name ="Label30"
                            Caption ="5"
                            FontName ="MS Sans Serif"
                            LayoutCachedLeft =1320
                            LayoutCachedTop =3780
                            LayoutCachedWidth =1485
                            LayoutCachedHeight =4020
                        End
                    End
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =3900
                    Top =3780
                    TabIndex =20
                    Name ="Cr5Online"
                    ControlSource ="Cr5Online"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =3780
                    LayoutCachedWidth =4160
                    LayoutCachedHeight =4020
                End
                Begin CheckBox
                    OverlapFlags =85
                    Left =5040
                    Top =3780
                    TabIndex =21
                    Name ="Cr5CheckOut"
                    ControlSource ="Cr5CheckOut"

                    LayoutCachedLeft =5040
                    LayoutCachedTop =3780
                    LayoutCachedWidth =5300
                    LayoutCachedHeight =4020
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    Enabled = NotDefault
                    Locked = NotDefault
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1440
                    Left =1620
                    Top =3780
                    Width =900
                    TabIndex =22
                    BackColor =8421631
                    ForeColor =-2147483640
                    Name ="Combo33"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT CashSys.ID, CashSys.ComputerName FROM CashSys WHERE (((CashSys.ID)=5)); "
                    ColumnWidths ="0;1440"
                    DefaultValue ="4"
                    FontName ="MS Sans Serif"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =3780
                    LayoutCachedWidth =2520
                    LayoutCachedHeight =4020
                End
                Begin CheckBox
                    OverlapFlags =247
                    Left =3900
                    Top =2700
                    TabIndex =23
                    Name ="Cr2Online"
                    ControlSource ="Cr2Online"

                    LayoutCachedLeft =3900
                    LayoutCachedTop =2700
                    LayoutCachedWidth =4160
                    LayoutCachedHeight =2940
                End
            End
        End
    End
End
CodeBehindForm
' See "CashRegister.cls"
