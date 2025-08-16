Version =20
VersionRequired =20
Begin Form
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12480
    DatasheetFontHeight =10
    ItemSuffix =24
    Top =1020
    Bottom =1020
    Filter ="((TimeClockPunches.PunchDate=#8/30/2005#))"
    RecSrcDt = Begin
        0x9e86a9d276dce240
    End
    RecordSource ="qryPunchDatesSummary"
    OnCurrent ="[Event Procedure]"
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    DatasheetFontName ="Arial"
    OnLoad ="[Event Procedure]"
    FilterOnLoad =0
    ShowPageMargins =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            BackStyle =0
            FontName ="Arial"
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
        Begin Subform
            SpecialEffect =2
            BorderLineStyle =0
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin PageHeader
            DisplayWhen =1
            Height =360
            BackColor =-2147483633
            Name ="PageHeaderSection"
            Begin
                Begin TextBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    Left =1500
                    Top =120
                    Width =1860
                    ColumnWidth =1680
                    Name ="PunchDate"
                    ControlSource ="PunchDate"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            Left =60
                            Top =120
                            Width =915
                            Height =240
                            Name ="Label0"
                            Caption ="PunchDate:"
                            FontName ="Tahoma"
                        End
                    End
                End
            End
        End
        Begin Section
            CanGrow = NotDefault
            Height =6300
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Subform
                    OverlapFlags =247
                    Left =240
                    Top =660
                    Width =10320
                    Height =4140
                    TabIndex =1
                    Name ="TimeClockPunches"
                    SourceObject ="Form.TimeClockPunchesByDateSub"
                    LinkChildFields ="PunchDate"
                    LinkMasterFields ="PunchDate"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            OverlapFlags =93
                            Left =10680
                            Top =900
                            Width =1560
                            Height =2280
                            Name ="Label9"
                            Caption ="1:00 pm =  13:00 2:00 pm =  14:00     3:00 pm =  15:00 4:00 pm =  16:00  5:00 pm"
                                " =  17:00  6:00 pm =  18:00  7:00 pm =  19:00  8:00 pm =  20:00  9:00 pm =  21:0"
                                "0 10:00 pm= 22:00  11:00 pm= 23:00  "
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =93
                    TextAlign =2
                    IMESentenceMode =3
                    Left =3000
                    Top =4980
                    Height =420
                    FontSize =16
                    TabIndex =2
                    Name ="Text2"
                    ControlSource ="=TimeClockPunches.Form!SumOfHoursWorked"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =2040
                            Top =4980
                            Width =900
                            Height =240
                            Name ="Label3"
                            Caption ="Total Hours"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin ComboBox
                    LimitToList = NotDefault
                    OverlapFlags =93
                    IMESentenceMode =3
                    ListWidth =1440
                    Left =1500
                    Top =120
                    Width =1680
                    Height =345
                    FontSize =12
                    Name ="cmbPunchDate"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT qryPunchDatesSummary.PunchDate FROM qryPunchDatesSummary ORDER BY qryPunc"
                        "hDatesSummary.PunchDate DESC; "
                    ColumnWidths ="1440"
                    AfterUpdate ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =360
                            Top =120
                            Width =975
                            Height =245
                            Name ="PunchDate1_Label"
                            Caption ="Select Date:"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin CommandButton
                    TabStop = NotDefault
                    OverlapFlags =93
                    Left =720
                    Top =5040
                    Width =1275
                    Height =405
                    TabIndex =3
                    Name ="cmdAddNew"
                    Caption ="Add New Date"
                    OnClick ="[Event Procedure]"

                    Overlaps =1
                End
                Begin CustomControl
                    Enabled = NotDefault
                    TabStop = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =93
                    Left =5700
                    Top =120
                    Width =1560
                    Height =300
                    AutoActivate =1
                    TabIndex =4
                    Name ="DTPicker0"
                    OleData = Begin
                        0x000e0000d0cf11e0a1b11ae1000000000000000000000000000000003e000300 ,
                        0xfeff090006000000000000000000000001000000020000000000000000100000 ,
                        0x0400000001000000feffffff0000000003000000ffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000009e1bdd20c487d1118be30000 ,
                        0xf8754da100000000000000000000000080deba45a9adc5010700000040010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000bc000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fefffffffdfffffffffffffffffffffffffffffffffffffffeffffff ,
                        0xfeffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff52006f006f007400200045006e007400720079000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000016000500ffffffffffffffff010000009e1bdd20c487d1118be30000 ,
                        0xf8754da1000000000000000000000000b01440b9a9adc5010500000040010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000bc000000 ,
                        0x0000000003004100630063006500730073004f0062006a005300690074006500 ,
                        0x4400610074006100000000000000000000000000000000000000000000000000 ,
                        0x0000000026000200ffffffffffffffffffffffff000000000000000000000000 ,
                        0x000000000000000000000000000000000000000000000000000000005c000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000ffffffffffffffffffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x00000000fffffffffffffffffefffffffdfffffffefffffffeffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff01000000feffffff0300000004000000feffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffffff ,
                        0xffffffff5c000000000000000100000000000000000000000000000000000000 ,
                        0x2400000038000000000000000000000000000000000000000000000036353141 ,
                        0x383934302d383743352d313164312d384245332d303030304638373534444131 ,
                        0x442a8476ffffffff04000000780000002800000000000200ffffffff302a8476 ,
                        0x000000009e1bdd20c487d1118be30000f8754da12143341208000000c00a0000 ,
                        0x1102000041287f0e0000060024000000d5070800020010000000000000000000 ,
                        0x0f270c0005001f00000000000000000041060100010001000000000000000000 ,
                        0x080000000100000001efcdab00000500b4f5130006005200ffffffffffffffff ,
                        0x0000000000000000ffffff001fdeecbd01000500e1f513000352e30b918fce11 ,
                        0x9de300aa004bb851010000009001444201000d4d532053616e73205365726966 ,
                        0xd0298476000000000800000054010000380000000000000000100000c0298476 ,
                        0x010000000400000058010000440000000000000000100000b029847600000000 ,
                        0x040000005c010000480000000000000000100000a42984766400000004000000 ,
                        0x600100004000000000000000001000008c298476000000000400000068010000 ,
                        0x3000000001000000001000007829847600000000040000006801000030000000 ,
                        0x01000000080000005c2984760000000004000000680100003000000001000000 ,
                        0x0001000000000000
                    End
                    OLEClass ="DTPicker"
                    Class ="MSComCtl2.DTPicker.2"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4380
                            Top =120
                            Width =1245
                            Height =240
                            Name ="Label11"
                            Caption ="Display calendar"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    Left =360
                    Top =5520
                    Width =7905
                    Height =300
                    Name ="Label12"
                    Caption ="Please Note:  Punch records are created when a user clicks on the \"Punch Clock\""
                        " button in the cash register"
                End
                Begin TextBox
                    TabStop = NotDefault
                    DecimalPlaces =2
                    OverlapFlags =223
                    TextAlign =2
                    IMESentenceMode =3
                    Left =5700
                    Top =4980
                    Height =420
                    FontSize =16
                    TabIndex =5
                    Name ="TotalWages"
                    ControlSource ="=TimeClockPunches.Form!SumOfWages"
                    Format ="Standard"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            Left =4740
                            Top =4980
                            Width =975
                            Height =240
                            Name ="Label18"
                            Caption ="Total Wages"
                            FontName ="Tahoma"
                        End
                    End
                End
                Begin Label
                    OverlapFlags =93
                    Left =360
                    Top =5940
                    Width =7905
                    Height =300
                    Name ="Label19"
                    Caption ="Time Reports are located in the Menu option D-3.  Use report #36 and 37."
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8460
                    Top =4980
                    Width =3540
                    Height =495
                    TabIndex =6
                    Name ="cmdHourly"
                    Caption ="Goto Print reports Option #36 Hours Worked By Date By Cashier"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =8460
                    LayoutCachedTop =4980
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =5475
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    Left =8460
                    Top =5640
                    Width =3540
                    Height =495
                    TabIndex =7
                    Name ="cmdHourly37"
                    Caption ="Goto Print reports Option #37 Hours Worked By Cashier By Date"
                    OnClick ="[Event Procedure]"
                    ImageData = Begin
                        0x00000000
                    End

                    LayoutCachedLeft =8460
                    LayoutCachedTop =5640
                    LayoutCachedWidth =12000
                    LayoutCachedHeight =6135
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =60
                    Top =60
                    Width =12420
                    Height =6240
                    Name ="Box5"
                    LayoutCachedLeft =60
                    LayoutCachedTop =60
                    LayoutCachedWidth =12480
                    LayoutCachedHeight =6300
                End
            End
        End
        Begin PageFooter
            DisplayWhen =1
            Height =0
            BackColor =-2147483633
            Name ="PageFooterSection"
        End
    End
End
CodeBehindForm
' See "TimeClockPunchesByDate.cls"
