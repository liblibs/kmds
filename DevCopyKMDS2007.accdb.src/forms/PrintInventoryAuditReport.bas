Version =20
VersionRequired =20
Begin Form
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    DividingLines = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    ScrollBars =0
    TabularFamily =0
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    DatasheetFontHeight =10
    ItemSuffix =21
    Left =5505
    Top =1860
    Right =12450
    Bottom =5490
    HelpContextId =4700
    RecSrcDt = Begin
        0x23cb48681da6e240
    End
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
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
        Begin CheckBox
            SpecialEffect =2
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin ComboBox
            SpecialEffect =2
            BorderLineStyle =0
            FontName ="Tahoma"
        End
        Begin CustomControl
            SpecialEffect =2
        End
        Begin Section
            Height =3900
            BackColor =-2147483633
            Name ="Detail"
            Begin
                Begin Label
                    OverlapFlags =85
                    Left =240
                    Top =120
                    Width =5040
                    Height =420
                    Name ="Label0"
                    Caption ="Print Inventory  "
                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =1
                    Left =240
                    Top =660
                    Width =5040
                    Height =420
                    Name ="Label1"
                    Caption ="Select range of inventory to print.."
                End
                Begin ComboBox
                    OverlapFlags =93
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2310
                    Left =2760
                    Top =1440
                    Width =1380
                    Height =330
                    Name ="cboSalesDepartmentIDBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [SalesDepartment].[ID], [SalesDepartment].[Description] FROM SalesDepartm"
                        "ent ORDER BY [ID]; "
                    ColumnWidths ="360;1950"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =630
                            Top =1440
                            Width =2055
                            Height =240
                            Name ="From Sales Department Number_Label"
                            Caption ="Range of Sales Department"
                            EventProcPrefix ="From_Sales_Department_Number_Label"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =87
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =2310
                    Left =4260
                    Top =1440
                    Width =1380
                    Height =330
                    TabIndex =1
                    Name ="cboSalesDepartmentIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [SalesDepartment].[ID], [SalesDepartment].[Description] FROM SalesDepartm"
                        "ent ORDER BY [ID]; "
                    ColumnWidths ="360;1950"
                    DefaultValue ="99"

                    Begin
                        Begin Label
                            OverlapFlags =93
                            TextAlign =3
                            Left =4920
                            Top =1200
                            Width =255
                            Height =240
                            Name ="Label6"
                            Caption ="To"
                        End
                    End
                End
                Begin CommandButton
                    OverlapFlags =85
                    Left =3360
                    Top =3420
                    Width =1455
                    Height =405
                    TabIndex =2
                    Name ="cmdReport"
                    Caption ="Generate Report"
                    OnClick ="[Event Procedure]"

                End
                Begin Label
                    OverlapFlags =87
                    TextAlign =3
                    Left =3240
                    Top =1200
                    Width =435
                    Height =240
                    Name ="Label12"
                    Caption ="From"
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1785
                    Left =2820
                    Top =1920
                    Width =1380
                    Height =330
                    TabIndex =3
                    Name ="cboCashiersIDBegin"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Cashier].[ID], [Cashier].[Name] FROM Cashier ORDER BY [ID]; "
                    ColumnWidths ="345;1440"
                    DefaultValue ="1"

                    Begin
                        Begin Label
                            OverlapFlags =85
                            TextAlign =3
                            Left =1380
                            Top =1920
                            Width =1320
                            Height =240
                            Name ="Cashier_Label"
                            Caption ="Cashiers"
                        End
                    End
                End
                Begin ComboBox
                    OverlapFlags =85
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1785
                    Left =4320
                    Top =1920
                    Width =1380
                    Height =330
                    TabIndex =4
                    Name ="cboCashiersIDEnd"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT [Cashier].[ID], [Cashier].[Name] FROM Cashier ORDER BY [ID]; "
                    ColumnWidths ="345;1440"
                    DefaultValue ="99"

                End
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =2820
                    Top =2400
                    Width =1380
                    Height =330
                    AutoActivate =1
                    TabIndex =5
                    Name ="dtpDateBegin"
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
                        0xf8754da1000000000000000000000000902b3c601c66c4010700000040010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000200000090000000 ,
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
                        0xf8754da1000000000000000000000000d0a43fab9f3cc7010500000040010000 ,
                        0x0000000003004f006c0065004f0062006a006500630074004400610074006100 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000001e000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000b8000000 ,
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
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x000000009e1bdd20c487d1118be30000f8754da1214334120800000082090000 ,
                        0x4602000041287f0e0000060020000000d70701000600140006001f000e000000 ,
                        0x0f270c0005001f00000000000000000041060100010001000000000000000000 ,
                        0x080000000000000001efcdab00000500b0e0130006000000ffffffffffffffff ,
                        0x00000000000000001fdeecbd01000500dde013000352e30b918fce119de300aa ,
                        0x004bb851010000009001444201000d4d532053616e7320536572696600000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000002000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OnUpdated ="[Event Procedure]"
                    OLEClass ="DTPicker"
                    Class ="MSComCtl2.DTPicker.2"

                End
                Begin Label
                    OverlapFlags =85
                    TextAlign =3
                    Left =1380
                    Top =2460
                    Width =1320
                    Height =240
                    Name ="Label18"
                    Caption ="Dates"
                End
                Begin CustomControl
                    Enabled = NotDefault
                    SizeMode =1
                    SpecialEffect =0
                    OverlapFlags =85
                    Left =4320
                    Top =2400
                    Width =1380
                    Height =330
                    AutoActivate =1
                    TabIndex =6
                    Name ="dtpDateEnd"
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
                        0xf8754da100000000000000000000000010b51a7a1a66c4010700000000010000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000200000080000000 ,
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
                        0xf8754da1000000000000000000000000d01542ab9f3cc7010500000040010000 ,
                        0x0000000043006f006e00740065006e0074007300000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000000000000000000000000000 ,
                        0x0000000012000201ffffffff02000000ffffffff000000000000000000000000 ,
                        0x00000000000000000000000000000000000000000000000002000000a8000000 ,
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
                        0xf4f1130006000000ffffffffffffffff00000000000000001fdeecbd01000500 ,
                        0x20f213002143341208000000820900004602000041287f0e0000060020000000 ,
                        0xd70701000000150006001f000e0000000f270c0005001f000000000000000000 ,
                        0x41060100010001000000000000000000080000000000000001efcdab00000500 ,
                        0xd4e0130006000000ffffffffffffffff00000000000000001fdeecbd01000500 ,
                        0x01e113000352e30b918fce119de300aa004bb851010000009001444201000d4d ,
                        0x532053616e732053657269663c0000000000000000000000ffffffff00000000 ,
                        0x00000000ffffffff000000000000000000000000000000000000000050010000 ,
                        0x18f0b812af17741b00000000000000001f000000000002005c8d30094c8f3009 ,
                        0x00000000e09e7a1bf09e7a1b0000000000000000000000000000000000000000 ,
                        0x0000000000000000000000000000000000000000f8000000606db712af17741b ,
                        0x000000000000000005000000000005005c8d30094c8f300900000000009f7a1b ,
                        0x189f7a1b00000000000000000000000000000000000000000000000000000000 ,
                        0x0000000000000000
                    End
                    OLEClass ="DTPicker"
                    Class ="MSComCtl2.DTPicker.2"

                End
                Begin CommandButton
                    Cancel = NotDefault
                    OverlapFlags =85
                    Left =5400
                    Top =3060
                    Width =576
                    Height =576
                    TabIndex =7
                    Name ="cmdClose"
                    Caption ="Command20"
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

                End
            End
        End
    End
End
CodeBehindForm
' See "PrintInventoryAuditReport.cls"
