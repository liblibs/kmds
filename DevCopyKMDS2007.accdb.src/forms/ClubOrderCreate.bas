Version =20
VersionRequired =20
Begin Form
    AutoResize = NotDefault
    RecordSelectors = NotDefault
    AutoCenter = NotDefault
    NavigationButtons = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =255
    BorderStyle =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =12720
    DatasheetFontHeight =10
    ItemSuffix =57
    Left =5520
    Top =1410
    Right =18825
    Bottom =9270
    HelpContextId =2100
    RecSrcDt = Begin
        0x78c28751688ae240
    End
    Caption =" KMDS Software ©  KMDS Software ©  KMDS Software ©  KMDS Software © "
    OnOpen ="[Event Procedure]"
    OnClose ="[Event Procedure]"
    HelpFile ="KMDSHelp.chm"
    DatasheetFontName ="Arial"
    OnClick ="[Event Procedure]"
    OnDblClick ="[Event Procedure]"
    OnLoad ="[Event Procedure]"
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
        Begin CustomControl
            SpecialEffect =2
        End
        Begin FormHeader
            Height =0
            BackColor =-2147483633
            Name ="FormHeader"
        End
        Begin Section
            Height =7140
            BackColor =-2147483633
            Name ="Detail"
            OnClick ="[Event Procedure]"
            Begin
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =85
                    Left =240
                    Top =6060
                    Width =2400
                    Height =405
                    TabIndex =3
                    ForeColor =0
                    Name ="cmdUseOrderGuide"
                    Caption ="&Use Order Guide"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =85

                    LayoutCachedLeft =240
                    LayoutCachedTop =6060
                    LayoutCachedWidth =2640
                    LayoutCachedHeight =6465
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =93
                    AccessKey =68
                    Left =240
                    Top =6540
                    Width =2580
                    Height =405
                    TabIndex =4
                    ForeColor =0
                    Name ="cmdNoOrderGuide"
                    Caption ="&Do NOT Use Order Guide (Blank)"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =68

                    LayoutCachedLeft =240
                    LayoutCachedTop =6540
                    LayoutCachedWidth =2820
                    LayoutCachedHeight =6945
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =223
                    Left =120
                    Top =60
                    Width =12600
                    Height =6960
                    Name ="boxScreen"
                End
                Begin ListBox
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =3
                    Left =480
                    Top =1560
                    Width =4920
                    Height =3960
                    FontSize =10
                    BackColor =10092543
                    Name ="lstClubs"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Club.ID, Club.Name, Format([Phone1],\"(@@@) @@@-@@@@\") AS Phone, Trim([A"
                        "ddress1]) & \", \" & Trim([city]) & \", \" & Trim([st]) & \" \" & Trim([zip]) AS"
                        " googleAdd FROM Club WHERE (((Club.NonLiquor)=False)) ORDER BY Club.Name; "
                    ColumnWidths ="0;3240;1440"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =480
                            Top =1080
                            Width =2115
                            Height =435
                            Name ="First Select Club Name_Label"
                            Caption ="To start a new order,  Select Licensee Name."
                            EventProcPrefix ="First_Select_Club_Name_Label"
                        End
                    End
                End
                Begin TextBox
                    TabStop = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =1080
                    Top =5640
                    Width =4020
                    Height =300
                    FontSize =10
                    TabIndex =1
                    BackColor =16777164
                    Name ="txtSelectedClub"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =240
                            Top =5640
                            Width =735
                            Height =435
                            Name ="Label28"
                            Caption ="Selected Club"
                        End
                    End
                End
                Begin ListBox
                    ColumnHeads = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =7
                    Left =5580
                    Top =1560
                    Width =7080
                    Height =3960
                    FontSize =10
                    TabIndex =5
                    BackColor =16777164
                    Name ="lstInvoices"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT OrdHead.ID, IIf(([qryClubOrderLineTotals].[ExtPrice]-[qryClubOrderLineTot"
                        "als].[ExtDiscount])=[OrdDollars],[Club].[Name],\"**** REPRINT **** \" & [cLUB].["
                        "NAME]) AS ClubName, Format([InvDate],\"mm/dd\") AS [Date], OrdHead.InvoiceNo AS "
                        "[Invoice #], qryClubOrderLineTotals.ExtPrice-qryClubOrderLineTotals.ExtDiscount "
                        "AS Amount, OrdHead.PaidFlag, Club.ID FROM (Club INNER JOIN OrdHead ON Club.ID=Or"
                        "dHead.ClubID) INNER JOIN qryClubOrderLineTotals ON OrdHead.ID=qryClubOrderLineTo"
                        "tals.OrdheadID ORDER BY IIf(([qryClubOrderLineTotals].[ExtPrice]-[qryClubOrderLi"
                        "neTotals].[ExtDiscount])=[OrdDollars],[Club].[Name],\"**** REPRINT **** \" & [cL"
                        "UB].[NAME]), Format([InvDate],\"mm/dd/yy\"); "
                    ColumnWidths ="0;3600;720;1080;1800;0;0;0"
                    BeforeUpdate ="[Event Procedure]"
                    AfterUpdate ="[Event Procedure]"
                    OnDblClick ="[Event Procedure]"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            Left =5580
                            Top =1080
                            Width =2220
                            Height =420
                            Name ="Outstanding Orders_Label"
                            Caption ="To Edit an existing order, Select an order below"
                            EventProcPrefix ="Outstanding_Orders_Label"
                        End
                    End
                End
                Begin TextBox
                    Visible = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =223
                    IMESentenceMode =3
                    Left =6480
                    Top =5640
                    Width =4920
                    Height =300
                    FontSize =10
                    TabIndex =2
                    BackColor =16777164
                    Name ="txtSelectedInvoice"

                    Begin
                        Begin Label
                            OverlapFlags =223
                            TextAlign =3
                            Left =5640
                            Top =5640
                            Width =735
                            Height =435
                            Name ="Label34"
                            Caption ="Selected Invoice"
                        End
                    End
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =223
                    AccessKey =69
                    Left =5580
                    Top =6060
                    Width =2040
                    Height =405
                    TabIndex =6
                    ForeColor =0
                    Name ="cmdEditThisOrder"
                    Caption ="&Edit this order"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =69

                    LayoutCachedLeft =5580
                    LayoutCachedTop =6060
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =6465
                    Overlaps =1
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =1020
                    Width =5400
                    Height =6000
                    Name ="Box37"
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =5520
                    Top =1020
                    Width =7200
                    Height =6000
                    Name ="Box38"
                End
                Begin CommandButton
                    Cancel = NotDefault
                    TabStop = NotDefault
                    OverlapFlags =247
                    Left =11940
                    Top =6360
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

                    LayoutCachedLeft =11940
                    LayoutCachedTop =6360
                    LayoutCachedWidth =12606
                    LayoutCachedHeight =6966
                    Overlaps =1
                    Begin
                        Begin Label
                            OverlapFlags =247
                            TextAlign =2
                            Left =11940
                            Top =6000
                            Width =615
                            Height =405
                            Name ="Label65"
                            Caption ="Close Form"
                            LayoutCachedLeft =11940
                            LayoutCachedTop =6000
                            LayoutCachedWidth =12555
                            LayoutCachedHeight =6405
                        End
                    End
                End
                Begin Rectangle
                    OverlapFlags =255
                    Left =120
                    Top =60
                    Width =12600
                    Height =5520
                    Name ="Box39"
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =5580
                    Top =6540
                    Width =2040
                    Height =405
                    TabIndex =8
                    Name ="cmdOrderGuide"
                    Caption ="Print Guide for #5495"
                    OnClick ="[Event Procedure]"
                    Tag ="5495"

                    LayoutCachedLeft =5580
                    LayoutCachedTop =6540
                    LayoutCachedWidth =7620
                    LayoutCachedHeight =6945
                    Overlaps =1
                End
                Begin ComboBox
                    Visible = NotDefault
                    OverlapFlags =247
                    IMESentenceMode =3
                    ColumnCount =2
                    ListWidth =1920
                    Left =3000
                    Top =6660
                    Width =2400
                    TabIndex =9
                    Name ="cboClubClone"
                    RowSourceType ="Table/Query"
                    RowSource ="SELECT Club.ID, Club.Name FROM Club ORDER BY [Name]; "
                    ColumnWidths ="0;1920"
                    AfterUpdate ="[Event Procedure]"

                    LayoutCachedLeft =3000
                    LayoutCachedTop =6660
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =6900
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    AccessKey =71
                    Left =2700
                    Top =6060
                    Width =1500
                    Height =420
                    TabIndex =10
                    ForeColor =0
                    Name ="cmdClone"
                    Caption ="Clone Order &Guide"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =71

                    LayoutCachedLeft =2700
                    LayoutCachedTop =6060
                    LayoutCachedWidth =4200
                    LayoutCachedHeight =6480
                    Overlaps =1
                End
                Begin Label
                    OverlapFlags =247
                    TextAlign =2
                    Left =165
                    Top =120
                    Width =5235
                    Height =960
                    FontSize =36
                    Name ="Label0"
                    Caption ="Licensee Order"
                    FontName ="Trebuchet MS"
                    LayoutCachedLeft =165
                    LayoutCachedTop =120
                    LayoutCachedWidth =5400
                    LayoutCachedHeight =1080
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    AccessKey =68
                    Left =4260
                    Top =6060
                    Width =1200
                    Height =405
                    TabIndex =11
                    ForeColor =0
                    Name ="cmdPDADownload"
                    Caption ="P&DA Download"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    UnicodeAccessKey =68

                    LayoutCachedLeft =4260
                    LayoutCachedTop =6060
                    LayoutCachedWidth =5460
                    LayoutCachedHeight =6465
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =6480
                    Top =360
                    Width =1800
                    Height =420
                    TabIndex =12
                    Name ="cmdEmailOrders"
                    Caption ="Process e-Mails"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =6480
                    LayoutCachedTop =360
                    LayoutCachedWidth =8280
                    LayoutCachedHeight =780
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    AccessKey =72
                    Left =7740
                    Top =6060
                    Width =2520
                    Height =420
                    TabIndex =13
                    ForeColor =0
                    Name ="cmdOGfromHistory"
                    Caption ="Create Order fom Sales &History"
                    OnClick ="[Event Procedure]"
                    FontName ="Trebuchet MS"
                    ControlTipText ="Create Order of EVERYTHING this club has bought."
                    UnicodeAccessKey =72

                    LayoutCachedLeft =7740
                    LayoutCachedTop =6060
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =6480
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    AccessKey =80
                    Left =7740
                    Top =6540
                    Width =2520
                    Height =420
                    TabIndex =14
                    Name ="cmdPrintClubHistory"
                    Caption ="&Purchase History Report"
                    OnClick ="[Event Procedure]"
                    UnicodeAccessKey =80

                    LayoutCachedLeft =7740
                    LayoutCachedTop =6540
                    LayoutCachedWidth =10260
                    LayoutCachedHeight =6960
                    Overlaps =1
                End
                Begin CommandButton
                    Visible = NotDefault
                    OverlapFlags =247
                    Left =9420
                    Top =360
                    Width =1800
                    Height =420
                    TabIndex =15
                    Name ="cmdWebOrders"
                    Caption ="Process Web Orders"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =9420
                    LayoutCachedTop =360
                    LayoutCachedWidth =11220
                    LayoutCachedHeight =780
                    Overlaps =1
                End
                Begin CommandButton
                    OverlapFlags =247
                    Left =10440
                    Top =6060
                    Width =1200
                    Height =900
                    TabIndex =16
                    Name ="cmdCompare"
                    Caption ="Compare Licensee Orders"
                    OnClick ="[Event Procedure]"

                    LayoutCachedLeft =10440
                    LayoutCachedTop =6060
                    LayoutCachedWidth =11640
                    LayoutCachedHeight =6960
                    Overlaps =1
                End
            End
        End
        Begin FormFooter
            Height =720
            BackColor =-2147483633
            Name ="FormFooter"
            Begin
                Begin Label
                    Visible = NotDefault
                    OverlapFlags =93
                    TextAlign =2
                    Left =240
                    Top =60
                    Width =9795
                    Height =570
                    FontSize =20
                    FontWeight =700
                    ForeColor =255
                    Name ="lbRetr"
                    Caption ="Retrieving Dispenser Order"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    FontUnderline = NotDefault
                    OverlapFlags =215
                    TextAlign =2
                    Left =240
                    Top =60
                    Width =6120
                    Height =300
                    FontSize =10
                    BackColor =-2147483633
                    ForeColor =1279872587
                    Name ="lblMapQuest"
                    Caption ="Click here for &Google Map to AMERICAN LEGION"
                    FontName ="MS Sans Serif"
                    Tag ="http://maps.google.com/maps"
                    ControlTipText ="http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q=2129+2ND+ST,"
                        "+BAKER+CITY,+OR+97814&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
                    HyperlinkAddress ="http://maps.google.com/maps?f=q&hl=en&geocode=&time=&date=&ttype=&q=2129+2ND+ST,"
                        "+BAKER+CITY,+OR+97814&sll=&sspn=&ie=UTF8&z=16&iwloc=addr&om=1"
                End
                Begin Label
                    OverlapFlags =215
                    TextAlign =2
                    Left =240
                    Top =420
                    Width =6120
                    Height =300
                    FontSize =10
                    BackColor =-2147483633
                    Name ="Label49"
                    Caption ="2129 2ND ST, BAKER CITY, OR 97814"
                    FontName ="MS Sans Serif"
                    Tag ="http://maps.google.com/maps"
                End
            End
        End
    End
End
CodeBehindForm
' See "ClubOrderCreate.cls"
