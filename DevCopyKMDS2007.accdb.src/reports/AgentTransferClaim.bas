Version =20
VersionRequired =20
Begin Report
    LayoutForPrint = NotDefault
    OrderByOn = NotDefault
    AllowDesignChanges = NotDefault
    DefaultView =0
    TabularFamily =9
    DateGrouping =1
    GrpKeepTogether =1
    PictureAlignment =2
    DatasheetGridlinesBehavior =3
    GridX =24
    GridY =24
    Width =14280
    DatasheetFontHeight =10
    ItemSuffix =264
    Top =1755
    OnPage ="[Event Procedure]"
    OrderBy ="LineID"
    RecSrcDt = Begin
        0xe28a9e46cab2e540
    End
    RecordSource ="qryAgentTransferLinesAndSysfile"
    Caption ="AgentTransferClaim"
    OnOpen ="[Event Procedure]"
    DatasheetFontName ="Arial"
    FilterOnLoad =0
    DatasheetGridlinesColor12 =12632256
    Begin
        Begin Label
            FontItalic = NotDefault
            BackStyle =0
            TextAlign =1
            TextFontFamily =18
            FontSize =11
            FontWeight =700
            ForeColor =8388608
            FontName ="Times New Roman"
        End
        Begin Rectangle
            BackStyle =0
            BorderWidth =1
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Line
            BorderLineStyle =0
            BorderColor =8388608
        End
        Begin Image
            OldBorderStyle =0
            BorderLineStyle =0
            PictureAlignment =2
        End
        Begin CheckBox
            BorderLineStyle =0
            LabelX =230
            LabelY =-30
        End
        Begin TextBox
            FELineBreak = NotDefault
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
            AsianLineBreak =255
            ShowDatePicker =0
        End
        Begin ListBox
            OldBorderStyle =0
            BorderLineStyle =0
            FontName ="Arial"
        End
        Begin ComboBox
            OldBorderStyle =0
            BorderLineStyle =0
            BackStyle =0
            FontName ="Arial"
        End
        Begin Subform
            OldBorderStyle =0
            BorderLineStyle =0
        End
        Begin BreakLevel
            GroupHeader = NotDefault
            GroupFooter = NotDefault
            ControlSource ="ID"
        End
        Begin FormHeader
            KeepTogether = NotDefault
            Height =2640
            OnFormat ="[Event Procedure]"
            Name ="ReportHeader"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =90
                    Top =120
                    Width =14085
                    Height =420
                    FontSize =14
                    Name ="lblTransfer"
                    Caption ="Oregon Liquor Control Commission"
                    FontName ="Trebuchet MS"
                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5760
                    Top =1620
                    Width =720
                    ColumnOrder =0
                    Name ="txtPages"
                    ControlSource ="=[Pages]"

                End
                Begin TextBox
                    Visible = NotDefault
                    IMESentenceMode =3
                    Left =5580
                    Top =1320
                    Width =600
                    ColumnOrder =1
                    TabIndex =1
                    Name ="Status"
                    ControlSource ="Status"

                End
                Begin TextBox
                    TextAlign =2
                    IMESentenceMode =3
                    Left =90
                    Top =900
                    Width =14070
                    Height =360
                    ColumnOrder =2
                    FontSize =12
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text175"
                    ControlSource ="=\"A-\" & [AgentNumber] & \" \" & [ThisAgentName]"
                    FontName ="Trebuchet MS"

                End
                Begin Rectangle
                    Left =60
                    Top =1260
                    Width =4620
                    Height =1380
                    Name ="Box57"
                End
                Begin Label
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =34
                    Left =90
                    Top =540
                    Width =14085
                    Height =375
                    FontSize =14
                    Name ="Label197"
                    Caption ="Store Claims Report"
                    FontName ="Trebuchet MS"
                End
                Begin Label
                    TextAlign =0
                    Left =540
                    Top =1380
                    Width =3705
                    Height =255
                    FontSize =9
                    Name ="Label201"
                    Caption ="Broken In Case; broken and dry, conc. damage"
                End
                Begin Rectangle
                    Left =180
                    Top =1380
                    Width =240
                    Height =240
                    Name ="Box202"
                End
                Begin Label
                    TextAlign =0
                    Left =540
                    Top =1680
                    Width =3705
                    Height =255
                    FontSize =9
                    Name ="Label203"
                    Caption ="Store/Customer breakage, pilferage"
                End
                Begin Rectangle
                    Left =180
                    Top =1680
                    Width =240
                    Height =240
                    Name ="Box204"
                End
                Begin Label
                    TextAlign =0
                    Left =540
                    Top =1980
                    Width =3705
                    Height =255
                    FontSize =9
                    Name ="Label205"
                    Caption ="Transportation claim; see OS + D Form"
                End
                Begin Rectangle
                    Left =180
                    Top =1980
                    Width =240
                    Height =240
                    Name ="Box206"
                End
                Begin Label
                    TextAlign =0
                    Left =540
                    Top =2280
                    Width =3705
                    Height =255
                    FontSize =9
                    Name ="Label207"
                    Caption ="Distillery Claim"
                End
                Begin Rectangle
                    Left =180
                    Top =2280
                    Width =240
                    Height =240
                    Name ="Box208"
                End
                Begin Rectangle
                    Left =8760
                    Top =1260
                    Width =5460
                    Height =1380
                    Name ="Box209"
                End
                Begin Label
                    Left =8820
                    Top =1380
                    Width =5340
                    Height =240
                    FontSize =9
                    Name ="Label210"
                    Caption ="Do not list more than one classification on this claim report."
                End
                Begin Label
                    Left =8820
                    Top =1680
                    Width =5340
                    Height =240
                    FontSize =9
                    Name ="Label211"
                    Caption ="Transportation claims are to be mailed to Distribution Division with"
                End
                Begin Label
                    Left =8820
                    Top =1980
                    Width =5340
                    Height =240
                    FontSize =9
                    Name ="Label212"
                    Caption ="OS + D form immediately.  All other claims are filed with Distribution"
                End
                Begin Label
                    Left =8820
                    Top =2280
                    Width =5340
                    Height =240
                    FontSize =9
                    Name ="Label213"
                    Caption =" Division by OLCC District Representative or Distillery Representative."
                End
            End
        End
        Begin PageHeader
            Height =660
            OnFormat ="[Event Procedure]"
            Name ="PageHeaderSection"
            Begin
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7380
                    Top =360
                    Width =1380
                    Height =300
                    Name ="Label88"
                    Caption ="Bottle Count"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =8760
                    Top =360
                    Width =2100
                    Height =300
                    Name ="Label214"
                    Caption ="Total Amt Claimed"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =6120
                    Top =60
                    Width =1260
                    Height =600
                    Name ="Label185"
                    Caption ="Retail Amt. Claimed"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =5220
                    Top =60
                    Width =900
                    Height =600
                    Name ="Label184"
                    Caption ="Price"
                    LayoutCachedLeft =5220
                    LayoutCachedTop =60
                    LayoutCachedWidth =6120
                    LayoutCachedHeight =660
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =120
                    Top =60
                    Width =1380
                    Height =600
                    Name ="Label83"
                    Caption ="Code"
                    LayoutCachedLeft =120
                    LayoutCachedTop =60
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =660
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =2040
                    Top =60
                    Width =3180
                    Height =600
                    Name ="Label84"
                    Caption ="Brand Name"
                    LayoutCachedLeft =2040
                    LayoutCachedTop =60
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =660
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =1500
                    Top =60
                    Width =540
                    Height =600
                    Name ="Label85"
                    Caption ="Qty"
                    LayoutCachedLeft =1500
                    LayoutCachedTop =60
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =660
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =10860
                    Top =360
                    Width =1440
                    Height =300
                    Name ="Label215"
                    Caption ="T-Claim No."
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =12300
                    Top =60
                    Width =1920
                    Height =600
                    Name ="Label216"
                    Caption ="Explanation"
                End
                Begin Label
                    OldBorderStyle =1
                    TextAlign =2
                    Left =7380
                    Top =60
                    Width =4920
                    Height =300
                    Name ="Label217"
                    Caption ="Office Use Only"
                End
            End
        End
        Begin BreakHeader
            KeepTogether = NotDefault
            Height =0
            Name ="GroupHeader0"
        End
        Begin Section
            KeepTogether = NotDefault
            Height =375
            Name ="Detail"
            Begin
                Begin TextBox
                    BackStyle =1
                    Left =2100
                    Width =3120
                    FontSize =9
                    TabIndex =1
                    Name ="Description"
                    ControlSource ="Description"

                    LayoutCachedLeft =2100
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =1
                    BackStyle =1
                    Left =180
                    Top =15
                    Width =1200
                    Height =225
                    Name ="ItemID"
                    ControlSource ="ItemID"
                    FontName ="Arial Narrow"

                    LayoutCachedLeft =180
                    LayoutCachedTop =15
                    LayoutCachedWidth =1380
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    TextAlign =3
                    BackStyle =1
                    Left =1560
                    Width =420
                    FontSize =9
                    TabIndex =2
                    Name ="QuantityOrdered"
                    ControlSource ="QuantityShipped"

                    LayoutCachedLeft =1560
                    LayoutCachedWidth =1980
                    LayoutCachedHeight =240
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =1
                    BackStyle =1
                    Left =12420
                    Top =60
                    Width =1680
                    FontSize =9
                    TabIndex =3
                    Name ="ExtPrice"
                    ControlSource ="ClaimCode"
                    Format ="Standard"

                End
                Begin Line
                    LineSlant = NotDefault
                    Left =1500
                    Width =0
                    Height =360
                    Name ="Line91"
                    LayoutCachedLeft =1500
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =2040
                    Width =0
                    Height =360
                    Name ="Line92"
                    LayoutCachedLeft =2040
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =6120
                    Width =0
                    Height =360
                    Name ="Line96"
                End
                Begin Line
                    Left =7380
                    Width =0
                    Height =360
                    Name ="Line99"
                End
                Begin Line
                    Left =5220
                    Width =0
                    Height =360
                    Name ="Line133"
                    LayoutCachedLeft =5220
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =360
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =5340
                    Width =720
                    Height =255
                    FontSize =9
                    TabIndex =4
                    Name ="Text186"
                    ControlSource ="Price"
                    Format ="Standard"

                    LayoutCachedLeft =5340
                    LayoutCachedWidth =6060
                    LayoutCachedHeight =255
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =6180
                    Width =1140
                    Height =225
                    FontSize =9
                    TabIndex =5
                    Name ="Text187"
                    ControlSource ="ExtPrice"
                    Format ="Standard"

                End
                Begin Line
                    Left =8760
                    Width =0
                    Height =360
                    Name ="Line188"
                End
                Begin Line
                    Left =12300
                    Width =0
                    Height =360
                    Name ="Line196"
                End
                Begin Line
                    Left =10860
                    Width =0
                    Height =360
                    Name ="Line218"
                End
                Begin Line
                    Left =14220
                    Width =0
                    Height =360
                    Name ="Line219"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =120
                    Width =0
                    Height =360
                    Name ="Line220"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =120
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =120
                    Top =360
                    Width =14100
                    Name ="Line117"
                End
            End
        End
        Begin BreakFooter
            KeepTogether = NotDefault
            Height =600
            Name ="GroupFooter1"
            Begin
                Begin Label
                    FontItalic = NotDefault
                    Left =480
                    Top =60
                    Width =525
                    Height =270
                    FontSize =10
                    ForeColor =0
                    Name ="Label39"
                    Caption ="Total"
                    LayoutCachedLeft =480
                    LayoutCachedTop =60
                    LayoutCachedWidth =1005
                    LayoutCachedHeight =330
                End
                Begin TextBox
                    TextAlign =2
                    BackStyle =1
                    Left =1620
                    Top =60
                    Width =585
                    FontSize =10
                    FontWeight =700
                    Name ="Text119"
                    ControlSource ="=Sum([QuantityShipped])"

                    LayoutCachedLeft =1620
                    LayoutCachedTop =60
                    LayoutCachedWidth =2205
                    LayoutCachedHeight =300
                End
                Begin TextBox
                    DecimalPlaces =2
                    TextAlign =3
                    BackStyle =1
                    Left =6180
                    Top =60
                    Width =1140
                    FontSize =10
                    FontWeight =700
                    TabIndex =1
                    Name ="Text190"
                    ControlSource ="=Sum([ExtPrice])"
                    Format ="Standard"

                End
                Begin Line
                    Left =1500
                    Width =0
                    Height =360
                    Name ="Line225"
                    LayoutCachedLeft =1500
                    LayoutCachedWidth =1500
                    LayoutCachedHeight =360
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =2040
                    Width =0
                    Height =360
                    Name ="Line226"
                    LayoutCachedLeft =2040
                    LayoutCachedWidth =2040
                    LayoutCachedHeight =360
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =6120
                    Width =0
                    Height =360
                    Name ="Line227"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =7380
                    Width =0
                    Height =360
                    Name ="Line228"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =5220
                    Width =0
                    Height =360
                    Name ="Line229"
                    LayoutCachedLeft =5220
                    LayoutCachedWidth =5220
                    LayoutCachedHeight =360
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =8760
                    Width =0
                    Height =360
                    Name ="Line232"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =12300
                    Width =0
                    Height =360
                    Name ="Line233"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =10860
                    Width =0
                    Height =360
                    Name ="Line234"
                End
                Begin Line
                    LineSlant = NotDefault
                    Left =14220
                    Width =0
                    Height =360
                    Name ="Line235"
                End
                Begin Line
                    Left =120
                    Width =0
                    Height =360
                    Name ="Line236"
                    LayoutCachedLeft =120
                    LayoutCachedWidth =120
                    LayoutCachedHeight =360
                End
                Begin Line
                    Left =120
                    Top =360
                    Width =14100
                    Name ="Line238"
                End
                Begin Line
                    Left =120
                    Top =420
                    Width =14100
                    Name ="Line239"
                End
            End
        End
        Begin PageFooter
            Height =480
            Name ="PageFooterSection"
            Begin
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Top =240
                    Width =5400
                    FontSize =9
                    FontWeight =700
                    ForeColor =8388608
                    Name ="Text48"
                    ControlSource ="=Now()"
                    Format ="Long Time"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Top =240
                    Width =8520
                    FontSize =9
                    FontWeight =700
                    TabIndex =1
                    ForeColor =8388608
                    Name ="Text49"
                    ControlSource ="=\"Page \" & [Page] & \" of \" & [Pages]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =1
                    TextFontFamily =18
                    Left =60
                    Width =8520
                    FontSize =9
                    FontWeight =700
                    TabIndex =2
                    ForeColor =8388608
                    Name ="Text63"
                    ControlSource ="=Now()"
                    Format ="Long Date"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =3
                    TextFontFamily =18
                    Left =5700
                    Width =8520
                    FontSize =9
                    FontWeight =700
                    TabIndex =3
                    ForeColor =8388608
                    Name ="Text194"
                    ControlSource ="=[Name]"
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Top =240
                    Width =14160
                    FontSize =6
                    FontWeight =700
                    TabIndex =4
                    ForeColor =8388608
                    Name ="Text195"
                    ControlSource ="=\"Kelly Myers Data Systems\""
                    FontName ="Times New Roman"

                End
                Begin TextBox
                    FontItalic = NotDefault
                    TextAlign =2
                    TextFontFamily =18
                    Left =60
                    Width =14160
                    FontSize =9
                    FontWeight =700
                    TabIndex =5
                    ForeColor =8388608
                    Name ="OLCCName"
                    ControlSource ="=GetStoreName()"
                    FontName ="Times New Roman"

                End
            End
        End
        Begin FormFooter
            KeepTogether = NotDefault
            Height =2340
            Name ="ReportFooter"
            Begin
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10380
                    Top =1200
                    Width =3780
                    Height =300
                    FontSize =11
                    Name ="AgentNumber"
                    ControlSource ="AgentNumber"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =9300
                            Top =1200
                            Width =1050
                            Height =300
                            Name ="Label178"
                            Caption ="Store No."
                        End
                    End
                End
                Begin Line
                    Left =10380
                    Top =480
                    Width =3840
                    Name ="Line179"
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10380
                    Top =900
                    Width =3780
                    Height =300
                    FontSize =11
                    TabIndex =1
                    Name ="Text79"
                    ControlSource ="OrdDate"
                    Format ="Medium Date"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =9300
                            Top =900
                            Width =1050
                            Height =300
                            Name ="Label80"
                            Caption ="Date"
                        End
                    End
                End
                Begin TextBox
                    OldBorderStyle =1
                    BorderWidth =1
                    TextAlign =2
                    IMESentenceMode =3
                    Left =10380
                    Top =1500
                    Width =3780
                    Height =300
                    FontSize =11
                    TabIndex =2
                    Name ="Text240"
                    ControlSource ="=\"A-\" & [AgentNumber] & \" \" & [ThisAgentName]"

                    Begin
                        Begin Label
                            OldBorderStyle =1
                            BorderWidth =1
                            Left =9300
                            Top =1500
                            Width =1050
                            Height =300
                            Name ="Label242"
                            Caption ="Location"
                        End
                    End
                End
                Begin Label
                    BorderWidth =1
                    Left =10380
                    Top =540
                    Width =1290
                    Height =255
                    FontSize =9
                    Name ="Label241"
                    Caption ="Agent Signature"
                End
                Begin Label
                    Left =240
                    Top =60
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label243"
                    Caption ="Distillery Claims Codes"
                End
                Begin Label
                    Left =420
                    Top =360
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label244"
                    Caption ="40 - Defective Cap"
                End
                Begin Label
                    Left =420
                    Top =600
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label245"
                    Caption ="42 - Seal Broken"
                End
                Begin Label
                    Left =420
                    Top =840
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label246"
                    Caption ="43 - No Bottle Cap"
                End
                Begin Label
                    Left =420
                    Top =1080
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label247"
                    Caption ="44 - Missing in Case"
                End
                Begin Label
                    Left =420
                    Top =1320
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label248"
                    Caption ="45 - Complaint"
                End
                Begin Label
                    Left =420
                    Top =1560
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label249"
                    Caption ="46 - Defective Bottle"
                End
                Begin Label
                    Left =420
                    Top =1800
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label250"
                    Caption ="47 - Foreign Matter"
                End
                Begin Label
                    Left =420
                    Top =2040
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label251"
                    Caption ="48 - Labels"
                End
                Begin Label
                    Left =2220
                    Top =360
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label252"
                    Caption ="49 - Short-Low-Empty"
                End
                Begin Label
                    Left =2220
                    Top =600
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label253"
                    Caption ="50 - Improper Packaging"
                End
                Begin Label
                    Left =2220
                    Top =840
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label254"
                    Caption ="51 - Solidified"
                End
                Begin Label
                    Left =2220
                    Top =1080
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label255"
                    Caption ="52 - Miscoded Carton"
                End
                Begin Label
                    Left =2220
                    Top =1320
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label256"
                    Caption ="53 - Not Listed in Oregon"
                End
                Begin Label
                    Left =2220
                    Top =1560
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label257"
                    Caption ="58 - Supplier Request"
                End
                Begin Label
                    Left =2220
                    Top =1800
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label258"
                    Caption ="60 - Picked by Dist. Rep."
                End
                Begin Label
                    Left =2220
                    Top =2040
                    Width =2520
                    Height =240
                    FontSize =9
                    Name ="Label259"
                    Caption ="62 - Wrong Product"
                End
                Begin Line
                    Left =6420
                    Top =2220
                    Width =7680
                    Name ="Line260"
                End
                Begin Label
                    Left =4860
                    Top =2040
                    Width =1500
                    Height =240
                    FontSize =9
                    Name ="Label261"
                    Caption ="Additional remarks"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    Left =5580
                    Top =120
                    Width =3240
                    Height =720
                    Name ="Label262"
                    Caption ="Warehouse OK"
                End
                Begin Label
                    OldBorderStyle =1
                    BorderWidth =1
                    Left =5580
                    Top =900
                    Width =3240
                    Height =840
                    Name ="Label263"
                    Caption ="DR or DRO OK"
                End
            End
        End
    End
End
CodeBehindForm
' See "AgentTransferClaim.cls"
