VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form rentpaymentform 
   Caption         =   "Rent Payment"
   ClientHeight    =   10485
   ClientLeft      =   2610
   ClientTop       =   420
   ClientWidth     =   13800
   BeginProperty Font 
      Name            =   "DigifaceWide"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "rentpaymentform.frx":0000
   ScaleHeight     =   10485
   ScaleWidth      =   13800
   Begin VB.TextBox Text15 
      DataField       =   "House_Type"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10920
      TabIndex        =   44
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Frame Frame2 
      Caption         =   "Choose House Type"
      Height          =   2295
      Left            =   10920
      TabIndex        =   43
      Top             =   1080
      Width           =   2295
      Begin VB.OptionButton Option6 
         Caption         =   "Big"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   47
         Top             =   1560
         Width           =   2055
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Medium"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   46
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Small"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   2055
      End
   End
   Begin VB.TextBox Text13 
      DataField       =   "Current_Date"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   42
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text14 
      DataField       =   "Amount_Paid"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   40
      Top             =   4560
      Width           =   2175
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   38
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox Text12 
      DataField       =   "No of days defaulted"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   36
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Search"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   10920
      TabIndex        =   35
      Top             =   7320
      Width           =   1695
   End
   Begin VB.TextBox Text11 
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8280
      TabIndex        =   34
      Top             =   7320
      Width           =   2415
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   10920
      Top             =   9120
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   873
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"rentpaymentform.frx":2939A
      OLEDBString     =   $"rentpaymentform.frx":29455
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Rent_Payment_Table"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text3 
      DataField       =   "Mode_of_Payment"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   8160
      TabIndex        =   32
      Top             =   4800
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Payment Mode"
      Height          =   1815
      Left            =   5280
      TabIndex        =   28
      Top             =   4680
      Width           =   2535
      Begin VB.OptionButton Option3 
         Caption         =   "Credit Card"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   2175
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cheque"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2415
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Cash"
         BeginProperty Font 
            Name            =   "DigifaceWide"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Exit Form"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      TabIndex        =   27
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   26
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   25
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   5400
      TabIndex        =   24
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   23
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   22
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      TabIndex        =   21
      Top             =   9480
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8040
      TabIndex        =   20
      Top             =   8520
      Width           =   1575
   End
   Begin VB.TextBox Text10 
      DataField       =   "Remarks"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   19
      Top             =   7080
      Width           =   2175
   End
   Begin VB.TextBox Text9 
      DataField       =   "House_No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   18
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox Text8 
      DataField       =   "Estate_No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   17
      Top             =   1920
      Width           =   2175
   End
   Begin VB.TextBox Text7 
      DataField       =   "Date_of_Payment"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "dd/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   3
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8160
      TabIndex        =   16
      Top             =   2880
      Width           =   2175
   End
   Begin VB.TextBox Text6 
      DataField       =   "Receipt_No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2520
      TabIndex        =   15
      Top             =   6240
      Width           =   2175
   End
   Begin VB.TextBox Text5 
      DataField       =   "Balance"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   14
      Top             =   5520
      Width           =   2175
   End
   Begin VB.TextBox Text4 
      DataField       =   "Amount_Owed"
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "#,##0.00"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   2057
         SubFormatType   =   1
      EndProperty
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   13
      Top             =   3720
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      DataField       =   "Agent_No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   12
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      DataField       =   "Tenant_No"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      TabIndex        =   11
      Top             =   1200
      Width           =   2175
   End
   Begin VB.Label Label14 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Current Date"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   41
      Top             =   2880
      Width           =   1575
   End
   Begin VB.Label Label12 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount Owed"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   39
      Top             =   3720
      Width           =   1575
   End
   Begin VB.Label Label13 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "NO. of Days Defaulted"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      TabIndex        =   37
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Key in Tenant No."
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   33
      Top             =   6720
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      Height          =   1815
      Left            =   8160
      Shape           =   4  'Rounded Rectangle
      Top             =   6600
      Width           =   4695
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   13080
      TabIndex        =   4
      Top             =   8400
      Visible         =   0   'False
      Width           =   495
      URL             =   "C:\Gerald Riwo KCSE 2011 Project\Other Project files\evanescence - fields of innoscence.mp3"
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   873
      _cy             =   873
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Remarks"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   7080
      Width           =   1575
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent No."
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estate No."
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   8
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "House No."
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Receipt No."
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   6
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Amount Paid"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   3
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Due"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5280
      TabIndex        =   2
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tenant No."
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "RENT PAYMENT"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "rentpaymentform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
MsgBox ("This is the Last record")
End If
End Sub

Private Sub Command10_Click()
On Error GoTo HANDLER
Dim required, paid, pending, balance As Double
required = Val(Text4.Text)
paid = Val(Text14.Text)
pending = required - paid
a$ = DateDiff("d", Text7.Text, Text13.Text)
days = a$
If days < 0 Then
days = 0
End If
balance = pending * (1.005 ^ (days))
Text5.Text = balance
Exit Sub
HANDLER:
MsgBox "An error has occured. Please ensure that you have filled all fields as required.", vbOKOnly + vbExclamation, "Error!"
End Sub

Private Sub Command2_Click()
'If MsgBox("Add New Record?", vbYesNo + vbQuestion, "Add New?") = vbYes Then
'Command5.Enabled = False
'Command7.Enabled = False
'Command1.Enabled = False
'Command3.Enabled = False
'Command8.Enabled = False
Adodc1.Recordset.AddNew
'End If
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.BOF = True Then
MsgBox ("This is the first record")
End If
End Sub

Private Sub Command4_Click()
With Adodc1.Recordset
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "Please fill in all the details", vbExclamation + vbOKOnly, "Error"

Else
Adodc1.Recordset.Save
MsgBox "Record saved", vbOKOnly + vbInformation, "SAVED"

'Command5.Enabled = True
'Command7.Enabled = True
'Command1.Enabled = True
'Command3.Enabled = True
'Command8.Enabled = True
End If
End With

End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
MsgBox ("No more records")
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.Delete
On Error GoTo HANDLER
With rsDelete
If MsgBox("Are you sure you want to Delete this record?", vbOKCancel, "DELETE") = vbOK Then
.Delete
MsgBox ("Remember, you cannot retrieve the data once deleted!" & "DELETE")
End If
Exit Sub
HANDLER:
MsgBox "The record has been successfully deleted."
End With

End Sub

Private Sub Command7_Click()
Adodc1.Recordset.MovePrevious
With Adodc1.Recordset
If .BOF = True Then
MsgBox ("First record")
.MoveNext
End If
End With
End Sub

Private Sub Command8_Click()
Unload Me

menuform.Show
End Sub

Private Sub Command9_Click()
Dim Valsearch As String
search = Text11.Text
search = Trim$(search)

If search <> "" Then
    With Adodc1.Recordset
       .MoveFirst
       .Find "[Tenant_No]='" & search & "'"
            If .EOF Then
                MsgBox "The record you specified was not found. Please ensure that you typed the correct tenant number", vbOKOnly + vbExclamation, "Search result"
                Adodc1.Refresh
            Else
                MsgBox "Record found!"
            End If
            Text11.Text = ""
    End With
End If

End Sub

Private Sub Option1_Click()
If Option1.Enabled = True Then
Text3.Text = "Cash"
End If
End Sub

Private Sub Option2_Click()
If Option2.Enabled = True Then
Text3.Text = "Cheque"
End If
End Sub

Private Sub Option3_Click()
If Option3.Enabled = True Then
Text3.Text = "Credit Card"
End If
End Sub

Private Sub Option4_Click()
If Option4.Enabled = True Then
Text15.Text = "Small"
End If
End Sub

Private Sub Option5_Click()
If Option5.Enabled = True Then
Text15.Text = "Medium"
End If
End Sub

Private Sub Option6_Click()
If Option6.Enabled = True Then
Text15.Text = "Big"
End If
End Sub
