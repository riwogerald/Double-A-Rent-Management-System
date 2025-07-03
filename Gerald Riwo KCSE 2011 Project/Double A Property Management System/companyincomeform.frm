VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form companyincomeform 
   Caption         =   "Company Income"
   ClientHeight    =   8460
   ClientLeft      =   3765
   ClientTop       =   1575
   ClientWidth     =   11985
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
   Picture         =   "companyincomeform.frx":0000
   ScaleHeight     =   8460
   ScaleWidth      =   11985
   Begin VB.CommandButton Command10 
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
      Height          =   615
      Left            =   8520
      TabIndex        =   21
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6240
      TabIndex        =   20
      Top             =   3120
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   3000
      Top             =   5400
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
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
      Connect         =   $"companyincomeform.frx":7FED
      OLEDBString     =   $"companyincomeform.frx":80A8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Company_Income_Table"
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
   Begin VB.CommandButton Command9 
      Caption         =   "Exit Form"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   17
      Top             =   7560
      Width           =   2055
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   16
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   15
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   14
      Top             =   7560
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Last"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9000
      TabIndex        =   13
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
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
      Height          =   495
      Left            =   2400
      TabIndex        =   12
      Top             =   3720
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6000
      TabIndex        =   11
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Previous"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3120
      TabIndex        =   10
      Top             =   6480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "First"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text5 
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
      Left            =   2280
      TabIndex        =   8
      Top             =   4560
      Width           =   2895
   End
   Begin VB.TextBox Text3 
      DataField       =   "Date_Collected"
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
      Left            =   8400
      TabIndex        =   7
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      DataField       =   "Commission_Earned"
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
      Left            =   2280
      TabIndex        =   6
      Top             =   3000
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      DataField       =   "Rent_Amount_Collected"
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
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Key in Agent No."
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
      Left            =   6240
      TabIndex        =   19
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      Height          =   1695
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   2400
      Width           =   4575
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   7560
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   495
      URL             =   "C:\Gerald Riwo KCSE 2011 Project\Other Project files\Linkin Park The Catalyst_ToxicWap.com.mp3"
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
   Begin VB.Label Label6 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Agent No"
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
      Left            =   120
      TabIndex        =   4
      Top             =   4680
      Width           =   1335
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Commission Earned"
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
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1815
   End
   Begin VB.Label Label3 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Date Collected"
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
      Left            =   6120
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rent Amount Collected"
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
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "COMPANY INCOME"
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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   5055
   End
End
Attribute VB_Name = "companyincomeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.BOF = True Then
MsgBox ("This is the first record")
End If
End Sub

Private Sub Command10_Click()
Dim Valsearch As String
search = Text6.Text
search = Trim$(search)

If search <> "" Then
    With Adodc1.Recordset
       .MoveFirst
       .Find "[Agent_No]='" & search & "'"
            If .EOF Then
                MsgBox "The record you specified was not found. Please ensure that you typed the correct agent number", vbOKOnly + vbExclamation, "Search result"
                Adodc1.Refresh
            Else
                MsgBox "Record found!"
            End If
            Text6.Text = ""
    End With
End If
End Sub

Private Sub Command2_Click()
Adodc1.Recordset.MovePrevious
With Adodc1.Recordset
If .BOF = True Then
MsgBox ("First record")
.MoveNext
End If
End With
End Sub

Private Sub Command3_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
MsgBox ("No more records")
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command4_Click()
Dim one, result As Double
one = Val(Text1.Text)
result = one * 0.04
Text2.Text = result
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
MsgBox ("This is the Last record")
End If
End Sub

Private Sub Command6_Click()
'If MsgBox("Add New Record?", vbYesNo + vbQuestion, "Add New?") = vbYes Then
'Command3.Enabled = False
'Command2.Enabled = False
'Command5.Enabled = False
'Command1.Enabled = False
'Command9.Enabled = False
Adodc1.Recordset.AddNew
'End If
End Sub

Private Sub Command7_Click()
With Adodc1.Recordset
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Then
MsgBox "Please fill in all the details", vbExclamation + vbOKOnly, "Error"

Else
Adodc1.Recordset.Save
MsgBox "Record saved", vbOKOnly + vbInformation, "SAVED"

'Command3.Enabled = True
'Command2.Enabled = True
'Command5.Enabled = True
'Command1.Enabled = True
'Command9.Enabled = True
End If
End With
End Sub

Private Sub Command8_Click()
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

Private Sub Command9_Click()
Unload Me

menuform.Show
End Sub
