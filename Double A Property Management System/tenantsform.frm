VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form tenantsform 
   BorderStyle     =   0  'None
   Caption         =   "Tenants"
   ClientHeight    =   9630
   ClientLeft      =   2490
   ClientTop       =   750
   ClientWidth     =   14490
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
   Picture         =   "tenantsform.frx":0000
   ScaleHeight     =   9630
   ScaleWidth      =   14490
   ShowInTaskbar   =   0   'False
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
      Height          =   615
      Left            =   4200
      TabIndex        =   35
      Top             =   7080
      Width           =   1455
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
      Height          =   615
      Left            =   720
      TabIndex        =   34
      Top             =   7080
      Width           =   3135
   End
   Begin VB.TextBox Text10 
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
      Left            =   8640
      TabIndex        =   29
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose House Type"
      Height          =   1455
      Left            =   6360
      TabIndex        =   28
      Top             =   3960
      Width           =   2175
      Begin VB.OptionButton Option3 
         Caption         =   "Big"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Medium"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Small"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   2055
      End
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   330
      Left            =   9360
      Top             =   6840
      Visible         =   0   'False
      Width           =   1200
      _ExtentX        =   2117
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
      Connect         =   $"tenantsform.frx":370CE
      OLEDBString     =   $"tenantsform.frx":3717F
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Tenants_Table"
      Caption         =   ""
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
   Begin VB.CommandButton Command8 
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
      Height          =   615
      Left            =   9960
      TabIndex        =   26
      Top             =   8040
      Width           =   1455
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
      Height          =   615
      Left            =   4080
      TabIndex        =   25
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
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
      Height          =   615
      Left            =   6960
      TabIndex        =   24
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command5 
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
      Height          =   615
      Left            =   4080
      TabIndex        =   23
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   1080
      TabIndex        =   22
      Top             =   8040
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
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
      Left            =   9960
      TabIndex        =   21
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
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
      Height          =   615
      Left            =   6960
      TabIndex        =   20
      Top             =   8880
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      Height          =   615
      Left            =   1080
      TabIndex        =   19
      Top             =   8880
      Width           =   1455
   End
   Begin VB.TextBox Text9 
      DataField       =   "Next_of_Kin"
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
      Left            =   8640
      TabIndex        =   18
      Top             =   3240
      Width           =   3135
   End
   Begin VB.TextBox Text8 
      DataField       =   "Contacts"
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
      Left            =   8640
      TabIndex        =   17
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      DataField       =   "Tenant_ID"
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
      Left            =   8640
      TabIndex        =   16
      Top             =   1080
      Width           =   3135
   End
   Begin VB.TextBox Text6 
      DataField       =   "Rent_Amount"
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
      Left            =   8640
      TabIndex        =   12
      Top             =   5760
      Width           =   3135
   End
   Begin VB.TextBox Text5 
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
      Left            =   2640
      TabIndex        =   11
      Top             =   5760
      Width           =   3135
   End
   Begin VB.TextBox Text4 
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
      Left            =   2640
      TabIndex        =   10
      Top             =   4800
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      DataField       =   "Kins_Contact"
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
      Left            =   2640
      TabIndex        =   9
      Top             =   3360
      Width           =   3135
   End
   Begin VB.TextBox Text2 
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
      Left            =   2640
      TabIndex        =   8
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      DataField       =   "Tenant_Name"
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
      Left            =   2640
      TabIndex        =   7
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label11 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Type in Tenant No."
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
      Left            =   720
      TabIndex        =   33
      Top             =   6600
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      Height          =   1455
      Left            =   600
      Shape           =   4  'Rounded Rectangle
      Top             =   6480
      Width           =   5295
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   7320
      TabIndex        =   27
      Top             =   240
      Visible         =   0   'False
      Width           =   615
      URL             =   "C:\Gerald Riwo KCSE 2011 Project\Other Project files\CHANGES MEDLEY.mp3"
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
      _cx             =   1085
      _cy             =   873
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tenant ID"
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
      Left            =   6360
      TabIndex        =   15
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Kin's Contact"
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
      TabIndex        =   14
      Top             =   3360
      Width           =   1695
   End
   Begin VB.Label Label8 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Next of Kin"
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
      Left            =   6360
      TabIndex        =   13
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Rent Amount"
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
      Left            =   6360
      TabIndex        =   6
      Top             =   5640
      Width           =   1215
   End
   Begin VB.Label Label6 
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
      Height          =   495
      Left            =   240
      TabIndex        =   5
      Top             =   5880
      Width           =   1575
   End
   Begin VB.Label Label5 
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
      Left            =   240
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Contacts"
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
      Left            =   6360
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label3 
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
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Tenant Name"
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
      TabIndex        =   1
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "TENANTS"
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
      Width           =   2415
   End
End
Attribute VB_Name = "tenantsform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'If MsgBox("Add New Record?", vbYesNo + vbQuestion, "Add New?") = vbYes Then
'Command6.Enabled = False
'Command5.Enabled = False
'Command8.Enabled = False
'Command4.Enabled = False
'Command3.Enabled = False
Adodc1.Recordset.AddNew
'End If
End Sub

Private Sub Command2_Click()
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

Private Sub Command3_Click()
Unload Me

menuform.Show
End Sub

Private Sub Command4_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.BOF = True Then
MsgBox ("This is the first record")
End If
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MovePrevious
With Adodc1.Recordset
If .BOF = True Then
MsgBox ("First record")
.MoveNext
End If
End With
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
MsgBox ("No more records")
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command7_Click()
With Adodc1.Recordset
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Or Text9.Text = "" Or Text10.Text = "" Then
MsgBox "Please fill in all the details", vbExclamation + vbOKOnly, "Error"

Else
Adodc1.Recordset.Save
MsgBox "Record saved", vbOKOnly + vbInformation, "SAVED"

'Command6.Enabled = True
'Command5.Enabled = True
'Command8.Enabled = True
'Command4.Enabled = True
'Command3.Enabled = True
End If
End With
End Sub

Private Sub Command8_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
MsgBox ("This is the Last record")
End If
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
Text10.Text = "Small"
End If
End Sub

Private Sub Option2_Click()
If Option2.Enabled = True Then
Text10.Text = "Medium"
End If
End Sub

Private Sub Option3_Click()
If Option3.Enabled = True Then
Text10.Text = "Big"
End If
End Sub
