VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form estateform 
   BorderStyle     =   0  'None
   Caption         =   "Estates"
   ClientHeight    =   10110
   ClientLeft      =   2685
   ClientTop       =   555
   ClientWidth     =   13440
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
   Picture         =   "estatesform.frx":0000
   ScaleHeight     =   10110
   ScaleWidth      =   13440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command10 
      Caption         =   "Check For Vacant Houses"
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
      Left            =   5760
      TabIndex        =   32
      Top             =   6960
      Width           =   4575
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
      Height          =   615
      Left            =   10080
      TabIndex        =   31
      Top             =   3960
      Width           =   1455
   End
   Begin VB.TextBox Text9 
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
      Left            =   7080
      TabIndex        =   30
      Top             =   3960
      Width           =   2535
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11640
      Top             =   6960
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
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
      Connect         =   $"estatesform.frx":26843
      OLEDBString     =   $"estatesform.frx":268F4
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Estates_Table"
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "DigifaceWide"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.TextBox Text7 
      DataField       =   "Occupied_Houses"
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
      Height          =   615
      Left            =   2040
      TabIndex        =   28
      Top             =   4440
      Width           =   2655
   End
   Begin VB.TextBox Text8 
      DataField       =   "Estate_Type"
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
      Height          =   615
      Left            =   2040
      TabIndex        =   26
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Choose Estate Type"
      Height          =   1335
      Left            =   2040
      TabIndex        =   22
      Top             =   1920
      Width           =   2655
      Begin VB.OptionButton Option3 
         Caption         =   "High Income"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   960
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Middle Income"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   2175
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Low Income"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
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
      Left            =   9720
      TabIndex        =   20
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton Command7 
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
      Left            =   6480
      TabIndex        =   19
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton Command6 
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
      Left            =   3120
      TabIndex        =   18
      Top             =   8280
      Width           =   2055
   End
   Begin VB.CommandButton Command5 
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
      Left            =   240
      TabIndex        =   17
      Top             =   8400
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
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
      Height          =   615
      Left            =   9720
      TabIndex        =   16
      Top             =   9360
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
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
      Left            =   6480
      TabIndex        =   15
      Top             =   9360
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
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
      Left            =   3120
      TabIndex        =   14
      Top             =   9360
      Width           =   2055
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
      Left            =   240
      TabIndex        =   13
      Top             =   9360
      Width           =   2055
   End
   Begin VB.TextBox Text6 
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
      Height          =   615
      Left            =   2040
      TabIndex        =   12
      Top             =   7200
      Width           =   2655
   End
   Begin VB.TextBox Text5 
      DataField       =   "Vacant_Houses"
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
      Height          =   615
      Left            =   7080
      TabIndex        =   11
      Top             =   5520
      Width           =   2535
   End
   Begin VB.TextBox Text4 
      DataField       =   "Total_No_of_Houses"
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
      Height          =   615
      Left            =   2040
      TabIndex        =   10
      Top             =   5520
      Width           =   2655
   End
   Begin VB.TextBox Text3 
      DataField       =   "Estate_Location"
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
      Height          =   615
      Left            =   7080
      TabIndex        =   9
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox Text2 
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
      Height          =   615
      Left            =   7080
      TabIndex        =   8
      Top             =   960
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      DataField       =   "Estate_Name"
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
      Height          =   615
      Left            =   2040
      TabIndex        =   7
      Top             =   960
      Width           =   2655
   End
   Begin VB.Label Label10 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Key in Estate No."
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
      Left            =   6960
      TabIndex        =   29
      Top             =   3360
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      Height          =   1575
      Left            =   6720
      Shape           =   4  'Rounded Rectangle
      Top             =   3240
      Width           =   5175
   End
   Begin VB.Label Label9 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Occupied Houses"
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
      TabIndex        =   27
      Top             =   4440
      Width           =   1455
   End
   Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
      Height          =   495
      Left            =   10200
      TabIndex        =   21
      Top             =   600
      Visible         =   0   'False
      Width           =   615
      URL             =   "C:\Gerald Riwo KCSE 2011 Project\Other Project files\Miss Independent instumental.wav"
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
   Begin VB.Label Label8 
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
      Height          =   615
      Left            =   240
      TabIndex        =   6
      Top             =   7200
      Width           =   1215
   End
   Begin VB.Label Label7 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "No. of Vacant Houses"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   4920
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Total No. of Houses"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   4
      Top             =   5520
      Width           =   1455
   End
   Begin VB.Label Label4 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estate Location"
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
      Left            =   4920
      TabIndex        =   3
      Top             =   2280
      Width           =   1455
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
      Height          =   615
      Left            =   4920
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Estate Name"
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
      Caption         =   "ESTATES"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "estateform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
'If MsgBox("Add New Record?", vbYesNo + vbQuestion, "Add New?") = vbYes Then
'Command7.Enabled = False
'Command6.Enabled = False
'Command8.Enabled = False
'Command5.Enabled = False
'Command4.Enabled = False
Adodc1.Recordset.AddNew
'End If
End Sub

Private Sub Command10_Click()
Dim vacant, occupied, total As Single
occupied = Val(Text7.Text)
total = Val(Text4.Text)
vacant = total - occupied
Text5.Text = vacant
End Sub

Private Sub Command2_Click()
With Adodc1.Recordset
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Or Text5.Text = "" Or Text6.Text = "" Or Text7.Text = "" Or Text8.Text = "" Then
MsgBox "Please fill in all the details", vbExclamation + vbOKOnly, "Error"

Else
Adodc1.Recordset.Save
MsgBox "Record saved", vbOKOnly + vbInformation, "SAVED"

'Command7.Enabled = True
'Command6.Enabled = True
'Command8.Enabled = True
'Command5.Enabled = True
'Command4.Enabled = True
End If
End With
End Sub

Private Sub Command3_Click()
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

Private Sub Command4_Click()
Unload Me

menuform.Show
End Sub

Private Sub Command5_Click()
Adodc1.Recordset.MoveFirst
If Adodc1.Recordset.BOF = True Then
MsgBox ("This is the first record")
End If
End Sub

Private Sub Command6_Click()
Adodc1.Recordset.MovePrevious
With Adodc1.Recordset
If .BOF = True Then
MsgBox ("First record")
.MoveNext
End If
End With
End Sub

Private Sub Command7_Click()
Adodc1.Recordset.MoveNext
If Adodc1.Recordset.EOF = True Then
MsgBox ("No more records")
Adodc1.Recordset.MoveFirst
End If
End Sub

Private Sub Command8_Click()
Adodc1.Recordset.MoveLast
If Adodc1.Recordset.EOF = True Then
MsgBox ("This is the Last record")
End If
End Sub

Private Sub Command9_Click()
Dim Valsearch As String
search = Text9.Text
search = Trim$(search)

If search <> "" Then
    With Adodc1.Recordset
       .MoveFirst
       .Find "[Estate_No]='" & search & "'"
            If .EOF Then
                MsgBox "The record you specified was not found. Please ensure that you typed the correct estate number", vbOKOnly + vbExclamation, "Search result"
                Adodc1.Refresh
            Else
                MsgBox "Record found!"
            End If
            Text9.Text = ""
    End With
End If

End Sub

Private Sub Option1_Click()
If Option1.Enabled = True Then
Text8.Text = "Low Income"
End If
End Sub

Private Sub Option2_Click()
If Option2.Enabled = True Then
Text8.Text = "Middle Income"
End If
End Sub

Private Sub Option3_Click()
If Option3.Enabled = True Then
Text8.Text = "High Income"
End If
End Sub
