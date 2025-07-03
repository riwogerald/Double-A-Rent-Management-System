VERSION 5.00
Begin VB.Form splashform 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   9540
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   14595
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "DigifaceWide"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "splashform.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "splashform.frx":000C
   ScaleHeight     =   9540
   ScaleWidth      =   14595
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   120
      Picture         =   "splashform.frx":138BF
      ScaleHeight     =   1275
      ScaleWidth      =   1365
      TabIndex        =   6
      Top             =   240
      Width           =   1425
   End
   Begin VB.PictureBox ProgressBar1 
      Height          =   375
      Left            =   3120
      ScaleHeight     =   315
      ScaleWidth      =   10995
      TabIndex        =   4
      Top             =   8640
      Width           =   11055
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   12600
      Top             =   2160
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights reserved 2011©"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   11400
      TabIndex        =   5
      Top             =   6600
      Width           =   2895
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Obudho Gerald Riwo"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   4320
      TabIndex        =   3
      Top             =   1800
      Width           =   4575
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by:"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   3615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Please Wait Loading..........."
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   480
      TabIndex        =   1
      Top             =   8400
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Double A Property Management System"
      BeginProperty Font 
         Name            =   "DigifaceWide"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   1680
      TabIndex        =   0
      Top             =   360
      Width           =   12735
   End
End
Attribute VB_Name = "splashform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Timer1_Timer()
ProgressBar1.Value = ProgressBar1.Value + 0.78125
If ProgressBar1.Value >= 100 Then
Timer1.Interval = 0
Timer1.Enabled = False
Unload Me
Load menuform
menuform.Show
End If
End Sub

