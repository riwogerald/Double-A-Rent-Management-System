VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.MDIForm menuform 
   BackColor       =   &H8000000C&
   Caption         =   "Menu"
   ClientHeight    =   8490
   ClientLeft      =   2805
   ClientTop       =   1875
   ClientWidth     =   13620
   LinkTopic       =   "MDIForm1"
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   8535
      Left            =   0
      Picture         =   "menuform.frx":0000
      ScaleHeight     =   8475
      ScaleWidth      =   13560
      TabIndex        =   0
      Top             =   0
      Width           =   13620
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Double A Property Management Company"
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
         Height          =   1215
         Left            =   7320
         TabIndex        =   2
         Top             =   6840
         Width           =   4455
      End
      Begin WMPLibCtl.WindowsMediaPlayer WindowsMediaPlayer1 
         Height          =   495
         Left            =   1560
         TabIndex        =   1
         Top             =   7200
         Visible         =   0   'False
         Width           =   495
         URL             =   "C:\Rent System Project\Background Music\Delibes - Lakmé - Viens, Mallika, les liane.mp3"
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
   End
   Begin VB.Menu mnuforms 
      Caption         =   "&Forms"
      Begin VB.Menu mnuagents 
         Caption         =   "&Agents"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuagentearnings 
         Caption         =   "&Agent Earnings"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnucompanyexpenses 
         Caption         =   "&Company Expenses"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnucompanyincome 
         Caption         =   "&Company Income"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuemployees 
         Caption         =   "&Employees"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuestate 
         Caption         =   "&Estates"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnulandlords 
         Caption         =   "&Landlords"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuremittance 
         Caption         =   "&Remittance"
         Shortcut        =   ^H
      End
      Begin VB.Menu mnurentpayment 
         Caption         =   "&Rent Payment"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnutenancyagreement 
         Caption         =   "&Tenancy Agreement"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnutenants 
         Caption         =   "&Tenants"
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuapplications 
      Caption         =   "&Applications"
      Begin VB.Menu mnucalc 
         Caption         =   "&Calculator"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnumsword 
         Caption         =   "&Microsoft Word"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnumsexcel 
         Caption         =   "&Microsoft Excel"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnumsaccess 
         Caption         =   "&Microsoft Access"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnunotepad 
         Caption         =   "&Notepad"
         Shortcut        =   {F8}
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuaboutsystem 
         Caption         =   "&About the System"
         Shortcut        =   {F12}
      End
   End
   Begin VB.Menu mnuexit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "menuform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub mnuaboutsystem_Click()
aboutform.Show

Unload Me
End Sub

Private Sub mnuagentearnings_Click()
agentearningsform.Show

Unload Me
End Sub

Private Sub mnuagentearningsrpt_Click()
agentearningsreport.Show
End Sub

Private Sub mnuagents_Click()
agentsform.Show

Unload Me
End Sub

Private Sub mnuagentsrpt_Click()
agentsreport.Show
End Sub

Private Sub mnucalc_Click()
Shell "calc.exe"
End Sub

Private Sub mnucompanyexpenses_Click()
companyexpensesform.Show

Unload Me
End Sub

Private Sub mnucompanyexpensesrpt_Click()
companyexpensesreport.Show
End Sub

Private Sub mnucompanyincome_Click()
companyincomeform.Show

Unload Me
End Sub

Private Sub mnucompanyincomerpt_Click()
companyincomereport.Show
End Sub

Private Sub mnuemployees_Click()
employeesform.Show

Unload Me
End Sub

Private Sub mnuemployeesrpt_Click()
employeesreport.Show
End Sub

Private Sub mnuestate_Click()
Unload Me

estateform.Show
End Sub

Private Sub mnuestatesrpt_Click()
estatesreport.Show
End Sub

Private Sub mnuexit_Click()
If MsgBox("Are you sure want to exit the Double A Property Management System?", vbYesNo + vbQuestion, "Exit the System?") = vbYes Then
        Unload Me
        End If
End Sub

Private Sub mnulandlords_Click()
Unload Me

landlordform.Show
End Sub

Private Sub mnulandlordsrpt_Click()
landlordreport.Show
End Sub

Private Sub mnumsaccess_Click()
Shell "C:\Program Files\Microsoft Office\Office12\msaccess"
End Sub

Private Sub mnumsexcel_Click()
Shell "C:\Program Files\Microsoft Office\Office15\excel"
End Sub

Private Sub mnumsword_Click()
Shell "C:\Program Files\Microsoft Office\Office15\winword"
End Sub

Private Sub mnunotepad_Click()
Shell "notepad.exe"
End Sub

Private Sub mnuremittance_Click()
Unload Me

remittanceform.Show
End Sub

Private Sub mnuremittancerpt_Click()
remittancereport.Show
End Sub

Private Sub mnurentpayment_Click()
Unload Me

rentpaymentform.Show
End Sub

Private Sub mnurentpaymentrpt_Click()
rentpaymentreport.Show
End Sub

Private Sub mnutenancyagreement_Click()
Unload Me

tenancyagreementform.Show
End Sub

Private Sub mnutenancyagreementrpt_Click()
tenancyagreementreport.Show
End Sub

Private Sub mnutenants_Click()
Unload Me

tenantsform.Show
End Sub

Private Sub mnutenantsrpt_Click()
tenantsreport.Show
End Sub
