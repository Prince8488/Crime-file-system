VERSION 5.00
Begin VB.MDIForm frmMDI 
   BackColor       =   &H8000000C&
   Caption         =   "Crime File System"
   ClientHeight    =   7905
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   15045
   LinkTopic       =   "MDIForm1"
   Picture         =   "frmMDI.frx":0000
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Menu mnuAdmin 
      Caption         =   "Administrator"
      Begin VB.Menu mnuLogin 
         Caption         =   "Login"
      End
      Begin VB.Menu mnuLogoff 
         Caption         =   "Logoff"
      End
      Begin VB.Menu mnud 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAdminAddU 
         Caption         =   "Add User"
      End
      Begin VB.Menu mnuAdminDelU 
         Caption         =   "Delete User"
      End
      Begin VB.Menu mnuAdminAddComSt 
         Caption         =   "Add Complaint Status"
      End
   End
   Begin VB.Menu mnuActivReg 
      Caption         =   "Registers"
      Begin VB.Menu mnuActivRegCrim 
         Caption         =   "Criminal Register"
      End
      Begin VB.Menu mnuActivRegPris 
         Caption         =   "Prisoners register"
      End
      Begin VB.Menu mnuActivRegPost 
         Caption         =   "Postmortem"
      End
      Begin VB.Menu mnuActRegMst 
         Caption         =   "Most Wanted"
      End
   End
   Begin VB.Menu mnuActiv 
      Caption         =   "Activities"
      Begin VB.Menu mnuActivFir 
         Caption         =   "FIR"
      End
      Begin VB.Menu mnuActivChrgSht 
         Caption         =   "Charge Sheet"
      End
   End
   Begin VB.Menu mnuActivComp 
      Caption         =   "Complaint"
      Begin VB.Menu mnuActivCompStat 
         Caption         =   "Complaint status"
      End
      Begin VB.Menu mnuActivCompReg 
         Caption         =   "Complaint Registration"
      End
   End
   Begin VB.Menu mnuActivRpt 
      Caption         =   "Reports"
      Begin VB.Menu mnuActivRptCrim 
         Caption         =   "Crime Report"
      End
      Begin VB.Menu mnuActivRptMst 
         Caption         =   "Most Wanted"
      End
      Begin VB.Menu mnuRptMortem 
         Caption         =   "Post Mortem"
      End
      Begin VB.Menu mnuRptFIR 
         Caption         =   "FIR"
      End
   End
End
Attribute VB_Name = "frmMDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub dbconnection()
Set con = New ADODB.Connection
Set res = New ADODB.Recordset
With con

   .ConnectionString = "Driver=(MySQL ODBC 5.1.13 Driver);SERVER=localhost;PWD=root;UID=root;PORT=3306;Data Source=crime"
   .CursorLocation = adUseClient
   .Open
  
End With
End Sub

 Private Sub MDIForm_Load()
 mnuLogoff.Enabled = False
 mnuAdminAddU.Enabled = False
 mnuAdminDelU.Enabled = False
 mnuActiv.Enabled = False
 mnuAdminAddComSt.Enabled = False
 mnuActivReg.Enabled = False
 mnuActivComp.Enabled = False
 mnuActivRpt.Enabled = False
 End Sub


'DataEnvironment1.Connection1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=F:\New folder\bca\5th sem\crime file\CrimeFile\CrimeFile\CrimeFile\Crime File System\crimefile.mdb;Persist Security Info=False"
'End Sub

Private Sub mnuActivChrgSht_Click()
Load frmChargeSheet
frmChargeSheet.Show
End Sub

Private Sub mnuActivCompReg_Click()
Load frmCompltReg
frmCompltReg.Show

End Sub

Private Sub mnuActivCompStat_Click()
Load frmCompStatus
frmCompStatus.Show
End Sub

Private Sub mnuActivCsHist_Click()
Load frmHistSheet
frmHistSheet.Show
End Sub

Private Sub mnuActivFir_Click()
Load frmFIR
frmFIR.Show
End Sub

Private Sub mnuActivRegCrim_Click()
Load frmCriminalRegister
frmCriminalRegister.Show
End Sub

Private Sub mnuActivRegPost_Click()
Load frmPostmortem
frmPostmortem.Show
End Sub

Private Sub mnuActivRegPris_Click()
Load frmPrisonerReg
frmPrisonerReg.Show
End Sub

Private Sub mnuActivRptCrim_Click()
'Load DataReportCrimeRpt
'DataReportCrimeRpt.Show
DataReportCriminalReport.Show
End Sub

Private Sub mnuActivRptMst_Click()
'Load DataReportMostWanted
DataReportMostWanted.Show
End Sub

Private Sub mnuActRegMst_Click()
Load frmMostWanted
frmMostWanted.Show
End Sub

Private Sub mnuAdminAddComSt_Click()
Load frmAddCompltStatus
frmAddCompltStatus.Show
End Sub

Private Sub mnuAdminAddU_Click()
Load frmAddUser
frmAddUser.Show
End Sub

Private Sub mnuAdminDelU_Click()
Load frmDeleteUser
frmDeleteUser.Show
End Sub

Private Sub mnuLogin_Click()
Load frmLogin
frmLogin.Show

End Sub

Private Sub mnuLogoff_Click()
mnuLogoff.Enabled = False
mnuAdminAddU.Enabled = False
mnuAdminDelU.Enabled = False
mnuActiv.Enabled = False
mnuLogin.Enabled = True
mnuAdminAddComSt.Enabled = False
mnuActivReg.Enabled = False
mnuActivComp.Enabled = False
mnuActivRpt.Enabled = False
End Sub

Private Sub mnuRptFIR_Click()
'Load frmFIRrpt
'frmFIRrpt.Show
DataReportFIR.Show
End Sub

Private Sub mnuRptMortem_Click()
'Load frmPMortemRpt
'frmPMortemRpt.Show
DataReportMortageReport.Show
End Sub
