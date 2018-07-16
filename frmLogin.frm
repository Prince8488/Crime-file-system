VERSION 5.00
Begin VB.Form frmLogin 
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Login"
   ClientHeight    =   3630
   ClientLeft      =   3780
   ClientTop       =   3165
   ClientWidth     =   6705
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   6705
   Begin VB.CommandButton cmdUserLog 
      Caption         =   "User Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2745
      TabIndex        =   7
      Top             =   2970
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H80000009&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4215
      TabIndex        =   6
      Top             =   2955
      Width           =   1230
   End
   Begin VB.CommandButton cmdAdminLog 
      BackColor       =   &H80000009&
      Caption         =   "Admin Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1260
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   5
      Top             =   2955
      Width           =   1230
   End
   Begin VB.TextBox txtPass 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      IMEMode         =   3  'DISABLE
      Left            =   3255
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   2025
      Width           =   1905
   End
   Begin VB.TextBox txtUname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3255
      TabIndex        =   3
      Text            =   " "
      Top             =   1125
      Width           =   1905
   End
   Begin VB.Label Label3 
      BackColor       =   &H80000005&
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1500
      TabIndex        =   2
      Top             =   2025
      Width           =   1230
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
      Caption         =   "Username"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   1590
      TabIndex        =   1
      Top             =   1125
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000005&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   2872
      TabIndex        =   0
      Top             =   225
      Width           =   960
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdminLog_Click()
Set rs = con.Execute("Select * from AdminLogin where Username='" + txtUname.Text + "' and Password='" + txtPass.Text + "'")
If (Not rs.EOF) Then
    
    MsgBox "Login Success", vbInformation, "Crime File System"
    frmMDI.mnuLogoff.Enabled = True
    frmMDI.mnuAdminAddU.Enabled = True
    frmMDI.mnuAdminDelU.Enabled = True
    frmMDI.mnuActiv.Enabled = True
    frmMDI.mnuLogin.Enabled = False
    frmMDI.mnuAdminAddComSt.Enabled = True
    frmMDI.mnuActivReg.Enabled = True
    frmMDI.mnuActivComp.Enabled = True
    frmMDI.mnuActivRpt.Enabled = True
    Unload Me
Else
    MsgBox "Failure", vbCritical, "Crime File System"
End If
'rs.Close
Unload Me
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdUserLog_Click()
Set rs = con.Execute("Select * from UserLogin where Username='" + txtUname.Text + "' and Password='" + txtPass.Text + "'")
If (Not rs.EOF) Then
    
    MsgBox "Login Success", vbInformation, "Crime File System"
    frmMDI.mnuLogoff.Enabled = True
    frmMDI.mnuLogin.Enabled = False

    frmMDI.mnuActiv.Enabled = True
    frmMDI.mnuActivReg.Enabled = True
    frmMDI.mnuActivComp.Enabled = True
    frmMDI.mnuActivRpt.Enabled = True
    Unload Me
Else
    MsgBox "Failure", vbCritical, "Crime File System"
End If
'rs.Close

End Sub

Private Sub Form_Load()
dbconnection
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close
End Sub

Private Sub txtUname_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
