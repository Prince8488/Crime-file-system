VERSION 5.00
Begin VB.Form frmAddUser 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add User"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   7215
   Begin VB.Frame frmLogin 
      BackColor       =   &H80000018&
      Height          =   4335
      Left            =   967
      TabIndex        =   0
      Top             =   270
      Width           =   5280
      Begin VB.CommandButton cmdCancel 
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
         Height          =   420
         Left            =   3015
         TabIndex        =   6
         Top             =   3285
         Width           =   1365
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   990
         TabIndex        =   5
         Top             =   3285
         Width           =   1230
      End
      Begin VB.TextBox txtPassword 
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
         Left            =   2340
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   2160
         Width           =   2490
      End
      Begin VB.TextBox txtUsername 
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
         Left            =   2340
         TabIndex        =   3
         Top             =   1485
         Width           =   2490
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H008080FF&
         Height          =   645
         Left            =   765
         TabIndex        =   1
         Top             =   315
         Width           =   3705
         Begin VB.Label Label1 
            BackColor       =   &H008080FF&
            Caption         =   "Add New User"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   15.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   465
            Left            =   900
            TabIndex        =   2
            Top             =   135
            Width           =   2175
         End
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
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
         Height          =   420
         Left            =   450
         TabIndex        =   8
         Top             =   2250
         Width           =   1860
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   450
         TabIndex        =   7
         Top             =   1485
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmAddUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
dbconnection
End Sub
Private Sub cmdAdd_Click()
Set rs = con.Execute("select * from UserLogin where Username='" + txtUsername.Text + "' and Password='" + txtPassword.Text + "'")
If (Not rs.EOF) Then
    MsgBox "Sorry!! User already exists. Try another username", vbCritical, "Crime File System"
    txtPassword.Text = ""
    txtUsername.Text = ""
    txtUsername.SetFocus
Else
    con.Execute ("insert into UserLogin values('" + txtUsername.Text + "','" + txtPassword.Text + "')")
    MsgBox "User added sucessfully", vbInformation, "Crime File System"
    txtPassword.Text = ""
    txtUsername.Text = ""
    txtUsername.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub txtUsername_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
