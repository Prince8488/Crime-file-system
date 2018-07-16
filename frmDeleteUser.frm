VERSION 5.00
Begin VB.Form frmDeleteUser 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Delete User"
   ClientHeight    =   4995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7050
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4995
   ScaleWidth      =   7050
   Begin VB.Frame frmLogin 
      BackColor       =   &H80000018&
      Height          =   4335
      Left            =   885
      TabIndex        =   0
      Top             =   270
      Width           =   5280
      Begin VB.Frame Frame2 
         BackColor       =   &H0080C0FF&
         Height          =   645
         Left            =   765
         TabIndex        =   4
         Top             =   315
         Width           =   3705
         Begin VB.Label Label1 
            BackColor       =   &H0080C0FF&
            Caption         =   "Delete User"
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
            TabIndex        =   5
            Top             =   135
            Width           =   2175
         End
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
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
         TabIndex        =   3
         Top             =   3285
         Width           =   1230
      End
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
         TabIndex        =   2
         Top             =   3285
         Width           =   1365
      End
      Begin VB.ComboBox cmbUsername 
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2250
         TabIndex        =   1
         Top             =   1845
         Width           =   2220
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
         Left            =   495
         TabIndex        =   6
         Top             =   1890
         Width           =   1905
      End
   End
End
Attribute VB_Name = "frmDeleteUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDelete_Click()
con.Execute ("delete from UserLogin where Username='" + cmbUsername.Text + "'")
MsgBox "User deleted sucessfully!!", vbInformation, "CFS"
cmbUsername.Text = ""
End Sub

Private Sub Form_Load()
dbconnection
Set rs = con.Execute("select * from UserLogin")
While (Not rs.EOF)
    cmbUsername.AddItem rs(0)
    rs.MoveNext
Wend
End Sub

