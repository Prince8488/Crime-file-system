VERSION 5.00
Begin VB.Form frmCompStatus 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Complaint Status"
   ClientHeight    =   6390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8880
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   5055
      Left            =   900
      TabIndex        =   0
      Top             =   675
      Width           =   7125
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
         Height          =   375
         Left            =   2880
         TabIndex        =   6
         Top             =   4545
         Width           =   1230
      End
      Begin VB.ComboBox cmbComplntNo 
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
         Left            =   3060
         TabIndex        =   5
         Top             =   1575
         Width           =   1455
      End
      Begin VB.CommandButton cmdViewStat 
         Caption         =   "View Status"
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
         Left            =   4770
         TabIndex        =   4
         Top             =   1530
         Width           =   1455
      End
      Begin VB.TextBox txtDetails 
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
         Height          =   1950
         Left            =   855
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   2385
         Width           =   5340
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "   Complaint Status"
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
         Left            =   2160
         TabIndex        =   3
         Top             =   360
         Width           =   2790
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Enter the complaint number"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   540
         TabIndex        =   2
         Top             =   1620
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmCompStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdViewStat_Click()
Set rs = con.Execute("select * from ComplntStatus where ComplntNo=" + cmbComplntNo.Text + "")
If (Not rs.EOF) Then
    txtDetails.Text = rs(1)
End If

End Sub

Private Sub Form_Load()
dbconnection
Set rs = con.Execute("select distinct(ComplntNo) from ComplntStatus")
While (Not rs.EOF)
    cmbComplntNo.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub

