VERSION 5.00
Begin VB.Form frmAddCompltStatus 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Add Complaint Status"
   ClientHeight    =   6990
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   9615
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   6135
      Left            =   900
      TabIndex        =   0
      Top             =   360
      Width           =   7800
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000013&
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
         Height          =   465
         Left            =   4200
         MaskColor       =   &H000000FF&
         TabIndex        =   11
         Top             =   5490
         Width           =   1545
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H80000013&
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
         Height          =   465
         Left            =   2280
         MaskColor       =   &H000000FF&
         TabIndex        =   10
         Top             =   5490
         Width           =   1545
      End
      Begin VB.TextBox txtStatus 
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
         Height          =   1365
         Left            =   1755
         MultiLine       =   -1  'True
         TabIndex        =   9
         Top             =   3915
         Width           =   4965
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H80000016&
         Height          =   1815
         Left            =   630
         TabIndex        =   6
         Top             =   1890
         Width           =   6495
         Begin VB.TextBox txtDetails 
            Appearance      =   0  'Flat
            Enabled         =   0   'False
            Height          =   1140
            Left            =   2430
            MultiLine       =   -1  'True
            TabIndex        =   8
            Top             =   315
            Width           =   3435
         End
         Begin VB.Label Label4 
            BackColor       =   &H80000016&
            Caption         =   "Details Of Suspect"
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
            Left            =   180
            TabIndex        =   7
            Top             =   630
            Width           =   1905
         End
      End
      Begin VB.CommandButton cmdDetails 
         BackColor       =   &H80000013&
         Caption         =   "Details"
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
         Left            =   5670
         MaskColor       =   &H000000FF&
         TabIndex        =   5
         Top             =   1350
         Width           =   1590
      End
      Begin VB.ComboBox cmbComplntNo 
         Height          =   315
         Left            =   3150
         TabIndex        =   4
         Top             =   1350
         Width           =   2265
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Status"
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
         Left            =   450
         TabIndex        =   3
         Top             =   4455
         Width           =   870
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Select Complaint No:"
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
         Left            =   450
         TabIndex        =   2
         Top             =   1305
         Width           =   2490
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Add Complaint Status"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2295
         TabIndex        =   1
         Top             =   360
         Width           =   2940
      End
   End
End
Attribute VB_Name = "frmAddCompltStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
If (cmbComplntNo.Text = "" Or txtStatus.Text = "") Then
    MsgBox "Missing Fields", vbInformation, "CFS"
Else
con.Execute ("insert into ComplntStatus values(" + cmbComplntNo.Text + ",'" + txtStatus.Text + "')")
MsgBox "Status Added Successfully", vbInformation, "Crime File System"
txtDetails.Text = ""
txtStatus.Text = ""
cmbComplntNo.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDetails_Click()
Set rs = con.Execute("select Details from ComplaintReg where ComplntNo=" + cmbComplntNo.Text + "")
If (Not rs.EOF) Then
    txtDetails.Text = rs(0)
    txtStatus.SetFocus
End If

End Sub

Private Sub Form_Load()
dbconnection
Set rs = con.Execute("select ComplntNo from ComplaintReg")
While (Not rs.EOF)
    cmbComplntNo.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

