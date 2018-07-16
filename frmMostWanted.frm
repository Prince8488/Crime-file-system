VERSION 5.00
Begin VB.Form frmMostWanted 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Most Wanted Criminals"
   ClientHeight    =   7995
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7995
   ScaleWidth      =   7710
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   6630
      Left            =   1013
      TabIndex        =   0
      Top             =   675
      Width           =   5685
      Begin VB.TextBox txthd 
         Height          =   285
         Left            =   2640
         TabIndex        =   14
         Top             =   960
         Width           =   1935
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
         Height          =   375
         Left            =   2790
         TabIndex        =   13
         Top             =   5670
         Width           =   1275
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Left            =   1305
         TabIndex        =   12
         Top             =   5670
         Width           =   1275
      End
      Begin VB.ComboBox cmbSex 
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
         ItemData        =   "frmMostWanted.frx":0000
         Left            =   2385
         List            =   "frmMostWanted.frx":000A
         TabIndex        =   11
         Text            =   "Male"
         Top             =   3948
         Width           =   2490
      End
      Begin VB.TextBox txtDesc 
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
         Left            =   2520
         TabIndex        =   10
         Top             =   4680
         Width           =   2490
      End
      Begin VB.TextBox txtAge 
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
         Left            =   2400
         MaxLength       =   2
         TabIndex        =   9
         Top             =   3112
         Width           =   2490
      End
      Begin VB.TextBox txtNickName 
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
         Left            =   2385
         TabIndex        =   8
         Top             =   2276
         Width           =   2490
      End
      Begin VB.TextBox txtName 
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
         Left            =   2385
         TabIndex        =   7
         Top             =   1440
         Width           =   2490
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Description"
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
         Left            =   675
         TabIndex        =   6
         Top             =   4725
         Width           =   1500
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Sex"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   630
         TabIndex        =   5
         Top             =   3870
         Width           =   1320
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Age"
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
         Left            =   630
         TabIndex        =   4
         Top             =   3060
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "Nick Name"
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
         Left            =   630
         TabIndex        =   3
         Top             =   2250
         Width           =   1320
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Name"
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
         Left            =   630
         TabIndex        =   2
         Top             =   1485
         Width           =   1410
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Most Wanted Criminals"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   1035
         TabIndex        =   1
         Top             =   315
         Width           =   3750
      End
   End
End
Attribute VB_Name = "frmMostWanted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer

Private Sub cmbSex_GotFocus()
Dim a As Integer
a = val(txtAge.Text)

If (a < 18) Or (a >= 100) Then
MsgBox "Invalid age"
txtAge.Text = " "
txtAge.SetFocus
End If
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If (txtAge.Text = "" Or txtDesc.Text = "" Or txtName.Text = "" Or txtNickName.Text = "") Then
    MsgBox "Missing fields!!, Please fill up all", vbInformation, "Crime File System"
Else
    Set rs = con.Execute("select count(*) from MostWanted")
    If (Not rs.EOF) Then
    X = rs(0)
        If (X = 0) Then
            X = 1
        Else
            X = X + 1
        End If
 txthd.Text = X
    End If
    If (Not IsNumeric(txtAge.Text)) Then
        MsgBox "Age should be number", vbInformation, "Crime File Syatem"
        txtAge.Text = ""
        txtAge.SetFocus
    Else
    
   con.Execute ("insert into MostWanted values(" + txthd.Text + ",'" + txtName.Text + _
                "','" + txtNickName.Text + "'," + txtAge.Text + ",'" + cmbSex.Text + "','" + txtDesc.Text + "')")
    MsgBox "Record Added Successfully", vbInformation, "Crime File system"
    txtName.Text = ""
    txtAge.Text = ""
    txtDesc.Text = ""
    txtNickName.Text = ""
    txtName.SetFocus
    End If
End If
End Sub

Private Sub Form_Load()
dbconnection
End Sub


Private Sub txtAge_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtNickName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
