VERSION 5.00
Begin VB.Form frmPrisonerReg 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Prisoners Register"
   ClientHeight    =   9435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9435
   ScaleWidth      =   9555
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   8760
      Left            =   945
      TabIndex        =   0
      Top             =   180
      Width           =   7740
      Begin VB.CommandButton Command2 
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
         Left            =   3735
         TabIndex        =   10
         Top             =   8055
         Width           =   1185
      End
      Begin VB.CommandButton Command1 
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
         Left            =   2205
         TabIndex        =   9
         Top             =   8055
         Width           =   1275
      End
      Begin VB.TextBox txtPrisNo 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
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
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   11
         Top             =   1050
         Width           =   2415
      End
      Begin VB.TextBox txtChgrNo 
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
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1837
         Width           =   2415
      End
      Begin VB.TextBox txtNikName 
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
         Left            =   4245
         TabIndex        =   2
         Top             =   2624
         Width           =   2415
      End
      Begin VB.TextBox txtCrmType 
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
         Left            =   4245
         TabIndex        =   3
         Top             =   3411
         Width           =   2415
      End
      Begin VB.TextBox txtFamMem 
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
         Left            =   4245
         TabIndex        =   4
         Top             =   4198
         Width           =   2415
      End
      Begin VB.TextBox txtIdenMark 
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
         Left            =   4245
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   4985
         Width           =   2415
      End
      Begin VB.TextBox txtHt 
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
         Left            =   4245
         MaxLength       =   1
         TabIndex        =   6
         Top             =   5772
         Width           =   2415
      End
      Begin VB.TextBox txtWt 
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
         Left            =   4245
         MaxLength       =   3
         TabIndex        =   7
         Top             =   6559
         Width           =   2415
      End
      Begin VB.TextBox txtColor 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4245
         TabIndex        =   8
         Top             =   7350
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Prisoners Register"
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
         Left            =   2550
         TabIndex        =   21
         Top             =   225
         Width           =   2865
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Prisoner's No:"
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
         Left            =   1260
         TabIndex        =   20
         Top             =   1095
         Width           =   1890
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "ChargeSheet No:"
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
         Left            =   1260
         TabIndex        =   19
         Top             =   1867
         Width           =   1890
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Nickname"
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
         Left            =   1260
         TabIndex        =   18
         Top             =   2685
         Width           =   1890
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Type Of Crime"
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
         Left            =   1260
         TabIndex        =   17
         Top             =   3411
         Width           =   1890
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Family Members"
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
         Left            =   1260
         TabIndex        =   16
         Top             =   4230
         Width           =   1890
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Identification Mark"
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
         Left            =   1260
         TabIndex        =   15
         Top             =   4995
         Width           =   1890
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Height in feet"
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
         Left            =   1260
         TabIndex        =   14
         Top             =   5775
         Width           =   1890
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Weight in kg"
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
         Left            =   1260
         TabIndex        =   13
         Top             =   6540
         Width           =   1890
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Color"
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
         Left            =   1260
         TabIndex        =   12
         Top             =   7320
         Width           =   1890
      End
   End
End
Attribute VB_Name = "frmPrisonerReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Private Sub Command1_Click()
If (txtChgrNo.Text = "" Or txtColor.Text = "" Or txtCrmType.Text = "" Or txtFamMem.Text = "" Or txtHt.Text = "" Or txtIdenMark.Text = "" Or txtNikName.Text = "" Or txtPrisNo.Text = "" Or txtWt.Text = "") Then
    MsgBox "Missing Fields!!, Please fill up all", vbInformation, "Crime File System"
Else
    If ((Not IsNumeric(txtChgrNo.Text)) Or (Not IsNumeric(txtHt.Text)) Or (Not IsNumeric(txtWt.Text))) Then
        MsgBox "Some of the values entered are not matching, Please check numbers are entered correctly", vbCritical, "Crime File System"
    Else
        Set rs = con.Execute("select * from PrisonersReg where ChargeSheetNo=" + txtChgrNo.Text + "")
        If (Not rs.EOF) Then
            MsgBox "Charge Sheet Number Already Exist, Please try another number", vbCritical, "Crime File System"
            txtChgrNo.Text = ""
            txtChgrNo.SetFocus
        Else
    con.Execute ("insert into PrisonersReg values(" + txtPrisNo.Text + "," + txtChgrNo.Text + ",'" + txtNikName.Text + _
    "','" + txtCrmType.Text + "','" + txtFamMem.Text + "','" + txtIdenMark.Text + "'," + txtHt.Text + _
    "," + txtWt.Text + ",'" + txtColor.Text + "')")
    con.Execute ("insert into prisonersTemp values(" + txtPrisNo.Text + ",'No')")
    MsgBox "Record Added Sucessfully", vbInformation, "CFS"
    txtPrisNo.Text = txtPrisNo.Text + 1
    txtChgrNo.Text = ""
    txtColor.Text = ""
    txtCrmType.Text = ""
    txtFamMem.Text = ""
    txtHt.Text = ""
    txtIdenMark.Text = ""
    txtNikName.Text = ""
    txtWt.Text = ""
    txtChgrNo.SetFocus
        End If
    End If
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
dbconnection
Set rs = con.Execute("select max(PrisonerNo) from PrisonersReg")
If (Not rs.EOF) Then
    'cnt = rs(0)
    If (cnt = 0) Then
        cnt = 1
    Else
        cnt = cnt + 1
    End If
    txtPrisNo.Text = cnt
End If
rs.Close
End Sub

Private Sub txtChgrNo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtColor_GotFocus()
Dim a As Integer
a = val(txtWt.Text)

If (a < 40) Or (a >= 200) Then
MsgBox "Invalid Weight"
txtWt.Text = " "
txtWt.SetFocus
End If
End Sub

Private Sub txtColor_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtCrmType_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtFamMem_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtHt_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtIdenMark_Change()
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtNikName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtPrisNo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtWt_GotFocus()
Dim a As Integer
a = val(txtHt.Text)

If (a < 4) Or (a >= 8) Then
MsgBox "Invalid age"
txtHt.Text = " "
txtHt.SetFocus
End If
End Sub

Private Sub txtWt_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub
