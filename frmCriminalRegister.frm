VERSION 5.00
Begin VB.Form frmCriminalRegister 
   BackColor       =   &H80000009&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   9075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   9450
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   8340
      Left            =   1500
      TabIndex        =   0
      Top             =   270
      Width           =   6450
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
         Height          =   405
         Left            =   3330
         TabIndex        =   20
         Top             =   7605
         Width           =   1215
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000013&
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
         Height          =   405
         Left            =   1755
         TabIndex        =   10
         Top             =   7605
         Width           =   1215
      End
      Begin VB.TextBox txtCriminalNo 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
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
         Height          =   360
         Left            =   3150
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1035
         Width           =   2370
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         TabIndex        =   2
         Top             =   1890
         Width           =   2370
      End
      Begin VB.TextBox txtNickName 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         TabIndex        =   3
         Top             =   2679
         Width           =   2370
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         MaxLength       =   2
         TabIndex        =   4
         Top             =   3501
         Width           =   2370
      End
      Begin VB.TextBox txtOccupation 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         TabIndex        =   5
         Top             =   4323
         Width           =   2370
      End
      Begin VB.TextBox txtCrimeType 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         TabIndex        =   6
         Top             =   5145
         Width           =   2370
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000E&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3150
         TabIndex        =   7
         Top             =   5970
         Width           =   2370
      End
      Begin VB.OptionButton optyes 
         BackColor       =   &H8000000E&
         Caption         =   "Yes"
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
         Left            =   3240
         TabIndex        =   8
         Top             =   6720
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton optno 
         BackColor       =   &H8000000E&
         Caption         =   "No"
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
         Left            =   4665
         TabIndex        =   9
         Top             =   6750
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "    Criminal Register"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   1980
         TabIndex        =   19
         Top             =   225
         Width           =   3090
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Criminal No:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   540
         TabIndex        =   18
         Top             =   1095
         Width           =   1440
      End
      Begin VB.Label Label3 
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
         Height          =   540
         Left            =   540
         TabIndex        =   17
         Top             =   1920
         Width           =   1440
      End
      Begin VB.Label Label4 
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
         Height          =   540
         Left            =   540
         TabIndex        =   16
         Top             =   2745
         Width           =   1440
      End
      Begin VB.Label Label5 
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
         Height          =   540
         Left            =   540
         TabIndex        =   15
         Top             =   3570
         Width           =   1440
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Occupation"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   540
         TabIndex        =   14
         Top             =   4395
         Width           =   1440
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Crime Type"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   540
         TabIndex        =   13
         Top             =   5220
         Width           =   1440
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   540
         TabIndex        =   12
         Top             =   6045
         Width           =   1440
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Most Wanted"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   540
         TabIndex        =   11
         Top             =   6720
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmCriminalRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Dim X As String

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If (optyes.Value = True) Then
    X = "Yes"
Else
    X = "No"
End If
If (txtAddress.Text = "" Or txtAge.Text = "" Or txtCrimeType.Text = "" Or txtName.Text = "" Or txtNickName.Text = "" Or txtOccupation.Text = "") Then
    MsgBox "Missing Fields!!,Please fill up all", vbInformation, "Crime File System"
ElseIf (Not IsNumeric(txtAge.Text)) Then
    MsgBox "Age should be numeric"
Else
    con.Execute ("insert into CriminalReg values(" + txtCriminalNo.Text + _
    ",'" + txtName.Text + "','" + txtNickName.Text + "'," + txtAge.Text + _
    ",'" + txtOccupation.Text + "','" + txtCrimeType.Text + "','" + txtAddress.Text + _
    "','" + X + "')")
    MsgBox "Record Added successfully", vbInformation, "Crime File System"
    txtCriminalNo.Text = txtCriminalNo.Text + 1
    txtAddress.Text = ""
    txtAge.Text = ""
    txtCrimeType.Text = ""
    txtName.Text = ""
    txtNickName.Text = ""
    txtOccupation.Text = ""
    txtName.SetFocus
End If
End Sub

Private Sub Form_Load()
dbconnection
Set rs = con.Execute("select max(CriminalNo) from CriminalReg")
If (Not rs.EOF) Then
    cnt = rs(0)
    If (cnt = 0) Then
        cnt = 1
    Else
        cnt = cnt + 1
    End If
    txtCriminalNo.Text = cnt
End If
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub



Private Sub txtAge_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtCrimeType_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtCriminalNo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
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

Private Sub txtOccupation_GotFocus()
Dim a As Integer
a = val(txtAge.Text)

If (a < 18) Or (a >= 100) Then
MsgBox "Invalid age"
txtAge.Text = " "
txtAge.SetFocus
End If
End Sub

Private Sub txtOccupation_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
