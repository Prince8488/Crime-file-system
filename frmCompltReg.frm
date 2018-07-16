VERSION 5.00
Begin VB.Form frmCompltReg 
   BackColor       =   &H8000000E&
   Caption         =   "Complaint Registration"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15240
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10095
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.OptionButton optFemale 
      BackColor       =   &H8000000E&
      Caption         =   "Female"
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
      Left            =   9765
      TabIndex        =   22
      Top             =   5940
      Width           =   1230
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H8000000E&
      Caption         =   "Male"
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
      Left            =   8325
      TabIndex        =   21
      Top             =   5940
      Value           =   -1  'True
      Width           =   1230
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8880
      TabIndex        =   20
      Top             =   9270
      Width           =   1230
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7230
      TabIndex        =   19
      Top             =   9270
      Width           =   1230
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5595
      TabIndex        =   18
      Top             =   9270
      Width           =   1230
   End
   Begin VB.TextBox txtNationality 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8783
      TabIndex        =   17
      Top             =   8505
      Width           =   2220
   End
   Begin VB.TextBox txtDate 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8775
      TabIndex        =   16
      Top             =   7650
      Width           =   2220
   End
   Begin VB.TextBox txtHname 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8783
      TabIndex        =   15
      Top             =   6795
      Width           =   2220
   End
   Begin VB.TextBox txtAge 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8738
      MaxLength       =   3
      TabIndex        =   14
      Top             =   5055
      Width           =   2220
   End
   Begin VB.TextBox txtDetails 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   8738
      MultiLine       =   -1  'True
      TabIndex        =   13
      Top             =   3960
      Width           =   2220
   End
   Begin VB.TextBox txtOccupation 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8738
      TabIndex        =   12
      Top             =   3330
      Width           =   2220
   End
   Begin VB.TextBox txtName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8738
      TabIndex        =   11
      Top             =   2430
      Width           =   2220
   End
   Begin VB.TextBox txtCno 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8738
      MaxLength       =   3
      TabIndex        =   10
      Top             =   1620
      Width           =   2220
   End
   Begin VB.Label Label10 
      BackColor       =   &H8000000E&
      Caption         =   "Nationality"
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
      Left            =   4238
      TabIndex        =   9
      Top             =   8550
      Width           =   2580
   End
   Begin VB.Label Label9 
      BackColor       =   &H8000000E&
      Caption         =   "Complaint Date"
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
      Left            =   4238
      TabIndex        =   8
      Top             =   7665
      Width           =   2580
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000E&
      Caption         =   "Father's/ Husband's Name"
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
      Left            =   4238
      TabIndex        =   7
      Top             =   6795
      Width           =   2580
   End
   Begin VB.Label Label7 
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
      Height          =   330
      Left            =   4238
      TabIndex        =   6
      Top             =   5925
      Width           =   2580
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000E&
      Caption         =   "Complaint No:"
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
      Left            =   4238
      TabIndex        =   5
      Top             =   1665
      Width           =   2580
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
      Height          =   330
      Left            =   4238
      TabIndex        =   4
      Top             =   5085
      Width           =   2580
   End
   Begin VB.Label Label4 
      BackColor       =   &H8000000E&
      Caption         =   "Details Of Complaint"
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
      Left            =   4238
      TabIndex        =   3
      Top             =   4215
      Width           =   2580
   End
   Begin VB.Label Label3 
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
      Height          =   330
      Left            =   4238
      TabIndex        =   2
      Top             =   3345
      Width           =   2580
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
      Height          =   330
      Left            =   4238
      TabIndex        =   1
      Top             =   2475
      Width           =   2580
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000E&
      Caption         =   "Complaint Registration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6135
      TabIndex        =   0
      Top             =   450
      Width           =   2985
   End
End
Attribute VB_Name = "frmCompltReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim X As Integer
Dim age As Integer


Private Sub cmdAddNew_Click()
txtName.Text = ""
txtAge.Text = ""
txtDate.Text = ""
txtDetails.Text = ""
txtHname.Text = ""
txtNationality.Text = ""
txtOccupation.Text = ""
Set rs = con.Execute("select count(*) from ComplaintReg")
X = rs(0)
X = X + 1
txtCno.Text = X

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
Dim sex As String
Dim age As Integer

If (optMale.Value = True) Then
        sex = "Male"
    Else
        sex = "Female"
End If
If (Trim(txtAge.Text) = "" Or txtDate.Text = "" Or txtDetails.Text = "" Or txtHname.Text = "" Or txtName.Text = "" Or txtNationality.Text = "" Or txtOccupation.Text = "") Then
    MsgBox "Missing Fields, Please fill up all the fields", vbInformation, "Crime File"
    Exit Sub
ElseIf (txtAge.Text <> "") Then
    If (Not IsNumeric(txtAge.Text)) Then
        MsgBox "Age Should be Number", vbInformation
            Exit Sub
    End If

End If
    'age = CInt(txtAge.Text)
    
    con.Execute ("insert into ComplaintReg values(" + txtCno.Text + ",'" + txtName.Text + "','" + txtOccupation.Text + "','" + txtDetails.Text + "'," + txtAge.Text + ",'" + sex + "','" + txtHname.Text + "','" + txtDate.Text + "','" + txtNationality.Text + "')")
    
    con.Execute ("insert into ComplntTemp values(" + txtCno.Text + ",'No')")
    
    MsgBox "Record Added Successfully"

End Sub

Private Sub Form_Load()

dbconnection
Set rs = con.Execute("select count(*) from ComplaintReg")
X = rs(0)
txtCno.Text = X + 1
rs.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
con.Close

End Sub

Private Sub optFemale_GotFocus()
Dim a As Integer
a = val(txtAge.Text)

If (a < 18) Or (a >= 100) Then
MsgBox "Invalid age"
txtAge.Text = " "
txtAge.SetFocus
End If

End Sub

Private Sub optMale_GotFocus()
Dim a As Integer
a = val(txtAge.Text)

If (a < 18) Or (a >= 100) Then
MsgBox "Invalid age"
txtAge.Text = " "
txtAge.SetFocus
End If

End Sub

Private Sub txtAge_KeyPress(KeyAscii As Integer)

If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then

KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If


End Sub

Private Sub txtCno_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtDate_Change()
txtDate.Text = Date
End Sub

Private Sub txtDetails_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtHname_KeyPress(KeyAscii As Integer)
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

Private Sub txtNationality_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtOccupation_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
