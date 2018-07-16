VERSION 5.00
Begin VB.Form frmPostmortem 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Postmortem"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   8295
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000E&
      Height          =   8025
      Left            =   720
      TabIndex        =   0
      Top             =   315
      Width           =   6855
      Begin VB.ComboBox cmbFirrNo 
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
         Left            =   3645
         TabIndex        =   21
         Top             =   1350
         Width           =   2040
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
         ItemData        =   "frmPostmortem.frx":0000
         Left            =   3690
         List            =   "frmPostmortem.frx":000A
         TabIndex        =   2
         Top             =   3330
         Width           =   1995
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
         Left            =   3600
         TabIndex        =   10
         Top             =   7335
         Width           =   1230
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
         Left            =   1800
         TabIndex        =   9
         Top             =   7335
         Width           =   1185
      End
      Begin VB.TextBox txtPostNo 
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
         Height          =   390
         Left            =   3690
         MaxLength       =   3
         TabIndex        =   7
         Top             =   705
         Width           =   1965
      End
      Begin VB.TextBox txtRslt 
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
         Height          =   1110
         Left            =   3705
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   1950
         Width           =   1965
      End
      Begin VB.TextBox txtDate 
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
         Height          =   390
         Left            =   3705
         TabIndex        =   3
         Top             =   3915
         Width           =   1965
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
         Height          =   390
         Left            =   3705
         TabIndex        =   4
         Top             =   4560
         Width           =   1965
      End
      Begin VB.TextBox txtHouseName 
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
         Height          =   390
         Left            =   3705
         TabIndex        =   5
         Top             =   5205
         Width           =   1965
      End
      Begin VB.TextBox txtDoctName 
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
         Height          =   390
         Left            =   3705
         TabIndex        =   6
         Top             =   5850
         Width           =   1965
      End
      Begin VB.TextBox txtPoliceSt 
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
         Height          =   390
         Left            =   3705
         TabIndex        =   8
         Top             =   6510
         Width           =   1965
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000E&
         Caption         =   "Postmortem"
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
         Left            =   2445
         TabIndex        =   20
         Top             =   135
         Width           =   1665
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000E&
         Caption         =   "Postmortem No:"
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
         Left            =   1080
         TabIndex        =   19
         Top             =   780
         Width           =   1740
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000E&
         Caption         =   "FIR No:"
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
         Left            =   1080
         TabIndex        =   18
         Top             =   1395
         Width           =   1740
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000E&
         Caption         =   "Result Of Mortem"
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
         Left            =   1080
         TabIndex        =   17
         Top             =   2070
         Width           =   1740
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
         Height          =   390
         Left            =   1080
         TabIndex        =   16
         Top             =   3360
         Width           =   1740
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         Caption         =   "Date Of Death"
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
         Left            =   1080
         TabIndex        =   15
         Top             =   3990
         Width           =   1740
      End
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         Caption         =   "Description Of Case"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   4635
         Width           =   1740
      End
      Begin VB.Label Label8 
         BackColor       =   &H8000000E&
         Caption         =   "House Name"
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
         Left            =   1080
         TabIndex        =   13
         Top             =   5280
         Width           =   1740
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000E&
         Caption         =   "Doctor's Name"
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
         Left            =   1080
         TabIndex        =   12
         Top             =   5925
         Width           =   1740
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000E&
         Caption         =   "Police Station"
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
         Left            =   1080
         TabIndex        =   11
         Top             =   6585
         Width           =   1740
      End
   End
End
Attribute VB_Name = "frmPostmortem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnt As Integer
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If (txtDate.Text = "" Or txtDesc.Text = "" Or txtDoctName.Text = "" Or cmbFirrNo.Text = "" Or txtHouseName.Text = "" Or txtPoliceSt.Text = "" Or txtPostNo.Text = "" Or txtRslt.Text = "") Then
    MsgBox "Missing Fields!! Please Fill Up all", vbInformation, "Crime File System"
Else
    Set rs = con.Execute("select * from Postmortem where FirNo=" + cmbFirrNo.Text + "")
    If (Not rs.EOF) Then
        MsgBox "Duplication is not allowded, Please try another FIR number", vbCritical, "Crime File System"
        cmbFirrNo.SetFocus
        
    Else
        con.Execute ("insert into postmortem values(" + txtPostNo.Text + "," + cmbFirrNo.Text + ",'" + txtRslt.Text + _
        "','" + cmbSex.Text + "','" + txtDate.Text + "','" + txtDesc.Text + _
        "','" + txtHouseName.Text + "','" + txtDoctName.Text + "','" + txtPoliceSt.Text + "')")
        MsgBox "Record Added Sucessfully", vbInformation, "Crime File System"
        txtPostNo.Text = txtPostNo + 1
        txtDate.Text = ""
        txtDesc.Text = ""
        txtDoctName.Text = ""
        txtHouseName.Text = ""
        txtPoliceSt.Text = ""
        txtRslt.Text = ""
    End If
End If
End Sub

Private Sub Form_Load()
dbconnection
cmbSex.Text = "Male"
Set rs = con.Execute("select max(PMortemNo) from Postmortem")
If (Not rs.EOF) Then
'    cnt = rs(0)
    If (cnt = 0) Then
        cnt = 1
    Else
        cnt = cnt + 1
    End If
    txtPostNo.Text = cnt
End If
rs.Close
Set rs = con.Execute("select Firno from FIR")
While (Not rs.EOF)
    cmbFirrNo.AddItem rs(0)
    rs.MoveNext
Wend
rs.Close
End Sub

Private Sub txtDate_Change()
txtDate.Text = Date
End Sub

Private Sub txtDesc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtDoctName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtHouseName_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtPoliceSt_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtPostNo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtRslt_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
