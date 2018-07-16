VERSION 5.00
Begin VB.Form frmFIR 
   BackColor       =   &H80000009&
   Caption         =   "FIR"
   ClientHeight    =   10185
   ClientLeft      =   75
   ClientTop       =   465
   ClientWidth     =   14130
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10185
   ScaleWidth      =   14130
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   9105
      Left            =   5115
      TabIndex        =   0
      Top             =   855
      Width           =   9690
      Begin VB.ComboBox cmbComplantno 
         Height          =   315
         Left            =   7380
         TabIndex        =   2
         Top             =   495
         Width           =   1770
      End
      Begin VB.ComboBox cmbT3 
         Height          =   315
         ItemData        =   "frmFIR.frx":0000
         Left            =   4500
         List            =   "frmFIR.frx":000A
         TabIndex        =   14
         Top             =   6675
         Width           =   690
      End
      Begin VB.ComboBox cmbT1 
         Height          =   315
         ItemData        =   "frmFIR.frx":0016
         Left            =   4590
         List            =   "frmFIR.frx":0020
         TabIndex        =   33
         Top             =   1665
         Width           =   690
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
         Left            =   5100
         TabIndex        =   17
         Top             =   8025
         Width           =   1410
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
         Left            =   2985
         TabIndex        =   16
         Top             =   8025
         Width           =   1320
      End
      Begin VB.ComboBox cmbAct 
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
         ItemData        =   "frmFIR.frx":002C
         Left            =   3645
         List            =   "frmFIR.frx":003C
         TabIndex        =   8
         Top             =   3825
         Width           =   2220
      End
      Begin VB.ComboBox cmbForLoc 
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
         ItemData        =   "frmFIR.frx":006F
         Left            =   3645
         List            =   "frmFIR.frx":0079
         TabIndex        =   7
         Top             =   3320
         Width           =   2220
      End
      Begin VB.TextBox txtTime 
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
         Height          =   330
         Left            =   3645
         MaxLength       =   5
         TabIndex        =   4
         Top             =   1652
         Width           =   735
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
         Height          =   330
         Left            =   3645
         TabIndex        =   3
         Top             =   1080
         Width           =   2220
      End
      Begin VB.TextBox txtPlaceOcc 
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
         Height          =   330
         Left            =   3645
         TabIndex        =   6
         Top             =   2764
         Width           =   2220
      End
      Begin VB.TextBox txtFirNo 
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
         Height          =   330
         Left            =   3645
         MaxLength       =   3
         TabIndex        =   1
         Top             =   540
         Width           =   2220
      End
      Begin VB.TextBox txtTypInfo 
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
         Height          =   330
         Left            =   3645
         TabIndex        =   5
         Top             =   2208
         Width           =   2220
      End
      Begin VB.TextBox txtInfoRcd 
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
         Height          =   330
         Left            =   3645
         TabIndex        =   15
         Top             =   7215
         Width           =   2220
      End
      Begin VB.TextBox txtRcdTime 
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
         Height          =   330
         Left            =   3645
         MaxLength       =   5
         TabIndex        =   13
         Top             =   6660
         Width           =   645
      End
      Begin VB.TextBox txtPolice 
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
         Height          =   330
         Left            =   3645
         TabIndex        =   12
         Top             =   6105
         Width           =   2220
      End
      Begin VB.TextBox txtPassNo 
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
         Height          =   330
         Left            =   3645
         MaxLength       =   6
         TabIndex        =   11
         Top             =   5550
         Width           =   2220
      End
      Begin VB.TextBox txtInfoAdd 
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
         Height          =   330
         Left            =   3645
         TabIndex        =   10
         Top             =   4995
         Width           =   2220
      End
      Begin VB.TextBox txtDistrict 
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
         Height          =   330
         Left            =   3645
         TabIndex        =   9
         Top             =   4402
         Width           =   2220
      End
      Begin VB.Label Label17 
         BackColor       =   &H80000018&
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
         Height          =   375
         Left            =   6075
         TabIndex        =   32
         Top             =   540
         Width           =   1320
      End
      Begin VB.Label Label15 
         BackColor       =   &H80000018&
         Caption         =   "Information Received"
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
         Left            =   810
         TabIndex        =   31
         Top             =   7305
         Width           =   2445
      End
      Begin VB.Label Label14 
         BackColor       =   &H80000018&
         Caption         =   "Received Time"
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
         Left            =   810
         TabIndex        =   30
         Top             =   6765
         Width           =   2445
      End
      Begin VB.Label Label13 
         BackColor       =   &H80000018&
         Caption         =   "Name of Police"
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
         Left            =   810
         TabIndex        =   29
         Top             =   6195
         Width           =   2445
      End
      Begin VB.Label Label12 
         BackColor       =   &H80000018&
         Caption         =   "Passport No:"
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
         Left            =   810
         TabIndex        =   28
         Top             =   5640
         Width           =   2445
      End
      Begin VB.Label Label11 
         BackColor       =   &H80000018&
         Caption         =   "Informant Address"
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
         Left            =   810
         TabIndex        =   27
         Top             =   5085
         Width           =   2445
      End
      Begin VB.Label Label9 
         BackColor       =   &H80000018&
         Caption         =   "District"
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
         Left            =   810
         TabIndex        =   26
         Top             =   4491
         Width           =   2445
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000018&
         Caption         =   "Act"
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
         Left            =   810
         TabIndex        =   25
         Top             =   3933
         Width           =   2445
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000018&
         Caption         =   "Foreign/Local"
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
         Left            =   810
         TabIndex        =   24
         Top             =   3375
         Width           =   2445
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000018&
         Caption         =   "Place of occurance"
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
         Left            =   810
         TabIndex        =   23
         Top             =   2817
         Width           =   2445
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000018&
         Caption         =   "Type of information"
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
         Left            =   810
         TabIndex        =   22
         Top             =   2259
         Width           =   2445
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000018&
         Caption         =   "Time"
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
         Left            =   810
         TabIndex        =   21
         Top             =   1701
         Width           =   2445
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000018&
         Caption         =   "Date"
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
         Left            =   810
         TabIndex        =   20
         Top             =   1143
         Width           =   2445
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000018&
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
         Height          =   375
         Left            =   810
         TabIndex        =   19
         Top             =   585
         Width           =   2445
      End
   End
   Begin VB.Image Image1 
      Height          =   3990
      Left            =   0
      Picture         =   "frmFIR.frx":008D
      Top             =   135
      Width           =   5010
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000014&
      Caption         =   "FIR"
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
      Left            =   9690
      TabIndex        =   18
      Top             =   180
      Width           =   735
   End
End
Attribute VB_Name = "frmFIR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdSave_Click()
If (txtDate.Text = "" Or txtDistrict.Text = "" Or txtFirNo.Text = "" Or txtInfoAdd.Text = "" Or txtInfoRcd.Text = "" Or txtPassNo.Text = "" Or txtPlaceOcc.Text = "" Or txtPolice.Text = "" Or txtRcdTime.Text = "" Or txtTime.Text = "" Or txtTypInfo.Text = "" Or cmbComplantno.Text = "") Then
        MsgBox "Missing Fields!!", vbInformation, "Crime File System"
    Else
    Set rs = con.Execute("select * from FIR where FirNo=" + txtFirNo.Text + " or ComplntNo=" + cmbComplantno.Text + "")
    If (Not rs.EOF) Then
            MsgBox "Sorry!! FIR Or Complaint Number already exists. Try another number", vbCritical, "CFS"
            txtFirNo.Text = ""
            txtFirNo.SetFocus
    Else
    Dim t1, t2, t3 As String
    t1 = txtTime.Text + cmbT1.Text
    't2 = txtTime2.Text + cmbT2.Text
    t3 = txtRcdTime.Text + cmbT3.Text
        If ((Not IsNumeric(txtPassNo.Text)) Or (Not IsNumeric(txtFirNo.Text))) Then
            MsgBox "Please check the number fields", vbInformation, "Crime File System"
            'txtPassNo.Text = ""
            'txtPassNo.SetFocus
        Else
        con.Execute ("insert into FIR values(" + txtFirNo.Text + "," + cmbComplantno.Text + _
                    ",'" + txtDate.Text + "','" + t1 + "','" + txtTypInfo.Text + _
                    "','" + txtPlaceOcc.Text + "','" + cmbForLoc.Text + "','" + cmbAct.Text + _
                    "','" + txtDistrict.Text + "','" + txtInfoAdd.Text + "'," + txtPassNo.Text + ",'" + txtPolice.Text + _
                    "','" + t3 + "','" + txtInfoRcd.Text + "')")
        con.Execute ("UPDATE ComplntTemp set Status='Yes' where ComplntNo=" + cmbComplantno.Text + "")
                MsgBox "Record Added", vbInformation, "Crime File system"
                txtDate.Text = ""
                txtDistrict.Text = ""
                
                txtFirNo.Text = ""
                txtInfoAdd.Text = ""
                txtInfoRcd.Text = ""
                txtPassNo.Text = ""
                txtPlaceOcc.Text = ""
                txtPolice.Text = ""
                txtRcdTime.Text = ""
                
                txtTime.Text = ""
                txtTypInfo.Text = ""
                txtFirNo.SetFocus
                End If
            End If
    End If
    
End Sub

Private Sub Form_Load()
cmbT1.Text = "AM"
cmbT3.Text = "AM"
cmbAct.Text = "Murder"
cmbForLoc.Text = "Foreign"
dbconnection
Set rs = con.Execute("Select ComplntNo from ComplntTemp where Status='No'")
While (Not rs.EOF)
    cmbComplantno.AddItem rs(0)
    rs.MoveNext
Wend
End Sub

Private Sub txtDate_Change()
txtDate.Text = Date
End Sub

Private Sub txtDistrict_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtFirNo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtInfoAdd_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtInfoRcd_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtPassNo_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub

Private Sub txtPlaceOcc_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtPolice_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub

Private Sub txtRcdTime_Change()
txtRcdTime.Text = Time
End Sub

Private Sub txtTime_Change()
txtTime.Text = Time
End Sub

Private Sub txtTime_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57 Or KeyAscii = 32) Then
KeyAscii = 0
MsgBox ("Please enter only Numbers")
End If
End Sub



Private Sub txtTypInfo_KeyPress(KeyAscii As Integer)

If (KeyAscii >= 48 And KeyAscii <= 57) Then
KeyAscii = 0
MsgBox ("Please enter only alphabets")
End If
End Sub
